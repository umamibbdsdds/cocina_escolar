// ============================================================
//  server.js  —  API UMAMI Restaurante
//  Backend para Railway (MySQL) + compatible con mesas1.py
//
//  Estructura basada en el server.js de Cocina Escolar INFRAMEN:
//   - Pool de conexiones compatible con Railway (MYSQL_URL) o locales
//   - Logger detallado en cada solicitud
//   - Exportación Excel y PDF
//
//  Endpoints:
//   GET  /categorias         → lista de categorías del menú
//   GET  /platillos          → todos los platillos (o ?categoria=X)
//   GET  /platillo/:id       → un platillo por ID
//   POST /pedido             → crear pedido para una mesa
//   GET  /pedido/:mesa       → pedido activo de la mesa
//   PUT  /pedido/:mesa/estado→ cambiar estado del pedido
//   GET  /mesas              → lista de mesas y estados
//   GET  /mesa/:numero       → estado de una mesa
//   POST /solicitar-mesero  → solicitar mesero para mesa
//   POST /solicitar-cuenta  → solicitar la cuenta/pago
//   PUT  /pago/:mesa/estado  → cambiar estado del pago
//   GET  /pago/:mesa         → estado del pago de la mesa
//   GET  /estado-qr/:mesa    → verificar si QR fue escaneado
//   GET  /informe             → resumen de ventas del día (JSON)
//   GET  /informe/:fecha      → resumen histórico
//   GET  /informe/exportar/excel[/:fecha]
//   GET  /informe/exportar/pdf[/:fecha]
// ============================================================

const express  = require("express");
const mysql    = require("mysql2/promise");
const cors     = require("cors");
const ExcelJS  = require("exceljs");
const PDFDocument = require("pdfkit-table");
const os       = require("os");

const app = express();

// [ORDEN CRÍTICO] cors y json ANTES del logger
app.use(cors());
app.use(express.json());

// ── Pool de conexiones ──
// Soporta 3 escenarios sin tocar el código:
//   1) Railway con URL completa  → MYSQL_URL / DATABASE_URL / MYSQL_PUBLIC_URL
//   2) Railway por variables sueltas → MYSQLHOST, MYSQLPORT, MYSQLUSER...
//   3) Variables genéricas / local  → DB_HOST, DB_PORT, DB_USER... (o defaults)
const connectionUrl =
    process.env.MYSQL_URL       ||
    process.env.DATABASE_URL    ||
    process.env.MYSQL_PUBLIC_URL ||
    null;

const pool = connectionUrl
    ? mysql.createPool({
          uri: connectionUrl,
          waitForConnections: true,
          connectionLimit: 10,
      })
    : mysql.createPool({
          host:     process.env.MYSQLHOST     || process.env.DB_HOST     || "localhost",
          port:     Number(process.env.MYSQLPORT || process.env.DB_PORT) || 3306,
          user:     process.env.MYSQLUSER    || process.env.DB_USER     || "root",
          password: process.env.MYSQLPASSWORD || process.env.DB_PASSWORD || "230223",
          database: process.env.MYSQLDATABASE || process.env.DB_NAME    || "UMAMI_DB",
          waitForConnections: true,
          connectionLimit: 10,
      });

// ── Logger de solicitudes ──
app.use((req, _res, next) => {
    console.log(`\n[REQ] [${new Date().toLocaleTimeString()}] ${req.method} ${req.url}`);
    if (req.body && Object.keys(req.body).length > 0) {
        console.log("   [BODY]", JSON.stringify(req.body));
    }
    next();
});

// ── Utilidades ──
function hoyLocal() {
    const d = new Date();
    const p = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function fechaValida(f) {
    return /^\d{4}-\d{2}-\d{2}$/.test(f);
}

function responderError(res, contexto, err, status = 500) {
    console.error(`[ERROR] ${contexto}:`, err.message);
    if (res.headersSent) return res.end();
    res.status(status).json({ status: "error", mensaje: err.message });
}

// ── Asegurar esquema mínimo ──
async function asegurarEsquema() {
    try {
        await pool.query(`
            CREATE TABLE IF NOT EXISTS categorias (
                id       INT AUTO_INCREMENT PRIMARY KEY,
                nombre   VARCHAR(60) NOT NULL UNIQUE,
                orden    INT NOT NULL DEFAULT 0,
                creado   TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        `);
        await pool.query(`
            CREATE TABLE IF NOT EXISTS platillos (
                id           INT AUTO_INCREMENT PRIMARY KEY,
                id_categoria INT NOT NULL,
                nombre       VARCHAR(120) NOT NULL,
                descripcion  TEXT,
                precio       DECIMAL(10,2) NOT NULL DEFAULT 0,
                foto_src     VARCHAR(500) DEFAULT NULL,
                disponible   TINYINT(1) NOT NULL DEFAULT 1,
                creado       TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (id_categoria) REFERENCES categorias(id) ON DELETE RESTRICT
            )
        `);
        await pool.query(`
            CREATE TABLE IF NOT EXISTS mesas (
                numero       INT PRIMARY KEY,
                capacidad    INT NOT NULL DEFAULT 4,
                estado       ENUM('libre','ordenando','en_cocina','lista','cuenta_pedida','pagada')
                             NOT NULL DEFAULT 'libre',
                qr_token     VARCHAR(100) DEFAULT NULL,
                qr_escaneado TINYINT(1) DEFAULT 0,
                creado       TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        `);
        await pool.query(`
            CREATE TABLE IF NOT EXISTS pedidos (
                id           INT AUTO_INCREMENT PRIMARY KEY,
                mesa_numero  INT NOT NULL,
                estado       ENUM('recibido','preparando','listo','entregado','cancelado')
                             NOT NULL DEFAULT 'recibido',
                total        DECIMAL(10,2) DEFAULT 0,
                creado       TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                actualizado  TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                FOREIGN KEY (mesa_numero) REFERENCES mesas(numero) ON DELETE CASCADE
            )
        `);
        await pool.query(`
            CREATE TABLE IF NOT EXISTS pedido_items (
                id              INT AUTO_INCREMENT PRIMARY KEY,
                pedido_id       INT NOT NULL,
                platillo_id     INT NOT NULL,
                cantidad        INT NOT NULL DEFAULT 1,
                nota            VARCHAR(200) DEFAULT NULL,
                estado          ENUM('recibido','preparando','listo','entregado','cancelado')
                                NOT NULL DEFAULT 'recibido',
                precio_unitario DECIMAL(10,2) NOT NULL,
                FOREIGN KEY (pedido_id)   REFERENCES pedidos(id) ON DELETE CASCADE,
                FOREIGN KEY (platillo_id) REFERENCES platillos(id) ON DELETE RESTRICT
            )
        `);
        await pool.query(`
            CREATE TABLE IF NOT EXISTS meseros (
                id       INT AUTO_INCREMENT PRIMARY KEY,
                nombre   VARCHAR(100) NOT NULL,
                activo   TINYINT(1) NOT NULL DEFAULT 1,
                creado   TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        `);
        await pool.query(`
            CREATE TABLE IF NOT EXISTS pagos (
                id           INT AUTO_INCREMENT PRIMARY KEY,
                mesa_numero  INT NOT NULL,
                pedido_id    INT NOT NULL,
                monto        DECIMAL(10,2) NOT NULL,
                metodo       ENUM('efectivo','tarjeta','transferencia','mixto') NOT NULL DEFAULT 'efectivo',
                estado       ENUM('pendiente','solicitado','autorizado','pagado') NOT NULL DEFAULT 'pendiente',
                creado       TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                actualizado  TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                FOREIGN KEY (mesa_numero) REFERENCES mesas(numero) ON DELETE CASCADE,
                FOREIGN KEY (pedido_id)   REFERENCES pedidos(id) ON DELETE CASCADE
            )
        `);
        console.log("[OK] [ESQUEMA] Tablas verificadas/creadas correctamente");
    } catch (err) {
        console.error("[ERROR] [ESQUEMA] No se pudo verificar el esquema:", err.message);
    }
}

// ── Salud del servidor ──
app.get("/", (_req, res) => {
    console.log("   [OK] Ping de salud respondido");
    res.json({ status: "ok", servicio: "UMAMI Restaurante API - Server 1" });
});

// ============================================================
//  CATEGORÍAS
// ============================================================
app.get("/categorias", async (_req, res) => {
    console.log("   [LISTAR] Listando categorías...");
    try {
        const [rows] = await pool.query("SELECT id, nombre, orden FROM categorias ORDER BY orden");
        console.log(`   [OK] ${rows.length} categoría(s) encontrada(s)`);
        res.json(rows);
    } catch (err) {
        responderError(res, "Fallo al listar categorías", err);
    }
});

// ============================================================
//  PLATILLOS
// ============================================================
app.get("/platillos", async (req, res) => {
    const cat = req.query.categoria;
    console.log(`   [LISTAR] Listando platillos${cat ? ` (categoría: ${cat})` : ""}...`);
    try {
        let sql = `
            SELECT p.id, p.id_categoria, p.nombre, p.descripcion,
                   p.precio, p.foto_src, p.disponible,
                   c.nombre AS categoria_nombre
            FROM platillos p
            JOIN categorias c ON p.id_categoria = c.id
            WHERE p.disponible = 1
        `;
        const params = [];
        if (cat) {
            sql += " AND p.id_categoria = ?";
            params.push(cat);
        }
        sql += " ORDER BY c.orden, p.nombre";

        const [rows] = await pool.query(sql, params);
        console.log(`   [OK] ${rows.length} platillo(s) encontrado(s)`);
        res.json(rows);
    } catch (err) {
        responderError(res, "Fallo al listar platillos", err);
    }
});

app.get("/platillo/:id", async (req, res) => {
    const id = req.params.id;
    console.log(`   [BUSCAR] Buscando platillo ID: ${id}`);
    try {
        const [rows] = await pool.query(
            `SELECT p.*, c.nombre AS categoria_nombre
             FROM platillos p
             JOIN categorias c ON p.id_categoria = c.id
             WHERE p.id = ?`,
            [id]
        );
        if (rows.length === 0) {
            console.log(`   [AVISO] Platillo ID ${id} no encontrado`);
            return res.status(404).json({ status: "error", mensaje: "Platillo no encontrado" });
        }
        console.log(`   [OK] Platillo encontrado: ${rows[0].nombre}`);
        res.json(rows[0]);
    } catch (err) {
        responderError(res, `Fallo al buscar platillo ID ${id}`, err);
    }
});

// ============================================================
//  PEDIDOS
// ============================================================

// Crear pedido (enviar orden desde la mesa)
app.post("/pedido", async (req, res) => {
    const { mesa, items } = req.body;
    console.log(`   [PEDIDO] Nuevo pedido para mesa ${mesa}, ${items ? items.length : 0} item(s)`);

    if (!mesa || !items || !Array.isArray(items) || items.length === 0) {
        console.log("   [AVISO] Datos incompletos para crear pedido");
        return res.status(400).json({ status: "error", mensaje: "Faltan datos (mesa o items)" });
    }

    const connection = await pool.getConnection();
    try {
        await connection.beginTransaction();

        // 1) Obtener precios de los platillos
        const platilloIds = items.map(i => i.platillo_id);
        const [platillos] = await connection.query(
            "SELECT id, precio FROM platillos WHERE id IN (?)",
            [platilloIds]
        );

        if (platillos.length !== platilloIds.length) {
            await connection.rollback();
            connection.release();
            console.log("   [AVISO] Uno o más platillos no existen");
            return res.status(400).json({ status: "error", mensaje: "Uno o más platillos no existen" });
        }

        const precioMap = {};
        platillos.forEach(p => { precioMap[p.id] = Number(p.precio); });

        // 2) Calcular total
        let total = 0;
        const enrichedItems = items.map(i => {
            const pu = precioMap[i.platillo_id];
            total += pu * (i.cantidad || 1);
            return { ...i, precio_unitario: pu };
        });

        // 3) Crear pedido
        const [pedidoRes] = await connection.query(
            "INSERT INTO pedidos (mesa_numero, estado, total) VALUES (?, 'recibido', ?)",
            [mesa, total]
        );
        const pedidoId = pedidoRes.insertId;

        // 4) Insertar items
        for (const item of enrichedItems) {
            await connection.query(
                "INSERT INTO pedido_items (pedido_id, platillo_id, cantidad, nota, estado, precio_unitario) VALUES (?, ?, ?, ?, 'recibido', ?)",
                [pedidoId, item.platillo_id, item.cantidad || 1, item.nota || null, item.precio_unitario]
            );
        }

        // 5) Actualizar estado de la mesa
        await connection.query(
            "UPDATE mesas SET estado = 'en_cocina' WHERE numero = ?",
            [mesa]
        );

        await connection.commit();
        connection.release();

        console.log("   [OK] Pedido creado: ID " + pedidoId + ", total=$" + total.toFixed(2) + ", mesa " + mesa + " → en_cocina");
        res.status(201).json({
            status: "ok",
            mensaje: "Orden enviada",
            pedido_id: pedidoId,
            total
        });
    } catch (err) {
        await connection.rollback();
        connection.release();
        responderError(res, "Fallo al crear pedido", err);
    }
});

// Obtener pedido activo de una mesa
app.get("/pedido/:mesa", async (req, res) => {
    const mesa = req.params.mesa;
    console.log(`   [BUSCAR] Consultando pedido activo de mesa ${mesa}`);
    try {
        const [pedidos] = await pool.query(
            `SELECT p.id, p.estado, p.total, p.creado, p.actualizado,
                    (SELECT JSON_ARRAYAGG(
                        JSON_OBJECT('id', pi.id, 'platillo_id', pi.platillo_id,
                                    'nombre', pl.nombre, 'cantidad', pi.cantidad,
                                    'nota', pi.nota, 'estado', pi.estado,
                                    'precio_unitario', pi.precio_unitario)
                     ) FROM pedido_items pi
                     JOIN platillos pl ON pi.platillo_id = pl.id
                     WHERE pi.pedido_id = p.id) AS items
             FROM pedidos p
             WHERE p.mesa_numero = ? AND p.estado NOT IN ('entregado','cancelado')
             ORDER BY p.creado DESC LIMIT 1`,
            [mesa]
        );
        if (pedidos.length === 0) {
            console.log(`   [AVISO] No hay pedido activo para mesa ${mesa}`);
            return res.json({ status: "ok", pedido: null });
        }
        const pedido = pedidos[0];
        // Parsear items si viene como string (MySQL puede devolver string para JSON)
        if (typeof pedido.items === "string") {
            try { pedido.items = JSON.parse(pedido.items); } catch(e) { pedido.items = []; }
        }
        console.log(`   [OK] Pedido encontrado: ID ${pedido.id}, estado=${pedido.estado}, ${pedido.items ? pedido.items.length : 0} items`);
        res.json({ status: "ok", pedido });
    } catch (err) {
        responderError(res, `Fallo al consultar pedido de mesa ${mesa}`, err);
    }
});

// Cambiar estado de un pedido (cocina → listo, mesero → entregado, etc.)
app.put("/pedido/:mesa/estado", async (req, res) => {
    const mesa = Number(req.params.mesa);
    const { estado } = req.body;
    const estadosValidos = ["recibido","preparando","listo","entregado","cancelado"];
    console.log(`   [ESTADO] Cambiando estado de pedido (mesa ${mesa}) → "${estado}"`);

    if (!estado || !estadosValidos.includes(estado)) {
        console.log("   [AVISO] Estado inválido:", estado);
        return res.status(400).json({ status: "error", mensaje: `Estado inválido. Válidos: ${estadosValidos.join(", ")}` });
    }

    const connection = await pool.getConnection();
    try {
        await connection.beginTransaction();

        const [pedidos] = await connection.query(
            "SELECT id, estado FROM pedidos WHERE mesa_numero = ? AND estado NOT IN ('entregado','cancelado') ORDER BY creado DESC LIMIT 1",
            [mesa]
        );
        if (pedidos.length === 0) {
            await connection.rollback();
            connection.release();
            console.log(`   [AVISO] No hay pedido activo para mesa ${mesa}`);
            return res.status(404).json({ status: "error", mensaje: "No hay pedido activo para esta mesa" });
        }

        const pedidoId = pedidos[0].id;
        await connection.query(
            "UPDATE pedidos SET estado = ? WHERE id = ?",
            [estado, pedidoId]
        );

        // Si todos los items se entregan, actualizar mesa
        let nuevoEstadoMesa = null;
        if (estado === "entregado") {
            nuevoEstadoMesa = "lista";
        } else if (estado === "preparando") {
            nuevoEstadoMesa = "en_cocina";
        } else if (estado === "cancelado") {
            // Verificar si quedan otros pedidos
            const [otros] = await connection.query(
                "SELECT COUNT(*) AS total FROM pedidos WHERE mesa_numero = ? AND estado NOT IN ('entregado','cancelado') AND id != ?",
                [mesa, pedidoId]
            );
            nuevoEstadoMesa = otros[0].total > 0 ? "en_cocina" : "libre";
        }

        if (nuevoEstadoMesa) {
            await connection.query(
                "UPDATE mesas SET estado = ? WHERE numero = ?",
                [nuevoEstadoMesa, mesa]
            );
            console.log(`   [OK] Mesa ${mesa} → estado "${nuevoEstadoMesa}"`);
        }

        await connection.commit();
        connection.release();

        console.log(`   [OK] Pedido ${pedidoId} (mesa ${mesa}) → "${estado}"`);
        res.json({ status: "ok", mensaje: `Pedido actualizado a ${estado}` });
    } catch (err) {
        await connection.rollback();
        connection.release();
        responderError(res, `Fallo al cambiar estado de pedido (mesa ${mesa})`, err);
    }
});

// ============================================================
//  MESAS
// ============================================================
app.get("/mesas", async (_req, res) => {
    console.log("   [LISTAR] Listando mesas...");
    try {
        const [rows] = await pool.query("SELECT numero, capacidad, estado, qr_escaneado FROM mesas ORDER BY numero");
        console.log(`   [OK] ${rows.length} mesa(s) encontrada(s)`);
        res.json(rows);
    } catch (err) {
        responderError(res, "Fallo al listar mesas", err);
    }
});

app.get("/mesa/:numero", async (req, res) => {
    const numero = req.params.numero;
    console.log(`   [BUSCAR] Consultando mesa ${numero}`);
    try {
        const [rows] = await pool.query("SELECT numero, capacidad, estado, qr_escaneado FROM mesas WHERE numero = ?", [numero]);
        if (rows.length === 0) {
            console.log(`   [AVISO] Mesa ${numero} no encontrada`);
            return res.status(404).json({ status: "error", mensaje: "Mesa no encontrada" });
        }
        console.log(`   [OK] Mesa ${numero}: estado=${rows[0].estado}`);
        res.json(rows[0]);
    } catch (err) {
        responderError(res, `Fallo al consultar mesa ${numero}`, err);
    }
});

// ============================================================
//  MESERO
// ============================================================
app.post("/solicitar-mesero", async (req, res) => {
    const { mesa } = req.body;
    console.log(`   [MESERO] Solicitud de mesero para mesa ${mesa}`);

    if (!mesa) {
        console.log("   [AVISO] Falta número de mesa");
        return res.status(400).json({ status: "error", mensaje: "Falta número de mesa" });
    }

    try {
        // Buscar primer mesero activo disponible
        const [meseros] = await pool.query(
            "SELECT id, nombre FROM meseros WHERE activo = 1 ORDER BY id LIMIT 1"
        );
        if (meseros.length === 0) {
            console.log("   [AVISO] No hay meseros activos");
            return res.status(404).json({ status: "error", mensaje: "No hay meseros disponibles" });
        }
        console.log(`   [OK] Mesero asignado: ${meseros[0].nombre} para mesa ${mesa}`);
        res.json({ status: "ok", mensaje: "Mesero en camino", mesero: meseros[0] });
    } catch (err) {
        responderError(res, `Fallo al solicitar mesero para mesa ${mesa}`, err);
    }
});

// ============================================================
//  PAGO / CUENTA
// ============================================================

// Solicitar la cuenta
app.post("/solicitar-cuenta", async (req, res) => {
    const { mesa } = req.body;
    console.log(`   [CUENTA] Solicitud de cuenta para mesa ${mesa}`);

    if (!mesa) {
        console.log("   [AVISO] Falta número de mesa");
        return res.status(400).json({ status: "error", mensaje: "Falta número de mesa" });
    }

    const connection = await pool.getConnection();
    try {
        await connection.beginTransaction();

        // Obtener pedido activo
        const [pedidos] = await connection.query(
            "SELECT id, total FROM pedidos WHERE mesa_numero = ? AND estado NOT IN ('entregado','cancelado') ORDER BY creado DESC LIMIT 1",
            [mesa]
        );
        if (pedidos.length === 0) {
            await connection.rollback();
            connection.release();
            console.log(`   [AVISO] No hay pedido activo para mesa ${mesa}`);
            return res.status(404).json({ status: "error", mensaje: "No hay pedido activo" });
        }

        const pedidoId = pedidos[0].id;
        const total = Number(pedidos[0].total);

        // Crear o actualizar registro de pago
        const [pagos] = await connection.query(
            "SELECT id, estado FROM pagos WHERE mesa_numero = ? AND pedido_id = ? AND estado IN ('pendiente','solicitado') LIMIT 1",
            [mesa, pedidoId]
        );

        let pagoId;
        if (pagos.length > 0) {
            await connection.query(
                "UPDATE pagos SET estado = 'solicitado', actualizado = NOW() WHERE id = ?",
                [pagos[0].id]
            );
            pagoId = pagos[0].id;
            console.log(`   [OK] Pago existente ${pagoId} actualizado a 'solicitado'`);
        } else {
            const [pagoRes] = await connection.query(
                "INSERT INTO pagos (mesa_numero, pedido_id, monto, metodo, estado) VALUES (?, ?, ?, 'efectivo', 'solicitado')",
                [mesa, pedidoId, total]
            );
            pagoId = pagoRes.insertId;
            console.log("   [OK] Pago creado: ID " + pagoId + ", monto=$" + total.toFixed(2));
        }

        // Actualizar mesa
        await connection.query(
            "UPDATE mesas SET estado = 'cuenta_pedida' WHERE numero = ?",
            [mesa]
        );

        await connection.commit();
        connection.release();

        console.log(`   [OK] Cuenta solicitada: mesa ${mesa}, pago_id=${pagoId}`);
        res.json({
            status: "ok",
            mensaje: "Cuenta solicitada",
            pago_id: pagoId,
            total
        });
    } catch (err) {
        await connection.rollback();
        connection.release();
        responderError(res, `Fallo al solicitar cuenta para mesa ${mesa}`, err);
    }
});

// Cambiar estado del pago (admin autoriza, marca pagado)
app.put("/pago/:mesa/estado", async (req, res) => {
    const mesa = Number(req.params.mesa);
    const { estado, metodo } = req.body;
    const estadosValidos = ["pendiente","solicitado","autorizado","pagado"];
    console.log(`   [PAGO] Cambiando estado de pago (mesa ${mesa}) → "${estado}"`);

    if (!estado || !estadosValidos.includes(estado)) {
        console.log("   [AVISO] Estado de pago inválido:", estado);
        return res.status(400).json({ status: "error", mensaje: `Estado inválido. Válidos: ${estadosValidos.join(", ")}` });
    }

    const connection = await pool.getConnection();
    try {
        await connection.beginTransaction();

        const [pagos] = await connection.query(
            "SELECT id, pedido_id FROM pagos WHERE mesa_numero = ? AND estado IN ('pendiente','solicitado','autorizado') ORDER BY creado DESC LIMIT 1",
            [mesa]
        );
        if (pagos.length === 0) {
            await connection.rollback();
            connection.release();
            console.log(`   [AVISO] No hay pago pendiente para mesa ${mesa}`);
            return res.status(404).json({ status: "error", mensaje: "No hay pago pendiente" });
        }

        const pagoId = pagos[0].id;
        const updateFields = [estado];
        let sql = "UPDATE pagos SET estado = ?";
        if (metodo) {
            sql += ", metodo = ?";
            updateFields.push(metodo);
        }
        sql += ", actualizado = NOW() WHERE id = ?";
        updateFields.push(pagoId);

        await connection.query(sql, updateFields);

        // Si se marca como pagado, liberar la mesa
        if (estado === "pagado") {
            await connection.query(
                "UPDATE mesas SET estado = 'pagada', qr_escaneado = 0 WHERE numero = ?",
                [mesa]
            );
            // También marcar el pedido como entregado
            await connection.query(
                "UPDATE pedidos SET estado = 'entregado' WHERE id = ?",
                [pagos[0].pedido_id]
            );
            console.log(`   [OK] Mesa ${mesa} liberada (pagada)`);
        } else if (estado === "autorizado") {
            await connection.query(
                "UPDATE mesas SET estado = 'cuenta_pedida' WHERE numero = ?",
                [mesa]
            );
        }

        await connection.commit();
        connection.release();

        console.log(`   [OK] Pago ${pagoId} (mesa ${mesa}) → "${estado}"`);
        res.json({ status: "ok", mensaje: `Pago actualizado a ${estado}` });
    } catch (err) {
        await connection.rollback();
        connection.release();
        responderError(res, `Fallo al cambiar estado de pago (mesa ${mesa})`, err);
    }
});

// Obtener estado del pago de una mesa
app.get("/pago/:mesa", async (req, res) => {
    const mesa = req.params.mesa;
    console.log(`   [BUSCAR] Consultando estado de pago, mesa ${mesa}`);
    try {
        const [rows] = await pool.query(
            `SELECT p.id, p.monto, p.metodo, p.estado, p.creado, p.actualizado
             FROM pagos p
             WHERE p.mesa_numero = ?
             ORDER BY p.creado DESC LIMIT 1`,
            [mesa]
        );
        if (rows.length === 0) {
            console.log(`   [AVISO] No hay pago para mesa ${mesa}`);
            return res.json({ status: "ok", pago: null });
        }
        console.log("   [OK] Pago encontrado: ID " + rows[0].id + ", estado=" + rows[0].estado + ", monto=$" + Number(rows[0].monto).toFixed(2));
        res.json({ status: "ok", pago: rows[0] });
    } catch (err) {
        responderError(res, `Fallo al consultar pago de mesa ${mesa}`, err);
    }
});

// ============================================================
//  QR
// ============================================================
app.get("/estado-qr/:mesa", async (req, res) => {
    const mesa = req.params.mesa;
    console.log(`   [QR] Verificando QR escaneado, mesa ${mesa}`);
    try {
        const [rows] = await pool.query(
            "SELECT qr_escaneado FROM mesas WHERE numero = ?",
            [mesa]
        );
        if (rows.length === 0) {
            console.log(`   [AVISO] Mesa ${mesa} no encontrada`);
            return res.status(404).json({ status: "error", mensaje: "Mesa no encontrada" });
        }
        const escaneado = rows[0].qr_escaneado ? true : false;
        console.log(`   [OK] Mesa ${mesa} QR escaneado: ${escaneado}`);
        res.json({ status: "ok", escaneado, mesa: Number(mesa) });
    } catch (err) {
        responderError(res, `Fallo al verificar QR mesa ${mesa}`, err);
    }
});

// Registrar escaneo de QR
app.post("/escanear-qr/:mesa", async (req, res) => {
    const mesa = req.params.mesa;
    console.log(`   [QR] Registrando escaneo QR, mesa ${mesa}`);
    try {
        const [result] = await pool.query(
            "UPDATE mesas SET qr_escaneado = 1, estado = 'ordenando' WHERE numero = ? AND (estado = 'libre' OR estado = 'pagada')",
            [mesa]
        );
        if (result.affectedRows === 0) {
            console.log(`   [AVISO] Mesa ${mesa} no se pudo actualizar (quizá ya ocupada)`);
            return res.json({ status: "ok", mensaje: "Mesa ya activa", ya_activa: true });
        }
        console.log(`   [OK] Mesa ${mesa} QR escaneado, estado → 'ordenando'`);
        res.json({ status: "ok", mensaje: "QR registrado", mesa: Number(mesa) });
    } catch (err) {
        responderError(res, `Fallo al registrar QR mesa ${mesa}`, err);
    }
});

// ============================================================
//  INFORMES Y EXPORTACIÓN
//  [AVISO] Las rutas /informe/exportar/* ANTES de /informe/:fecha
// ============================================================

async function consultarInforme(fecha) {
    const [rows] = await pool.query(
        `SELECT c.nombre AS categoria,
                COUNT(pi.id) AS items_vendidos,
                SUM(pi.cantidad * pi.precio_unitario) AS total_vendido,
                SUM(pi.cantidad) AS unidades
         FROM pedido_items pi
         JOIN pedidos p      ON pi.pedido_id = p.id
         JOIN platillos pl   ON pi.platillo_id = pl.id
         JOIN categorias c   ON pl.id_categoria = c.id
         WHERE DATE(p.creado) = ?
           AND p.estado NOT IN ('cancelado')
           AND pi.estado NOT IN ('cancelado')
         GROUP BY c.nombre, c.orden
         ORDER BY c.orden`,
        [fecha]
    );
    return rows;
}

// Exportar Excel - hoy
app.get("/informe/exportar/excel", async (_req, res) => {
    const fecha = hoyLocal();
    console.log(`   [EXCEL] Exportando Excel del día: ${fecha}`);
    try {
        const rows = await consultarInforme(fecha);
        await enviarExcel(res, rows, "Informe UMAMI", `Informe_UMAMI_${fecha}.xlsx`);
        console.log(`   [OK] Excel generado con ${rows.length} fila(s)`);
    } catch (err) {
        responderError(res, "Fallo al exportar a Excel", err);
    }
});

// Exportar Excel - fecha específica
app.get("/informe/exportar/excel/:fecha", async (req, res) => {
    const fecha = req.params.fecha;
    console.log(`   [EXCEL] Exportando Excel histórico: ${fecha}`);
    if (!fechaValida(fecha)) return res.status(400).json({ status: "error", mensaje: "Formato de fecha inválido" });
    try {
        const rows = await consultarInforme(fecha);
        await enviarExcel(res, rows, `Informe UMAMI ${fecha}`, `Informe_UMAMI_${fecha}.xlsx`);
        console.log(`   [OK] Excel histórico generado con ${rows.length} fila(s)`);
    } catch (err) {
        responderError(res, "Fallo al exportar informe histórico a Excel", err);
    }
});

// Exportar PDF - hoy
app.get("/informe/exportar/pdf", async (_req, res) => {
    const fecha = hoyLocal();
    console.log(`   [PDF] Exportando PDF del día: ${fecha}`);
    try {
        const rows = await consultarInforme(fecha);
        await enviarPDF(res, rows, `Informe Diario UMAMI (${fecha})`, "Resumen de Ventas", `Informe_UMAMI_${fecha}.pdf`);
        console.log(`   [OK] PDF generado con ${rows.length} fila(s)`);
    } catch (err) {
        responderError(res, "Fallo al exportar a PDF", err);
    }
});

// Exportar PDF - fecha específica
app.get("/informe/exportar/pdf/:fecha", async (req, res) => {
    const fecha = req.params.fecha;
    console.log(`   [PDF] Exportando PDF histórico: ${fecha}`);
    if (!fechaValida(fecha)) return res.status(400).json({ status: "error", mensaje: "Formato de fecha inválido" });
    try {
        const rows = await consultarInforme(fecha);
        await enviarPDF(res, rows, `Informe UMAMI (${fecha})`, `Resumen de Ventas - ${fecha}`, `Informe_UMAMI_${fecha}.pdf`);
        console.log(`   [OK] PDF histórico generado con ${rows.length} fila(s)`);
    } catch (err) {
        responderError(res, "Fallo al exportar informe histórico a PDF", err);
    }
});

// Informe de hoy (JSON)
app.get("/informe", async (_req, res) => {
    const fecha = hoyLocal();
    console.log(`   [INFORME] Informe del día solicitado: ${fecha}`);
    try {
        const rows = await consultarInforme(fecha);
        console.log(`   [OK] Informe generado: ${rows.length} categoría(s)`);
        res.json({ status: "ok", fecha, informe: rows });
    } catch (err) {
        responderError(res, "Fallo al generar el informe diario", err);
    }
});

// Informe histórico por fecha — DESPUÉS de /informe/exportar/*
app.get("/informe/:fecha", async (req, res) => {
    const fechaParam = req.params.fecha;
    console.log(`   [INFORME] Informe histórico solicitado para: ${fechaParam}`);
    if (!fechaValida(fechaParam)) {
        console.log("   [AVISO] Formato de fecha inválido:", fechaParam);
        return res.status(400).json({ status: "error", mensaje: "Formato de fecha inválido (use AAAA-MM-DD)" });
    }
    try {
        const rows = await consultarInforme(fechaParam);
        console.log(`   [OK] Informe histórico: ${rows.length} categoría(s) para ${fechaParam}`);
        res.json({ status: "ok", fecha: fechaParam, informe: rows });
    } catch (err) {
        responderError(res, "Fallo al generar el informe histórico", err);
    }
});

// ── Funciones de exportación ──
async function enviarExcel(res, rows, nombreHoja, nombreArchivo) {
    const workbook  = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(nombreHoja);

    worksheet.columns = [
        { header: "Categoría",     key: "categoria",     width: 25 },
        { header: "Items Vendidos", key: "items_vendidos", width: 18 },
        { header: "Unidades",      key: "unidades",       width: 14 },
        { header: "Total Vendido", key: "total_vendido",  width: 18 },
    ];
    worksheet.getRow(1).font = { bold: true };

    rows.forEach((row) => {
        worksheet.addRow({
            categoria:     String(row.categoria || ""),
            items_vendidos: Number(row.items_vendidos) || 0,
            unidades:       Number(row.unidades) || 0,
            total_vendido:  Number(row.total_vendido) || 0,
        });
    });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename=${nombreArchivo}`);
    await workbook.xlsx.write(res);
    res.end();
}

async function enviarPDF(res, rows, titulo, tituloTabla, nombreArchivo) {
    const doc = new PDFDocument({ margin: 30, size: "A4" });

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename=${nombreArchivo}`);
    doc.pipe(res);

    doc.fontSize(18).text(titulo, { align: "center" });
    doc.moveDown();

    const table = {
        title: tituloTabla,
        headers: ["Categoría", "Items Vendidos", "Unidades", "Total Vendido"],
        rows: rows.map((row) => [
            String(row.categoria || ""),
            String(Number(row.items_vendidos) || 0),
            String(Number(row.unidades) || 0),
            "$" + (Number(row.total_vendido) || 0).toFixed(2),
        ]),
    };

    await doc.table(table, {
        prepareHeader: () => doc.font("Helvetica-Bold").fontSize(12),
        prepareRow: () => doc.font("Helvetica").fontSize(10),
    });

    doc.end();
}

// ── Iniciar servidor ──
// Railway inyecta PORT. En local usa 3000.
const PORT = process.env.PORT || 3000;

asegurarEsquema().finally(() => {
    app.listen(PORT, "0.0.0.0", () => {
        console.log("\n=== Servidor UMAMI Restaurante ===");
        console.log("    Puerto: " + PORT);
        console.log("==================================");
        console.log("\n[APP FLUTTER / mesas1.py] Usa una de estas IPs:");
        const ifaces = os.networkInterfaces();
        Object.keys(ifaces).forEach(function(name) {
            ifaces[name].forEach(function(iface) {
                if (iface.family === "IPv4" && !iface.internal) {
                    console.log(`    http://${iface.address}:${PORT}   (interfaz: ${name})`);
                }
            });
        });
        console.log("\n    Copia la IP al kBaseUrl en Flutter o a SERVER_IP en mesas1.py");
        console.log("==================================\n");
    });
});
