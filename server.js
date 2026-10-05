// ============================================================
//  server.js  -  API Cocina Escolar INFRAMEN
//  CORRECCIONES:
//   - Rutas de exportación movidas ANTES de /informe/:fecha (evita conflicto)
//   - Pendientes busca por codigo_barra O carnet (igual que estudiante)
//   - Logs detallados en cada acción de la app
// ============================================================
const express = require("express");
const mysql = require("mysql2/promise");
const cors = require("cors");
const ExcelJS = require("exceljs");
const PDFDocument = require("pdfkit-table");
const os = require("os");

const app = express();

// [AVISO]  ORDEN CRÍTICO: cors y json deben ir ANTES del logger.
// Si el logger va primero, req.body llega undefined en POST/PUT.
app.use(cors());
app.use(express.json());

// ---------- Pool de conexiones ----------
// [HOSTING] Soporta 3 escenarios sin tocar el codigo:
//   1) Railway con URL completa  -> MYSQL_URL / DATABASE_URL / MYSQL_PUBLIC_URL
//   2) Railway por variables sueltas -> MYSQLHOST, MYSQLPORT, MYSQLUSER...
//   3) Variables genericas / local  -> DB_HOST, DB_PORT, DB_USER... (o defaults)
const connectionUrl =
    process.env.MYSQL_URL ||
    process.env.DATABASE_URL ||
    process.env.MYSQL_PUBLIC_URL ||
    null;

const pool = connectionUrl
    ? mysql.createPool({
          uri: connectionUrl,
          waitForConnections: true,
          connectionLimit: 10,
      })
    : mysql.createPool({
          host: process.env.MYSQLHOST || process.env.DB_HOST || "localhost",
          port: Number(process.env.MYSQLPORT || process.env.DB_PORT) || 3306,
          user: process.env.MYSQLUSER || process.env.DB_USER || "root",
          password: process.env.MYSQLPASSWORD || process.env.DB_PASSWORD || "230223",
          database: process.env.MYSQLDATABASE || process.env.DB_NAME || "cocina_escolar",
          waitForConnections: true,
          connectionLimit: 10,
      });

// ---------- Logger de solicitudes ----------
app.use((req, _res, next) => {
    console.log(`\n[REQ] [${new Date().toLocaleTimeString()}] ${req.method} ${req.url}`);
    if (req.body && Object.keys(req.body).length > 0) {
        console.log("   [BODY] Body:", JSON.stringify(req.body));
    }
    next();
});

// ---------- Utilidades ----------
function hoyLocal() {
    const d = new Date();
    const p = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function fechaValida(f) {
    return /^\d{4}-\d{2}-\d{2}$/.test(f);
}

async function consultarInforme(fecha) {
    const [rows] = await pool.query(
        `SELECT u.tipo,
                COUNT(*) AS entregados,
                SUM(m.estado = 'DEVUELTO') AS devueltos,
                SUM(m.estado = 'RETIRADO') AS pendientes
         FROM movimientos m
         JOIN utensilios u ON m.utensilio_id = u.id
         WHERE DATE(m.fecha_retiro) = ?
         GROUP BY u.tipo`,
        [fecha]
    );
    return rows;
}

function responderError(res, contexto, err, status = 500) {
    console.error(`[ERROR] [ERROR] ${contexto}:`, err.message);
    if (res.headersSent) return res.end();
    res.status(status).json({ status: "error", mensaje: err.message });
}

// Asegura que la tabla utensilios tenga la columna "cantidad"
async function asegurarEsquema() {
    try {
        const [cols] = await pool.query("SHOW COLUMNS FROM utensilios LIKE 'cantidad'");
        if (cols.length === 0) {
            await pool.query("ALTER TABLE utensilios ADD COLUMN cantidad INT NOT NULL DEFAULT 0");
            console.log("[OK] [ESQUEMA] Columna 'cantidad' creada en la tabla utensilios");
        } else {
            console.log("[OK] [ESQUEMA] Columna 'cantidad' ya existe en utensilios");
        }
    } catch (err) {
        console.error("[ERROR] [ESQUEMA] No se pudo verificar el esquema de utensilios:", err.message);
    }
}

// ---------- Salud del servidor ----------
app.get("/", (_req, res) => {
    console.log("   [OK] Ping de salud respondido");
    res.json({ status: "ok", servicio: "Cocina Escolar API - Server 3" });
});

// ============================================================
//  ESTUDIANTES
// ============================================================

// Buscar estudiante por código de barras o carnet
app.get("/estudiante/:codigo", async (req, res) => {
    const codigo = req.params.codigo;
    console.log(`   [BUSCAR] Buscando estudiante con código: "${codigo}"`);
    try {
        const [rows] = await pool.query(
            "SELECT id, nombre, carnet, grado FROM estudiantes WHERE codigo_barra = ? OR carnet = ?",
            [codigo, codigo]
        );

        if (rows.length === 0) {
            console.log(`   [AVISO]  Estudiante NO encontrado para código: "${codigo}"`);
            return res.status(404).json({ status: "error", mensaje: "Estudiante no encontrado" });
        }

        console.log(`   [OK] Estudiante encontrado: ${rows[0].nombre} (ID: ${rows[0].id}, carnet: ${rows[0].carnet})`);
        res.json({ status: "ok", estudiante: rows[0] });
    } catch (err) {
        responderError(res, "Fallo al buscar estudiante", err);
    }
});

// Registrar nuevo estudiante
app.post("/estudiante", async (req, res) => {
    const { nombre, carnet, codigo_barra, grado } = req.body;
    console.log(`   [REGISTRAR] Registrando estudiante: nombre="${nombre}", carnet="${carnet}", codigo_barra="${codigo_barra}", grado="${grado}"`);

    if (!nombre || !carnet || !codigo_barra) {
        console.log("   [AVISO]  Datos incompletos al registrar estudiante");
        return res.status(400).json({
            status: "error",
            mensaje: "Faltan datos obligatorios (nombre, carnet o codigo_barra)",
        });
    }

    try {
        const [result] = await pool.query(
            "INSERT INTO estudiantes (nombre, carnet, codigo_barra, grado) VALUES (?, ?, ?, ?)",
            [nombre, carnet, codigo_barra, grado || ""]
        );
        console.log(`   [OK] Estudiante registrado con ID: ${result.insertId}`);
        res.status(201).json({
            status: "ok",
            mensaje: "Estudiante registrado con éxito",
            estudiante_id: result.insertId,
        });
    } catch (err) {
        responderError(res, "Fallo al registrar estudiante", err);
    }
});

// ============================================================
//  UTENSILIOS (CRUD)
// ============================================================

// Listar todos los utensilios
app.get("/utensilios", async (_req, res) => {
    console.log("   [LISTAR] Listando inventario de utensilios...");
    try {
        const [rows] = await pool.query("SELECT id, tipo, cantidad FROM utensilios ORDER BY id");
        console.log(`   [OK] ${rows.length} utensilio(s) encontrado(s)`);
        res.json(rows);
    } catch (err) {
        responderError(res, "Fallo al listar utensilios", err);
    }
});

// Agregar utensilio
app.post("/utensilios", async (req, res) => {
    const tipo = String(req.body.tipo || "").trim();
    const cantidad = Number(req.body.cantidad);
    console.log(`   [AGREGAR] Agregando utensilio: tipo="${tipo}", cantidad=${cantidad}`);

    if (!tipo || !Number.isInteger(cantidad) || cantidad < 0) {
        console.log("   [AVISO]  Datos inválidos al agregar utensilio");
        return res.status(400).json({
            status: "error",
            mensaje: "Datos inválidos (tipo y cantidad entera >= 0)",
        });
    }

    try {
        const [result] = await pool.query(
            "INSERT INTO utensilios (tipo, cantidad) VALUES (?, ?)",
            [tipo, cantidad]
        );
        console.log(`   [OK] Utensilio agregado con ID: ${result.insertId}`);
        res.status(201).json({ status: "ok", mensaje: "Utensilio agregado", id: result.insertId });
    } catch (err) {
        responderError(res, "Fallo al agregar utensilio", err);
    }
});

// Editar utensilio
app.put("/utensilios/:id", async (req, res) => {
    const id = req.params.id;
    const tipo = String(req.body.tipo || "").trim();
    const cantidad = Number(req.body.cantidad);
    console.log(`   [EDITAR]  Editando utensilio ID ${id}: tipo="${tipo}", cantidad=${cantidad}`);

    if (!tipo || !Number.isInteger(cantidad) || cantidad < 0) {
        console.log("   [AVISO]  Datos inválidos al editar utensilio");
        return res.status(400).json({
            status: "error",
            mensaje: "Datos inválidos (tipo y cantidad entera >= 0)",
        });
    }

    try {
        const [result] = await pool.query(
            "UPDATE utensilios SET tipo = ?, cantidad = ? WHERE id = ?",
            [tipo, cantidad, id]
        );
        if (result.affectedRows === 0) {
            console.log(`   [AVISO]  Utensilio ID ${id} no existe`);
            return res.status(404).json({ status: "error", mensaje: "Utensilio no existe" });
        }
        console.log(`   [OK] Utensilio ID ${id} actualizado correctamente`);
        res.json({ status: "ok", mensaje: "Utensilio actualizado" });
    } catch (err) {
        responderError(res, "Fallo al editar utensilio", err);
    }
});

// Eliminar utensilio
app.delete("/utensilios/:id", async (req, res) => {
    const id = req.params.id;
    console.log(`   [ELIMINAR]  Eliminando utensilio ID: ${id}`);
    try {
        const [result] = await pool.query("DELETE FROM utensilios WHERE id = ?", [id]);
        if (result.affectedRows === 0) {
            console.log(`   [AVISO]  Utensilio ID ${id} no existe`);
            return res.status(404).json({ status: "error", mensaje: "Utensilio no existe" });
        }
        console.log(`   [OK] Utensilio ID ${id} eliminado`);
        res.json({ status: "ok", mensaje: "Utensilio eliminado" });
    } catch (err) {
        if (err.code === "ER_ROW_IS_REFERENCED_2" || err.code === "ER_ROW_IS_REFERENCED") {
            console.log(`   [AVISO]  No se puede eliminar utensilio ID ${id}: tiene movimientos registrados`);
            return res.status(409).json({
                status: "error",
                mensaje: "No se puede eliminar: el utensilio tiene movimientos registrados",
            });
        }
        responderError(res, "Fallo al eliminar utensilio", err);
    }
});

// ============================================================
//  MOVIMIENTOS (RETIROS / DEVOLUCIONES)
// ============================================================

// Registrar retiro
app.post("/retiro", async (req, res) => {
    const { estudiante_id, tipo } = req.body;
    console.log(`   [RETIRO] Retiro solicitado: estudiante_id=${estudiante_id}, tipo="${tipo}"`);

    if (!estudiante_id || !tipo) {
        console.log("   [AVISO]  Faltan datos para registrar retiro");
        return res.status(400).json({
            status: "error",
            mensaje: "Faltan datos obligatorios (estudiante_id o tipo)",
        });
    }

    const connection = await pool.getConnection();
    try {
        await connection.beginTransaction();

        // Buscar utensilio y verificar stock
        const [uten] = await connection.query(
            "SELECT id, cantidad FROM utensilios WHERE LOWER(TRIM(tipo)) = LOWER(TRIM(?)) LIMIT 1",
            [tipo]
        );

        if (uten.length === 0) {
            await connection.rollback();
            connection.release();
            console.log(`   [AVISO]  Tipo de utensilio inválido: "${tipo}"`);
            return res.status(400).json({ status: "error", mensaje: `Tipo de utensilio inválido: ${tipo}` });
        }

        const utensilioId = uten[0].id;
        const stockTotal = uten[0].cantidad;

        // Contar cuántos están retirados actualmente
        const [pendientesRes] = await connection.query(
            "SELECT COUNT(*) AS total_pendientes FROM movimientos WHERE utensilio_id = ? AND estado = 'RETIRADO'",
            [utensilioId]
        );
        const totalPendientes = pendientesRes[0].total_pendientes;
        const disponibles = stockTotal - totalPendientes;

        console.log(`   [STOCK] Stock "${tipo}": total=${stockTotal}, en uso=${totalPendientes}, disponibles=${disponibles}`);

        if (totalPendientes >= stockTotal) {
            await connection.rollback();
            connection.release();
            console.log(`   [ERROR] Sin stock disponible de "${tipo}"`);
            return res.status(400).json({
                status: "error",
                mensaje: `No hay stock disponible de '${tipo}' (Todos están en uso o retirados)`,
            });
        }

        // Registrar el movimiento
        const [result] = await connection.query(
            "INSERT INTO movimientos (estudiante_id, utensilio_id, fecha_retiro, estado) VALUES (?, ?, NOW(), 'RETIRADO')",
            [estudiante_id, utensilioId]
        );

        await connection.commit();
        connection.release();

        console.log(`   [OK] Retiro registrado: movimiento_id=${result.insertId}, estudiante_id=${estudiante_id}, tipo="${tipo}"`);
        res.status(201).json({ status: "ok", mensaje: "Retiro registrado", movimiento_id: result.insertId });
    } catch (err) {
        await connection.rollback();
        connection.release();
        responderError(res, "Fallo al registrar retiro en base de datos", err);
    }
});

// Registrar devolución
app.put("/devolucion/:id", async (req, res) => {
    const movimientoId = req.params.id;
    console.log(`   [DEVOLUCION] Devolución solicitada para movimiento ID: ${movimientoId}`);
    try {
        const [result] = await pool.query(
            "UPDATE movimientos SET fecha_devolucion = NOW(), estado = 'DEVUELTO' WHERE id = ? AND estado = 'RETIRADO'",
            [movimientoId]
        );

        if (result.affectedRows === 0) {
            console.log(`   [AVISO]  Movimiento ID ${movimientoId}: no existe o ya fue devuelto`);
            return res.status(404).json({ status: "error", mensaje: "Movimiento no existe o ya fue devuelto" });
        }

        console.log(`   [OK] Devolución exitosa para movimiento ID: ${movimientoId}`);
        res.json({ status: "ok", mensaje: "Devolución registrada" });
    } catch (err) {
        responderError(res, `Fallo al ejecutar devolución para ID ${movimientoId}`, err);
    }
});

// ============================================================
//  PENDIENTES
//  CORRECCIÓN: Busca por codigo_barra O carnet (igual que /estudiante/:codigo)
//  Esto soluciona el problema cuando la app escanea el código de barras
//  pero el servidor antes solo buscaba por número de carnet.
// ============================================================
app.get("/pendientes/:codigo", async (req, res) => {
    const codigo = req.params.codigo;
    console.log(`   [BUSCAR] Consultando pendientes para código/carnet: "${codigo}"`);
    try {
        const [rows] = await pool.query(
            `SELECT m.id, u.tipo, m.fecha_retiro
             FROM movimientos m
             JOIN estudiantes e ON m.estudiante_id = e.id
             JOIN utensilios u ON m.utensilio_id = u.id
             WHERE (e.carnet = ? OR e.codigo_barra = ?) AND m.estado = 'RETIRADO'`,
            [codigo, codigo]
        );
        console.log(`   [OK] ${rows.length} préstamo(s) pendiente(s) para "${codigo}"`);
        if (rows.length > 0) {
            rows.forEach(r => console.log(`      - ID: ${r.id}, tipo: ${r.tipo}, retirado: ${r.fecha_retiro}`));
        }
        res.json(rows);
    } catch (err) {
        responderError(res, `Fallo al consultar pendientes del código ${codigo}`, err);
    }
});

// ============================================================
//  INFORMES Y EXPORTACIÓN
//  [AVISO]  IMPORTANTE: Las rutas /informe/exportar/* deben ir ANTES de /informe/:fecha
//  Si van después, Express interpreta "exportar" como una fecha y nunca las alcanza.
// ============================================================

// Exportar Excel - hoy
app.get("/informe/exportar/excel", async (_req, res) => {
    const fecha = hoyLocal();
    console.log(`   [STOCK] Exportando Excel del día: ${fecha}`);
    try {
        const rows = await consultarInforme(fecha);
        console.log(`   [OK] Excel generado con ${rows.length} fila(s)`);
        await enviarExcel(res, rows, "Informe Diario", `Informe_Cocina_${fecha}.xlsx`);
    } catch (err) {
        responderError(res, "Fallo al exportar a Excel", err);
    }
});

// Exportar Excel - fecha específica
app.get("/informe/exportar/excel/:fecha", async (req, res) => {
    const fecha = req.params.fecha;
    console.log(`   [STOCK] Exportando Excel histórico: ${fecha}`);
    if (!fechaValida(fecha)) return res.status(400).json({ status: "error", mensaje: "Formato de fecha inválido" });
    try {
        const rows = await consultarInforme(fecha);
        console.log(`   [OK] Excel histórico generado con ${rows.length} fila(s)`);
        await enviarExcel(res, rows, `Informe ${fecha}`, `Informe_Cocina_${fecha}.xlsx`);
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
        console.log(`   [OK] PDF generado con ${rows.length} fila(s)`);
        await enviarPDF(res, rows, `Informe Diario de Cocina Escolar (${fecha})`, "Resumen de Movimientos", `Informe_Cocina_${fecha}.pdf`);
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
        console.log(`   [OK] PDF histórico generado con ${rows.length} fila(s)`);
        await enviarPDF(res, rows, `Informe de Cocina Escolar (${fecha})`, `Resumen de Movimientos - ${fecha}`, `Informe_Cocina_${fecha}.pdf`);
    } catch (err) {
        responderError(res, "Fallo al exportar informe histórico a PDF", err);
    }
});

// Informe de hoy (JSON)
app.get("/informe", async (_req, res) => {
    const fechaHoy = hoyLocal();
    console.log(`   [INFORME] Informe del día solicitado: ${fechaHoy}`);
    try {
        const rows = await consultarInforme(fechaHoy);
        console.log(`   [OK] Informe generado: ${rows.length} tipo(s) de utensilio(s)`);
        res.json({ status: "ok", fecha: fechaHoy, informe: rows });
    } catch (err) {
        responderError(res, "Fallo al generar el informe diario", err);
    }
});

// Informe histórico por fecha (JSON) — [AVISO] Debe ir DESPUÉS de /informe/exportar/*
app.get("/informe/:fecha", async (req, res) => {
    const fechaParam = req.params.fecha;
    console.log(`   [INFORME] Informe histórico solicitado para: ${fechaParam}`);
    if (!fechaValida(fechaParam)) {
        console.log("   [AVISO]  Formato de fecha inválido:", fechaParam);
        return res.status(400).json({ status: "error", mensaje: "Formato de fecha inválido (use AAAA-MM-DD)" });
    }
    try {
        const rows = await consultarInforme(fechaParam);
        console.log(`   [OK] Informe histórico: ${rows.length} tipo(s) para ${fechaParam}`);
        res.json({ status: "ok", fecha: fechaParam, informe: rows });
    } catch (err) {
        responderError(res, "Fallo al generar el informe histórico", err);
    }
});

// ---------- Funciones de exportación ----------
async function enviarExcel(res, rows, nombreHoja, nombreArchivo) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(nombreHoja);

    worksheet.columns = [
        { header: "Tipo de Utensilio", key: "tipo", width: 25 },
        { header: "Entregados", key: "entregados", width: 15 },
        { header: "Devueltos", key: "devueltos", width: 15 },
        { header: "Pendientes", key: "pendientes", width: 15 },
    ];
    worksheet.getRow(1).font = { bold: true };

    rows.forEach((row) => {
        worksheet.addRow({
            tipo: row.tipo,
            entregados: Number(row.entregados) || 0,
            devueltos: Number(row.devueltos) || 0,
            pendientes: Number(row.pendientes) || 0,
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
        headers: ["Utensilio", "Entregados", "Devueltos", "Pendientes"],
        rows: rows.map((row) => [
            String(row.tipo),
            String(Number(row.entregados) || 0),
            String(Number(row.devueltos) || 0),
            String(Number(row.pendientes) || 0),
        ]),
    };

    await doc.table(table, {
        prepareHeader: () => doc.font("Helvetica-Bold").fontSize(12),
        prepareRow: () => doc.font("Helvetica").fontSize(10),
    });

    doc.end();
}

// ---------- Iniciar servidor ----------
// [HOSTING] Puerto dinámico: AlwaysData (o cualquier hosting) inyecta el
// puerto por la variable de entorno PORT. En tu PC local sigue usando 3000.
const PORT = process.env.PORT || 3000;

asegurarEsquema().finally(() => {
    app.listen(PORT, "0.0.0.0", () => {
        console.log("\n=== Servidor Cocina Escolar INFRAMEN ===");
        console.log("    Puerto: " + PORT);
        console.log("========================================");
        console.log("\n[APP FLUTTER] Usa una de estas IPs en kBaseUrl:");
        const ifaces = os.networkInterfaces();
        Object.keys(ifaces).forEach(function(name) {
            ifaces[name].forEach(function(iface) {
                if (iface.family === "IPv4" && !iface.internal) {
                    console.log("    http://" + iface.address + ":" + PORT + "   (interfaz: " + name + ")");
                }
            });
        });
        console.log("\n    Copia la IP de arriba al kBaseUrl en main.dart");
        console.log("========================================\n");
    });
});