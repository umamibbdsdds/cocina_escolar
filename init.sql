-- ============================================================
--  init.sql  -  Crea las tablas de Control de Vajilla Escolar
--  Ejecutalo en la pestana "Data / Query" del servicio MySQL de Railway
--  (o conectandote con MySQL Workbench a la URL publica de Railway).
-- ============================================================

-- 1) ESTUDIANTES
CREATE TABLE IF NOT EXISTS estudiantes (
    id           INT AUTO_INCREMENT PRIMARY KEY,
    nombre       VARCHAR(150) NOT NULL,
    carnet       VARCHAR(50)  NOT NULL,
    codigo_barra VARCHAR(50)  NOT NULL,
    grado        VARCHAR(50)  DEFAULT '',
    UNIQUE KEY uq_carnet (carnet),
    UNIQUE KEY uq_codigo_barra (codigo_barra)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 2) UTENSILIOS
CREATE TABLE IF NOT EXISTS utensilios (
    id       INT AUTO_INCREMENT PRIMARY KEY,
    tipo     VARCHAR(100) NOT NULL,
    cantidad INT NOT NULL DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 3) MOVIMIENTOS (retiros / devoluciones)
CREATE TABLE IF NOT EXISTS movimientos (
    id               INT AUTO_INCREMENT PRIMARY KEY,
    estudiante_id    INT NOT NULL,
    utensilio_id     INT NOT NULL,
    fecha_retiro     DATETIME NOT NULL,
    fecha_devolucion DATETIME NULL,
    estado           ENUM('RETIRADO','DEVUELTO') NOT NULL DEFAULT 'RETIRADO',
    CONSTRAINT fk_mov_estudiante
        FOREIGN KEY (estudiante_id) REFERENCES estudiantes(id),
    CONSTRAINT fk_mov_utensilio
        FOREIGN KEY (utensilio_id) REFERENCES utensilios(id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- Datos iniciales de ejemplo (opcional)
INSERT INTO utensilios (tipo, cantidad) VALUES
  ('Plato',    50),
  ('Vaso',     50),
  ('Cuchara',  50),
  ('Tenedor',  50);
