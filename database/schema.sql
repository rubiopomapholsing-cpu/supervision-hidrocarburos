CREATE TABLE IF NOT EXISTS supervisiones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    instalacion TEXT NOT NULL,
    fecha TEXT NOT NULL,
    responsable TEXT NOT NULL,
    estado TEXT NOT NULL DEFAULT 'Pendiente'
        CHECK (estado IN ('Pendiente', 'En proceso', 'Completada')),
    observaciones TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS registros (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    oficina_regional TEXT,
    mes_presenta TEXT,
    accion TEXT,
    accion_especifica TEXT,
    sub_accion TEXT,
    tipo_agente TEXT,
    actividad TEXT,
    provincia TEXT,
    distrito TEXT,
    tipo_transporte TEXT,
    num_supervisores INTEGER,
    empresa_supervisora TEXT,
    calidad_entregable TEXT,
    expediente_siged TEXT,
    carta_linea TEXT,
    observaciones TEXT
);
