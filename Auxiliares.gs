const carpeta = "https://drive.google.com/file/d/19JlPNv17CWp8RnF1bdiA7C25K2G17U11/view?usp=drive_link"
/**
 * setupDatabase
 * Ejecuta esta función para inicializar la estructura de tablas.
 * Respeta el esquema relacional para el Módulo de Configuración.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Definición de las tablas del sistema
  const tables = {
    // --- MÓDULO DE CONFIGURACIÓN (Templates) ---
    "CONF_ETAPAS": ["id", "nombre_etapa", "orden", "color_hex", "descripcion", "created_at"],
    "CONF_TAREAS": ["id", "etapa_id", "nombre_tarea", "requiere_evidencia", "created_at"],
    
    "CONF_PROFESIONALES": [
      "id", 
      "nombre_completo", 
      "especialidad", 
      "rol",           
      "telefono",      
      "email", 
      "costo_hora",    
      "estado",        
      "created_at"
    ],
    
    "CONF_CHECKLISTS": ["id", "nombre_checklist", "config_json", "created_at"], 
    "CONF_GENERAL": ["parametro", "valor", "descripcion", "updated_at"],

    // --- MÓDULO OPERATIVO (Datos Reales) ---
    // Nueva tabla para alojar los proyectos creados
    "DB_PROYECTOS": [
      "id",               // UUID único del sistema
      "codigo",           // Código legible (ej: OBR-2024-001)
      "nombre_obra",      // Nombre del proyecto
      "cliente",          // Cliente principal
      "ubicacion",        // Dirección / Ciudad
      "fecha_inicio",     
      "fecha_fin",        // Fecha estimada de entrega
      "estado",           // Planificación | En Ejecución | Finalizado | Detenido
      "drive_folder_id",  // ID de la carpeta en Google Drive (CRUCIAL)
      "created_at"
    ],
    "DB_EJECUCION": [
      "id",               // UUID único de la instancia
      "proyecto_id",      // Vinculación con el Proyecto
      "etapa_id",         // Para agrupar (copiado de config)
      "nombre_tarea",     // Copiado de config (para snapshot)
      "requiere_evidencia",
      "estado",           // Pendiente | En Proceso | Aprobado | Rechazado
      "responsable_id",   // Quién debe hacerlo (opcional)
      "evidencia_url",    // Link a la foto/archivo en Drive
      "comentarios",      // Observaciones de obra
      "updated_at"
    ]
  };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // 15 segundos de seguridad

    Object.keys(tables).forEach(tableName => {
      let sheet = ss.getSheetByName(tableName);
      if (!sheet) {
        sheet = ss.insertSheet(tableName);
        console.log(`Hoja creada: ${tableName}`);
      }

      // Configurar encabezados
      const headers = tables[tableName];
      
      // Estilizado profesional de encabezados
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange
           .setValues([headers])
           .setBackground("#556B2F") // Color base Lock
           .setFontColor("white")
           .setFontWeight("bold")
           .setHorizontalAlignment("center");
      
      // Ajuste visual extra
      sheet.setFrozenRows(1);
      // Auto-resize solo si es nueva (para no molestar visualmente si ya tiene datos)
      if (sheet.getLastRow() <= 1) sheet.autoResizeColumns(1, headers.length);
    });

    // Inicializar parámetro de Drive si no existe
    const confSheet = ss.getSheetByName("CONF_GENERAL");
    const driveParam = "DRIVE_ROOT_FOLDER_URL";
    const data = confSheet.getDataRange().getValues();
    const exists = data.some(row => row[0] === driveParam);

    if (!exists) {
      confSheet.appendRow([driveParam, "", "URL raíz para almacenamiento de evidencias", new Date()]);
    }

    SpreadsheetApp.getUi().alert("✅ Base de datos actualizada. Tabla DB_PROYECTOS lista.");
    
  } catch (e) {
    console.error("Error en setupDatabase: " + e.toString());
    throw new Error("No se pudo inicializar la base de datos.");
  } finally {
    lock.releaseLock();
  }
}

/**
 * Función de utilidad para generar IDs únicos (UUIDv4 simplificado)
 */
function generateUUID() {
  return Utilities.getUuid();
}

