const carpeta = "https://drive.google.com/file/d/19JlPNv17CWp8RnF1bdiA7C25K2G17U11/view?usp=drive_link"
/**
 * setupDatabase
 * Ejecuta esta función para inicializar la estructura de tablas.
 * Respeta el esquema relacional para el Módulo de Configuración.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Definición de las tablas maestras de configuración
  const tables = {
    "CONF_ETAPAS": ["id", "nombre_etapa", "orden", "color_hex", "descripcion", "created_at"],
    "CONF_TAREAS": ["id", "etapa_id", "nombre_tarea", "requiere_evidencia", "created_at"],
    "CONF_PROFESIONALES": ["id", "nombre_completo", "especialidad", "email", "created_at"],
    "CONF_CHECKLISTS": ["id", "nombre_checklist", "config_json", "created_at"], // config_json guardará tipos y orden
    "CONF_GENERAL": ["parametro", "valor", "descripcion", "updated_at"] // Para la URL de Drive
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

      // Configurar encabezados si la hoja está vacía o si queremos resetear (con precaución)
      const headers = tables[tableName];
      sheet.getRange(1, 1, 1, headers.length)
           .setValues([headers])
           .setBackground("#556B2F") // Color base GAS Expert
           .setFontColor("white")
           .setFontWeight("bold");
      
      // Congelar la primera fila
      sheet.setFrozenRows(1);
    });

    // Inicializar parámetro de Drive si no existe
    const confSheet = ss.getSheetByName("CONF_GENERAL");
    const driveParam = "DRIVE_ROOT_FOLDER_URL";
    const data = confSheet.getDataRange().getValues();
    const exists = data.some(row => row[0] === driveParam);

    if (!exists) {
      confSheet.appendRow([driveParam, "", "URL raíz para almacenamiento de evidencias", new Date()]);
    }

    SpreadsheetApp.getUi().alert("✅ Base de datos inicializada con éxito.");
    
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

