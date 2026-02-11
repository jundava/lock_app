const carpeta = "https://drive.google.com/file/d/19JlPNv17CWp8RnF1bdiA7C25K2G17U11/view?usp=drive_link"
/**
 * @file setupDatabase.gs
 * @summary Inicializa y normaliza la infraestructura de base de datos en Google Sheets.
 * @author GAS Expert
 */

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const COLOR_BASE = "#556B2F";

  // 1. DEFINICIÓN DEL ESQUEMA MAESTRO (Basado en análisis de CSVs)
  const schema = {
    "CONF_GENERAL": ["parametro", "valor", "descripcion", "updated_at"],
    "CONF_TIPO_PROYECTO": ["id", "nombre_tipo", "descripcion", "color_representativo", "created_at"],
    "CONF_ETAPAS": ["id", "nombre_etapa", "orden", "color_hex", "descripcion", "id_tipo_proyecto", "created_at"],
    "CONF_TAREAS": ["id", "etapa_id", "nombre_tarea", "requiere_evidencia", "tipo_entrada", "checklist_id", "created_at"],
    "CONF_PROFESIONALES": ["id", "nombre_completo", "especialidad", "rol", "telefono", "email", "costo_hora", "estado", "created_at", "perfil_sistema"],
    "CONF_CHECKLISTS": ["id", "nombre_checklist", "config_json", "id_tipo_proyecto", "created_at"],
    "CONF_REL_ASIGNACIONES": ["id", "id_proyecto", "id_profesional", "created_at"],
    "DB_PROYECTOS": ["id", "codigo", "nombre_obra", "cliente", "ubicacion", "fecha_inicio", "fecha_fin", "estado", "drive_folder_id", "drive_url", "id_tipo_proyecto", "created_at"],
    "DB_EJECUCION": ["id", "proyecto_id", "etapa_id", "nombre_tarea", "requiere_evidencia", "tipo_entrada", "checklist_id", "estado", "responsable_id", "datos_evidencia", "comentarios", "updated_at", "datos_checklist"]
  };

  const lock = LockService.getScriptLock();
  
  try {
    // Bloqueo por 30 segundos para operaciones de estructura
    lock.waitLock(30000);
    console.log("🚀 Iniciando construcción de infraestructura relacional...");

    Object.keys(schema).forEach(tableName => {
      let sheet = ss.getSheetByName(tableName);
      const expectedHeaders = schema[tableName];

      if (!sheet) {
        sheet = ss.insertSheet(tableName);
        console.log(`✅ Hoja creada: ${tableName}`);
      }

      const currentHeaders = sheet.getLastColumn() > 0 
        ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] 
        : [];

      // Sincronización de columnas (IDEMPOTENCIA)
      expectedHeaders.forEach((header, index) => {
        const colIndex = currentHeaders.indexOf(header);
        if (colIndex === -1) {
          // Si la columna no existe, la agregamos
          const newColPos = index + 1;
          sheet.insertColumnBefore(newColPos);
          const cell = sheet.getRange(1, newColPos);
          cell.setValue(header);
          
          // Estilo Profesional
          cell.setBackground(COLOR_BASE)
              .setFontColor("white")
              .setFontWeight("bold")
              .setHorizontalAlignment("center");
          
          console.log(`  + Columna agregada en ${tableName}: ${header}`);
        }
      });

      // Formato de tabla
      sheet.setFrozenRows(1);
      if (sheet.getLastColumn() > 0) {
        sheet.autoResizeColumns(1, sheet.getLastColumn());
      }
    });

    // 2. INICIALIZACIÓN DE DATOS SEMILLA (Opcional - CONF_GENERAL)
    const confSheet = ss.getSheetByName("CONF_GENERAL");
    const existingParams = confSheet.getDataRange().getValues().map(r => r[0]);
    
    const seedData = [
      ["DRIVE_ROOT_FOLDER_URL", carpeta, "URL raíz para evidencias", new Date()],
      ["TIPOS_PROYECTO", "Siroque, Expansión, Especial", "Listado de tipos de obra", new Date()]
    ];

    seedData.forEach(row => {
      if (!existingParams.includes(row[0])) {
        confSheet.appendRow(row);
      }
    });

    // 3. LIMPIEZA DE CACHE
    CacheService.getScriptCache().removeAll(["indices_db", "config_app"]);
    
    console.log("🎯 Base de datos consolidada con éxito.");
    
    // Feedback al usuario (Solo si hay UI)
    try {
      SpreadsheetApp.getUi().alert("✅ Infraestructura Lista", "La base de datos ha sido normalizada y las columnas relacionales están activas.", SpreadsheetApp.getUi().ButtonSet.OK);
    } catch(e) {}

  } catch (e) {
    console.error("Fatal Error en setupDatabase: " + e.message);
    throw "Error de Infraestructura: Contacte al administrador.";
  } finally {
    lock.releaseLock();
  }
}
