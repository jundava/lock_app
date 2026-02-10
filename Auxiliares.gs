const carpeta = "https://drive.google.com/file/d/19JlPNv17CWp8RnF1bdiA7C25K2G17U11/view?usp=drive_link"
/**
 * setupDatabase
 * Ejecuta esta función para inicializar la estructura de tablas.
 * Respeta el esquema relacional para el Módulo de Configuración.
 */
function setupDatabaseA() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Definición de las tablas del sistema
  const tables = {
    // --- MÓDULO DE CONFIGURACIÓN (Templates) ---
    "CONF_ETAPAS": ["id", "nombre_etapa", "orden", "color_hex", "descripcion", "created_at"],
    
    "CONF_TAREAS": [
      "id", 
      "etapa_id", 
      "nombre_tarea", 
      "requiere_evidencia", 
      "tipo_entrada",       // <--- NUEVO: 'text' | 'textarea'
      "checklist_id",       // <--- NUEVO: ID del checklist opcional
      "created_at"
    ],
    
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
    "DB_PROYECTOS": [
      "id",               // UUID único
      "codigo",           // Código legible
      "nombre_obra",      // Nombre
      "cliente",          // Cliente
      "ubicacion",        // Dirección
      "fecha_inicio",     
      "fecha_fin",        
      "estado",           
      "drive_folder_id",  // ID carpeta
      "drive_url",        // <--- NUEVO: URL directa (evita errores de regeneración)
      "created_at"
    ],
    
    "DB_EJECUCION": [
      "id",
      "proyecto_id",
      "etapa_id",
      "nombre_tarea",
      "requiere_evidencia",
      "tipo_entrada",       // <--- NUEVO: Copiado de CONF_TAREAS
      "checklist_id",       // <--- NUEVO: Copiado de CONF_TAREAS
      "estado",             // Pendiente | Completado
      "responsable_id",
      "datos_evidencia",    // JSON respuestas o URL Foto
      "comentarios",
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
      // Auto-resize para mejorar legibilidad
      sheet.autoResizeColumns(1, headers.length);
    });

    // Inicializar parámetro de Drive si no existe
    const confSheet = ss.getSheetByName("CONF_GENERAL");
    const driveParam = "DRIVE_ROOT_FOLDER_URL";
    const data = confSheet.getDataRange().getValues();
    const exists = data.some(row => row[0] === driveParam);

    if (!exists) {
      confSheet.appendRow([driveParam, "", "URL raíz para almacenamiento de evidencias", new Date()]);
    }

    SpreadsheetApp.getUi().alert("✅ Base de datos actualizada con nuevas columnas (Tipo y Checklist).");
    
  } catch (e) {
    console.error("Error en setupDatabase: " + e.toString());
    throw new Error("No se pudo inicializar la base de datos.");
  } finally {
    lock.releaseLock();
  }
}

function setupDatabaseB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const CONFIG_COL = "tipo_proyecto";
  
  // 1. Definición de Hojas y su configuración de inserción
  const esquemas = [
    { nombre: "CONF_ETAPAS", desc: "Ciclos de vida" },
    { nombre: "CONF_CHECKLISTS", desc: "Protocolos técnicos" },
    { nombre: "DB_PROYECTOS", desc: "Maestro de obras" }
  ];

  console.log("🚀 Iniciando consolidación de infraestructura...");

  esquemas.forEach(esquema => {
    const sheet = ss.getSheetByName(esquema.nombre);
    if (!sheet) {
      console.error(`❌ Error: La hoja ${esquema.nombre} no fue encontrada.`);
      return;
    }

    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // Verificamos si la columna ya existe para garantizar idempotencia
    if (headers.indexOf(CONFIG_COL) === -1) {
      
      // Buscamos la columna de timestamp para insertar la nueva columna justo antes
      let targetIndex = lastCol + 1;
      const timeColIndex = headers.findIndex(h => h.toLowerCase().includes('_at'));
      
      if (timeColIndex !== -1) {
        targetIndex = timeColIndex + 1;
        sheet.insertColumnBefore(targetIndex);
      } else {
        // Si no hay timestamp, al final
        targetIndex = lastCol + 1;
      }

      // Seteamos el encabezado
      const headerCell = sheet.getRange(1, targetIndex);
      headerCell.setValue(CONFIG_COL);

      // --- ESTÉTICA PROFESIONAL ---
      // Copiamos el estilo de la primera celda de la cabecera (Color #556B2F, etc)
      const styleTemplate = sheet.getRange(1, 1);
      headerCell.setBackground(styleTemplate.getBackground())
                .setFontColor(styleTemplate.getFontColor())
                .setFontWeight(styleTemplate.getFontWeight())
                .setHorizontalAlignment("center");
      
      console.log(`✅ Hoja ${esquema.nombre}: Columna '${CONFIG_COL}' creada en posición ${targetIndex}.`);
      
      // --- REGLA DE NEGOCIO: VALOR POR DEFECTO ---
      // Para no romper la app, asignamos "PROPIO" a las filas existentes
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const range = sheet.getRange(2, targetIndex, lastRow - 1, 1);
        range.setValue("PROPIO");
        console.log(`   - Se asignó valor 'PROPIO' a ${lastRow - 1} filas existentes.`);
      }

    } else {
      console.log(`ℹ️ Hoja ${esquema.nombre}: Ya cuenta con el discriminador.`);
    }
  });

  // 2. Actualización de Parámetros Globales en CONF_GENERAL
  const confGeneral = ss.getSheetByName("CONF_GENERAL");
  if (confGeneral) {
    const data = confGeneral.getDataRange().getValues();
    const parametroKey = "TIPOS_PROYECTO";
    const existe = data.some(row => row[0] === parametroKey);

    if (!existe) {
      // Usamos appendRow para respetar la integridad de la tabla de parámetros
      confGeneral.appendRow([
        parametroKey, 
        "PROPIO, AJENO", 
        "Define los ciclos independientes de etapas y tareas", 
        new Date()
      ]);
      console.log(`✅ CONF_GENERAL: Parámetro '${parametroKey}' inicializado.`);
    }
  }

  // 3. Limpieza de Cache para forzar la lectura de la nueva estructura
  CacheService.getScriptCache().removeAll(["indices_db", "config_app"]);
  
  console.log("🎯 Proceso finalizado. El sistema ahora soporta ramificación por tipo de proyecto.");
  
}

function setupDatabaseC() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const TIPO_COL_NAME = "tipo_proyecto";
  
  console.log("🚀 Iniciando consolidación de infraestructura maestra...");

  // --- 1. CREACIÓN DE HOJA MAESTRA DE TIPOS ---
  let sheetTipo = ss.getSheetByName("CONF_TIPO_PROYECTO");
  if (!sheetTipo) {
    sheetTipo = ss.insertSheet("CONF_TIPO_PROYECTO");
    const headers = ["id", "nombre_tipo", "descripcion", "color_representativo", "created_at"];
    
    // Formateo de cabecera profesional
    sheetTipo.getRange(1, 1, 1, headers.length)
             .setValues([headers])
             .setBackground("#556B2F")
             .setFontColor("#FFFFFF")
             .setFontWeight("bold")
             .setHorizontalAlignment("center");
    
    // Datos iniciales para habilitar el CRUD
    const initialData = [
      [Utilities.getUuid(), "PROPIO", "Ciclo estándar para proyectos internos", "#556B2F", new Date()],
      [Utilities.getUuid(), "AJENO", "Ciclo simplificado para servicios externos", "#176282", new Date()]
    ];
    sheetTipo.getRange(2, 1, initialData.length, initialData[0].length).setValues(initialData);
    
    // Ajuste de columnas
    sheetTipo.setFrozenRows(1);
    sheetTipo.autoResizeColumns(1, headers.length);
    console.log("✅ Hoja CONF_TIPO_PROYECTO creada con registros base.");
  }

  // --- 2. ASEGURAR COLUMNAS RELACIONALES ---
  // Listado de hojas que deben apuntar a un tipo de proyecto
  const esquemas = [
    { nombre: "CONF_ETAPAS", desc: "Ciclos de vida" },
    { nombre: "CONF_CHECKLISTS", desc: "Protocolos técnicos" },
    { nombre: "DB_PROYECTOS", desc: "Maestro de obras" }
  ];

  esquemas.forEach(esquema => {
    const sheet = ss.getSheetByName(esquema.nombre);
    if (!sheet) {
      console.warn(`⚠️ Hoja ${esquema.nombre} no encontrada.`);
      return;
    }

    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    if (headers.indexOf(TIPO_COL_NAME) === -1) {
      // Posicionamiento inteligente: antes de las columnas de fecha '_at'
      let targetIndex = lastCol + 1;
      const timeColIndex = headers.findIndex(h => String(h).toLowerCase().includes('_at'));
      
      if (timeColIndex !== -1) {
        targetIndex = timeColIndex + 1;
        sheet.insertColumnBefore(targetIndex);
      }
      
      const headerCell = sheet.getRange(1, targetIndex);
      headerCell.setValue(TIPO_COL_NAME);
      
      // Aplicar estilo de la hoja (tomando como referencia la primera celda)
      const styleTemplate = sheet.getRange(1, 1);
      headerCell.setBackground(styleTemplate.getBackground())
                .setFontColor(styleTemplate.getFontColor())
                .setFontWeight(styleTemplate.getFontWeight())
                .setHorizontalAlignment("center");

      // Migración de datos existentes a 'PROPIO'
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, targetIndex, lastRow - 1, 1).setValue("PROPIO");
      }
      console.log(`✅ Columna '${TIPO_COL_NAME}' añadida a ${esquema.nombre}.`);
    }
  });

  // --- 3. ACTUALIZAR PARÁMETROS EN CONF_GENERAL ---
  const confGeneral = ss.getSheetByName("CONF_GENERAL");
  if (confGeneral) {
    const data = confGeneral.getDataRange().getValues();
    if (!data.some(row => row[0] === "TIPOS_PROYECTO")) {
      confGeneral.appendRow([
        "TIPOS_PROYECTO", 
        "PROPIO, AJENO", 
        "Define las familias de ciclos de trabajo", 
        new Date()
      ]);
    }
  }

  // --- 4. FEEDBACK FINAL SEGURO ---
  console.log("🎯 Infraestructura consolidada con éxito.");
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Base de Datos Lista", 
             "Se ha creado la hoja CONF_TIPO_PROYECTO y se han vinculado las tablas relacionales.", 
             ui.ButtonSet.OK);
  } catch (e) {
    // Si se corre desde el editor sin UI activa, solo logeamos
    console.log("Aviso: Ejecución terminada (UI no disponible).");
  }
}

function setupDatabaseD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const TIPO_ID_COL = "id_tipo_proyecto"; // Nombre técnico de la FK
  
  console.log("🚀 Iniciando consolidación de infraestructura relacional...");

  // 1. CREACIÓN Y/O VALIDACIÓN DE HOJA MAESTRA CONF_TIPO_PROYECTO
  let sheetMaestra = ss.getSheetByName("CONF_TIPO_PROYECTO");
  let defaultTypeId = "";

  if (!sheetMaestra) {
    sheetMaestra = ss.insertSheet("CONF_TIPO_PROYECTO");
    const headers = ["id", "nombre_tipo", "descripcion", "color_representativo", "created_at"];
    
    // Estilo profesional
    sheetMaestra.getRange(1, 1, 1, headers.length)
                .setValues([headers])
                .setBackground("#556B2F")
                .setFontColor("#FFFFFF")
                .setFontWeight("bold")
                .setHorizontalAlignment("center");

    // Registro inicial por defecto
    defaultTypeId = Utilities.getUuid();
    const initialData = [
      [defaultTypeId, "PROPIO", "Proyectos de ejecución interna", "#556B2F", new Date()]
    ];
    sheetMaestra.getRange(2, 1, 1, headers.length).setValues(initialData);
    sheetMaestra.setFrozenRows(1);
    console.log(`✅ Hoja CONF_TIPO_PROYECTO creada. ID asignado a PROPIO: ${defaultTypeId}`);
  } else {
    // Si ya existe, obtenemos el ID del tipo "PROPIO" para la migración
    const data = sheetMaestra.getDataRange().getValues();
    const rowPropio = data.find(r => r[1] === "PROPIO");
    defaultTypeId = rowPropio ? rowPropio[0] : Utilities.getUuid();
  }

  // 2. ACTUALIZACIÓN DE TABLAS RELACIONADAS (FK Implementation)
  const esquemas = [
    { nombre: "CONF_ETAPAS", desc: "Ciclos de vida" },
    { nombre: "CONF_CHECKLISTS", desc: "Protocolos técnicos" },
    { nombre: "DB_PROYECTOS", desc: "Maestro de obras" }
  ];

  esquemas.forEach(esquema => {
    const sheet = ss.getSheetByName(esquema.nombre);
    if (!sheet) return;

    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // Si existe la columna vieja "tipo_proyecto", la renombramos a "id_tipo_proyecto"
    // Si no existe ninguna, la creamos.
    let targetCol;
    const oldIndex = headers.indexOf("tipo_proyecto");
    const newIndex = headers.indexOf(TIPO_ID_COL);

    if (newIndex === -1) {
      if (oldIndex !== -1) {
        // Renombrar
        targetCol = oldIndex + 1;
        sheet.getRange(1, targetCol).setValue(TIPO_ID_COL);
        console.log(`📝 Columna renombrada en ${esquema.nombre}`);
      } else {
        // Crear antes de marcas de tiempo
        const timeColIndex = headers.findIndex(h => String(h).toLowerCase().includes('_at'));
        targetCol = (timeColIndex !== -1) ? timeColIndex + 1 : lastCol + 1;
        sheet.insertColumnBefore(targetCol);
        sheet.getRange(1, targetCol).setValue(TIPO_ID_COL);
        
        // Aplicar estilo de cabecera
        const styleTemplate = sheet.getRange(1, 1);
        sheet.getRange(1, targetCol).setBackground(styleTemplate.getBackground())
                                    .setFontColor(styleTemplate.getFontColor())
                                    .setFontWeight(styleTemplate.getFontWeight())
                                    .setHorizontalAlignment("center");
      }

      // 3. MIGRACIÓN DE DATOS (DATA INTEGRITY)
      // Asignamos el UUID del tipo 'PROPIO' a las filas existentes
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const range = sheet.getRange(2, targetCol, lastRow - 1, 1);
        range.setValue(defaultTypeId);
        console.log(`🔗 ${esquema.nombre}: Vinculadas ${lastRow - 1} filas al ID maestro.`);
      }
    }
  });

  // 4. LIMPIEZA Y FEEDBACK
  CacheService.getScriptCache().removeAll(["indices_db", "config_app"]);
  
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Sincronización Relacional Exitosa", 
             "Se ha establecido la hoja CONF_TIPO_PROYECTO como maestro y se han actualizado las llaves foráneas (id_tipo_proyecto).", 
             ui.ButtonSet.OK);
  } catch (e) {
    console.log("Infraestructura lista (Sin UI activa).");
  }
}

/**
 * Función de utilidad para generar IDs únicos (UUIDv4 simplificado)
 */
function generateUUID() {
  return Utilities.getUuid();
}

