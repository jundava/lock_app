function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Lock - Gestión de Proyectos de Obras')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
/**
 * DriveManager: Gestiona la creación de carpetas por proyecto
 */
const DriveManager = {
  
  /**
   * Obtiene o crea la subcarpeta de un proyecto
   * @param {string} proyectoId El UUID o nombre del proyecto
   * @return {string} ID de la carpeta de Google Drive
   */
  getOrCreateProjectFolder: function(proyectoId) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const confSheet = ss.getSheetByName("CONF_GENERAL");
      const data = confSheet.getDataRange().getValues();
      
      // 1. Obtener la URL raíz de la configuración
      const rootUrlRow = data.find(row => row[0] === "DRIVE_ROOT_FOLDER_URL");
      if (!rootUrlRow || !rootUrlRow[1]) {
        throw new Error("La URL raíz de Google Drive no está configurada.");
      }
      
      const rootFolderId = this.extractIdFromUrl(rootUrlRow[1]);
      const rootFolder = DriveApp.getFolderById(rootFolderId);
      
      // 2. Buscar si ya existe la subcarpeta (por nombre o ID)
      const folderIterator = rootFolder.getFoldersByName(proyectoId);
      
      if (folderIterator.hasNext()) {
        return folderIterator.next().getId();
      } else {
        // 3. Crear la carpeta si no existe
        const newFolder = rootFolder.createFolder(proyectoId);
        return newFolder.getId();
      }
      
    } catch (e) {
      console.error("Error en DriveManager: " + e.toString());
      throw e;
    } finally {
      lock.releaseLock();
    }
  },

  /**
   * Helper para extraer ID de una URL de carpeta de Drive
   */
  extractIdFromUrl: function(url) {
    const match = url.match(/[-\w]{25,}/);
    return match ? match[0] : url;
  }
}

/**
 * CRUD SERVICE - Módulo de Configuración
 * Especialista: GAS Expert
 */

const CONFIG_SHEETS = {
  ETAPAS: "CONF_ETAPAS",
  TAREAS: "CONF_TAREAS",
  PROFESIONALES: "CONF_PROFESIONALES",
  CHECKLISTS: "CONF_CHECKLISTS",
  GENERAL: "CONF_GENERAL",
  TIPOS: "CONF_TIPO_PROYECTO"
};


/**
 * Borrado físico de configuración
 */
function deleteConfigRecord(tableName, id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tableName);
    const data = sheet.getDataRange().getValues();
    const idColIndex = data[0].indexOf("id");
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] === id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Registro eliminado" };
      }
    }
    return { success: false, error: "ID no encontrado" };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Específico para guardar la URL de Drive en CONF_GENERAL
 */
function updateDriveUrl(url) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEETS.GENERAL);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === "DRIVE_ROOT_FOLDER_URL") {
      sheet.getRange(i + 1, 2).setValue(url);
      sheet.getRange(i + 1, 4).setValue(new Date());
      return { success: true };
    }
  }
  return { success: false, error: "Parámetro no encontrado" };
}

/**
 * Obtiene las tareas vinculadas a una etapa específica
 * Usa Map para optimizar si fuera necesario en el futuro
 */
function getTareasByEtapa(etapaId) {
  const todasLasTareas = readConfig(CONFIG_SHEETS.TAREAS);
  return todasLasTareas.filter(t => t.etapa_id === etapaId);
}

/**
 * Lectura robusta: Normaliza encabezados, tipos de datos y SANITIZA FECHAS.
 */
function readConfig(tableName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tableName);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues(); 
  if (data.length <= 1) return [];

  const originalHeaders = data.shift();
  // Normalizamos headers a minúsculas y guiones bajos
  const headers = originalHeaders.map(h => h.toString().trim().toLowerCase().replace(/\s+/g, '_'));
  
  return data.map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      if(header) {
        let value = row[i];
        
        // 1. Normalización de Booleanos
        if (typeof value === 'string') {
          if (value.toUpperCase() === 'TRUE') value = true;
          if (value.toUpperCase() === 'FALSE') value = false;
        }
        
        // 2. SANITIZACIÓN DE FECHAS (CRÍTICO PARA QUE SE VEAN LOS DATOS)
        // Convertimos objetos Date a String ISO para que viajen seguros al HTML
        if (value instanceof Date) {
           value = value.toISOString(); 
        }
        
        obj[header] = value;
      }
    });
    return obj;
  });
}

/**
 * Escritura robusta: Busca la columna correcta ignorando mayúsculas
 */
function saveConfigRecord(tableName, item) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(tableName);
    
    // Leer encabezados y normalizarlos para buscar correspondencias
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) throw new Error("La tabla " + tableName + " está vacía.");
    
    const headers = data[0].map(h => h.toString().trim().toLowerCase());
    
    // Generar nuevo ID si no existe
    if (!item.id) {
      item.id = Utilities.getUuid();
      item.created_at = new Date(); // Se guardará como fecha objeto en GAS
      
      // Mapeamos los datos del item a las columnas del Sheet
      const newRow = headers.map(h => {
        // Buscamos la clave en el objeto item que coincida con el header (case insensitive)
        const itemKey = Object.keys(item).find(k => k.toLowerCase() === h);
        return itemKey ? item[itemKey] : "";
      });
      
      sheet.appendRow(newRow);
      return { success: true, message: "Registro creado", data: item };
    } 
    
    // Lógica de Actualización (Update)
    const idColIndex = headers.indexOf("id");
    if (idColIndex === -1) throw new Error("No se encuentra la columna 'id'");
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] === item.id) {
        const rowRange = sheet.getRange(i + 1, 1, 1, headers.length);
        
        // Preparamos la fila actualizada
        const updatedRow = headers.map((h, colIdx) => {
           const itemKey = Object.keys(item).find(k => k.toLowerCase() === h);
           // Si el item tiene el dato, lo usamos. Si no, mantenemos el valor actual del sheet
           return itemKey !== undefined ? item[itemKey] : data[i][colIdx];
        });
        
        rowRange.setValues([updatedRow]);
        return { success: true, message: "Registro actualizado" };
      }
    }
    
    throw new Error("No se encontró el ID para actualizar.");
    
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}


/**
 * ELIMINACIÓN EN CASCADA DE PROYECTO
 * Elimina: Drive Folder + Tareas de Ejecución + Registro de Proyecto
 */
function deleteProjectFull(projectId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. BORRAR CARPETA DE DRIVE
    const sheetProyectos = ss.getSheetByName("DB_PROYECTOS");
    const dataProyectos = sheetProyectos.getDataRange().getValues();
    const idIndex = dataProyectos[0].indexOf("id");
    const driveIndex = dataProyectos[0].indexOf("drive_folder_id");
    
    let driveId = null;
    let rowIndex = -1;

    // Buscar el proyecto y su ID de Drive
    for (let i = 1; i < dataProyectos.length; i++) {
      if (dataProyectos[i][idIndex] === projectId) {
        driveId = dataProyectos[i][driveIndex];
        rowIndex = i + 1; // Fila real en la hoja (1-based)
        break;
      }
    }

    if (driveId && driveId !== "NO_CONFIGURADO") {
      try {
        DriveApp.getFolderById(driveId).setTrashed(true); // Enviamos a la papelera
      } catch (e) {
        console.warn("No se pudo borrar la carpeta Drive (quizás ya no existe): " + e.message);
      }
    }

    // 2. BORRAR TAREAS DE EJECUCIÓN (DB_EJECUCION)
    const sheetEjecucion = ss.getSheetByName("DB_EJECUCION");
    const dataEjecucion = sheetEjecucion.getDataRange().getValues();
    const projIdIndexEjec = dataEjecucion[0].indexOf("proyecto_id");
    
    // Recorremos hacia atrás para borrar sin afectar los índices
    for (let i = dataEjecucion.length - 1; i >= 1; i--) {
      if (dataEjecucion[i][projIdIndexEjec] === projectId) {
        sheetEjecucion.deleteRow(i + 1);
      }
    }

        // 3. BORRAR RELACION RESPONSABLES (CONF_REL_ASIGNACIONES)
    const sheetAsign = ss.getSheetByName("CONF_REL_ASIGNACIONES");
    const dataAsign = sheetAsign.getDataRange().getValues();
    const projIdIndexAsign = dataAsign[0].indexOf("id_proyecto");
    
    // Recorremos hacia atrás para borrar sin afectar los índices
    for (let i = dataAsign.length - 1; i >= 1; i--) {
      if (dataAsign[i][projIdIndexAsign] === projectId) {
        sheetAsign.deleteRow(i + 1);
      }
    }

    // 4. BORRAR REGISTRO DE PROYECTO (DB_PROYECTOS)
    if (rowIndex !== -1) {
      sheetProyectos.deleteRow(rowIndex);
    } else {
      throw new Error("Proyecto no encontrado");
    }

    return { success: true };

  } catch (e) {
    console.error(e);
    throw new Error("Error al eliminar proyecto: " + e.toString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * Obtiene la configuración detallada (Árbol: Tipo -> Etapas -> Tareas)
 * Optimizado para renderizado directo en Vue.js (Preview y Configuración).
 * * @param {string} idTipo - El UUID del tipo de proyecto (CONF_TIPO_PROYECTO)
 * @returns {Object} Objeto con metadatos del tipo, resumen y estructura jerárquica.
 */
function getTipoDetailedConfig(idTipo) {
  // No necesitamos LockService aquí porque es una operación de LECTURA pura.
  
  try {
    // 1. Carga de datos en paralelo (conceptualmente)
    // Asumimos que readConfig devuelve arrays de objetos limpios
    const tipos = readConfig(CONFIG_SHEETS.TIPOS || "CONF_TIPO_PROYECTO");
    const etapasRaw = readConfig(CONFIG_SHEETS.ETAPAS || "CONF_ETAPAS");
    const tareasRaw = readConfig(CONFIG_SHEETS.TAREAS || "CONF_TAREAS");

    // 2. Validación de Existencia
    const tipoInfo = tipos.find(t => t.id === idTipo);
    if (!tipoInfo) {
      throw new Error(`Tipo de proyecto con ID ${idTipo} no encontrado.`);
    }

    // 3. Filtrado y Ordenamiento de Etapas (O(n))
    // Convertimos orden a Number para asegurar sort correcto (1, 2, 10 en vez de 1, 10, 2)
    const misEtapas = etapasRaw
      .filter(e => e.id_tipo_proyecto === idTipo)
      .sort((a, b) => (Number(a.orden) || 999) - (Number(b.orden) || 999));

    // Creamos un Set de IDs para búsqueda O(1) en el paso siguiente
    const idsEtapasSet = new Set(misEtapas.map(e => e.id));

    // 4. Filtrado de Tareas (O(m))
    const misTareas = tareasRaw.filter(t => idsEtapasSet.has(t.etapa_id));

    // 5. Construcción del Árbol Jerárquico (Nesting)
    // Esto facilita enormemente la vida al Frontend
    const estructuraCicloVida = misEtapas.map(etapa => {
      // Filtramos las tareas de esta etapa específica
      const tareasDeEtapa = misTareas.filter(t => t.etapa_id === etapa.id);
      
      return {
        ...etapa, // Heredamos id, nombre, color, orden
        tareas: tareasDeEtapa, // Array anidado
        stats: {
          cantidad_tareas: tareasDeEtapa.length,
          con_evidencia: tareasDeEtapa.filter(t => String(t.requiere_evidencia).toLowerCase() === 'true').length
        }
      };
    });

    // 6. Retorno de Payload Completo
    return {
      success: true,
      data: {
        info: {
          id: tipoInfo.id,
          nombre: tipoInfo.nombre_tipo,
          descripcion: tipoInfo.descripcion,
          color: tipoInfo.color_representativo
        },
        resumen: {
          total_etapas: misEtapas.length,
          total_tareas: misTareas.length,
          estimacion_complejidad: misTareas.length > 20 ? 'Alta' : 'Baja'
        },
        ciclo_vida: estructuraCicloVida // Array listo para v-for en Vue
      }
    };

  } catch (e) {
    console.error("Error en getTipoDetailedConfig:", e);
    return { 
      success: false, 
      error: e.message 
    };
  }
}

// ==========================================
// MÓDULO DE EJECUCIÓN (COCKPIT)
// ==========================================

/**
 * 1. Obtiene las tareas operativas de un proyecto específico
 * BLINDADA: Convierte Fechas a ISOString para evitar error de servidor.
 */
function getProjectExecutionData(projectId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DB_EJECUCION");
  
  if (!sheet) return []; 

  // Usamos getValues para respetar tipos (Booleanos), pero debemos cuidar las Fechas
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return []; // Solo encabezados o vacía

  const originalHeaders = data.shift();
  // Normalizamos encabezados para que coincidan con lo que espera Vue
  const headers = originalHeaders.map(h => h.toString().trim().toLowerCase().replace(/\s+/g, '_'));

  // Convertimos la matriz en objetos JSON seguros
  const allTasks = data.map(row => {
    let task = {};
    headers.forEach((header, index) => {
      let value = row[index];

      // --- SANITIZACIÓN CRÍTICA (Evita el error de servidor) ---
      if (value instanceof Date) {
        // Si hay una fecha válida, la pasamos a texto. Si es inválida, cadena vacía.
        value = !isNaN(value) ? value.toISOString() : ""; 
      }
      // ---------------------------------------------------------

      // Normalización de Booleanos (Legacy Data)
      if (typeof value === 'string') {
          if (value.toUpperCase() === 'TRUE') value = true;
          if (value.toUpperCase() === 'FALSE') value = false;
      }

      task[header] = value;
    });
    return task;
  });

  // Filtramos por el ID del proyecto solicitado
  return allTasks.filter(t => t.proyecto_id === projectId);
}

/**
 * Sube una evidencia al Drive y actualiza la tarea con la URL.
 * Valida tipos: PDF, DOCX, XLSX, PPTX, TXT, Imágenes.
 */
function uploadTaskEvidence(taskId, fileData, fileName, mimeType) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Validaciones
    const allowedMimes = [
      'application/pdf', 
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // docx
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // xlsx
      'application/vnd.openxmlformats-officedocument.presentationml.presentation', // pptx
      'text/plain',
      'image/jpeg', 'image/png'
    ];
    
    if (!allowedMimes.includes(mimeType)) {
      throw new Error("Formato no permitido. Solo PDF, Word, Excel, PPT, TXT o Imágenes.");
    }

    // 2. Obtener Proyecto y Carpeta
    const taskInfo = getTaskInfo(taskId); 
    if (!taskInfo || !taskInfo.drive_folder_id) throw new Error("No se encontró la carpeta del proyecto.");

    const folder = DriveApp.getFolderById(taskInfo.drive_folder_id);
    
    // 3. Crear Blob y Archivo
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName);
    const file = folder.createFile(blob);
    const fileUrl = file.getUrl();

    // 4. Actualizar DB_EJECUCION (Llamamos a la función global)
    // Liberamos el lock actual momentáneamente o confiamos en que updateExecutionTask maneje su propio lock rápido
    // Nota: Como updateExecutionTask tiene su propio lock, es seguro llamarla.
    
    updateExecutionTask({
      id: taskId,
      datos_evidencia: fileUrl
    });

    return { success: true, url: fileUrl };

  } catch (e) {
    console.error("Error upload:", e);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/**
 * updateExecutionTask (VERSIÓN GLOBAL CORREGIDA)
 * Esta función DEBE estar fuera de cualquier otra para ser accesible desde el HTML.
 */
function updateExecutionTask(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("DB_EJECUCION");
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const map = {};
    headers.forEach((h, i) => map[h.toString().toLowerCase().trim()] = i + 1);

    if (!map['id']) throw new Error("Columna ID no encontrada en DB.");

    // Búsqueda eficiente
    const ids = sheet.getRange(2, map['id'], sheet.getLastRow() - 1, 1).getValues().flat();
    const rowIndex = ids.indexOf(data.id);

    if (rowIndex === -1) return { success: false, message: "Tarea no existe." };
    const rowNumber = rowIndex + 2;

    // --- ESCRITURA SELECTIVA ---

    // 1. Estado y Comentarios
    if (data.estado !== undefined && map['estado']) 
      sheet.getRange(rowNumber, map['estado']).setValue(data.estado);
    
    if (data.comentarios !== undefined && map['comentarios']) 
      sheet.getRange(rowNumber, map['comentarios']).setValue(data.comentarios);

    // 2. Checklist (Nueva Columna)
    if (data.datos_checklist !== undefined && map['datos_checklist']) {
      sheet.getRange(rowNumber, map['datos_checklist']).setValue(data.datos_checklist);
    }

    // 3. Evidencia (URL)
    if (data.datos_evidencia !== undefined && map['datos_evidencia']) {
      sheet.getRange(rowNumber, map['datos_evidencia']).setValue(data.datos_evidencia);
    }

    if (map['updated_at']) 
      sheet.getRange(rowNumber, map['updated_at']).setValue(new Date());

    return { success: true };

  } catch (e) {
    console.error("Error update:", e);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// Helper para buscar ID de carpeta rápido
function getTaskInfo(taskId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Buscar tarea
  const sheetEjec = ss.getSheetByName("DB_EJECUCION");
  const dataEjec = sheetEjec.getDataRange().getValues();
  // Asumimos orden dinámico, buscamos índices
  const hEjec = dataEjec[0];
  const idxId = hEjec.indexOf("id");
  const idxProj = hEjec.indexOf("proyecto_id");
  
  const rowTask = dataEjec.find(r => r[idxId] === taskId);
  if (!rowTask) return null;
  
  const projectId = rowTask[idxProj];
  
  // Buscar Proyecto
  const sheetProj = ss.getSheetByName("DB_PROYECTOS");
  const dataProj = sheetProj.getDataRange().getValues();
  const hProj = dataProj[0];
  const idxIdP = hProj.indexOf("id");
  const idxFolder = hProj.indexOf("drive_folder_id");
  
  const rowProj = dataProj.find(r => r[idxIdP] === projectId);
  
  return rowProj ? { drive_folder_id: rowProj[idxFolder] } : null;
}

/**
 * MOTOR DE TAREAS (Helper)
 * Borra tareas viejas de un proyecto y crea nuevas basadas en el Tipo.
 * NO USA LOCK INTERNO (para ser llamada por funciones que ya tienen lock)
 */
function regenerateProjectTasks(projectId, tipoId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEjecucion = ss.getSheetByName("DB_EJECUCION");
  const sheetEtapas = ss.getSheetByName("CONF_ETAPAS");
  const sheetTareas = ss.getSheetByName("CONF_TAREAS");

  // A. Obtener IDs de Etapas para este Tipo de Proyecto
  // CONF_ETAPAS: id(0)... id_tipo_proyecto(columna variable, buscar índice)
  const etapasData = sheetEtapas.getDataRange().getValues();
  const etapasHeaders = etapasData.shift();
  const idxTipoEnEtapa = etapasHeaders.indexOf("id_tipo_proyecto");
  const idxIdEtapa = 0;

  if (idxTipoEnEtapa === -1) throw new Error("No hay id_tipo_proyecto en CONF_ETAPAS");

  const etapasIds = etapasData
    .filter(r => String(r[idxTipoEnEtapa]) === String(tipoId))
    .map(r => r[idxIdEtapa]);

  if (etapasIds.length === 0) {
    console.warn("No hay etapas configuradas para el tipo: " + tipoId);
    return; // No hay nada que generar
  }

  // B. Obtener Tareas Plantilla asociadas a esas Etapas
  // CONF_TAREAS: id(0), etapa_id(1)...
  const tareasData = sheetTareas.getDataRange().getValues();
  const tareasHeaders = tareasData.shift(); // Quitamos header
  const idxEtapaEnTarea = tareasHeaders.indexOf("etapa_id"); 
  
  // Filtramos las tareas que pertenecen a las etapas encontradas
  const tareasTemplate = tareasData.filter(r => etapasIds.includes(r[idxEtapaEnTarea]));

  // C. Preparar nuevas filas para DB_EJECUCION
  // Estructura DB_EJECUCION: 
  // id, proyecto_id, etapa_id, nombre_tarea, requiere_evidencia, tipo_entrada, checklist_id, estado, responsable_id, datos_evidencia, comentarios, updated_at, datos_checklist
  
  // Mapeamos dinámicamente según los índices de CONF_TAREAS
  const tIdxName = tareasHeaders.indexOf("nombre_tarea");
  const tIdxEvidencia = tareasHeaders.indexOf("requiere_evidencia");
  const tIdxInput = tareasHeaders.indexOf("tipo_entrada");
  const tIdxChecklist = tareasHeaders.indexOf("checklist_id");

  const newRows = tareasTemplate.map(t => [
    Utilities.getUuid(), // id
    projectId,           // proyecto_id
    t[idxEtapaEnTarea],  // etapa_id
    t[tIdxName],         // nombre_tarea
    t[tIdxEvidencia],    // requiere_evidencia
    t[tIdxInput],        // tipo_entrada
    t[tIdxChecklist],    // checklist_id
    "Pendiente",         // estado inicial
    "",                  // responsable_id (vacío)
    "",                  // datos_evidencia
    "",                  // comentarios
    new Date(),          // updated_at
    ""                   // datos_checklist
  ]);

  // D. Transacción en DB_EJECUCION
  const dataEjec = sheetEjecucion.getDataRange().getValues();
  const headerEjec = dataEjec.shift(); // Guardar cabecera
  const idxProyEjec = headerEjec.indexOf("proyecto_id");

  // 1. Filtramos para borrar las viejas de este proyecto
  const dataCleaned = dataEjec.filter(r => String(r[idxProyEjec]) !== String(projectId));

  // 2. Combinamos (Viejas limpias + Nuevas generadas)
  const finalEjecucion = [...dataCleaned, ...newRows];

  // 3. Escribir
  sheetEjecucion.clearContents();
  sheetEjecucion.appendRow(headerEjec); // Poner cabecera
  if (finalEjecucion.length > 0) {
    sheetEjecucion.getRange(2, 1, finalEjecucion.length, finalEjecucion[0].length).setValues(finalEjecucion);
  }
}

/**
 * Calcula Estadísticas Globales.
 * Versión segura: Evita errores con nulos en datos_evidencia.
 */
function getGlobalProgressStats() {
  const tasks = readConfig("DB_EJECUCION");
  const etapas = readConfig("CONF_ETAPAS").sort((a, b) => (Number(a.orden) || 0) - (Number(b.orden) || 0));
  
  const map = {};
  
  tasks.forEach(t => {
    if (!t.proyecto_id) return;
    
    if (!map[t.proyecto_id]) {
      map[t.proyecto_id] = { 
        totalTasks: 0, completedTasks: 0,
        totalEvidence: 0, completedEvidence: 0,
        stages: {} 
      };
    }
    
    // 1. Tareas
    map[t.proyecto_id].totalTasks++;
    if (t.estado === 'Completado') map[t.proyecto_id].completedTasks++;

    // 2. Evidencias (Manejo seguro de tipos)
    const requiere = String(t.requiere_evidencia).toLowerCase() === 'true';
    if (requiere) {
      map[t.proyecto_id].totalEvidence++;
      // Validación segura: Existe y tiene longitud > 5
      if (t.datos_evidencia && t.datos_evidencia.toString().length > 5) {
        map[t.proyecto_id].completedEvidence++;
      }
    }

    // 3. Etapas
    if (!map[t.proyecto_id].stages[t.etapa_id]) {
      map[t.proyecto_id].stages[t.etapa_id] = { total: 0, completed: 0 };
    }
    map[t.proyecto_id].stages[t.etapa_id].total++;
    if (t.estado === 'Completado') {
      map[t.proyecto_id].stages[t.etapa_id].completed++;
    }
  });

  const result = {};
  Object.keys(map).forEach(pid => {
    const data = map[pid];
    const percent = data.totalTasks > 0 ? Math.round((data.completedTasks / data.totalTasks) * 100) : 0;

    let currentStageName = "Planificación";
    let currentStageColor = "#999";
    let found = false;

    for (let i = 0; i < etapas.length; i++) {
      const etapa = etapas[i];
      const stageData = data.stages[etapa.id];
      if (stageData && stageData.completed < stageData.total) {
          currentStageName = etapa.nombre_etapa;
          currentStageColor = etapa.color_hex;
          found = true;
          break;
      }
    }

    if (!found && data.totalTasks > 0 && percent === 100) {
      currentStageName = "Finalizado";
      currentStageColor = "#198754";
    }

    result[pid] = {
      progress: percent,
      stageText: currentStageName,
      stageColor: currentStageColor,
      taskCount: { total: data.totalTasks, pending: data.totalTasks - data.completedTasks },
      evidenceCount: { total: data.totalEvidence, pending: data.totalEvidence - data.completedEvidence }
    };
  });

  return result;
}

/**
 * Acción Destructiva: Cambia el tipo y resetea tareas
 */
function updateProjectTypeAndReset(projectId, newTypeId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Actualizar el TIPO en DB_PROYECTOS
    const sheetProy = ss.getSheetByName("DB_PROYECTOS");
    const data = sheetProy.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf("id");
    const typeIdx = headers.indexOf("id_tipo_proyecto"); 
    
    if (idIdx === -1 || typeIdx === -1) throw new Error("Estructura DB_PROYECTOS inválida");

    let foundRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(projectId)) {
        foundRow = i + 1; // +1 por base 1
        break;
      }
    }

    if (foundRow === -1) throw new Error("Proyecto no encontrado");

    // Actualizamos solo la celda del tipo
    sheetProy.getRange(foundRow, typeIdx + 1).setValue(newTypeId);

    // 2. Regenerar Tareas (Borra viejas, crea nuevas)
    regenerateProjectTasks(projectId, newTypeId);

    return { success: true };

  } catch (e) {
    console.error(e);
    throw new Error("Error al resetear proyecto: " + e.message);
  } finally {
    lock.releaseLock();
  }
}
function addChecklistDataColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DB_EJECUCION");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Normalizar headers para buscar
  const cleanHeaders = headers.map(h => h.toString().toLowerCase().trim());
  
  if (!cleanHeaders.includes("datos_checklist")) {
    // Insertamos la columna después de checklist_id (o al final si prefieres)
    // Buscamos un buen lugar, por ejemplo antes de 'estado' o 'updated_at'
    const targetCol = headers.length + 1; 
    
    sheet.getRange(1, targetCol).setValue("datos_checklist")
         .setBackground("#556B2F")
         .setFontColor("white")
         .setFontWeight("bold");
         
    console.log("✅ Columna 'datos_checklist' creada en DB_EJECUCION");
  } else {
    console.log("ℹ️ La columna 'datos_checklist' ya existe.");
  }
}

/* ==========================================================================
   SEGURIDAD Y ASIGNACIONES (NUEVO)
   ========================================================================== */

/**
 * 1. Obtiene el perfil del usuario actual desde CONF_PROFESIONALES
 */
function getCurrentUserProfile() {
  const email = Session.getActiveUser().getEmail();
  // Leemos la tabla de profesionales
  const raw = readConfig("CONF_PROFESIONALES"); 
  
  // Buscamos al usuario por email (normalizando a minúsculas)
  const user = raw.find(p => p.email && p.email.trim().toLowerCase() === email.toLowerCase());
  
  if (!user) {
    // CAMBIO CRÍTICO: Retornamos null.
    // Esto le indica al Frontend que debe mostrar la pantalla de "Acceso Denegado".
    return null; 
  }

  return {
    id: user.id,
    nombre: user.nombre_completo,
    email: user.email,
    rol_sistema: user.perfil_sistema || 'Operador' // Si está vacío en Excel, es Operador
  };
}

/**
 * 2. Lectura Segura de Proyectos (Row-Level Security)
 * Reemplaza el uso directo de readConfig("DB_PROYECTOS") en el frontend
 */
function getSecureProjects() {
  const user = getCurrentUserProfile();
  const allProjects = readConfig("DB_PROYECTOS");

  // A. Si es Admin, ve todo
  if (user.rol_sistema === 'Admin') {
    return allProjects;
  }

  // B. Si es Operador, filtramos por asignación
  if (user.rol_sistema === 'Operador') {
    const asignaciones = readConfig("CONF_REL_ASIGNACIONES");
    
    // Obtenemos IDs de proyectos asignados a este usuario
    const misProyectosIds = asignaciones
      .filter(a => a.id_profesional === user.id)
      .map(a => a.id_proyecto);

    return allProjects.filter(p => misProyectosIds.includes(p.id));
  }

  // C. Si no es nada (Invitado), no ve nada
  return [];
}

/**
 * VERSIÓN COMPLETA Y CORREGIDA
 * Guarda proyecto + Drive + Asignaciones + Tareas
 */
function saveProjectWithAssignments(projectData, asignadosIds) {
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(30000);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // --- 1. DETECCIÓN DE ESTADO ---
    const isNew = !projectData.id; 
    if (isNew) {
      projectData.id = Utilities.getUuid();
      projectData.created_at = new Date();
    }

    // --- 2. CREACIÓN DE ESTRUCTURA EN DRIVE (SOLO SI ES NUEVO) ---
    if (isNew) {
      let driveUrl = "";
      let driveId = "";
      
      try {
        const sheetGeneral = ss.getSheetByName("CONF_GENERAL");
        const generalData = sheetGeneral.getDataRange().getValues();
        const rootUrlRow = generalData.find(r => r[0] === "DRIVE_ROOT_FOLDER_URL");
        
        if (rootUrlRow && rootUrlRow[1]) {
          // Extraer ID de la URL raíz
          const match = rootUrlRow[1].match(/[-\w]{25,}/);
          const rootFolderId = match ? match[0] : null;
          
          if (rootFolderId) {
            const rootFolder = DriveApp.getFolderById(rootFolderId);
            const folderName = `${projectData.codigo} - ${projectData.nombre_obra}`;
            
            // Crear carpeta principal del proyecto
            const projectFolder = rootFolder.createFolder(folderName);
            driveId = projectFolder.getId();
            driveUrl = projectFolder.getUrl();
            
            // Crear subcarpetas de etapas
            const etapasRaw = readConfig("CONF_ETAPAS");
            const etapasTipo = etapasRaw
              .filter(e => e.id_tipo_proyecto === projectData.id_tipo_proyecto)
              .sort((a, b) => (Number(a.orden) || 999) - (Number(b.orden) || 999));
            
            etapasTipo.forEach(e => {
              if (e.nombre_etapa) {
                projectFolder.createFolder(`${e.orden}. ${e.nombre_etapa}`);
              }
            });
          }
        }
      } catch (driveError) {
        console.warn("Advertencia: No se pudo crear estructura en Drive.", driveError);
        driveUrl = "ERROR_DRIVE";
        driveId = "NO_CREADO";
      }
      
      // Asignar URLs de Drive al objeto
      projectData.drive_folder_id = driveId;
      projectData.drive_url = driveUrl;
    }

    // --- 3. GUARDAR EN DB_PROYECTOS (SIN LOCK ANIDADO) ---
    const sheetProyectos = ss.getSheetByName("DB_PROYECTOS");
    if (!sheetProyectos) throw new Error("La hoja DB_PROYECTOS no existe");
    
    const dataProyectos = sheetProyectos.getDataRange().getValues();
    if (dataProyectos.length === 0) throw new Error("DB_PROYECTOS está vacía");
    
    const headersProyectos = dataProyectos[0].map(h => h.toString().trim().toLowerCase());
    
    if (isNew) {
      // CREAR NUEVO REGISTRO
      const newRow = headersProyectos.map(h => {
        const itemKey = Object.keys(projectData).find(k => k.toLowerCase() === h);
        return itemKey ? projectData[itemKey] : "";
      });
      sheetProyectos.appendRow(newRow);
      
    } else {
      // ACTUALIZAR REGISTRO EXISTENTE
      const idColIndex = headersProyectos.indexOf("id");
      if (idColIndex === -1) throw new Error("No se encuentra columna 'id'");
      
      let found = false;
      for (let i = 1; i < dataProyectos.length; i++) {
        if (dataProyectos[i][idColIndex] === projectData.id) {
          const rowRange = sheetProyectos.getRange(i + 1, 1, 1, headersProyectos.length);
          
          const updatedRow = headersProyectos.map((h, colIdx) => {
            const itemKey = Object.keys(projectData).find(k => k.toLowerCase() === h);
            return itemKey !== undefined ? projectData[itemKey] : dataProyectos[i][colIdx];
          });
          
          rowRange.setValues([updatedRow]);
          found = true;
          break;
        }
      }
      
      if (!found) throw new Error("No se encontró el proyecto para actualizar");
    }

    // --- 4. ACTUALIZAR ASIGNACIONES ---
    const sheetAsig = ss.getSheetByName("CONF_REL_ASIGNACIONES");
    if (!sheetAsig) throw new Error("La hoja CONF_REL_ASIGNACIONES no existe");
    
    const dataAsig = sheetAsig.getDataRange().getValues();
    
    let headersAsig = [];
    let bodyAsig = [];
    if (dataAsig.length > 0) {
      headersAsig = dataAsig.shift(); 
      bodyAsig = dataAsig;
    }

    const idxProj = 1; // Columna id_proyecto

    // Eliminar asignaciones viejas de este proyecto
    let finalData = bodyAsig.filter(row => row[idxProj] !== projectData.id);

    // Agregar nuevas asignaciones
    const now = new Date();
    if (asignadosIds && Array.isArray(asignadosIds)) {
      asignadosIds.forEach(idProf => {
        finalData.push([
          Utilities.getUuid(),
          projectData.id,
          idProf,
          now
        ]);
      });
    }

    // Reescribir tabla de asignaciones
    if (finalData.length > 0 && headersAsig.length > 0) {
      if (sheetAsig.getLastRow() > 1) {
         sheetAsig.getRange(2, 1, sheetAsig.getLastRow() - 1, headersAsig.length).clearContent();
      }
      sheetAsig.getRange(2, 1, finalData.length, headersAsig.length).setValues(finalData);
    }

    // --- 5. GENERACIÓN DE TAREAS SI ES NUEVO ---
    if (isNew && projectData.id_tipo_proyecto) {
      regenerateProjectTasksInternal(ss, projectData.id, projectData.id_tipo_proyecto);
    }
    
    SpreadsheetApp.flush(); // Forzar escritura
    
    return { success: true, projectId: projectData.id };

  } catch (e) {
    console.error("Error saveProjectWithAssignments:", e);
    throw new Error("Error al guardar proyecto: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Versión interna de regenerateProjectTasks (sin lock propio)
 */
function regenerateProjectTasksInternal(ss, projectId, tipoId) {
  const sheetEjecucion = ss.getSheetByName("DB_EJECUCION");
  const sheetEtapas = ss.getSheetByName("CONF_ETAPAS");
  const sheetTareas = ss.getSheetByName("CONF_TAREAS");

  const etapasData = sheetEtapas.getDataRange().getValues();
  const etapasHeaders = etapasData.shift();
  const idxTipoEnEtapa = etapasHeaders.indexOf("id_tipo_proyecto");
  const idxIdEtapa = 0;

  if (idxTipoEnEtapa === -1) throw new Error("No hay id_tipo_proyecto en CONF_ETAPAS");

  const etapasIds = etapasData
    .filter(r => String(r[idxTipoEnEtapa]) === String(tipoId))
    .map(r => r[idxIdEtapa]);

  if (etapasIds.length === 0) {
    console.warn("No hay etapas configuradas para el tipo: " + tipoId);
    return;
  }

  const tareasData = sheetTareas.getDataRange().getValues();
  const tareasHeaders = tareasData.shift();
  const idxEtapaEnTarea = tareasHeaders.indexOf("etapa_id"); 
  
  const tareasTemplate = tareasData.filter(r => etapasIds.includes(r[idxEtapaEnTarea]));

  const tIdxName = tareasHeaders.indexOf("nombre_tarea");
  const tIdxEvidencia = tareasHeaders.indexOf("requiere_evidencia");
  const tIdxInput = tareasHeaders.indexOf("tipo_entrada");
  const tIdxChecklist = tareasHeaders.indexOf("checklist_id");

  const newRows = tareasTemplate.map(t => [
    Utilities.getUuid(),
    projectId,
    t[idxEtapaEnTarea],
    t[tIdxName],
    t[tIdxEvidencia],
    t[tIdxInput],
    t[tIdxChecklist],
    "Pendiente",
    "",
    "",
    "",
    new Date(),
    ""
  ]);

  const dataEjec = sheetEjecucion.getDataRange().getValues();
  const headerEjec = dataEjec.shift();
  const idxProyEjec = headerEjec.indexOf("proyecto_id");

  const dataCleaned = dataEjec.filter(r => String(r[idxProyEjec]) !== String(projectId));
  const finalEjecucion = [...dataCleaned, ...newRows];

  sheetEjecucion.clearContents();
  sheetEjecucion.appendRow(headerEjec);
  if (finalEjecucion.length > 0) {
    sheetEjecucion.getRange(2, 1, finalEjecucion.length, finalEjecucion[0].length).setValues(finalEjecucion);
  }
}

/**
 * MAPA DE RELACIONES (Integridad Referencial)
 * Define qué tablas dependen de otras.
 * Clave: Tabla Padre (La que intentas borrar)
 * Valor: Array de objetos con la Tabla Hija y la Columna FK que apunta al padre.
 */
const SCHEMA_DEPENDENCIES = {
  "CONF_TIPO_PROYECTO": [
    { table: "CONF_ETAPAS", fk: "id_tipo_proyecto"},
    { table: "CONF_CHECKLISTS", fk: "id_tipo_proyecto"},
    { table: "DB_PROYECTOS", fk: "id_tipo_proyecto"}
  ],
  "CONF_ETAPAS": [
    { table: "CONF_TAREAS", fk: "etapa_id"},
    { table: "DB_EJECUCION", fk: "etapa_id"}
  ],
  "CONF_TAREAS": [
    { table: "DB_EJECUCION", fk: "nombre_tarea"}
  ],
  "CONF_CHECKLISTS": [
    { table: "CONF_TAREAS", fk: "checklist_id"},
    { table: "DB_EJECUCION", fk: "checklist_id"}
  ],
  "CONF_PROFESIONALES": [
    { table: "CONF_REL_ASIGNACIONES", fk: "id_profesional" },
    { table: "DB_EJECUCION", fk: "responsable_id"}
  ],
  "DB_PROYECTOS": [
    { table: "DB_EJECUCION", fk: "proyecto_id"},
    { table: "CONF_REL_ASIGNACIONES", fk: "id_proyecto"}
  ],
};

/**
 * Función Principal de Borrado Seguro
 * @param {string} sheetName - Nombre de la hoja (Tabla)
 * @param {string} id - UUID del registro a borrar
 */
function deleteConfigRecord(sheetName, id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(5000);

    // 1. VERIFICACIÓN DE DEPENDENCIAS
    const dependencyError = checkDependencies(ss, sheetName, id);
    if (dependencyError) {
      return { 
        success: false, 
        message: dependencyError 
      };
    }

    // 2. EJECUCIÓN DEL BORRADO
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`La hoja ${sheetName} no existe.`);

    const data = sheet.getDataRange().getValues();
    // Asumimos que la columna ID siempre es la primera (índice 0). 
    // Si no, habría que buscar el índice de la columna "id".
    const rowIndex = data.findIndex(row => row[0] == id);

    if (rowIndex === -1) {
      return { success: false, message: "Registro no encontrado." };
    }

    // rowIndex es base 0, deleteRow es base 1
    sheet.deleteRow(rowIndex + 1);
    
    return { success: true, message: "Registro eliminado correctamente." };

  } catch (e) {
    console.error(e);
    return { success: false, message: "Error del sistema: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper para revisar si el ID existe en tablas hijas
 */
function checkDependencies(ss, parentTable, id) {
  const dependencies = SCHEMA_DEPENDENCIES[parentTable];
  
  if (!dependencies) return null; // No tiene dependencias configuradas

  for (const dep of dependencies) {
    const sheet = ss.getSheetByName(dep.table);
    if (!sheet) continue; // Si la hoja no existe, saltamos (o logueamos error)

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) continue; // Solo encabezados

    const headers = data[0];
    const colIndex = headers.indexOf(dep.fk);

    if (colIndex === -1) {
      console.warn(`Columna FK '${dep.fk}' no encontrada en '${dep.table}'`);
      continue;
    }

    // Buscamos si el ID existe en la columna FK
    // Empezamos en i=1 para saltar el header
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colIndex]) === String(id)) {
        return `No se puede eliminar: Este registro está siendo usado en la tabla '${dep.table}'.`;
      }
    }
  }

  return null; // Todo limpio
}