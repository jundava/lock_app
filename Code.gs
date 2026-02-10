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

/* AGREGAR ESTO A TU Code.gs EXISTENTE */

function getUserInfo() {
  return {
    email: Session.getActiveUser().getEmail(),
    // En el futuro podemos buscar nombre y rol en la tabla de Profesionales
    role: 'Admin' 
  };
}

/**
 * CreateProjectFull - Versión Optimizada
 * Crea proyecto, estructura en Drive y clona tareas según el Tipo de Proyecto.
 */
function createProjectFull(projectData) {
  const lock = LockService.getScriptLock();
  
  try {
    // Esperamos el bloqueo para evitar duplicados de IDs o carpetas
    lock.waitLock(30000);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetProyectos = ss.getSheetByName("DB_PROYECTOS");
    const sheetEjecucion = ss.getSheetByName("DB_EJECUCION");
    const sheetGeneral = ss.getSheetByName("CONF_GENERAL");

    // --- 1. VALIDACIONES PREVIAS ---
    if (!projectData.id_tipo_proyecto) throw new Error("Falta el ID del Tipo de Proyecto.");
    
    // Generación de ID y Timestamps
    if (!projectData.id) projectData.id = Utilities.getUuid();
    const timestamp = new Date();
    projectData.created_at = timestamp;

    // --- 2. LECTURA DE CONFIGURACIÓN (JOIN EN MEMORIA) ---
    // Leemos etapas y filtramos por el tipo seleccionado
    const etapasRaw = readConfig("CONF_ETAPAS"); // Asumimos que readConfig devuelve objetos limpios
    const etapasTipo = etapasRaw
                        .filter(e => e.id_tipo_proyecto === projectData.id_tipo_proyecto)
                        .sort((a,b) => (Number(a.orden) || 999) - (Number(b.orden) || 999));

    if (etapasTipo.length === 0) {
      throw new Error("El Tipo de Proyecto seleccionado no tiene etapas configuradas. Configure CONF_ETAPAS primero.");
    }

    const idsEtapasValidas = etapasTipo.map(e => e.id);

    // Leemos tareas y filtramos solo las que coinciden con las etapas del tipo
    const tareasTemplate = readConfig("CONF_TAREAS")
                            .filter(t => idsEtapasValidas.includes(t.etapa_id));

    // --- 3. GESTIÓN ROBUSTA DE GOOGLE DRIVE ---
    // Primero obtenemos la configuración de la carpeta raíz
    let driveUrl = "";
    let driveId = "";
    
    try {
      const generalData = sheetGeneral.getDataRange().getValues();
      const rootUrlRow = generalData.find(r => r[0] === "DRIVE_ROOT_FOLDER_URL");
      
      if (rootUrlRow && rootUrlRow[1]) {
        // Extraemos ID de la URL
        const match = rootUrlRow[1].match(/[-\w]{25,}/);
        const rootFolderId = match ? match[0] : null;
        
        if (rootFolderId) {
          const rootFolder = DriveApp.getFolderById(rootFolderId);
          const folderName = `${projectData.codigo} - ${projectData.nombre_obra}`;
          
          // Crear carpeta del proyecto
          const projectFolder = rootFolder.createFolder(folderName);
          driveId = projectFolder.getId();
          driveUrl = projectFolder.getUrl();
          
          // Crear subcarpetas de etapas (Iteración limpia)
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
      // No detenemos el proceso, pero dejamos constancia
    }

    // Asignamos resultados de Drive al objeto de datos
    projectData.drive_folder_id = driveId;
    projectData.drive_url = driveUrl;

    // --- 4. GUARDAR EN DB_PROYECTOS (Mapeo Inteligente) ---
    // Función helper para normalizar encabezados: "Nombre Obra" -> "nombre_obra"
    const normalizeHeader = (h) => h.toString().toLowerCase().trim().replace(/\s+/g, '_');

    const headersProj = sheetProyectos.getRange(1, 1, 1, sheetProyectos.getLastColumn()).getValues()[0];
    const newRowProj = headersProj.map(header => {
      const key = normalizeHeader(header);
      // Buscamos la clave exacta o la normalizada en projectData
      return projectData[key] !== undefined ? projectData[key] : (projectData[header] || "");
    });
    
    sheetProyectos.appendRow(newRowProj);

    // --- 5. BATCH INSERT EN DB_EJECUCION ---
    if (tareasTemplate.length > 0) {
      const headersEjec = sheetEjecucion.getRange(1, 1, 1, sheetEjecucion.getLastColumn()).getValues()[0];
      
      const rowsToInsert = tareasTemplate.map(tpl => {
        // Objeto temporal de la nueva tarea
        const nuevaTarea = {
          id: Utilities.getUuid(),
          proyecto_id: projectData.id,
          etapa_id: tpl.etapa_id,
          nombre_tarea: tpl.nombre_tarea,
          requiere_evidencia: tpl.requiere_evidencia,
          tipo_entrada: tpl.tipo_entrada || 'text',
          checklist_id: tpl.checklist_id || '',
          estado: '',
          responsable_id: '',
          datos_evidencia: '',
          comentarios: '',
          updated_at: timestamp
        };

        // Mapeo seguro contra las columnas reales de la hoja
        return headersEjec.map(header => {
          const key = normalizeHeader(header);
          return nuevaTarea[key] !== undefined ? nuevaTarea[key] : "";
        });
      });

      // Escritura en bloque (Una sola llamada a API)
      sheetEjecucion.getRange(
        sheetEjecucion.getLastRow() + 1, 
        1, 
        rowsToInsert.length, 
        rowsToInsert[0].length
      ).setValues(rowsToInsert);
    }

    SpreadsheetApp.flush();
    return { success: true, message: `Proyecto [${projectData.nombre_obra}] y su ciclo de vida creados correctamente.` };

  } catch (e) {
    console.error("Error crítico en createProjectFull:", e);
    throw new Error(e.message || e.toString());
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

    // 3. BORRAR REGISTRO DE PROYECTO (DB_PROYECTOS)
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
 * REGENERACIÓN BLINDADA DE TAREAS
 * Borra las tareas actuales del proyecto y recarga SOLO las correspondientes a su id_tipo_proyecto.
 * Optimización: Batch Insert y Mapeo Dinámico de Columnas.
 */
function regenerateProjectTasks(projectId) {
  const lock = LockService.getScriptLock();
  try {
    // 1. Bloqueo de seguridad extendido (las regeneraciones son costosas)
    lock.waitLock(30000); 
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ----------------------------------------------------
    // PASO 1: OBTENER EL ADN DEL PROYECTO (TIPO)
    // ----------------------------------------------------
    const sheetProyectos = ss.getSheetByName("DB_PROYECTOS");
    const dataProyectos = sheetProyectos.getDataRange().getValues();
    const headersProy = dataProyectos[0];
    
    // Buscamos índices dinámicamente
    const idxIdProy = headersProy.indexOf("id");
    const idxTipoProy = headersProy.indexOf("id_tipo_proyecto");
    
    if (idxIdProy === -1 || idxTipoProy === -1) throw new Error("Estructura de DB_PROYECTOS inválida.");

    // Buscamos el proyecto
    let idTipo = null;
    let nombreObra = "";
    
    for (let i = 1; i < dataProyectos.length; i++) {
      if (dataProyectos[i][idxIdProy] === projectId) {
        idTipo = dataProyectos[i][idxTipoProy];
        nombreObra = dataProyectos[i][headersProy.indexOf("nombre_obra")] || "Sin nombre";
        break;
      }
    }

    if (!idTipo) throw new Error("El proyecto no tiene asignado un 'Tipo de Proyecto' o no existe.");

    // ----------------------------------------------------
    // PASO 2: OBTENER LA CONFIGURACIÓN FILTRADA (JOIN)
    // ----------------------------------------------------
    // Usamos readConfig para obtener objetos limpios de la configuración
    
    // A. Filtramos Etapas por Tipo
    const etapas = readConfig("CONF_ETAPAS")
                    .filter(e => e.id_tipo_proyecto === idTipo)
                    .sort((a,b) => a.orden - b.orden); // Respetamos el orden lógico
    
    if (etapas.length === 0) throw new Error("El tipo de proyecto asignado no tiene etapas configuradas.");
    
    const idsEtapasValidas = etapas.map(e => e.id);

    // B. Filtramos Tareas por Etapas válidas
    const tareasTemplate = readConfig("CONF_TAREAS")
                            .filter(t => idsEtapasValidas.includes(t.etapa_id));

    // ----------------------------------------------------
    // PASO 3: LIMPIEZA DE TAREAS VIEJAS (DELETE)
    // ----------------------------------------------------
    const sheetExec = ss.getSheetByName("DB_EJECUCION");
    const dataExec = sheetExec.getDataRange().getValues();
    const headersExec = dataExec[0];
    const idxProjIdExec = headersExec.indexOf("proyecto_id");

    if (idxProjIdExec === -1) throw new Error("No se encontró columna 'proyecto_id' en ejecución.");

    // Recorremos hacia atrás para borrar sin romper índices
    // Optimización: Solo borramos si encontramos coincidencia
    let rowsDeleted = 0;
    for (let i = dataExec.length - 1; i >= 1; i--) {
      if (dataExec[i][idxProjIdExec] === projectId) {
        sheetExec.deleteRow(i + 1);
        rowsDeleted++;
      }
    }
    console.log(`🧹 Se eliminaron ${rowsDeleted} tareas antiguas.`);

    // ----------------------------------------------------
    // PASO 4: INSERCIÓN MASIVA DE NUEVAS TAREAS (BATCH INSERT)
    // ----------------------------------------------------
    if (tareasTemplate.length > 0) {
      const timestamp = new Date();
      
      // Mapeamos los datos al orden real de columnas de DB_EJECUCION
      const rowsToInsert = tareasTemplate.map(tpl => {
        // Objeto temporal con los datos a insertar
        const rowData = {
          id: Utilities.getUuid(),
          proyecto_id: projectId,
          etapa_id: tpl.etapa_id,
          nombre_tarea: tpl.nombre_tarea,
          requiere_evidencia: tpl.requiere_evidencia,
          tipo_entrada: tpl.tipo_entrada || 'text', // Fallback default
          checklist_id: tpl.checklist_id || '',
          estado: 'Pendiente',
          responsable_id: '',
          datos_evidencia: '',
          comentarios: '',
          updated_at: timestamp
        };

        // Convertimos el objeto a array ordenado según los encabezados de la hoja
        return headersExec.map(headerName => {
            // Normalizamos keys a lowercase para matching seguro
            const key = Object.keys(rowData).find(k => k.toLowerCase() === headerName.toLowerCase().trim());
            return key ? rowData[key] : "";
        });
      });

      // Escritura en bloque (Una sola llamada a API)
      sheetExec.getRange(
        sheetExec.getLastRow() + 1, 
        1, 
        rowsToInsert.length, 
        rowsToInsert[0].length
      ).setValues(rowsToInsert);
    }

    SpreadsheetApp.flush();
    return { 
      success: true, 
      message: `Proyecto '${nombreObra}' regenerado: ${rowsDeleted} tareas borradas, ${tareasTemplate.length} nuevas insertadas según el tipo.` 
    };

  } catch (e) {
    console.error("🔥 Error crítico en regeneración:", e);
    throw new Error(e.message); // Re-lanzamos limpio para el Frontend
  } finally {
    lock.releaseLock();
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
 * Actualiza el Tipo de Proyecto y REGENERA toda su estructura de tareas.
 * ADVERTENCIA: Acción destructiva.
 */
function updateProjectTypeAndReset(projectId, newTypeId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Actualizar la cabecera del proyecto (DB_PROYECTOS)
    const sheetProy = ss.getSheetByName("DB_PROYECTOS");
    const data = sheetProy.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf("id");
    const typeIdx = headers.indexOf("id_tipo_proyecto"); // Asegúrate de que este nombre sea exacto en tu hoja
    
    if (idIdx === -1 || typeIdx === -1) throw new Error("Columnas ID o TIPO no encontradas en DB_PROYECTOS");

    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === projectId) {
        // Actualizamos celda directa (i+1 porque es base 1, typeIdx+1 porque es base 1)
        sheetProy.getRange(i + 1, typeIdx + 1).setValue(newTypeId);
        found = true;
        break;
      }
    }

    if (!found) throw new Error("Proyecto no encontrado");

    // 2. Liberar el lock aquí para permitir que la regeneración (que tiene su propio lock) funcione
    lock.releaseLock(); 

    // 3. Llamar a la regeneración (Esta función ya lee el nuevo tipo de la DB)
    return regenerateProjectTasks(projectId);

  } catch (e) {
    console.error(e);
    throw new Error("Error al cambiar tipo: " + e.message);
  } finally {
    // Seguridad extra por si falla antes del release manual
    try { lock.releaseLock(); } catch(e) {}
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