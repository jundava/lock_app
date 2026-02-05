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
  GENERAL: "CONF_GENERAL"
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
 * Lectura robusta: Normaliza encabezados a minúsculas y elimina espacios
 */
function readConfig(tableName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tableName);
  SpreadsheetApp.flush(); // Forzar escritura pendiente
  
  if (!sheet) return [];
  
  // Usamos getDisplayValues para obtener TODO como texto (evita problemas de fechas/números)
  const data = sheet.getDataRange().getDisplayValues();
  
  if (data.length <= 1) return []; // Solo encabezados o vacía
  
  // Normalizamos encabezados: "Nombre Etapa " -> "nombre_etapa" (si fuera el caso)
  // Aquí asumimos que queremos usar las claves tal cual vienen en la hoja pero limpias
  const originalHeaders = data.shift();
  const headers = originalHeaders.map(h => h.toString().trim().toLowerCase());
  
  return data.map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      // Mapeo seguro: si la columna tiene encabezado, guardamos el dato
      if(header) obj[header] = row[i];
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

