/**
 * ID de la hoja de cálculo proporcionada por el usuario.
 * @type {string}
 */
const SPREADSHEET_ID = '19YtSQIEuYCWWZ3F5UQeZhdaO7VogqRUPNWbuOUPigp0';

/**
 * Función principal que se ejecuta al acceder a la URL de la aplicación web.
 */
function doGet() {
  const html = HtmlService.createTemplateFromFile('index').evaluate();
  return html.setTitle('🛒 Gestor de Lista de Compras')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Tarea: Inicializa la Hoja de Cálculo (Spreadsheet) con las hojas y encabezados necesarios.
 * Debe ejecutarse una vez después de la configuración inicial en Apps Script.
 */
function initializeSheets() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // --- Hoja 'Lista' (Lista de Compras Actual) ---
    let listSheet = ss.getSheetByName('Lista');
    if (!listSheet) {
      listSheet = ss.insertSheet('Lista', 0);
    }
    // Encabezados de la Lista: Nombre, Cantidad, PrecioTotal, Categoría, Tienda, Comprado, ID_Item
    const listHeaders = ['Nombre', 'Cantidad', 'PrecioTotal', 'Categoria', 'Tienda', 'Comprado', 'ID_Item'];
    listSheet.getRange(1, 1, 1, listHeaders.length).setValues([listHeaders]).setFontWeight('bold');
    
    // --- Hoja 'Historial' (Registro de Gastos Semanales) ---
    let historySheet = ss.getSheetByName('Historial');
    if (!historySheet) {
      historySheet = ss.insertSheet('Historial', 1);
    }
    // Encabezados del Historial: Fecha, GastoTotal, GastoCasa, GastoBebe
    const historyHeaders = ['Fecha', 'GastoTotal', 'GastoCasa', 'GastoBebe'];
    historySheet.getRange(1, 1, 1, historyHeaders.length).setValues([historyHeaders]).setFontWeight('bold');
    
    return { success: true, message: "Hojas 'Lista' e 'Historial' inicializadas correctamente." };
  } catch (e) {
    Logger.log("Error al inicializar las hojas: " + e.toString());
    return { success: false, message: "Error al inicializar las hojas: " + e.toString() };
  }
}

/**
 * Tarea: Lee la lista de compras actual de la hoja 'Lista'.
 * @returns {Array<object>} La lista de items de compra.
 */
function getShoppingList() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Lista');
    if (!sheet) {
      return [];
    }

    const range = sheet.getDataRange();
    if (range.getNumRows() <= 1) {
      return []; // Solo hay encabezados
    }

    const data = range.getValues();
    const headers = data.shift(); // Elimina y obtiene los encabezados
    
    const list = data.map(row => {
      // Mapear los datos de la fila a un objeto Item
      const item = {
        id: row[6], // ID_Item
        store: row[4], // Tienda
        name: row[0], // Nombre
        quantityDisplay: row[1], // Cantidad
        // Al leer de Sheets, el precio puede ser un número o un string formateado (ej. "1,234.56")
        // Lo convertimos a string, quitamos las comas de miles y luego lo parseamos.
        totalPrice: parseFloat(String(row[2]).replace(/,/g, '')) || 0,
        expenseCategory: row[3], // Categoría
        // Hacemos la comprobación robusta para 'TRUE', true, 'true', etc.
        isPurchased: String(row[5]).toUpperCase() === 'TRUE', 
      };
      return item;
    });

    return list;
  } catch (e) {
    Logger.log("Error al obtener la lista: " + e.toString());
    return [];
  }
}

/**
 * Tarea: Sobrescribe la lista de compras actual en la hoja 'Lista'.
 * @param {Array<object>} list La nueva lista de items.
 */
function saveShoppingList(list) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Lista');
    if (!sheet) {
      throw new Error("La hoja 'Lista' no fue encontrada.");
    }

    // Limpiar solo el rango de datos (de la fila 2 hasta el final)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }

    if (!list || list.length === 0) {
      return true; // Lista vacía, no hay nada que escribir
    }

    const dataToWrite = list.map(item => [
      item.name,
      item.quantityDisplay,
      item.totalPrice,
      item.expenseCategory,
      item.store,
      item.isPurchased ? 'TRUE' : 'FALSE',
      item.id
    ]);

    if (dataToWrite.length > 0) {
      sheet.getRange(2, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
    }

    return true;
  } catch (e) {
    Logger.log("Error al guardar la lista: " + e.message);
    // Relanzar el error para que el cliente (onFailure) lo capture
    throw new Error("Error del servidor al guardar la lista: " + e.message);
  }
}

/**
 * Tarea: Registra un nuevo gasto histórico en la hoja 'Historial'.
 * @param {object} record El registro de gasto a guardar.
 */
function logHistory(record) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Historial');
    if (!sheet) return false;

    // El frontend envía: { totalSpent, breakdown: { Casa, Bebé } }
    const row = [
      new Date(), // Fecha
      record.totalSpent, // GastoTotal
      record.breakdown.Casa, // GastoCasa
      record.breakdown.Bebé // GastoBebe
    ];

    // Añadir la fila al final de la hoja
    sheet.appendRow(row);
    return true;
  } catch (e) {
    Logger.log("Error al registrar historial: " + e.toString());
    return false;
  }
}

/**
 * Tarea: Elimina un registro específico del historial por su timestamp
 * @param {number} timestamp El timestamp del registro a eliminar
 * @returns {boolean} true si se eliminó correctamente, false en caso contrario
 */
function deleteHistoryRecord(timestamp) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Historial');
    if (!sheet) return false;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return false; // Solo encabezados

    // Buscar la fila que coincida con el timestamp
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].getTime() === timestamp) {
        sheet.deleteRow(i + 1); // +1 porque los índices de hoja empiezan en 1
        return true;
      }
    }
    return false;
  } catch (e) {
    Logger.log("Error al eliminar registro del historial: " + e.toString());
    return false;
  }
}

/**
 * Tarea: Obtiene el historial reciente (últimos 5) de la hoja 'Historial'.
 * @returns {Array<object>} Los últimos 5 registros de historial.
 */
function getHistory() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Historial');
    if (!sheet) return [];

    const numRows = sheet.getLastRow();
    if (numRows <= 1) return []; // Solo encabezados

    // Leer hasta 5 filas (excluyendo encabezados)
    const rowsToRead = Math.min(5, numRows - 1);
    const startRow = numRows - rowsToRead + 1;

    // Leer los datos (columnas A:D)
    const data = sheet.getRange(startRow, 1, rowsToRead, 4).getValues();

    // Invertir el orden para que los más recientes estén al principio (como en Firestore)
    return data.reverse().map(row => ({
      // Mapear los datos de la fila a un objeto Historial
      timestamp: row[0].getTime(), // Usar milisegundos para simular el Timestamp
      totalSpent: row[1],
      breakdown: {
        Casa: row[2],
        Bebé: row[3]
      }
    }));
  } catch (e) {
    Logger.log("Error al obtener historial: " + e.toString());
    return [];
  }
}