/**
 * =========================================================================
 * MÓDULO DE ENRUTAMIENTO WEB (CONEXIÓN CON GITHUB PAGES)
 * =========================================================================
 */

// 1. Maneja las peticiones de lectura (GET). Envía los datos del edificio a GitHub.
function doGet(e) {
  try {
    var datos = obtenerDatosEdificio();
    return ContentService.createTextOutput(JSON.stringify(datos))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ exito: false, mensaje: error.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

// 2. Maneja las peticiones de escritura (POST). Recibe el voucher desde GitHub y lo procesa.
function doPost(e) {
  try {
    // Validar que existan datos de entrada
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No se recibieron datos en la petición POST.");
    }
    
    // Convertir el texto JSON que envía GitHub en un objeto JavaScript
    var parametros = JSON.parse(e.postData.contents);
    
    // Ejecutar la función lógica de registro
    var resultado = registrarPagoConArchivo(parametros); 
    
    return ContentService.createTextOutput(JSON.stringify(resultado))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ exito: false, mensaje: error.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * =========================================================================
 * LÓGICA DE NEGOCIO (PROCESAMIENTO DE DATOS)
 * =========================================================================
 */

// 3. Obtener datos con mapeo exacto de las columnas automatizadas
function obtenerDatosEdificio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('BD edificio MECM') || ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var objetoEdificio = {};

  var limpiarNum = function(valor) {
    if (!valor) return 0;
    var procesado = String(valor).replace(/[^0-9.]/g, '');
    return parseFloat(procesado) || 0;
  };

  for (var i = 1; i < data.length; i++) {
    var fila = data[i];
    var idDepto = String(fila[1]).trim(); // Columna B
    if (idDepto) {
      objetoEdificio[idDepto] = {
        inquilino: String(fila[2]),         // Columna C
        fono: String(fila[4]),              // Columna E
        alquiler: limpiarNum(fila[5]),      // Columna F
        mantenimiento: limpiarNum(fila[6]), // Columna G
        limpieza: limpiarNum(fila[7]),      // Columna H
        seguridad: limpiarNum(fila[8]),     // Columna I
        agua: limpiarNum(fila[9]),          // Columna J
        electricity: limpiarNum(fila[10]),  // Columna K
        totalMora: limpiarNum(fila[11]),    // Columna L (Total Inicial)
        montoRecibido: limpiarNum(fila[13]),// Columna N (Ingresado Manual)
        diferencia: limpiarNum(fila[14]),   // Columna O (Resta automatizada)
        estadoSheet: String(fila[15]).trim() // Columna P (Estado automatizado)
      };
    }
  }
  return objetoEdificio;
}

// 4. Registrar voucher en Google Drive y guardar el enlace en la columna M del Excel
function registrarPagoConArchivo(e) {
  try {
    if (!e) {
      throw new Error("Ejecución manual detectada. Esta función solo recibe datos desde la página web.");
    }
    
    if (!e.datosBase64 || !e.idDepto) {
      throw new Error("Los datos del comprobante o el número de departamento están incompletos.");
    }

    // --- CARPETA CONFIGURADA CON TU ID ---
    var idCarpetaVouchers = "1c4iFR6_4yHWoQDSkwuBSPaL1eOIN8xyx"; 
    var carpeta = DriveApp.getFolderById(idCarpetaVouchers);
    
    var bytes = Utilities.base64Decode(e.datosBase64);
    var blob = Utilities.newBlob(bytes, e.tipoMime, e.nombreArchivo);
    
    // Guardamos el archivo en Drive
    var archivoGuardado = carpeta.createFile(blob);
    
    // USAMOS UN BLOQUE DE SEGURIDAD: Intentar dar permisos públicos de lectura al archivo
    try {
      archivoGuardado.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (errPermiso) {
      // Si la cuenta tiene directivas que bloquean el acceso externo, no detiene la ejecución
      Logger.log("Nota: Permisos públicos omitidos debido a políticas de la cuenta.");
    }
    
    var urlDeLaFoto = archivoGuardado.getUrl();

    // URL de tu hoja de cálculo
    var urlExcel = "https://docs.google.com/spreadsheets/d/1g-QEovPN5mE11aB9rZOBLvxTuVtWDu2utQuRAAl-ZsM/edit"; 
    var libro = SpreadsheetApp.openByUrl(urlExcel);
    var hoja = libro.getSheetByName("BD edificio MECM") || libro.getSheets()[0];
    
    var datos = hoja.getDataRange().getValues();
    var filaEncontrada = -1;

    // Normalizar ID buscado quitando espacios y posibles decimales (.0)
    var idBuscado = String(e.idDepto).trim().split('.')[0];

    for (var i = 1; i < datos.length; i++) {
      if (datos[i][1] !== undefined && datos[i][1] !== "") {
        var idCeldaExcel = String(datos[i][1]).trim().split('.')[0];
        
        if (idCeldaExcel === idBuscado) {
          filaEncontrada = i + 1; // Ajuste por índice base 0 y fila de encabezado
          break;
        }
      }
    }

    if (filaEncontrada === -1) {
      throw new Error("No se pudo localizar el Departamento '" + idBuscado + "' en la columna B de la hoja.");
    }

    // Escribir enlace en la columna M (Columna 13)
    hoja.getRange(filaEncontrada, 13).setValue(urlDeLaFoto);
    
    // Forzar el vaciado de memoria inmediato para asegurar la escritura de los datos
    SpreadsheetApp.flush();

    return {
      exito: true,
      mensaje: "¡Comprobante subido con éxito y registrado en Excel!",
      urlFoto: urlDeLaFoto
    };

  } catch (error) {
    throw new Error(error.message);
  }
}