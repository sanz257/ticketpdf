/**
 * Constantes de Configuración
 * Por favor, verifica y actualiza estos valores según sea necesario.
 */
const SPREADSHEET_ID = '1xkOFm2zq-T5bzDhRwarrfzfzhwKNyASilOwqMsOsMWg';
const CLIENTE_SHEET_NAME = 'detalle_cliente';
const ORDEN_SHEET_NAME = 'detalle_orden';
const DRIVE_FOLDER_ID = '1vPvxoHmifC91D2XYa9ZIOSyXsLtG1ewi';
const IGV_RATE = 0.18; // Tasa de IGV (18%)

/**
 * Función principal que recibe la solicitud HTTP POST de AppSheet (Webhook).
 * @param {object} e El objeto de evento de la solicitud HTTP.
 * @returns {GoogleAppsScript.Content.TextOutput} Respuesta JSON para AppSheet.
 */
function doPost(e) {
  try {
    // 1. Obtener los parámetros enviados por AppSheet.
    // AppSheet típicamente envía un JSON en el cuerpo de la solicitud.
    if (!e.postData || e.postData.type !== 'application/json') {
      throw new Error('Tipo de contenido no soportado. Se espera JSON.');
    }

    const payload = JSON.parse(e.postData.contents);
    Logger.log('Payload recibido: ' + JSON.stringify(payload));

    // Validar y extraer los parámetros.
    const params = {
      id_orden: payload.id_orden,
      fecha: payload.fecha,
      hora: payload.hora,
      direccion: payload.direccion,
      observacion: payload.observacion,
      empleado: payload.empleado,
      tipo_movimiento: payload.tipo_movimiento,
      pago: payload.pago
    };

    if (!params.id_orden) {
      throw new Error('El parámetro id_orden es obligatorio.');
    }

    // Ejecutar la generación del PDF
    const result = generarPdfTicket(params);

    return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: 'PDF generado y guardado.', fileName: result.fileName, fileUrl: result.fileUrl }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('Error en doPost: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Fallo al generar PDF: ' + error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Lógica principal para generar el ticket PDF.
 * @param {object} params Los parámetros de la orden (id_orden, fecha, hora, etc.).
 * @returns {object} Un objeto con el nombre y URL del archivo generado.
 */
function generarPdfTicket(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 2, 3, 4. Recolectar datos del cliente y de la orden.
  const clienteData = getDetalleCliente(ss, params.id_orden);
  const ordenItems = getDetalleOrden(ss, params.id_orden);
  
  if (ordenItems.length === 0) {
    throw new Error(`No se encontraron productos para la ID de Orden: ${params.id_orden}`);
  }

  // 5. Calcular totales.
  const totales = calcularTotales(ordenItems);

  // Combinar todos los datos para pasarlos al HTML.
  const data = {
    ...params, // Incluye todos los parámetros iniciales
    cliente: clienteData,
    items: ordenItems,
    ...totales // Incluye op_gravadas, IGV, TOTALAPAGAR
  };
  
  // 6. Generar el contenido HTML a partir de la plantilla.
  const template = HtmlService.createTemplateFromFile('ticket');
  template.data = data; // Asignamos el objeto de datos a la plantilla.
  const htmlOutput = template.evaluate();

  // Opciones de impresión para formato de ticket (380px de ancho).
  const options = {
    landscape: false, // Vertical
    pageSize: 'A4',   // Aunque es A4, el CSS del HTML controla el ancho para simular el ticket.
    // Aunque no hay una opción directa para ancho, el CSS lo forzará.
    // Marginos mínimos para maximizar el área de impresión.
    margin: 10
  };
  
  const pdfBlob = htmlOutput.getAs(MimeType.PDF);
  
  // Nombrar el archivo.
  const fileName = `TICKET_${params.id_orden}_${data.fecha.replace(/\//g, '-')}.pdf`;

  // 7. Guardar el PDF en la carpeta de Google Drive.
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  pdfBlob.setName(fileName);
  const file = folder.createFile(pdfBlob);
  
  Logger.log(`PDF generado y guardado como: ${file.getName()} en ${folder.getName()}`);

  return { fileName: file.getName(), fileUrl: file.getUrl() };
}

/**
 * Busca los detalles del cliente filtrando por ID_ORDEN.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss El objeto del libro 'orden'.
 * @param {string} id_orden El ID de la orden a buscar.
 * @returns {object} Los datos del primer cliente encontrado.
 */
function getDetalleCliente(ss, id_orden) {
  const sheet = ss.getSheetByName(CLIENTE_SHEET_NAME);
  if (!sheet) {
    Logger.log(`Hoja no encontrada: ${CLIENTE_SHEET_NAME}`);
    return {};
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // Quitamos los encabezados
  const cliente = {};

  // La solicitud indica buscar donde DNI_RUC (columna B, índice 1) coincida con id_orden.
  // Columna A (ID id_orden) es índice 0. Columna B (dni_ruc) es índice 1.
  const ID_ORDEN_COL = headers.indexOf('ID id_orden'); // Índice 0
  const DNI_RUC_COL = headers.indexOf('dni_ruc'); // Índice 1

  if (DNI_RUC_COL === -1 || ID_ORDEN_COL === -1) {
    Logger.log('Columnas ID id_orden o dni_ruc no encontradas en detalle_cliente. Usando índices 0 y 1 por defecto.');
    // Si no se encuentran los nombres, usamos índices fijos para seguridad.
    // La solicitud pide filtrar por DNI_RUC (col 2) = id_orden. Esto se mantiene.
    // Nota: Se buscará el ID de orden en la columna DNI_RUC/RUC (columna B).
    
    // Si no encontramos los headers, asumimos la estructura: [0:ID id_orden, 1:dni_ruc, 2:nombrecomp, 3:razonsoc, 4:contacto, 5:direccionf]
    
    for (let i = 0; i < data.length; i++) {
      // Usamos el índice 1 (dni_ruc) para filtrar por el valor de id_orden (como se solicitó).
      if (data[i][1] && data[i][1].toString() === id_orden.toString()) { 
        cliente.Id_orden = data[i][0] || '';
        cliente.Dni_ruc = data[i][1] || '';
        cliente.Nombrecomp = data[i][2] || '';
        cliente.Razonsoc = data[i][3] || '';
        cliente.Contacto = data[i][4] || '';
        cliente.Direccionf = data[i][5] || '';
        return cliente; // Devolvemos el primer match
      }
    }
  } else {
    // Si se encuentran los headers, buscamos por ellos.
    for (let i = 0; i < data.length; i++) {
      // Búsqueda en la columna DNI_RUC
      if (data[i][DNI_RUC_COL] && data[i][DNI_RUC_COL].toString() === id_orden.toString()) {
        cliente.Id_orden = data[i][ID_ORDEN_COL] || '';
        cliente.Dni_ruc = data[i][DNI_RUC_COL] || '';
        cliente.Nombrecomp = data[i][headers.indexOf('nombrecomp')] || '';
        cliente.Razonsoc = data[i][headers.indexOf('razonsoc')] || '';
        cliente.Contacto = data[i][headers.indexOf('contacto')] || '';
        cliente.Direccionf = data[i][headers.indexOf('direccionf')] || '';
        return cliente; // Devolvemos el primer match
      }
    }
  }

  Logger.log(`No se encontró detalle de cliente para ID_ORDEN: ${id_orden}.`);
  return {};
}

/**
 * Busca los productos de la orden filtrando por id_orden.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss El objeto del libro 'orden'.
 * @param {string} id_orden El ID de la orden a buscar.
 * @returns {Array<object>} Una lista de objetos con los detalles de la orden.
 */
function getDetalleOrden(ss, id_orden) {
  const sheet = ss.getSheetByName(ORDEN_SHEET_NAME);
  if (!sheet) {
    Logger.log(`Hoja no encontrada: ${ORDEN_SHEET_NAME}`);
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // Quitamos los encabezados
  const items = [];

  // Asumimos que la columna 'id_orden' está en la posición 1 (índice B).
  const ID_ORDEN_COL = headers.indexOf('id_orden'); 
  const CODIGO_COL = headers.indexOf('codigo');
  const DESCRIPCION_COL = headers.indexOf('descripcion');
  const CANTIDAD_COL = headers.indexOf('cantidad');
  const PRECIO_UNIT_COL = headers.indexOf('preciounitario');
  const PRECIO_TOTAL_COL = headers.indexOf('preciototal');

  if (ID_ORDEN_COL === -1 || CODIGO_COL === -1) {
    Logger.log('Columnas clave no encontradas en detalle_orden. Usando índices fijos por defecto.');
    // Si no encontramos los headers, asumimos la estructura: [..., 1:id_orden, 2:codigo, 3:descripcion, 4:cantidad, 5:preciounitario, 6:preciototal]
    for (let i = 0; i < data.length; i++) {
      // Índice 1 es 'id_orden'
      if (data[i][1] && data[i][1].toString() === id_orden.toString()) {
        items.push({
          codigo: data[i][2],
          descripcion: data[i][3],
          cantidad: parseFloat(data[i][4]) || 0,
          preciounitario: parseFloat(data[i][5]) || 0,
          preciototal: parseFloat(data[i][6]) || 0
        });
      }
    }
  } else {
    // Si se encuentran los headers, buscamos por ellos.
    for (let i = 0; i < data.length; i++) {
      if (data[i][ID_ORDEN_COL] && data[i][ID_ORDEN_COL].toString() === id_orden.toString()) {
        items.push({
          codigo: data[i][CODIGO_COL],
          descripcion: data[i][DESCRIPCION_COL],
          cantidad: parseFloat(data[i][CANTIDAD_COL]) || 0,
          preciounitario: parseFloat(data[i][PRECIO_UNIT_COL]) || 0,
          preciototal: parseFloat(data[i][PRECIO_TOTAL_COL]) || 0
        });
      }
    }
  }

  return items;
}

/**
 * Calcula los totales (Total a Pagar, Op. Gravadas, IGV).
 * Asume que preciototal ya incluye el IGV.
 * @param {Array<object>} items Lista de productos de la orden.
 * @returns {object} Un objeto con los totales calculados.
 */
function calcularTotales(items) {
  let totalAPagar = items.reduce((sum, item) => sum + item.preciototal, 0);

  // Asumimos que Total a Pagar incluye el IGV
  const opGravadas = totalAPagar / (1 + IGV_RATE);
  const igv = totalAPagar - opGravadas;

  return {
    op_gravadas: opGravadas,
    IGV: igv,
    TOTALAPAGAR: totalAPagar
  };
}

/**
 * Función que permite cargar el HTML en el editor de Apps Script.
 * Se puede usar para pruebas.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('ticket').evaluate();
}
