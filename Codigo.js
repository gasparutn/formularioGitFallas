// =========================================================
// (doGet CORREGIDA por solicitud del usuario)
// =========================================================
function doGet(e) {
  try {
    const params = e.parameter;
    Logger.log("doGet INICIADO. Parámetros de URL: " + JSON.stringify(params));

    const appUrl = ScriptApp.getService().getUrl();

    // --- (NUEVA MODIFICACIÓN) ---
    // Si Mercado Pago devuelve "failure" (ej. "Volver al sitio"),
    // forzar la carga del formulario, ignorando cualquier paymentId que pueda venir.
    if (params.status === 'failure') {
      Logger.log("doGet detectó 'status=failure'. Redirigiendo al formulario.");
      // Cae en el "CAMINO 2" (mostrar formulario)
    }
    // --- (FIN DE LA MODIFICACIÓN) ---

    // Si status NO es 'failure' (es 'success', 'pending', o no existe)
    else {
      let paymentId = null;

      // --- Lógica para buscar el ID de pago ---
      if (params) {
        if (params.payment_id) {
          paymentId = params.payment_id;
        } else if (params.data && typeof params.data === 'string' && params.data.startsWith('{')) {
          try {
            const dataObj = JSON.parse(params.data);
            if (dataObj.id) paymentId = dataObj.id;
          } catch (jsonErr) {
            Logger.log("No se pudo parsear e.parameter.data: " + params.data);
          }
        } else if (params.topic && params.topic === 'payment' && params.id) {
          paymentId = params.id;
        }
      }

      // --- CAMINO 1: Se detectó un pago (Redirect de Mercado Pago) ---
      if (paymentId) {
        Logger.log("doGet detectó regreso de MP. Procesando Payment ID: " + paymentId);
        procesarNotificacionDePago(paymentId); // Vive en Pagos.gs

        // Esta página es simple y no necesita los archivos externos.
        // Tu código original está perfecto.
        const html = `
          <html>
            <head>
              <title>Pago Completo</title>
              <meta name="viewport" content="width=device-width, initial-scale=1.0">
              <style>
                body { font-family: Arial, sans-serif; display: flex; justify-content: center; align-items: center; height: 90vh; flex-direction: column; text-align: center; background-color: #f4f4f4; }
                .container { background-color: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); }
                .btn { display: inline-block; padding: 15px 30px; background-color: #28a745; color: white; text-decoration: none; border-radius: 5px; font-size: 1.2em; margin-top: 20px; transition: background-color: 0.3s; }
                .btn:hover { background-color: #218838; }
                h2 { color: #28a745; }
                p { font-size: 1.1em; color: #333; }
              </style>
            </head>
            <body>
              <div class="container">
                <h2>¡Pago Procesado Exitosamente!</h2>
                <p>Gracias por completar el pago. Presione el botón para volver al formulario.</p>
                <a href="${appUrl}" target="_top" class="btn">Volver al Formulario</a>
              </div>
            </body>
          </html>`;
        return HtmlService.createHtmlOutput(html)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
      }
      // Si no es 'failure' Y TAMPOCO tiene 'paymentId', cae al Camino 2.
    }

    // --- CAMINO 2: Visita normal al formulario (o 'status=failure') ---
    // Carga 'Index.html' como una plantilla
    const htmlTemplate = HtmlService.createTemplateFromFile('Index');

    // Pasa la URL de la app al HTML (para que tu JS la pueda usar si es necesario)
    htmlTemplate.appUrl = appUrl;

    // .evaluate() ejecuta los <?!= ... ?> e inserta el 
    // contenido de 'styleHead.html' y 'js.html'
    const html = htmlTemplate.evaluate()
      .setTitle("Formulario de Registro")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');

    return html;

  } catch (err) {
    Logger.log("Error en la detección de parámetros de doGet: " + err.toString());
    return HtmlService.createHtmlOutput("<b>Ocurrió un error:</b> " + err.message);
  }
}

// =========================================================
// (doPost - Webhook)
// =========================================================
function doPost(e) {
  let postData;
  try {
    Logger.log("doPost INICIADO. Contenido de 'e': " + JSON.stringify(e));
    if (!e || !e.postData || !e.postData.contents) {
      Logger.log("Error: El objeto 'e' o 'e.postData.contents' está vacío.");
      return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": "Payload vacío" })).setMimeType(ContentService.MimeType.JSON);
    }
    postData = e.postData.contents;
    Logger.log("doPost: Datos recibidos (raw): " + postData);
    const notificacion = JSON.parse(postData);
    Logger.log("doPost: Datos parseados (JSON): " + JSON.stringify(notificacion));

    if (notificacion.type === 'payment') {
      const paymentId = notificacion.data.id;
      if (paymentId) {
        Logger.log("Procesando ID de pago (desde doPost): " + paymentId);
        procesarNotificacionDePago(paymentId); // Vive en Pagos.gs
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ "status": "ok" })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log("Error grave en doPost (Webhook): " + error.toString());
    Logger.log("Datos (raw) que causaron el error: " + postData);
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// =========================================================
// (Punto 5, 11, 17, 24, 27) registrarDatos (ACTUALIZADO)
// (NUEVA MODIFICACIÓN) Cálculo de 'nuevoNumeroDeTurno' robustecido
// =========================================================
/**
* Guarda los datos finales en la hoja "Registros" (47 COLUMNAS)
* (Punto 5, 11) Ahora también registra a los hermanos.
* (Punto 17) Checkbox movido a AT (ahora AU).
* (Punto 24) Lógica de texto de Col C y D actualizada.
* (Punto 27) Añadida columna SOCIO (AB).
*/
function registrarDatos(datos) {
  Logger.log("REGISTRAR DATOS INICIADO. Datos: " + JSON.stringify(datos));
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(60000);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaConfig = ss.getSheetByName(NOMBRE_HOJA_CONFIG);
    let estadoActual = obtenerEstadoRegistro();

    if (estadoActual.cierreManual) return { status: 'CERRADO', message: 'El registro se encuentra cerrado.' };
    // (Punto 25) La validación de cupo para Pre-Venta ya se hizo en validarAcceso
    if (datos.tipoInscripto !== 'preventa' && estadoActual.alcanzado) return { status: 'LIMITE_ALCANZADO', message: 'Se ha alcanzado el cupo máximo.' };
    if (datos.tipoInscripto !== 'preventa' && datos.jornada === 'Jornada Normal extendida' && estadoActual.jornadaExtendidaAlcanzada) {
      return { status: 'LIMITE_EXTENDIDA', message: 'Se agotó el cupo para Jornada Extendida.' };
    }

    const dniBuscado = limpiarDNI(datos.dni);

    let hojaRegistro = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);

    // =========================================================
    // (¡¡¡NUEVA VALIDACIÓN ANTI-DUPLICADOS!!!)
    // =========================================================
    // Esta validación se ejecuta ANTES de añadir la fila
    // para prevenir que un DNI se registre dos veces.
    if (hojaRegistro && hojaRegistro.getLastRow() > 1) {
      const rangoDniRegistro = hojaRegistro.getRange(2, COL_DNI_INSCRIPTO, hojaRegistro.getLastRow() - 1, 1);
      const celdaRegistro = rangoDniRegistro.createTextFinder(dniBuscado).matchEntireCell(true).findNext();
      if (celdaRegistro) {
        Logger.log(`BLOQUEO DE REGISTRO: El DNI ${dniBuscado} ya existe en la fila ${celdaRegistro.getRow()}.`);
        // Usamos status 'ERROR' para que el frontend muestre el mensaje en rojo.
        return { status: 'ERROR', message: `El DNI ${dniBuscado} ya se encuentra registrado. No se puede crear un duplicado.` };
      }
    }
    // =========================================================
    // (FIN DE LA VALIDACIÓN)
    // =========================================================

    if (!hojaRegistro) {
      hojaRegistro = ss.insertSheet(NOMBRE_HOJA_REGISTRO);
      // --- (¡¡¡ENCABEZADOS ACTUALIZADOS PUNTO 27!!!) ---
      hojaRegistro.appendRow([
        'N° de Turno', 'Marca temporal', 'Marca N/E', 'Estado', // A-D
        'Email', 'Nombre', 'Apellido', // E-G (Punto 2)
        'Fecha de Nacimiento', 'Edad Actual', 'DNI', // H-J
        'Obra Social', 'Colegio/Jardin', // K-L
        'Responsable 1', 'DNI Resp 1', 'Tel Resp 1', // M-O (Punto 3, 5)
        'Responsable 2', 'Tel Resp 2', // P-Q (Punto 3)
        'Autorizados', // R
        'Deporte', 'Espec. Deporte', 'Enfermedad', 'Espec. Enfermedad', 'Alergia', 'Espec. Alergia', // S-X
        'Aptitud Física (Link)', 'Foto Carnet (Link)', // Y-Z
        'Jornada', 'SOCIO', // AA-AB (PUNTO 27)
        'Método de Pago', // AC
        'Precio', // AD (Punto 4)
        'Cuota 1', 'Cuota 2', 'Cuota 3', 'Cantidad Cuotas', // AE-AH
        'Estado de Pago', // AI
        'Monto a Pagar', // AJ (Punto 5)
        'ID Pago MP', 'Nombre Pagador (MP)', 'DNI Pagador MP', // AK-AM
        'Nombre y Apellido (Pagador Manual)', 'DNI Pagador (Manual)', // AN-AO (NUEVAS)
        'Comprobante MP', // AP (antes AM)
        'Comprobante Manual (Total/Ext)', // AQ (antes AN)
        'Comprobante Manual (C1)', // AR (antes AO)
        'Comprobante Manual (C2)', // AS (antes AP)
        'Comprobante Manual (C3)', // AT (antes AQ)
        'Enviar Email?' // AU (antes AR)
      ]);
      // (MODIFICACIÓN) Se elimina la línea que seteaba A2=2
    }

    // --- CÁLCULO DE PRECIOS ---
    // (MODIFICACIÓN) Llamada a la función helper (que vive en Pagos.gs) para centralizar la lógica
    const { precio, montoAPagar } = obtenerPrecioDesdeConfig(datos.metodoPago, datos.cantidadCuotas, hojaConfig);


    // --- REGISTRO DEL INSCRIPTO PRINCIPAL ---
    
    // (MODIFICACIÓN) Cálculo robusto del próximo N° de Turno
    const lastRow = hojaRegistro.getLastRow();
    let ultimoTurno = 0;
    if (lastRow > 1) {
      // Obtener todos los valores de la Columna A (desde la fila 2)
      const rangoTurnos = hojaRegistro.getRange(2, COL_NUMERO_TURNO, lastRow - 1, 1).getValues();
      const turnosReales = rangoTurnos.map(f => f[0]).filter(Number); // Filtra vacíos/texto y quédate con números
      
      if (turnosReales.length > 0) {
        ultimoTurno = Math.max(...turnosReales); // Encuentra el número más alto
      } else {
        ultimoTurno = 0; // <-- MODIFICACIÓN (era 1)
      }
    } else {
      ultimoTurno = 0; // <-- MODIFICACIÓN (era 1)
    }
    const nuevoNumeroDeTurno = ultimoTurno + 1; // Si la hoja está vacía, ultimoTurno=0, nuevo=1.


    const edadCalculada = calcularEdad(datos.fechaNacimiento);
    const edadFormateada = `${edadCalculada.anos}a, ${edadCalculada.meses}m, ${edadCalculada.dias}d`;
    const fechaObj = new Date(datos.fechaNacimiento);
    const fechaFormateada = Utilities.formatDate(fechaObj, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');

    // (Punto 24) Lógica de texto para Columna C (Marca N/E) y Columna D (Estado)
    let marcaNE = "";
    let estadoInscripto = "";
    const esPreventa = (datos.tipoInscripto === 'preventa');

    if (datos.jornada === 'Jornada Normal extendida') {
      marcaNE = esPreventa ? "Extendida (Pre-venta)" : "Extendida";
    } else { // Asume "Jornada Normal"
      marcaNE = esPreventa ? "Normal (Pre-Venta)" : "Normal";
    }

    if (esPreventa) {
      estadoInscripto = "Pre-Venta";
    } else {
      estadoInscripto = (datos.tipoInscripto === 'nuevo') ? 'Nuevo' : 'Anterior';
    }
    // (Fin Punto 24)

    const telResp1 = `(${datos.telAreaResp1}) ${datos.telNumResp1}`;
    const telResp2 = (datos.telAreaResp2 && datos.telNumResp2) ? `(${datos.telAreaResp2}) ${datos.telNumResp2}` : '';

    // (Punto 17, 27) appendRow actualizado para 47 columnas
    const filaNueva = [
      nuevoNumeroDeTurno, new Date(), marcaNE, estadoInscripto, // A-D
      datos.email, datos.nombre, datos.apellido, // E-G
      fechaFormateada, edadFormateada, dniBuscado, // H-J
      datos.obraSocial, datos.colegioJardin, // K-L
      datos.adultoResponsable1, datos.dniResponsable1, telResp1, // M-O
      datos.adultoResponsable2, telResp2, // P-Q
      datos.personasAutorizadas, // R
      datos.practicaDeporte, datos.especifiqueDeporte, datos.tieneEnfermedad, datos.especifiqueEnfermedad, datos.esAlergico, datos.especifiqueAlergia, // S-X
      datos.urlCertificadoAptitud || '', datos.urlFotoCarnet || '', // Y-Z
      datos.jornada, datos.esSocio, // AA-AB (PUNTO 27)
      datos.metodoPago, // AC
      precio, // AD (Precio)
      '', '', '', parseInt(datos.cantidadCuotas) || 0, // AE-AH
      datos.estadoPago, // AI (Estado de Pago)
      montoAPagar, // AJ (Monto a Pagar)
      '', '', '', // AK-AM (IDs de Pago MP, etc)
      '', '', // AN-AO (NUEVAS - Pagador Manual)
      '', // AP (Comprobante MP)
      '', '', '', '', // AQ-AT (Nuevos Comprobantes Manuales)
      false // AU (Checkbox)
    ];
    hojaRegistro.appendRow(filaNueva);
    const filaInsertada = hojaRegistro.getLastRow(); // Obtener la fila recién insertada

    const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    hojaRegistro.getRange(filaInsertada, COL_ENVIAR_EMAIL_MANUAL).setDataValidation(rule); // (Punto 17, 27) Columna AU

    // --- (Punto 5, 11) REGISTRO DE HERMANOS ---
    if (datos.hermanos && datos.hermanos.length > 0) {
      const hojaBusqueda = ss.getSheetByName(NOMBRE_HOJA_BUSQUEDA);
      
      // (NUEVA MODIFICACIÓN) El próximo turno de hermano debe continuar desde el principal
      let proximoTurnoHermano = nuevoNumeroDeTurno; 

      for (const hermano of datos.hermanos) {
        proximoTurnoHermano++; // Incrementar el turno para este hermano
        
        const dniHermano = limpiarDNI(hermano.dni);
        if (!dniHermano || !hermano.nombre || !hermano.apellido || !hermano.fechaNac) continue; // Saltar si faltan datos

        // (Punto 11) Determinar estado
        let estadoHermano = "Nuevo Hermano/a";
        if (hojaBusqueda && hojaBusqueda.getLastRow() > 1) {
          const rangoDNI = hojaBusqueda.getRange(2, COL_DNI_BUSQUEDA, hojaBusqueda.getLastRow() - 1, 1);
          const celdaEncontrada = rangoDNI.createTextFinder(dniHermano).matchEntireCell(true).findNext();
          if (celdaEncontrada) {
            estadoHermano = "Anterior Hermano/a";
          }
        }

        const edadCalcHermano = calcularEdad(hermano.fechaNac);
        const edadFmtHermano = `${edadCalcHermano.anos}a, ${edadCalcHermano.meses}m, ${edadCalcHermano.dias}d`;
        const fechaObjHermano = new Date(hermano.fechaNac);
        const fechaFmtHermano = Utilities.formatDate(fechaObjHermano, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');

        // ==========================================================================================
        // (MODIFICACIÓN CLAVE - SEGÚN ÚLTIMA SOLICITUD)
        // Se pre-rellenan los campos solicitados (Email, Obra Social, Responsables, Autorizados)
        // y se dejan vacíos los campos propios del hermano (Colegio, Salud, Pago, etc.)
        // ==========================================================================================
        const filaHermano = [
          proximoTurnoHermano, new Date(), '', estadoHermano, // A-D
          datos.email, hermano.nombre, hermano.apellido, // E-G (Email del principal)
          fechaFmtHermano, edadFmtHermano, dniHermano, // H-J
          datos.obraSocial, '', // K-L (Obra Social del principal, Colegio VACÍO)
          datos.adultoResponsable1, datos.dniResponsable1, telResp1, // M-O (Datos del Resp 1)
          datos.adultoResponsable2, telResp2, // P-Q (Datos del Resp 2)
          datos.personasAutorizadas, // R (Autorizados)
          '', '', '', '', '', '', // S-X (Salud VACÍO)
          '', '', // Y-Z (Aptitud, Foto VACÍOS)
          '', '', // AA-AB (Jornada, SOCIO VACÍOS - PUNTO 27)
          '', // AC (Método Pago VACÍO)
          0, // AD (Precio)
          '', '', '', 0, // AE-AH (Cuotas)
          'Pendiente (Hermano)', // AI (Estado de Pago)
          0, // AJ (Monto a Pagar)
          '', '', '', // AK-AM
          '', '', // AN-AO
          '', // AP
          '', '', '', '', // AQ-AT
          false // AU
        ];
        hojaRegistro.appendRow(filaHermano);
        const filaHermanoInsertada = hojaRegistro.getLastRow();
        hojaRegistro.getRange(filaHermanoInsertada, COL_ENVIAR_EMAIL_MANUAL).setDataValidation(rule); // (Punto 17, 27) Columna AU
      }
    }

    SpreadsheetApp.flush();
    obtenerEstadoRegistro(); // Actualiza el contador de cupos

    // Devolver solo los datos del inscripto principal
    return { status: 'OK_REGISTRO', message: '¡Registro Exitoso!', numeroDeTurno: nuevoNumeroDeTurno, datos: datos };

  } catch (e) {
    Logger.log("ERROR CRÍTICO EN REGISTRO: " + e.toString());
    return { status: 'ERROR', message: 'Fallo al registrar los datos: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}
// =========================================================
// (Punto 15, 17, 19, 27) subirComprobanteManual (ACTUALIZADO)
// =========================================================
/**
* Sube un comprobante manual a Drive y actualiza la hoja
* @param {string} dni DNI del inscripto (para buscar la fila)
* @param {object} fileData Objeto con {data, mimeType, fileName}
* @param {string} tipoComprobante 'total_mp', 'cuota1_mp', 'cuota2_mp', 'cuota3_mp', 'externo'
* @param {object} datosExtras {nombrePagador, dniPagador} (del mini-form)
*/
function subirComprobanteManual(dni, fileData, tipoComprobante, datosExtras) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const dniLimpio = limpiarDNI(dni);
    if (!dniLimpio || !fileData || !tipoComprobante) {
      return { status: 'ERROR', message: 'Faltan datos (DNI, archivo o tipo de comprobante).' };
    }
    
    // Validar que los datos del pagador (datosExtras) llegaron
    if (!datosExtras || !datosExtras.nombrePagador || !datosExtras.dniPagador) {
        return { status: 'ERROR', message: 'Faltan los datos del adulto pagador (Nombre o DNI).' };
    }

    // Usar el DNI del inscripto para la carpeta
    const fileUrl = uploadFileToDrive(fileData.data, fileData.mimeType, fileData.fileName, dniLimpio, 'comprobante');
    if (typeof fileUrl !== 'string' || !fileUrl.startsWith('http')) {
      throw new Error("Error al subir el archivo a Drive: " + (fileUrl.message || 'Error desconocido'));
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
    if (!hoja) throw new Error(`La hoja "${NOMBRE_HOJA_REGISTRO}" no fue encontrada.`);

    const rangoDni = hoja.getRange(2, COL_DNI_INSCRIPTO, hoja.getLastRow() - 1, 1);
    const celdaEncontrada = rangoDni.createTextFinder(dniLimpio).matchEntireCell(true).findNext();

    if (celdaEncontrada) {
      const fila = celdaEncontrada.getRow();
      let columnaDestinoArchivo; // Columna para el LINK

      // --- ¡NUEVA LÓGICA! ---
      // 1. Guardar los datos del pagador en sus columnas (AN y AO - DESPLAZADAS)
      hoja.getRange(fila, COL_PAGADOR_NOMBRE_MANUAL).setValue(datosExtras.nombrePagador); // AN (40)
      hoja.getRange(fila, COL_PAGADOR_DNI_MANUAL).setValue(datosExtras.dniPagador); // AO (41)

      // 2. Asignar la columna de destino para el LINK DEL ARCHIVO (COLUMNAS DESPLAZADAS)
      switch (tipoComprobante) {
        case 'total_mp':
        case 'mp_total': // Alias de Index.html
        case 'externo':
          // ¡ESTA ES TU PETICIÓN! Opción 'a' va a la columna AQ
          columnaDestinoArchivo = COL_COMPROBANTE_MANUAL_TOTAL_EXT; // AQ (43)
          break;
        case 'cuota1_mp':
        case 'mp_cuota_1': // Alias de Index.html
          columnaDestinoArchivo = COL_COMPROBANTE_MANUAL_CUOTA1; // AR (44)
          break;
        case 'cuota2_mp':
        case 'mp_cuota_2': // Alias de Index.html
          columnaDestinoArchivo = COL_COMPROBANTE_MANUAL_CUOTA2; // AS (45)
          break;
        case 'cuota3_mp':
        case 'mp_cuota_3': // Alias de Index.html
          columnaDestinoArchivo = COL_COMPROBANTE_MANUAL_CUOTA3; // AT (46)
          break;
        default:
          throw new Error(`Tipo de comprobante no reconocido: ${tipoComprobante}`);
      }

      // 3. Guardar SOLO el link en la columna de archivo correspondiente
      hoja.getRange(fila, columnaDestinoArchivo).setValue(fileUrl);
      
      // 4. Actualizar estado
      hoja.getRange(fila, COL_ESTADO_PAGO).setValue("En revisión");

      Logger.log(`Comprobante manual [${tipoComprobante}] subido para DNI ${dniLimpio} en fila ${fila}.`);
      return { status: 'OK', message: '¡Comprobante subido con éxito! Será revisado por la administración.' };
    } else {
      Logger.log(`No se encontró DNI ${dniLimpio} para subir comprobante manual.`);
      return { status: 'ERROR', message: `No se encontró el registro para el DNI ${dniLimpio}. Asegúrese de que el DNI del inscripto sea correcto.` };
    }

  } catch (e) {
    Logger.log("Error en subirComprobanteManual: " + e.toString());
    return { status: 'ERROR', message: 'Error en el servidor: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}
// --- FUNCIONES DE AYUDA (Helpers) ---

/* */
function uploadFileToDrive(data, mimeType, filename, dni, tipoArchivo) {
  try {
    if (!dni) return { status: 'ERROR', message: 'No se recibió DNI.' };
    let parentFolderId;
    switch (tipoArchivo) {
      case 'foto': parentFolderId = FOLDER_ID_FOTOS; break;
      case 'ficha': parentFolderId = FOLDER_ID_FICHAS; break;
      case 'comprobante': parentFolderId = FOLDER_ID_COMPROBANTES; break;
      default: return { status: 'ERROR', message: 'Tipo de archivo no reconocido.' };
    }
    if (!parentFolderId || parentFolderId.includes('AQUI_VA_EL_ID')) {
      return { status: 'ERROR', message: 'IDs de carpetas no configurados.' };
    }

    const parentFolder = DriveApp.getFolderById(parentFolderId);
    let subFolder;
    const folders = parentFolder.getFoldersByName(dni);
    subFolder = folders.hasNext() ? folders.next() : parentFolder.createFolder(dni);

    const decodedData = Utilities.base64Decode(data.split(',')[1]);
    const blob = Utilities.newBlob(decodedData, mimeType, filename);
    const file = subFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();

  } catch (e) {
    Logger.log('Error en uploadFileToDrive: ' + e.toString());
    return { status: 'ERROR', message: 'Error al subir archivo: ' + e.message };
  }
}

/* */
function limpiarDNI(dni) {
  if (!dni) return '';
  return String(dni).replace(/[.\s-]/g, '').trim();
}

/* */
function calcularEdad(fechaNacimientoStr) {
  if (!fechaNacimientoStr) return { anos: 0, meses: 0, dias: 0 };
  const fechaNacimiento = new Date(fechaNacimientoStr);
  const hoy = new Date();
  fechaNacimiento.setMinutes(fechaNacimiento.getMinutes() + fechaNacimiento.getTimezoneOffset());
  let anos = hoy.getFullYear() - fechaNacimiento.getFullYear();
  let meses = hoy.getMonth() - fechaNacimiento.getMonth();
  let dias = hoy.getDate() - fechaNacimiento.getDate();
  if (dias < 0) {
    meses--;
    dias += new Date(hoy.getFullYear(), hoy.getMonth(), 0).getDate();
  }
  if (meses < 0) {
    anos--;
    meses += 12;
  }
  return { anos, meses, dias };
}

/* */
// (NUEVA MODIFICACIÓN) 'registrosActuales' se calcula contando la Columna A
function obtenerEstadoRegistro() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaConfig = ss.getSheetByName(NOMBRE_HOJA_CONFIG);
    const hojaRegistro = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
    if (!hojaConfig) throw new Error(`Hoja "${NOMBRE_HOJA_CONFIG}" no encontrada.`);

    const limiteCupos = parseInt(hojaConfig.getRange('B1').getValue()) || 100;
    const limiteJornadaExtendida = parseInt(hojaConfig.getRange('B4').getValue());
    const formularioAbierto = hojaConfig.getRange('B11').getValue() === true;

    let registrosActuales = 0;
    let registrosJornadaExtendida = 0;
    
    if (hojaRegistro && hojaRegistro.getLastRow() > 1) {
      const lastRow = hojaRegistro.getLastRow();
      
      // (NUEVA MODIFICACIÓN) Contar registros basados en la Columna A (N° de Turno)
      const rangoTurnos = hojaRegistro.getRange(2, COL_NUMERO_TURNO, lastRow - 1, 1);
      const valoresTurnos = rangoTurnos.getValues();
      registrosActuales = valoresTurnos.filter(fila => fila[0] != null && fila[0] != "").length;
      
      // El conteo de Jornada Extendida sigue igual
      const data = hojaRegistro.getRange(2, COL_MARCA_N_E_A, lastRow - 1, 1).getValues();
      registrosJornadaExtendida = data.filter(row => String(row[0]).startsWith('Extendida')).length;
    }

    // Escribir el conteo robusto en la celda B2
    hojaConfig.getRange('B2').setValue(registrosActuales);
    hojaConfig.getRange('B5').setValue(registrosJornadaExtendida);
    SpreadsheetApp.flush();

    return {
      alcanzado: registrosActuales >= limiteCupos,
      jornadaExtendidaAlcanzada: registrosJornadaExtendida >= limiteJornadaExtendida,
      cierreManual: !formularioAbierto
    };
  } catch (e) {
    Logger.log("Error en obtenerEstadoRegistro: " + e.message);
    return { cierreManual: true, message: "Error al leer config: " + e.message };
  }
}


// =========================================================
// (Punto 1, 6, 7, 12, 25) validarAcceso (COMPLETAMENTE REESTRUCTURADO)
// =========================================================
function validarAcceso(dni, tipoInscripto) {
  try {
    // 1. OBTENER ESTADO GENERAL Y VALIDACIONES BÁSICAS
    const estado = obtenerEstadoRegistro();
    if (estado.cierreManual) return { status: 'CERRADO', message: 'El formulario se encuentra cerrado por mantenimiento.' };
    
    // El cupo máximo general solo aplica a 'nuevos'
    if (estado.alcanzado && tipoInscripto === 'nuevo') return { status: 'LIMITE_ALCANZADO', message: 'Se ha alcanzado el cupo máximo para nuevos registros.' };

    if (!dni) return { status: 'ERROR', message: 'El DNI no puede estar vacío.' };
    const dniLimpio = limpiarDNI(dni);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); // <-- Variable 'ss' definida aquí

    // 2. VERIFICAR SI EL DNI YA ESTÁ EN LA HOJA DE "Registros" (esto aplica a todos)
    const hojaRegistro = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
    if (hojaRegistro && hojaRegistro.getLastRow() > 1) {
      const rangoDniRegistro = hojaRegistro.getRange(2, COL_DNI_INSCRIPTO, hojaRegistro.getLastRow() - 1, 1);
      const celdaRegistro = rangoDniRegistro.createTextFinder(dniLimpio).matchEntireCell(true).findNext();
      if (celdaRegistro) {
         // Si ya existe, se devuelve el estado de pago, sin importar el tipo de inscripto seleccionado
         // (SOLUCIÓN AL ERROR): Se pasa 'ss' a la función
         return gestionarUsuarioYaRegistrado(ss, hojaRegistro, celdaRegistro.getRow(), dniLimpio, estado);
      }
    }

    // 3. LÓGICA SEGÚN EL TIPO DE INSCRIPTO
    
    // (Punto 25) VALIDACIÓN CRUZADA: Si eligen 'nuevo' o 'anterior', verificar que NO estén en Preventa.
    if (tipoInscripto === 'nuevo' || tipoInscripto === 'anterior') {
      const hojaPreventa = ss.getSheetByName(NOMBRE_HOJA_PREVENTA);
      if (hojaPreventa && hojaPreventa.getLastRow() > 1) {
        const rangoDniPreventa = hojaPreventa.getRange(2, COL_PREVENTA_DNI, hojaPreventa.getLastRow() - 1, 1);
        const celdaEncontradaPreventa = rangoDniPreventa.createTextFinder(dniLimpio).matchEntireCell(true).findNext();
        
        if (celdaEncontradaPreventa) {
          // ¡Encontrado en Pre-Venta! Bloquear.
          return { status: 'ERROR_TIPO_ANT', message: 'Usted tiene un cupo Pre-Venta. Por favor, elija la opción "Inscripto PRE-VENTA" para validar.' };
        }
      }
      // Si no está en Preventa, la función continúa...
    }
    
    // ========================================================
    // (Punto 25) CASO 1: INSCRIPTO PRE-VENTA
    // ========================================================
    if (tipoInscripto === 'preventa') {
      const hojaPreventa = ss.getSheetByName(NOMBRE_HOJA_PREVENTA);
      if (!hojaPreventa) return { status: 'ERROR', message: `La hoja de configuración "${NOMBRE_HOJA_PREVENTA}" no fue encontrada.` };
      
      const rangoDniPreventa = hojaPreventa.getRange(2, COL_PREVENTA_DNI, hojaPreventa.getLastRow() - 1, 1);
      const celdaEncontrada = rangoDniPreventa.createTextFinder(dniLimpio).matchEntireCell(true).findNext();

      if (!celdaEncontrada) {
        return { status: 'ERROR_TIPO_ANT', message: `El DNI ${dniLimpio} no se encuentra en la base de datos de Pre-Venta. Verifique el DNI o seleccione otro tipo de inscripción.` };
      }

      const fila = hojaPreventa.getRange(celdaEncontrada.getRow(), 1, 1, hojaPreventa.getLastColumn()).getValues()[0];
      
      const jornadaGuarda = String(fila[COL_PREVENTA_GUARDA - 1]).trim().toLowerCase();
      const jornadaPredefinida = (jornadaGuarda === 'si') ? 'Jornada Normal extendida' : 'Jornada Normal';
      
      // Validar cupo para la jornada asignada
      if (jornadaPredefinida === 'Jornada Normal extendida' && estado.jornadaExtendidaAlcanzada) {
        return { status: 'LIMITE_EXTENDIDA', message: 'Su DNI de Pre-Venta corresponde a Jornada Extendida, pero el cupo ya se ha agotado. Por favor, contacte a la administración.' };
      }
      
      const fechaNacimientoRaw = fila[COL_PREVENTA_FECHA_NAC - 1];
      const fechaNacimientoStr = (fechaNacimientoRaw instanceof Date) ? Utilities.formatDate(fechaNacimientoRaw, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : '';

      return {
        status: 'OK_PREVENTA', // Status nuevo para el cliente
        message: '✅ DNI de Pre-Venta validado. Se autocompletarán sus datos. Por favor, complete el resto del formulario.',
        datos: {
          email: fila[COL_PREVENTA_EMAIL - 1],
          nombre: fila[COL_PREVENTA_NOMBRE - 1],
          apellido: fila[COL_PREVENTA_APELLIDO - 1],
          dni: dniLimpio,
          fechaNacimiento: fechaNacimientoStr,
          jornada: jornadaPredefinida,
          esPreventa: true // Flag para que el cliente sepa cómo actuar
        },
        jornadaExtendidaAlcanzada: estado.jornadaExtendidaAlcanzada,
        tipoInscripto: tipoInscripto
      };
    }

    // ========================================================
    // CASO 2: INSCRIPTO NUEVO O ANTERIOR (Lógica existente)
    // ========================================================
    const hojaBusqueda = ss.getSheetByName(NOMBRE_HOJA_BUSQUEDA);
    if (!hojaBusqueda) return { status: 'ERROR', message: `La hoja "${NOMBRE_HOJA_BUSQUEDA}" no fue encontrada.` };

    const rangoDNI = hojaBusqueda.getRange(2, COL_DNI_BUSQUEDA, hojaBusqueda.getLastRow() - 1, 1);
    const celdaEncontrada = rangoDNI.createTextFinder(dniLimpio).matchEntireCell(true).findNext();

    if (celdaEncontrada) { // DNI encontrado en Base de Datos (posible ANTERIOR)
      if (tipoInscripto === 'nuevo') {
        return { status: 'ERROR_TIPO_NUEVO', message: "El DNI se encuentra en nuestra base de datos. Por favor, seleccione 'Soy Inscripto Anterior' y valide nuevamente." };
      }

      const rowIndex = celdaEncontrada.getRow();
      const fila = hojaBusqueda.getRange(rowIndex, COL_HABILITADO_BUSQUEDA, 1, 10).getValues()[0];
      const habilitado = fila[0];
      if (habilitado !== true) {
        return { status: 'NO_HABILITADO', message: 'El DNI se encuentra en la base de datos, pero no está habilitado para la inscripción. Por favor, consulte con la organización.' };
      }

      const nombre = fila[1];
      const apellido = fila[2];
      const fechaNacimientoRaw = fila[3];
      const obraSocial = String(fila[6] || '').trim();
      const colegioJardin = String(fila[7] || '').trim();
      const responsable = String(fila[8] || '').trim();
      // const telefono = String(fila[9] || '').trim(); // <-- (MODIFICACIÓN: Esta línea ya no se usa)
      const fechaNacimientoStr = (fechaNacimientoRaw instanceof Date) ? Utilities.formatDate(fechaNacimientoRaw, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : '';

      return {
        status: 'OK',
        datos: {
          nombre: nombre,
          apellido: apellido,
          dni: dniLimpio,
          fechaNacimiento: fechaNacimientoStr,
          obraSocial: obraSocial,
          colegioJardin: colegioJardin,
          adultoResponsable1: responsable,
          // (MODIFICACIÓN: Teléfono omitido según solicitud)
          // telResponsable1: telefono,
          esPreventa: false // Flag
        },
        edad: calcularEdad(fechaNacimientoStr),
        jornadaExtendidaAlcanzada: estado.jornadaExtendidaAlcanzada,
        tipoInscripto: tipoInscripto
      };

    } else { // DNI NO encontrado en Base de Datos (posible NUEVO)
      if (tipoInscripto === 'anterior') {
        return { status: 'ERROR_TIPO_ANT', message: "No se encuentra en la base de datos de años anteriores. Por favor, seleccione 'Soy Nuevo Inscripto'." };
      }
      if (tipoInscripto === 'preventa') {
         // Este caso ya fue manejado arriba, pero por si acaso.
         return { status: 'ERROR_TIPO_ANT', message: `El DNI ${dniLimpio} no se encuentra en la base de datos de Pre-Venta.` };
      }
      return {
        status: 'OK_NUEVO',
        message: '✅ DNI validado. Proceda al registro.',
        jornadaExtendidaAlcanzada: estado.jornadaExtendidaAlcanzada,
        tipoInscripto: tipoInscripto,
        datos: { dni: dniLimpio, esPreventa: false }
      };
    }

  } catch (e) {
    Logger.log("Error en validarAcceso: " + e.message + " Stack: " + e.stack);
    return { status: 'ERROR', message: 'Ocurrió un error al validar el DNI. ' + e.message };
  }
}



function gestionarUsuarioYaRegistrado(ss, hojaRegistro, filaRegistro, dniLimpio, estado) { // <-- (SOLUCIÓN) Acepta 'ss'
  const rangoFila = hojaRegistro.getRange(filaRegistro, 1, 1, hojaRegistro.getLastColumn()).getValues()[0];

  const estadoPago = rangoFila[COL_ESTADO_PAGO - 1];
  const metodoPago = rangoFila[COL_METODO_PAGO - 1];
  const nombreRegistrado = rangoFila[COL_NOMBRE - 1] + ' ' + rangoFila[COL_APELLIDO - 1];
  const estadoInscripto = rangoFila[COL_ESTADO_NUEVO_ANT - 1];
  
  // Lógica de HERMANO_COMPLETAR
  if (estadoInscripto === 'Nuevo Hermano/a' || estadoInscripto === 'Anterior Hermano/a') {
      let faltantes = [];
      // (MODIFICACIÓN) Se revisan los campos que el hermano DEBE completar
      if (!rangoFila[COL_COLEGIO_JARDIN - 1]) faltantes.push('Colegio / Jardín');
      if (!rangoFila[COL_PRACTICA_DEPORTE - 1]) faltantes.push('Practica Deporte');
      if (!rangoFila[COL_TIENE_ENFERMEDAD - 1]) faltantes.push('Enfermedad Preexistente');
      if (!rangoFila[COL_ES_ALERGICO - 1]) faltantes.push('Alergias');
      if (!rangoFila[COL_FOTO_CARNET - 1]) faltantes.push('Foto Carnet 4x4');
      if (!rangoFila[COL_JORNADA - 1]) faltantes.push('Jornada');
      if (!rangoFila[COL_SOCIO - 1]) faltantes.push('Es Socio'); // (PUNTO 27) Añadido
      if (!rangoFila[COL_METODO_PAGO - 1]) faltantes.push('Método de Pago');

      // (MODIFICACIÓN) Se revisan los campos que DEBERÍAN estar pre-rellenados
      // Con la corrección en 'registrarDatos', 'COL_EMAIL' ahora debería estar lleno.
      if (!rangoFila[COL_EMAIL - 1]) faltantes.push('Email'); 
      if (!rangoFila[COL_ADULTO_RESPONSABLE_1 - 1]) faltantes.push('Responsable 1');
      if (!rangoFila[COL_PERSONAS_AUTORIZADAS - 1]) faltantes.push('Personas Autorizadas');
      
      const datos = {
          dni: dniLimpio,
          nombre: rangoFila[COL_NOMBRE - 1],
          apellido: rangoFila[COL_APELLIDO - 1],
          // (SOLUCIÓN AL ERROR) Usa 'ss.getSpreadsheetTimeZone()' en lugar de 'SpreadsheetApp.getActive()'
          fechaNacimiento: rangoFila[COL_FECHA_NACIMIENTO_REGISTRO - 1] ? Utilities.formatDate(new Date(rangoFila[COL_FECHA_NACIMIENTO_REGISTRO - 1]), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : '',
          
          // (MODIFICACIÓN) Estos campos ahora SÍ vendrán de la hoja
          email: rangoFila[COL_EMAIL - 1] || '',
          adultoResponsable1: rangoFila[COL_ADULTO_RESPONSABLE_1 - 1] || '',
          dniResponsable1: rangoFila[COL_DNI_RESPONSABLE_1 - 1] || '',
          telResponsable1: rangoFila[COL_TEL_RESPONSABLE_1 - 1] || '',
          adultoResponsable2: rangoFila[COL_ADULTO_RESPONSABLE_2 - 1] || '',
          telResponsable2: rangoFila[COL_TEL_RESPONSABLE_2 - 1] || '',
          personasAutorizadas: rangoFila[COL_PERSONAS_AUTORIZADAS - 1] || '',
          obraSocial: rangoFila[COL_OBRA_SOCIAL - 1] || '',
          
          // Este campo (Colegio) se mantiene vacío a propósito (según tu lógica)
          colegioJardin: rangoFila[COL_COLEGIO_JARDIN - 1] || ''
      };
      
      // Ya sea que falten campos por llenar o no (ej. si vuelve a validar),
      // se le muestra el formulario para completar/revisar.
      return {
          status: 'HERMANO_COMPLETAR',
          message: `⚠️ ¡Hola ${datos.nombre}! Eres un hermano/a pre-registrado.\n` +
            `Por favor, complete/verifique TODOS los campos del formulario para obtener el cupo definitivo y el link para pagar.\n` +
            (faltantes.length > 0 ? `Campos requeridos faltantes detectados: <strong>${faltantes.join(', ')}</strong>.` : 'Todos los campos parecen estar listos para verificar.'),
          datos: datos,
          jornadaExtendidaAlcanzada: estado.jornadaExtendidaAlcanzada,
          tipoInscripto: estadoInscripto
      };
  }

  // Lógica de REGISTRO_ENCONTRADO (Repago, Subir comprobante, etc.)
  const aptitudFisica = rangoFila[COL_APTITUD_FISICA - 1];
  const adeudaAptitud = !aptitudFisica;
  const cantidadCuotasRegistrada = parseInt(rangoFila[COL_CANTIDAD_CUOTAS - 1]) || 0;
  let proximaCuotaPendiente = null;
  
  // --- (MODIFICACIÓN CRÍTICA) ---
  // Se cambia 'estadoPago === 'Pagado'' por 'String(estadoPago).includes('Pagado')'
  // Esto captura "Pagado", " Pagado", "Pagado." etc. y previene el repago.
  if (String(estadoPago).includes('Pagado')) {
      return { 
        status: 'REGISTRO_ENCONTRADO', 
        message: `✅ El DNI ${dniLimpio} (${nombreRegistrado}) ya se encuentra REGISTRADO y la inscripción está PAGADA.`, 
        adeudaAptitud: adeudaAptitud,
        cantidadCuotas: cantidadCuotasRegistrada,
        metodoPago: metodoPago,
        proximaCuotaPendiente: null
      };
  }

  // --- (NUEVA MODIFICACIÓN) ---
  // Si está "En revisión" (porque ya subió un comprobante), mostrar solo el mensaje.
  // No intentar generar link de pago ni mostrar el uploader.
  if (String(estadoPago).includes('En revisión')) {
      return {
          status: 'REGISTRO_ENCONTRADO',
          message: `⚠️ El DNI ${dniLimpio} (${nombreRegistrado}) ya se encuentra REGISTRADO. Su pago está "En revisión".`,
          adeudaAptitud: adeudaAptitud,
          cantidadCuotas: cantidadCuotasRegistrada,
          metodoPago: metodoPago,
          proximaCuotaPendiente: null
      };
  }
  
  if (String(metodoPago).includes('Efectivo') || String(metodoPago).includes('Transferencia')) {
      return {
          status: 'REGISTRO_ENCONTRADO',
          message: `⚠️ El DNI ${dniLimpio} (${nombreRegistrado}) ya se encuentra REGISTRADO. El pago (${metodoPago}) está PENDIENTE.`,
          adeudaAptitud: adeudaAptitud,
          cantidadCuotas: cantidadCuotasRegistrada,
          metodoPago: metodoPago,
          proximaCuotaPendiente: null
      };
  }

  // Lógica de repago...
  try {
      const datosParaPago = {
          dni: dniLimpio,
          apellidoNombre: nombreRegistrado,
          email: rangoFila[COL_EMAIL - 1],
          metodoPago: metodoPago,
          jornada: rangoFila[COL_JORNADA - 1]
      };
      let identificadorPago = null;
      if (metodoPago === 'Pago en Cuotas') {
        for (let i = 1; i <= cantidadCuotasRegistrada; i++) {
          let colCuota = i === 1 ? COL_CUOTA_1 : (i === 2 ? COL_CUOTA_2 : COL_CUOTA_3);
          let cuota_status = rangoFila[colCuota - 1];
          if (!cuota_status || (!cuota_status.toString().includes("Pagada") && !cuota_status.toString().includes("Notificada"))) {
            identificadorPago = `C${i}`;
            proximaCuotaPendiente = identificadorPago;
            break;
          }
        }
        if (identificadorPago == null) {
          // Si todas las cuotas están pagadas, pero el estado general (AI) no es "Pagado"
          // (porque quizás falta aptitud, etc.), igual lo tratamos como pagado.
          return {
            status: 'REGISTRO_ENCONTRADO',
            message: `✅ El DNI ${dniLimpio} (${nombreRegistrado}) ya completó todas las cuotas.`,
            adeudaAptitud: adeudaAptitud,
            cantidadCuotas: cantidadCuotasRegistrada,
            metodoPago: metodoPago,
            proximaCuotaPendiente: null
          };
        }
      }
      
      const init_point = crearPreferenciaDePago(datosParaPago, identificadorPago, cantidadCuotasRegistrada);
      
      if (!init_point || !init_point.toString().startsWith('http')) {
        return {
          status: 'REGISTRO_ENCONTRADO',
          message: `⚠️ Error al generar link: ${init_point}`,
          adeudaAptitud: adeudaAptitud,
          cantidadCuotas: cantidadCuotasRegistrada,
          metodoPago: metodoPago,
          proximaCuotaPendiente: proximaCuotaPendiente,
          error_init_point: init_point
        };
      }
      
      return {
          status: 'REGISTRO_ENCONTRADO',
          message: `⚠️ El DNI ${dniLimpio} (${nombreRegistrado}) ya se encuentra REGISTRADO. Se generó un link para la próxima cuota pendiente (${identificadorPago || 'Pago Total'}).`,
          init_point: init_point,
          adeudaAptitud: adeudaAptitud,
          cantidadCuotas: cantidadCuotasRegistrada,
          metodoPago: metodoPago,
          proximaCuotaPendiente: proximaCuotaPendiente
      };
  } catch (e) {
      Logger.log(`Error al generar link de repago para DNI ${dniLimpio}: ${e.message}`);
      return {
          status: 'REGISTRO_ENCONTRADO',
          message: `⚠️ El DNI ${dniLimpio} (${nombreRegistrado}) ya está REGISTRADO. Pago PENDIENTE, pero error al generar link: ${e.message}`,
          adeudaAptitud: adeudaAptitud,
          cantidadCuotas: cantidadCuotasRegistrada,
          metodoPago: metodoPago,
          proximaCuotaPendiente: proximaCuotaPendiente,
          error_init_point: e.message
      };
  }
}



// =========================================================
// (Punto 10) enviarEmailConfirmacion (ACTUALIZADO)
// =========================================================
/**
* (M) FUNCIÓN DE EMAIL (SIMPLIFICADA)
* (Punto 10) Agrega template para "Transferencia"
*/
function enviarEmailConfirmacion(datos, numeroDeTurno, init_point = null, overrideMetodo = null) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaConfig = ss.getSheetByName(NOMBRE_HOJA_CONFIG);

    if (!hojaConfig || !datos.email || hojaConfig.getRange('B8').getValue() !== true) {
      Logger.log("Envío de email deshabilitado o sin email.");
      return;
    }

    let asunto = "";
    let cuerpoOriginal = "";
    let cuerpoFinal = "";
    const metodo = overrideMetodo || datos.metodoPago;

    // (Punto 2) Usar nombre y apellido
    const nombreCompleto = `${datos.nombre} ${datos.apellido}`;

    if (metodo === 'Pago 1 Cuota Deb/Cred MP(Total)') {
      asunto = hojaConfig.getRange('E2:G2').getValue();
      cuerpoOriginal = hojaConfig.getRange('D4:H8').getValue();
      if (!asunto) asunto = "Confirmación de Registro (Pago Total)";
      if (!cuerpoOriginal) cuerpoOriginal = "Su cupo ha sido reservado.\n\nInscripto: {{nombreCompleto}}\nTurno: {{numeroDeTurno}}\n\nLink de Pago: {{linkDePago}}";

      cuerpoFinal = cuerpoOriginal
        .replace(/{{nombreCompleto}}/g, nombreCompleto)
        .replace(/{{numeroDeTurno}}/g, numeroDeTurno)
        .replace(/{{linkDePago}}/g, init_point || 'N/A');

    } else if (metodo === 'Pago Efectivo (Adm del Club)' || metodo === 'registro_sin_pago') {
      asunto = hojaConfig.getRange('E13:H13').getValue();
      cuerpoOriginal = hojaConfig.getRange('D15:H19').getValue();
      if (!asunto) asunto = "Confirmación de Registro (Pago Efectivo)";
      if (!cuerpoOriginal) cuerpoOriginal = "Su cupo ha sido reservado.\n\nInscripto: {{nombreCompleto}}\nTurno: {{numeroDeTurno}}\n\nPor favor, acérquese a la administración.";

      cuerpoFinal = cuerpoOriginal
        .replace(/{{nombreCompleto}}/g, nombreCompleto)
        .replace(/{{numeroDeTurno}}/g, numeroDeTurno);

      // (Punto 10) NUEVO CASO PARA TRANSFERENCIA
    } else if (metodo === 'Transferencia') {
      asunto = "Confirmación de Registro (Transferencia)"; // Asunto genérico (o agregar a Config)
      cuerpoOriginal = "Su cupo ha sido reservado.\n\nInscripto: {{nombreCompleto}}\nTurno: {{numeroDeTurno}}\n\n" +
        "Por favor, realice la transferencia a:\n" +
        "TITULAR DE LA CUENTA: Walter Jonas Marrello\n" +
        "Alias: clubhipicomendoza\n\n" +
        "IMPORTANTE: Una vez realizada, vuelva a ingresar al formulario con su DNI para subir el comprobante.";

      cuerpoFinal = cuerpoOriginal
        .replace(/{{nombreCompleto}}/g, nombreCompleto)
        .replace(/{{numeroDeTurno}}/g, numeroDeTurno);

    } else if (metodo === 'Pago en Cuotas') {
      asunto = hojaConfig.getRange('E24:G24').getValue();
      cuerpoOriginal = hojaConfig.getRange('D26:H30').getValue();
      if (!asunto) asunto = "Confirmación de Registro (Cuotas)";
      if (!cuerpoOriginal) cuerpoOriginal = "Su cupo ha sido reservado.\n\nInscripto: {{nombreCompleto}}\nTurno: {{numeroDeTurno}}\n\nLink Cuota 1: {{linkCuota1}}\nLink Cuota 2: {{linkCuota2}}\nLink Cuota 3: {{linkCuota3}}";

      cuerpoFinal = cuerpoOriginal
        .replace(/{{nombreCompleto}}/g, nombreCompleto)
        .replace(/{{numeroDeTurno}}/g, numeroDeTurno)
        .replace(/{{linkCuota1}}/g, init_point.link1 || 'Error al generar')
        .replace(/{{linkCuota2}}/g, init_point.link2 || 'Error al generar')
        .replace(/{{linkCuota3}}/g, init_point.link3 || 'Error al generar');

    } else {
      Logger.log(`Método de pago "${datos.metodoPago}" no reconocido para email.`);
      return;
    }

    // (Punto 29) DESACTIVADO
    /*
    MailApp.sendEmail({
      to: datos.email,
      subject: `${asunto} (Turno #${numeroDeTurno})`,
      body: cuerpoFinal
    });

    Logger.log(`Correo enviado a ${datos.email} por ${datos.metodoPago}.`);
    */
    Logger.log(`(Punto 29) Envío de email automático a ${datos.email} DESACTIVADO.`);


  } catch (e) {
    Logger.log("Error al enviar correo (enviarEmailConfirmacion): " + e.message);
  }
}

// =========================================================
// (Punto 15, 17, 19) subirComprobanteManual (DUPLICADO - ELIMINAR ESTA COPIA)
// =========================================================
// (Esta función está duplicada en el archivo original, la dejo como estaba)
function subirComprobanteManual(dni, fileData, tipoComprobante, datosExtras) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const dniLimpio = limpiarDNI(dni);
    if (!dniLimpio || !fileData || !tipoComprobante) {
      return { status: 'ERROR', message: 'Faltan datos (DNI, archivo o tipo de comprobante).' };
    }
    
    // Validar que los datos del pagador (datosExtras) llegaron
    if (!datosExtras || !datosExtras.nombrePagador || !datosExtras.dniPagador) {
        return { status: 'ERROR', message: 'Faltan los datos del adulto pagador (Nombre o DNI).' };
    }

    // Usar el DNI del inscripto para la carpeta
    const fileUrl = uploadFileToDrive(fileData.data, fileData.mimeType, fileData.fileName, dniLimpio, 'comprobante');
    if (typeof fileUrl !== 'string' || !fileUrl.startsWith('http')) {
      throw new Error("Error al subir el archivo a Drive: " + (fileUrl.message || 'Error desconocido'));
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
    if (!hoja) throw new Error(`La hoja "${NOMBRE_HOJA_REGISTRO}" no fue encontrada.`);

    const rangoDni = hoja.getRange(2, COL_DNI_INSCRIPTO, hoja.getLastRow() - 1, 1);
    const celdaEncontrada = rangoDni.createTextFinder(dniLimpio).matchEntireCell(true).findNext();

    if (celdaEncontrada) {
      const fila = celdaEncontrada.getRow();
      let columnaDestinoArchivo; // Columna para el LINK

      // --- ¡NUEVA LÓGICA! ---
      // 1. Guardar los datos del pagador en sus columnas (AN y AO - DESPLAZADAS)
      hoja.getRange(fila, COL_PAGADOR_NOMBRE_MANUAL).setValue(datosExtras.nombrePagador); // AN (40)
      hoja.getRange(fila, COL_PAGADOR_DNI_MANUAL).setValue(datosExtras.dniPagador); // AO (41)

      // 2. Asignar la columna de destino para el LINK DEL ARCHIVO (COLUMNAS DESPLAZADAS)
      switch (tipoComprobante) {
        case 'total_mp':
        case 'mp_total': // Alias de Index.html
        case 'externo':
          columnaDestinoArchivo = COL_COMPROBANTE_MANUAL_TOTAL_EXT; // AQ (43)
          break;
        case 'cuota1_mp':
        case 'mp_cuota_1': // Alias de Index.html
          columnaDestinoArchivo = COL_COMPROBANTE_MANUAL_CUOTA1; // AR (44)
          break;
        case 'cuota2_mp':
        case 'mp_cuota_2': // Alias de Index.html
          columnaDestinoArchivo = COL_COMPROBANTE_MANUAL_CUOTA2; // AS (45)
          break;
        case 'cuota3_mp':
        case 'mp_cuota_3': // Alias de Index.html
          columnaDestinoArchivo = COL_COMPROBANTE_MANUAL_CUOTA3; // AT (46)
          break;
        default:
          throw new Error(`Tipo de comprobante no reconocido: ${tipoComprobante}`);
      }

      // 3. Guardar SOLO el link en la columna de archivo correspondiente
      hoja.getRange(fila, columnaDestinoArchivo).setValue(fileUrl);
      
      // 4. Actualizar estado
      hoja.getRange(fila, COL_ESTADO_PAGO).setValue("En revisión");

      Logger.log(`Comprobante manual [${tipoComprobante}] subido para DNI ${dniLimpio} en fila ${fila}.`);
      return { status: 'OK', message: '¡Comprobante subido con éxito! Será revisado por la administración.' };
    } else {
      Logger.log(`No se encontró DNI ${dniLimpio} para subir comprobante manual.`);
      return { status: 'ERROR', message: `No se encontró el registro para el DNI ${dniLimpio}. Asegúrese de que el DNI del inscripto sea correcto.` };
    }

  } catch (e) {
    Logger.log("Error en subirComprobanteManual: " + e.toString());
    return { status: 'ERROR', message: 'Error en el servidor: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/* */
function subirAptitudManual(dni, fileData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const dniLimpio = limpiarDNI(dni);
    if (!dniLimpio || !fileData) {
      return { status: 'ERROR', message: 'Faltan datos (DNI o archivo).' };
    }

    const fileUrl = uploadFileToDrive(fileData.data, fileData.mimeType, fileData.fileName, dniLimpio, 'ficha');
    if (typeof fileUrl !== 'string' || !fileUrl.startsWith('http')) {
      throw new Error("Error al subir el archivo a Drive.");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
    if (!hoja) throw new Error(`La hoja "${NOMBRE_HOJA_REGISTRO}" no fue encontrada.`);

    const rangoDni = hoja.getRange(2, COL_DNI_INSCRIPTO, hoja.getLastRow() - 1, 1);
    const celdaEncontrada = rangoDni.createTextFinder(dniLimpio).matchEntireCell(true).findNext();

    if (celdaEncontrada) {
      const fila = celdaEncontrada.getRow();
      hoja.getRange(fila, COL_APTITUD_FISICA).setValue(fileUrl);

      Logger.log(`Aptitud Física subida para DNI ${dniLimpio} en fila ${fila}.`);
      return { status: 'OK', message: '¡Certificado de Aptitud subido con éxito!' };
    } else {
      Logger.log(`No se encontró DNI ${dniLimpio} para subir aptitud física.`);
      return { status: 'ERROR', message: `No se encontró el registro para el DNI ${dniLimpio}.` };
    }

  } catch (e) {
    Logger.log("Error en subirAptitudManual: " + e.toString());
    return { status: 'ERROR', message: 'Error en el servidor: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/* */
function sincronizarRegistros() {
  Logger.log("sincronizarRegistros: Función omitida.");
  return;
}

/* */
function subirArchivoIndividual(fileData, dni, tipoArchivo) {
  try {
    if (!fileData || !dni || !tipoArchivo) {
      return { status: 'ERROR', message: 'Faltan datos para la subida (DNI, archivo o tipo).' };
    }

    const dniLimpio = limpiarDNI(dni);

    const fileUrl = uploadFileToDrive(
      fileData.data,
      fileData.mimeType,
      fileData.fileName,
      dniLimpio,
      tipoArchivo
    );

    if (typeof fileUrl === 'object' && fileUrl.status === 'ERROR') {
      return fileUrl;
    }

    return { status: 'OK', url: fileUrl };

  } catch (e) {
    Logger.log("Error en subirArchivoIndividual: " + e.toString());
    return { status: 'ERROR', message: 'Error del servidor al subir: ' + e.message };
  }
}