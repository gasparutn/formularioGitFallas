/* */
// --- CONFIGURACIÓN DE MERCADO PAGO ---
/* */
const MERCADO_PAGO_ACCESS_TOKEN = 'APP_USR-7838602473992019-102318-99f752417a8ccd21a0d2eba48126da4d-2940898374'; // 🛑 Token V13
const MP_API_URL = 'https://api.mercadopago.com/checkout/preferences';

// =========================================================
// Las constantes (COL_EMAIL, COL_ESTADO_PAGO, etc.) se leen
// automáticamente desde el archivo 'Constantes.gs'.
// =========================================================

/**
* (PASO 1)
* (Punto 10) Añadida lógica para "Transferencia"
* (Punto 28) Lógica de "Pago en Cuotas" ajustada para "Pago en 3 Cuotas"
*/
function paso1_registrarRegistro(datos) {
  Logger.log("PASO 1 INICIADO. Datos recibidos: " + JSON.stringify(datos));
  try {
    if (!datos.urlFotoCarnet && !datos.esHermanoCompletando) { // (Punto 6) Los hermanos no suben foto en el registro inicial
      Logger.log("Error: El formulario se envió sin la URL de la Foto Carnet.");
      return { status: 'ERROR', message: 'Falta la Foto Carnet. Por favor, asegúrese de que el archivo se haya subido correctamente.' };
    }

    // (Punto 10) Nuevos estados de pago
    if (datos.metodoPago === 'Pago Efectivo (Adm del Club)') {
      datos.estadoPago = "Pendiente (Efectivo)";
    } else if (datos.metodoPago === 'Transferencia') {
      datos.estadoPago = "Pendiente (Transferencia)"; // NUEVO
    // (Punto 28) Ajuste de "Pago en Cuotas" a "Pago en 3 Cuotas" (o mantener "Pago en Cuotas" si el valor enviado no cambió)
    } else if (datos.metodoPago === 'Pago en Cuotas') { 
      datos.estadoPago = `Pendiente (${datos.cantidadCuotas} Cuotas)`; // (datos.cantidadCuotas será 3)
    } else { // 'Pago 1 Cuota Deb/Cred MP(Total)'
      datos.estadoPago = "Pendiente";
    }

    // (Punto 12) Si es un hermano completando, llamamos a una función diferente
    if (datos.esHermanoCompletando === true) {
       const respuestaUpdate = actualizarDatosHermano(datos);
       return respuestaUpdate;
    } else {
       // Si es registro normal, llamamos a registrarDatos (que ahora maneja hermanos)
       const respuestaRegistro = registrarDatos(datos); // registrarDatos() vive en codigo.gs
       return respuestaRegistro;
    }

  } catch (e) {
    Logger.log("Error en paso1_registrarRegistro: " + e.message);
    return { status: 'ERROR', message: 'Error general en el servidor (Paso 1): ' + e.message };
  }
}

// =========================================================
// (NUEVA FUNCIÓN HELPER para solucionar error de 'hermano')
// =========================================================
/**
 * Obtiene el precio y el monto a pagar desde la hoja de Config.
 * @param {string} metodoPago - El método de pago seleccionado.
 * @param {string|number} cantidadCuotasStr - La cantidad de cuotas (ej. "3").
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hojaConfig - La hoja de "Config".
 * @returns {{precio: number, montoAPagar: number}}
 */
function obtenerPrecioDesdeConfig(metodoPago, cantidadCuotasStr, hojaConfig) {
  let precio = 0;
  let montoAPagar = 0;
  try {
    const precioCuota = hojaConfig.getRange("B20").getValue();
    const precioTotal = hojaConfig.getRange("B14").getValue();

    if (metodoPago === 'Pago en Cuotas') {
      precio = precioCuota;
      montoAPagar = precio * (parseInt(cantidadCuotasStr) || 0);
    } else if (metodoPago === 'Pago 1 Cuota Deb/Cred MP(Total)') {
      precio = precioTotal;
      montoAPagar = precio;
    } else if (metodoPago === 'Pago Efectivo (Adm del Club)' || metodoPago === 'Transferencia') {
      precio = precioTotal; 
      montoAPagar = precio;
    }

    // Fallbacks
    if (precio === 0 && precioTotal > 0) {
      precio = precioTotal;
    }
    if (montoAPagar === 0 && precio > 0 && (metodoPago === 'Pago Efectivo (Adm del Club)' || metodoPago === 'Transferencia')) {
      montoAPagar = precio;
    }
    
    return { precio, montoAPagar };

  } catch (e) {
    Logger.log("Error en obtenerPrecioDesdeConfig: " + e.message);
    return { precio: 0, montoAPagar: 0 };
  }
}


/**
* (Punto 6, 12, 27) NUEVA FUNCIÓN para actualizar un hermano (ACTUALIZADA)
*/
function actualizarDatosHermano(datos) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(60000);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaRegistro = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
    const hojaConfig = ss.getSheetByName(NOMBRE_HOJA_CONFIG);
    const dniBuscado = limpiarDNI(datos.dni);
    
    if (!hojaRegistro) throw new Error("Hoja de Registros no encontrada");

    const rangoDni = hojaRegistro.getRange(2, COL_DNI_INSCRIPTO, hojaRegistro.getLastRow() - 1, 1);
    const celdaEncontrada = rangoDni.createTextFinder(dniBuscado).matchEntireCell(true).findNext();
    
    if (!celdaEncontrada) {
      return { status: 'ERROR', message: 'No se encontró el registro del hermano para actualizar.' };
    }
    
    const fila = celdaEncontrada.getRow();
    
    // --- CÁLCULO DE PRECIOS ---
    // (MODIFICACIÓN) Llamada a la nueva función helper para solucionar el error 'obtenerPrecioDesdeConfig is not defined'
    const { precio, montoAPagar } = obtenerPrecioDesdeConfig(datos.metodoPago, datos.cantidadCuotas, hojaConfig);

    const telResp1 = `(${datos.telAreaResp1}) ${datos.telNumResp1}`;
    const telResp2 = (datos.telAreaResp2 && datos.telNumResp2) ? `(${datos.telAreaResp2}) ${datos.telNumResp2}` : '';
    
    // --- (MODIFICACIÓN) ---
    // Reemplazada la lógica de 'E'/'N' por la lógica completa.
    const esPreventa = (datos.tipoInscripto === 'preventa');
    let marcaNE = "";
    if (datos.jornada === 'Jornada Normal extendida') {
      marcaNE = esPreventa ? "Extendida (Pre-venta)" : "Extendida";
    } else { // Asume "Jornada Normal"
      marcaNE = esPreventa ? "Normal (Pre-Venta)" : "Normal";
    }
    // --- (FIN MODIFICACIÓN) ---

    
    // (Punto 6, 27) Actualizar la fila del hermano con los datos completos
    hojaRegistro.getRange(fila, COL_MARCA_N_E_A).setValue(marcaNE);
    hojaRegistro.getRange(fila, COL_EMAIL).setValue(datos.email);
    hojaRegistro.getRange(fila, COL_OBRA_SOCIAL).setValue(datos.obraSocial);
    hojaRegistro.getRange(fila, COL_COLEGIO_JARDIN).setValue(datos.colegioJardin);
    hojaRegistro.getRange(fila, COL_ADULTO_RESPONSABLE_1).setValue(datos.adultoResponsable1);
    hojaRegistro.getRange(fila, COL_DNI_RESPONSABLE_1).setValue(datos.dniResponsable1);
    hojaRegistro.getRange(fila, COL_TEL_RESPONSABLE_1).setValue(telResp1);
    hojaRegistro.getRange(fila, COL_ADULTO_RESPONSABLE_2).setValue(datos.adultoResponsable2);
    hojaRegistro.getRange(fila, COL_TEL_RESPONSABLE_2).setValue(telResp2);
    hojaRegistro.getRange(fila, COL_PERSONAS_AUTORIZADAS).setValue(datos.personasAutorizadas);
    hojaRegistro.getRange(fila, COL_PRACTICA_DEPORTE).setValue(datos.practicaDeporte);
    hojaRegistro.getRange(fila, COL_ESPECIFIQUE_DEPORTE).setValue(datos.especifiqueDeporte);
    hojaRegistro.getRange(fila, COL_TIENE_ENFERMEDAD).setValue(datos.tieneEnfermedad);
    hojaRegistro.getRange(fila, COL_ESPECIFIQUE_ENFERMEDAD).setValue(datos.especifiqueEnfermedad);
    hojaRegistro.getRange(fila, COL_ES_ALERGICO).setValue(datos.esAlergico);
    hojaRegistro.getRange(fila, COL_ESPECIFIQUE_ALERGIA).setValue(datos.especifiqueAlergia);
    hojaRegistro.getRange(fila, COL_APTITUD_FISICA).setValue(datos.urlCertificadoAptitud || '');
    hojaRegistro.getRange(fila, COL_FOTO_CARNET).setValue(datos.urlFotoCarnet || '');
    hojaRegistro.getRange(fila, COL_JORNADA).setValue(datos.jornada);
    hojaRegistro.getRange(fila, COL_SOCIO).setValue(datos.esSocio); // (PUNTO 27) NUEVA LÍNEA
    hojaRegistro.getRange(fila, COL_METODO_PAGO).setValue(datos.metodoPago);
    hojaRegistro.getRange(fila, COL_PRECIO).setValue(precio);
    hojaRegistro.getRange(fila, COL_CANTIDAD_CUOTAS).setValue(parseInt(datos.cantidadCuotas) || 0); // (será 3)
    hojaRegistro.getRange(fila, COL_ESTADO_PAGO).setValue(datos.estadoPago);
    hojaRegistro.getRange(fila, COL_MONTO_A_PAGAR).setValue(montoAPagar);

    SpreadsheetApp.flush();
    
    // (Punto 2) Necesita nombre/apellido para el email
    datos.nombre = hojaRegistro.getRange(fila, COL_NOMBRE).getValue();
    datos.apellido = hojaRegistro.getRange(fila, COL_APELLIDO).getValue();
    
    return { status: 'OK_REGISTRO', message: '¡Registro de Hermano Actualizado!', numeroDeTurno: hojaRegistro.getRange(fila, COL_NUMERO_TURNO).getValue(), datos: datos };

  } catch (e) {
    Logger.log("Error en actualizarDatosHermano: " + e.message);
    return { status: 'ERROR', message: 'Error general en el servidor (Actualizar Hermano): ' + e.message };
  } finally {
    lock.releaseLock();
  }
}


/**
* (PASO 2)
* (Punto 10) Añadida lógica para "Transferencia"
* (Punto 29) Emails automáticos desactivados
*/
function paso2_crearPagoYEmail(datos, numeroDeTurno) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaConfig = ss.getSheetByName(NOMBRE_HOJA_CONFIG);
    const pagosHabilitados = hojaConfig.getRange('B23').getValue();

    if (pagosHabilitados === false) {
      Logger.log(`Pagos deshabilitados (Config B23). Registrando sin link!!`);
      // (Punto 29) Email automático desactivado.
      // enviarEmailConfirmacion(datos, numeroDeTurno, null, 'registro_sin_pago');
      return { status: 'OK_REGISTRO_SIN_PAGO', message: `¡Inscripción registrada!! Los pagos online están momentáneamente desactivados.\nPor favor, acérquese a la administración.` };
    }

    // (Punto 10) Manejar Efectivo y Transferencia
    if (datos.metodoPago === 'Pago Efectivo (Adm del Club)' || datos.metodoPago === 'Transferencia') {
      // (Punto 29) Email automático desactivado.
      // enviarEmailConfirmacion(datos, numeroDeTurno); // Enviar email (Efectivo o Transferencia)
      let message = (datos.metodoPago === 'Transferencia') ? 
          '¡Registro exitoso! Por favor, realice la transferencia y luego suba el comprobante.' :
          '¡Registro exitoso! Por favor, acérquese a la administración para completar el pago.';
      return { status: 'OK_EFECTIVO', message: message }; // Reutiliza OK_EFECTIVO
    }

    if (datos.metodoPago === 'Pago 1 Cuota Deb/Cred MP(Total)') {
      let init_point;
      try {
        // (Punto 2) Pasar nombre/apellido
        init_point = crearPreferenciaDePago(datos, null); 

        if (!init_point || !init_point.startsWith('http')) {
          return { status: 'OK_REGISTRO_SIN_LINK', message: init_point };
        }
      } catch (e) {
        Logger.log("Error al crear preferencia de pago (total): " + e.message);
        return { status: 'OK_REGISTRO_SIN_LINK', message: `¡Tu registro se guardó!! Pero falló la creación del link de pago.\nPor favor, contacte a la administración para abonar.` };
      }

      if (datos.email && init_point) {
        // (Punto 29) Email automático desactivado.
        // enviarEmailConfirmacion(datos, numeroDeTurno, init_point);
      }
      return { status: 'OK_PAGO', init_point: init_point };
    }

    // (Punto 28) Ajustado
    if (datos.metodoPago === 'Pago en Cuotas') {
      const cantidadCuotas = parseInt(datos.cantidadCuotas); // (Será 3)
      const emailLinks = {}; 

      try {
        // (Punto 10) Lógica de 2 o 3 cuotas
        const cuotasDisponibles = (cantidadCuotas === 2) ? [1, 2] : [1, 2, 3]; // (Será 3)
        
        for (let i = 1; i <= 3; i++) {
            if (cuotasDisponibles.includes(i)) {
                const link = crearPreferenciaDePago(datos, `C${i}`, cantidadCuotas);
                emailLinks[`link${i}`] = link;
            } else {
                emailLinks[`link${i}`] = 'N/A (No aplica)';
            }
        }

      } catch (e) {
        Logger.log("Error al crear preferencias de pago (cuotas): " + e.message);
        return { status: 'OK_REGISTRO_SIN_LINK', message: `¡Tu registro se guardó!! Pero falló la creación de los links de pago.\nPor favor, contacte a la administración.` };
      }

      if (datos.email) {
        // (Punto 29) Email automático desactivado.
        // enviarEmailConfirmacion(datos, numeroDeTurno, emailLinks);
      }

      const primerLink = emailLinks.link1;
      if (!primerLink || !primerLink.startsWith('http')) {
        return { status: 'OK_REGISTRO_SIN_LINK', message: `¡Registro guardado!! ${primerLink}` };
      }
      return { status: 'OK_PAGO', init_point: primerLink };
    }

  } catch (e) {
    Logger.log("Error en paso2_crearPagoYEmail: " + e.message);
    return { status: 'ERROR', message: 'Error general en el servidor (Paso 2): ' + e.message };
  }
}

// =========================================================
// crearPreferenciaDePago (ACTUALIZADO)
// =========================================================
function crearPreferenciaDePago(datos, cuotaIdentificador = null, cantidadTotalCuotas = 0) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaConfig = ss.getSheetByName(NOMBRE_HOJA_CONFIG);
    const hojaRegistro = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
    let precioInscripcion;
    let tituloPago;
    const dniLimpio = limpiarDNI(datos.dni);
    let externalReference;

    // --- LÓGICA DE BLOQUEO (Actualizada a nuevas columnas) ---
    // (PUNTO 27) ESTA LÓGICA NO SE VE AFECTADA PORQUE LEE CONSTANTES QUE SE ACTUALIZAN AUTOMÁTICAMENTE
    if (hojaRegistro && hojaRegistro.getLastRow() > 1) {
      const rangoDni = hojaRegistro.getRange(2, COL_DNI_INSCRIPTO, hojaRegistro.getLastRow() - 1, 1);
      const celdaEncontrada = rangoDni.createTextFinder(dniLimpio).matchEntireCell(true).findNext();

      if (celdaEncontrada) {
        const fila = celdaEncontrada.getRow();
        if (cuotaIdentificador) {
          const cuotaIndex = parseInt(cuotaIdentificador.replace('C',''));
          let colCuota;
          if (cuotaIndex === 1) colCuota = COL_CUOTA_1;
          else if (cuotaIndex === 2) colCuota = COL_CUOTA_2;
          else if (cuotaIndex === 3) colCuota = COL_CUOTA_3;

          if (colCuota) {
            const estadoCuota = hojaRegistro.getRange(fila, colCuota).getValue();
            const estadoCuotaStr = estadoCuota ? estadoCuota.toString() : '';
            if (estadoCuotaStr.includes("Pagada") || estadoCuotaStr.includes("Notificada")) {
              Logger.log(`Bloqueo: Cuota ${cuotaIndex} ya pagada/notificada para DNI ${dniLimpio}.`);
              throw new Error(`La Cuota ${cuotaIndex} ya fue registrada como pagada.`);
            }
          }
        } else {
          const estadoPagoGeneral = hojaRegistro.getRange(fila, COL_ESTADO_PAGO).getValue();
          const estadoPagoStr = estadoPagoGeneral ? estadoPagoGeneral.toString() : '';
          if (estadoPagoStr.includes("Pagado") || estadoPagoStr.includes("Notificada")) {
            Logger.log(`Bloqueo: Pago Total ya realizado para DNI ${dniLimpio}.`);
            throw new Error(`El pago total para este DNI ya fue registrado como pagado.`);
          }
        }
      }
    }
    // --- (FIN BLOQUEO) ---

    if (cuotaIdentificador) {
      precioInscripcion = hojaConfig.getRange("B20").getValue();
      tituloPago = `Inscripción Colonia 2025 - Cuota ${cuotaIdentificador.replace('C','')} de ${cantidadTotalCuotas}`;
      externalReference = `${dniLimpio}-${cuotaIdentificador}`;
    } else {
      precioInscripcion = hojaConfig.getRange("B14").getValue();
      tituloPago = "Inscripción Colonia 2025 - Total";
      externalReference = dniLimpio;
    }

    if (!precioInscripcion || isNaN(parseFloat(precioInscripcion)) || precioInscripcion <= 0) {
      throw new Error('No se pudo determinar el precio. Revise la Hoja Config (B14 o B20).');
    }

    const appUrl = ScriptApp.getService().getUrl();

    // (Punto 2) Usar nombre/apellido si existen, sino 'apellidoNombre' (para repagos)
    const nombrePayer = datos.nombre ? `${datos.nombre} ${datos.apellido}` : datos.apellidoNombre;

    const payload = {
      items: [{
        title: tituloPago,
        description: `Inscripción para DNI ${datos.dni}`,
        quantity: 1,
        currency_id: 'ARS',
        unit_price: parseFloat(precioInscripcion)
      }],
      payer: {
        name: nombrePayer,
        // email: datos.email // Sigue quitado
      },
      external_reference: externalReference,
      back_urls: {
        success: appUrl + '?status=success',
        pending: appUrl + '?status=pending',
        failure: appUrl + '?status=failure'
      },
      auto_return: 'approved'
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + MERCADO_PAGO_ACCESS_TOKEN },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(MP_API_URL, options);
    const data = JSON.parse(response.getContentText());

    if (response.getResponseCode() >= 400) {
      Logger.log("Error MP: " + response.getContentText());
      throw new Error('Falló la creación del link de pago: ' + data.message);
    }
    return data.init_point;

  } catch (e) {
    if (e.message && (e.message.startsWith("La Cuota") || e.message.startsWith("El pago total"))) {
      Logger.log(`Bloqueo de Re-pago aplicado: ${e.message}`);
      return e.message;
    }
    Logger.log("Error en crearPreferenciaDePago: " + e.message);
    throw e;
  }
}
// =========================================================
// (FIN DE LA CORRECCIÓN)
// =========================================================

// === FUNCIONES DE WEBHOOK ===

/* */
function procesarNotificacionDePago(paymentId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log("Procesamiento de pago " + paymentId + " ya en curso (lock).");
    return;
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
    if (hoja && hoja.getLastRow() > 1) {
      // (Punto 5) Columna ID Pago MP actualizada
      // (PUNTO 27) ESTA LÓGICA NO SE VE AFECTADA
      const rangoIds = hoja.getRange(2, COL_ID_PAGO_MP, hoja.getLastRow() - 1, 1);
      const finder = rangoIds.createTextFinder(String(paymentId).split(' ')[0]);
      const celdaEncontrada = finder.findNext();

      if (celdaEncontrada) {
        Logger.log(`Payment ID ${paymentId} ya fue procesado (encontrado en Fila ${celdaEncontrada.getRow()}). Ignorando webhook duplicado.`);
        lock.releaseLock();
        return;
      }
    }

    const url = "https://api.mercadopago.com/v1/payments/" + paymentId;
    const options = {
      'method': 'get',
      'headers': { 'Authorization': 'Bearer ' + MERCADO_PAGO_ACCESS_TOKEN },
      'muteHttpExceptions': true
    };
    const response = UrlFetchApp.fetch(url, options);
    const detallesDelPago = JSON.parse(response.getContentText());

    if (response.getResponseCode() !== 200) {
      Logger.log("Fallo al obtener info del pago " + paymentId + ". Respuesta: " + response.getContentText());
      return;
    }

    const externalRef = detallesDelPago.external_reference;
    if (!externalRef) {
      Logger.log("Pago " + paymentId + " no tiene external_reference. Ignorando.");
      return;
    }
    const refParts = externalRef.split('-');
    const dniInscripto = refParts[0];
    const cuotaNum = refParts.length > 1 ? refParts[1] : null;

    const estado = detallesDelPago.status;
    const paymentIdOperacion = detallesDelPago.id;
    const montoPagado = detallesDelPago.transaction_amount || 0; // (Punto 5) Capturar monto

    const pagador = detallesDelPago.payer || {};
    let pagadorNombre = `${pagador.first_name || ''} ${pagador.last_name || ''}`.trim();
    if (!pagadorNombre && detallesDelPago.card && detallesDelPago.card.holder && detallesDelPago.card.holder.name) {
      pagadorNombre = detallesDelPago.card.holder.name.trim();
    }
    if (!pagadorNombre && pagador.nickname) {
      pagadorNombre = pagador.nickname.trim();
    }
    if (!pagadorNombre && pagador.email) {
      pagadorNombre = pagador.email.trim();
    }
    const pagadorDni = (pagador.identification && pagador.identification.number) ? pagador.identification.number : 'N/D';

    const urlComprobante = (detallesDelPago.transaction_details && detallesDelPago.transaction_details.external_resource_url)
      ? detallesDelPago.transaction_details.external_resource_url
      : 'N/D';

    Logger.log(`Info pago -> Ref: ${externalRef}, DNI: ${dniInscripto}, Cuota: ${cuotaNum}, Estado: ${estado}, ID Op: ${paymentIdOperacion}, Pagador: ${pagadorNombre}`);

    if (estado === 'approved' && dniInscripto) {
      Logger.log("Pago APROBADO. Actualizando planilla...");
      const datosActualizacion = {
        cuotaNum: cuotaNum,
        idOperacion: paymentIdOperacion,
        nombrePagador: pagadorNombre || "N/A",
        dniPagador: pagadorDni,
        urlComprobante: urlComprobante,
        montoPagado: montoPagado // (Punto 5) Pasar monto
      };
      actualizarEstadoEnPlanilla(dniInscripto, datosActualizacion);
    } else {
      Logger.log(`Pago no aprobado (estado: ${estado}) o sin DNI inscripto. No se realizan cambios.`);
    }
  } catch (e) {
    Logger.log("Error fatal en procesarNotificacionDePago: " + e.toString());
  } finally {
    lock.releaseLock();
  }
}

// =========================================================
// (Punto 5, 27) actualizarEstadoEnPlanilla (ACTUALIZADO)
// =========================================================
function actualizarEstadoEnPlanilla(dni, datosActualizacion) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
    if (!hoja) {
      Logger.log(`La hoja "${NOMBRE_HOJA_REGISTRO}" no fue encontrada.`);
      return;
    }
    // (Punto 2) Columna DNI actualizada
    // (PUNTO 27) ESTA LÓGICA NO SE VE AFECTADA
    const rangoDni = hoja.getRange(2, COL_DNI_INSCRIPTO, hoja.getLastRow() - 1, 1);
    const celdaEncontrada = rangoDni.createTextFinder(String(dni).trim()).findNext();

    if (celdaEncontrada) {
      const fila = celdaEncontrada.getRow();
      const cuotaNum = datosActualizacion.cuotaNum; // "C1", "C2", "C3" o null

      const rangoFila = hoja.getRange(fila, 1, 1, hoja.getLastColumn());
      const rowData = rangoFila.getValues()[0];

      // --- BLOQUEO DE REPAGO ---
      // (PUNTO 27) ESTA LÓGICA NO SE VE AFECTADA
      if (cuotaNum == null) {
        const estadoActual = rowData[COL_ESTADO_PAGO - 1];
        if (estadoActual && (estadoActual.toString().includes("Pagado") || estadoActual.toString().includes("Notificada"))) {
          Logger.log(`REPAGO IGNORADO: Fila ${fila} (DNI ${dni}) ya tiene un Pago Total. Ignorando PaymentID ${datosActualizacion.idOperacion}.`);
          return;
        }
        hoja.getRange(fila, COL_ESTADO_PAGO).setValue("Pagado");
        // (Punto 5) Guardar monto pagado
        hoja.getRange(fila, COL_MONTO_A_PAGAR).setValue(datosActualizacion.montoPagado);
        Logger.log(`Éxito: Fila ${fila} (Pago Total) actualizada para DNI ${dni}.`);
        enviarEmailPagoConfirmado(rowData); // (Punto 29) Este email SÍ se envía

      } else {
        const cuotaIndex = parseInt(cuotaNum.replace('C',''));
        let columnaCuota;
        if (cuotaIndex === 1) columnaCuota = COL_CUOTA_1;
        else if (cuotaIndex === 2) columnaCuota = COL_CUOTA_2;
        else if (cuotaIndex === 3) columnaCuota = COL_CUOTA_3;
        else return;

        const estadoActualCuota = rowData[columnaCuota - 1];
        if (estadoActualCuota && (estadoActualCuota.toString().includes("Pagada") || estadoActualCuota.toString().includes("Notificada"))) {
          Logger.log(`REPAGO IGNORADO: Fila ${fila} (DNI ${dni}) ya tiene ${cuotaNum} pagada. Ignorando PaymentID ${datosActualizacion.idOperacion}.`);
          return;
        }
        hoja.getRange(fila, columnaCuota).setValue(`C${cuotaIndex} Pagada`);
        // (Punto 5) Anexar monto pagado
        const celdaMonto = hoja.getRange(fila, COL_MONTO_A_PAGAR);
        const montoAnterior = celdaMonto.getValue() || 0;
        celdaMonto.setValue(parseFloat(montoAnterior) + parseFloat(datosActualizacion.montoPagado));
        
        Logger.log(`Éxito: Fila ${fila} (${cuotaNum}) marcada como PAGADA para DNI ${dni}.`);
      }
      // --- (FIN BLOQUEO) ---

      const isCuotaPayment = cuotaNum !== null;

      function actualizarCelda_AnexarSempre(columna, nuevoValor) {
        const celda = hoja.getRange(fila, columna);
        let nuevoValorStr = String(nuevoValor).trim();

        if (columna === COL_ID_PAGO_MP) {
          const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
          nuevoValorStr = `${nuevoValorStr} (${timestamp})`;
        }

        if (isCuotaPayment) {
          const valorAntiguo = celda.getValue();
          const valorAntiguoStr = (valorAntiguo) ? String(valorAntiguo).trim() : "";
          if (valorAntiguoStr && valorAntiguoStr !== 'N/D') {
            const listaValoresAntiguos = valorAntiguoStr.split(',').map(s => s.trim());
            const idBaseNuevo = String(nuevoValor).trim();
            const idYaExiste = listaValoresAntiguos.some(v => v.startsWith(idBaseNuevo));
            if (!idYaExiste) {
              celda.setValue(`${valorAntiguoStr}, ${nuevoValorStr}`);
            }
          } else {
            celda.setValue(nuevoValorStr);
          }
        } else {
          celda.setValue(nuevoValorStr);
        }
      }

      function actualizarCelda_AnexarSiDiferente(columna, nuevoValor) {
        const celda = hoja.getRange(fila, columna);
        const nuevoValorStr = String(nuevoValor).trim();
        if (nuevoValorStr === 'N/D' && isCuotaPayment) return;
        if (isCuotaPayment) {
          const valorAntiguo = celda.getValue();
          const valorAntiguoStr = (valorAntiguo) ? String(valorAntiguo).trim() : "";
          if (!valorAntiguoStr || valorAntiguoStr === 'N/D') {
            celda.setValue(nuevoValorStr);
            return;
          }
          const listaValoresAntiguos = valorAntiguoStr.split(',').map(s => s.trim());
          if (!listaValoresAntiguos.includes(nuevoValorStr)) {
            celda.setValue(`${valorAntiguoStr}, ${nuevoValorStr}`);
          }
        } else {
          celda.setValue(nuevoValorStr);
        }
      }

      // Llamadas a los helpers (Columnas actualizadas)
      // (PUNTO 27) ESTA LÓGICA NO SE VE AFECTADA, LAS CONSTANTES SE DESPLAZAN SOLAS
      actualizarCelda_AnexarSempre(COL_ID_PAGO_MP, datosActualizacion.idOperacion); // AK (37)
      actualizarCelda_AnexarSempre(COL_COMPROBANTE_MP, datosActualizacion.urlComprobante); // AP (42)
      actualizarCelda_AnexarSiDiferente(COL_PAGADOR_NOMBRE, datosActualizacion.nombrePagador); // AL (38)
      // ¡CAMBIO CLAVE! Se usa la nueva constante para la columna AM (39)
      actualizarCelda_AnexarSiDiferente(COL_DNI_PAGADOR_MP, datosActualizacion.dniPagador); // AM (39)

      if (cuotaNum != null) {
        const cantidadCuotasRegistrada = parseInt(hoja.getRange(fila, COL_CANTIDAD_CUOTAS).getValue()) || 3;
        let cuotasPagadas = 0;
        const rowDataActualizada = hoja.getRange(fila, 1, 1, hoja.getLastColumn()).getValues()[0];

        for (let i = 1; i <= cantidadCuotasRegistrada; i++) {
          let colCuota = i === 1 ? COL_CUOTA_1 : (i === 2 ? COL_CUOTA_2 : COL_CUOTA_3);
          let cuota_status = rowDataActualizada[colCuota - 1];
          if (cuota_status && (cuota_status.toString().includes("Pagada") || cuota_status.toString().includes("Notificada"))) {
            cuotasPagadas++;
          }
        }

        if (cuotasPagadas === cantidadCuotasRegistrada) {
          hoja.getRange(fila, COL_ESTADO_PAGO).setValue("Pagado");
          Logger.log(`DNI ${dni}: ¡Todas las cuotas pagadas! Estado general actualizado.`);
          enviarEmailInscripcionCompleta(rowData); // (Punto 29) Este email SÍ se envía
        }
      }
    } else {
      Logger.log(`No se encontró registro con DNI ${dni} para actualizar.`);
    }
  } catch (e) {
    Logger.log(`Error al actualizar planilla para DNI ${dni}: ${e.toString()}`);
  } finally {
    lock.releaseLock();
  }
}
// ========================================================================
// (FUNCIONES DE EMAIL REVISADAS)
// (Punto 2) Usan nombre/apellido
// (PUNTO 27) ESTA LÓGICA NO SE VE AFECTADA
// (Punto 29) ESTAS FUNCIONES SIGUEN ACTIVAS
// ========================================================================

/* */
function enviarEmailPagoConfirmado(rowData) {
  try {
    const email = rowData[COL_EMAIL - 1];
    const responsable1 = rowData[COL_ADULTO_RESPONSABLE_1 - 1];
    if (!email || !responsable1) return;

    const asunto = "Pago confirmado!!";
    const cuerpo = `Hola Sr/a ${responsable1}\n\nEl pago de la inscripción se ha efectuado correctamente.\nBienvenido en la Escuela de Verano.`;

    MailApp.sendEmail(email, asunto, cuerpo);
  } catch (e) {
    Logger.log("Error en enviarEmailPagoConfirmado: " + e.message);
  }
}

/* */
function enviarEmailInscripcionCompleta(rowData) {
  try {
    const email = rowData[COL_EMAIL - 1];
    const responsable1 = rowData[COL_ADULTO_RESPONSABLE_1 - 1];
    if (!email || !responsable1) return;

    const asunto = "Inscripción COMPLETA y Confirmada";
    const cuerpo = `Hola Sr/a ${responsable1},\n\n¡FELICITACIONES! El pago de la inscripción se ha completado en su totalidad.\n\nEl cupo está 100% confirmado.\n¡Bienvenido/a en la Escuela de Verano!`;

    MailApp.sendEmail(email, asunto, cuerpo);
    Logger.log(`Email de Inscripción Completa enviado a ${email}.`);
  } catch (e) {
    Logger.log("Error en enviarEmailInscripcionCompleta: " + e.message);
  }
}