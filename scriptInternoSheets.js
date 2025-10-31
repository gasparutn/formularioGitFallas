/*

// =========================================================
// (¡¡¡CONSTANTES ACTUALIZADAS!!!) - 36 columnas total
// =========================================================
const NOMBRE_HOJA_REGISTRO = 'Registros';
const COL_NUMERO_TURNO = 1;       // A
const COL_ESTADO_NUEVO_ANT = 4;   // D 
const COL_EMAIL = 5;              // E
const COL_APELLIDO_NOMBRE = 6;    // F
const COL_DNI_INSCRIPTO = 9;      // I
const COL_ADULTO_RESPONSABLE_1 = 12;// L
const COL_METODO_PAGO = 25;       // Y
const COL_CUOTA_1 = 26;           // Z
const COL_CUOTA_2 = 27;           // AA
const COL_CUOTA_3 = 28;           // AB
const COL_CANTIDAD_CUOTAS = 29;   // AC
const COL_ESTADO_PAGO = 30;       // AD 
const COL_ENVIAR_EMAIL_MANUAL = 36;// AJ 
// (g) Columnas AK, AL, AM, AN eliminadas
// =========================================================


//* Se ejecuta CUANDO ABRES LA HOJA DE CÁLCULO.

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Escuela Hípico')
    .addSubMenu(ui.createMenu('✉️ Mailing').addItem('Enviar Emails de Bienvenida (General)', 'enviarEmailsBienvenidaManuales').addItem('Notificar Pagos de Cuota (Ventana)', 'mostrarDialogoNotificarCuotas'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🧰Utililidad').addItem('Eliminar Espaciados en Datos', 'limpiarColumnasPorHoja'))
    .addToUi();
}


 // Muestra la ventana emergente para seleccionar las cuotas a notificar.

function mostrarDialogoNotificarCuotas() {
  const html = HtmlService.createTemplateFromFile('NotificarCuotas').evaluate()
    .setWidth(450)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Notificar Cuotas Pagadas Pendientes');
}



 // (g) Rango de búsqueda actualizado a 36 columnas
 
function getServerDataAndShowDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRegistro = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
  const ui = SpreadsheetApp.getUi();

  const selection = ss.getSelection();
  const activeRange = selection.getActiveRange();

  if (!activeRange) {
    ui.alert("Atención", "Por favor, seleccione una o más celdas en las filas que desea notificar.", ui.ButtonSet.OK);
    return [];
  }

  const startRow = activeRange.getRow();
  const numRows = activeRange.getNumRows();

  if (startRow === 1) {
    ui.alert("Atención", "No seleccione la fila de encabezados.", ui.ButtonSet.OK);
    return [];
  }

  // (g) Rango máximo actualizado a la columna AJ = 36
  const range = hojaRegistro.getRange(startRow, 1, numRows, COL_ENVIAR_EMAIL_MANUAL);
  const data = range.getValues();

  const processedData = [];

  for (let i = 0; i < data.length; i++) {
    const rowData = data[i];
    const rowNum = startRow + i;

    if (rowData[COL_METODO_PAGO - 1] === 'Pago en Cuotas') {
      const cantidadCuotas = parseInt(rowData[COL_CANTIDAD_CUOTAS - 1]) || 3;
      const cuotasPendientesNotif = [];

      const cuotasColumns = [
        { statusCol: COL_CUOTA_1, index: 1 },
        { statusCol: COL_CUOTA_2, index: 2 },
        { statusCol: COL_CUOTA_3, index: 3 }
      ];

      for (const cuota of cuotasColumns) {
        if (cuota.index > cantidadCuotas) continue;

        const estadoCuota = rowData[cuota.statusCol - 1];
        const estadoCuotaStr = estadoCuota ? estadoCuota.toString() : '';

        if (estadoCuotaStr.includes("Pagada")) {
          cuotasPendientesNotif.push({
            index: cuota.index,
            nombreEstado: estadoCuotaStr,
            status: 'Pendiente Notificación'
          });
        }
      }

      if (cuotasPendientesNotif.length > 0) {
        processedData.push({
          row: rowNum,
          dni: rowData[COL_DNI_INSCRIPTO - 1],
          nombre: rowData[COL_APELLIDO_NOMBRE - 1],
          email: rowData[COL_EMAIL - 1],
          responsable: rowData[COL_ADULTO_RESPONSABLE_1 - 1],
          cuotas: cuotasPendientesNotif
        });
      }
    }
  }

  return processedData;
}


/**
 * Función de servidor: Envía los emails de las cuotas seleccionadas y actualiza el estado.

function enviarNotificacionCuotasSeleccionadas(cuotasSeleccionadas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRegistro = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
  let emailsEnviados = 0;
  let cuotasNotificadas = 0;
  let errores = 0;

  const emailsPorFila = new Map();

  for (const selection of cuotasSeleccionadas) {
    if (!emailsPorFila.has(selection.row)) {
      emailsPorFila.set(selection.row, {
        email: selection.email,
        responsable: selection.responsable,
        nombre: selection.nombre,
        cuotas: []
      });
    }
    emailsPorFila.get(selection.row).cuotas.push(selection.cuotaIndex);
  }

  for (const [fila, data] of emailsPorFila.entries()) {
    try {
      data.cuotas.sort((a, b) => a - b);

      const cuotasNotificadasStr = data.cuotas.map(c => `Cuota ${c}`).join(' y ');

      const asunto = `Confirmación de Pago - ${cuotasNotificadasStr}`;

      let cuerpo;
      if (data.cuotas.length === 1) {
        cuerpo = `Hola Sr/a ${data.responsable} (Responsable de ${data.nombre}),\n\nConfirmamos que hemos recibido el pago de la ${cuotasNotificadasStr}.\n\n¡Muchas gracias!`;
      } else {
        cuerpo = `Hola Sr/a ${data.responsable} (Responsable de ${data.nombre}),\n\nConfirmamos que hemos recibido el pago de las siguientes cuotas: ${cuotasNotificadasStr}.\n\n¡Muchas gracias!`;
      }

      MailApp.sendEmail(data.email, asunto, cuerpo);
      emailsEnviados++;

      data.cuotas.forEach(cuotaIndex => {
        let colCuota;
        if (cuotaIndex === 1) colCuota = COL_CUOTA_1;
        else if (cuotaIndex === 2) colCuota = COL_CUOTA_2;
        else if (cuotaIndex === 3) colCuota = COL_CUOTA_3;

        if (colCuota) {
          hojaRegistro.getRange(fila, colCuota).setValue(`C${cuotaIndex} Notificada`);
          cuotasNotificadas++;
        }
      });

    } catch (e) {
      errores++;
      Logger.log(`Error al enviar email en fila ${fila}: ${e.message}`);
    }
  }

  SpreadsheetApp.flush();

  return {
    enviados: emailsEnviados,
    marcados: cuotasNotificadas,
    errores: errores
  };
}



 // Función de utilidad para emails manuales (General)

function enviarEmailsBienvenidaManuales() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRegistro = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);

  if (!hojaRegistro) {
    ui.alert("Error", `No se encontró la hoja "${NOMBRE_HOJA_REGISTRO}".`, ui.ButtonSet.OK);
    return;
  }

  const rango = hojaRegistro.getRange(2, 1, hojaRegistro.getLastRow() - 1, hojaRegistro.getLastColumn());
  const data = rango.getValues();

  let enviados = 0;
  let omitidos = 0;
  let errores = 0;

  for (let i = 0; i < data.length; i++) {
    const rowData = data[i];
    const turno = rowData[COL_NUMERO_TURNO - 1];
    const filaActual = turno + 1; // (c) Fila = Turno + 1

    const checkbox = rowData[COL_ENVIAR_EMAIL_MANUAL - 1]; // Col AJ

    if (checkbox === true) {
      try {
        const estadoPago = rowData[COL_ESTADO_PAGO - 1];     // Col AD

        if (estadoPago === 'Pagado' || rowData[COL_METODO_PAGO - 1] === 'Pago Efectivo') {
          const email = rowData[COL_EMAIL - 1];                // Col E
          const responsable1 = rowData[COL_ADULTO_RESPONSABLE_1 - 1]; // Col L

          if (!email || !responsable1) {
            throw new Error("Fila " + filaActual + ": Falta email o nombre del responsable.");
          }

          const asunto = "Pago confirmado!!";
          const cuerpo = `Hola Sr/a ${responsable1}\n\nEl pago de la inscripción se ha efectuado correctamente.\nBienvenido en la Escuela de Verano.`;

          MailApp.sendEmail(email, asunto, cuerpo);

          hojaRegistro.getRange(filaActual, COL_ENVIAR_EMAIL_MANUAL).setValue(false);
          enviados++;

        } else {
          omitidos++;
        }
      } catch (e) {
        errores++;
        Logger.log(`Error al enviar email en fila ${filaActual}: ${e.message}`);
      }
    }
  }

  SpreadsheetApp.flush();
  ui.alert("Proceso Completado",
    `Emails de Bienvenida enviados: ${enviados}\n` +
    `Omitidos (no estaban "Pagado" o Efectivo): ${omitidos}\n` +
    `Errores (verificar logs): ${errores}`,
    ui.ButtonSet.OK);
}
*/