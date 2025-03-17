function enviarRecordatorios_DESACTIVADO() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tareas");
  if (!hoja) {
    Logger.log("ERROR: La hoja 'Tareas' no existe.");
    return;
  }

  var dataRange = hoja.getDataRange();
  var data = dataRange.getValues();

  var hoy = new Date();
  var diasAnticipacion = 2;

  for (var i = 1; i < data.length; i++) {
    var tarea = data[i][0];              
    var responsable = data[i][9];  //Correo ahora en la columna J      
    var estado = data[i][3];             
    var fechaFinalizacion = new Date(data[i][5]);  

    Logger.log("Procesando fila " + (i + 1) + ":");
    Logger.log("   - Tarea: " + tarea);
    Logger.log("   - Responsable: " + responsable);
    Logger.log("   - Estado: " + estado);
    Logger.log("   - Fecha de finalización (original): " + fechaFinalizacion);

    if (!(fechaFinalizacion instanceof Date) || isNaN(fechaFinalizacion.getTime())) {
      Logger.log("ERROR: La fecha no es válida en la fila " + (i + 1));
      continue;
    }

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(responsable)) {
      Logger.log("ERROR: Correo inválido en la fila " + (i + 1) + ": " + responsable);
      continue;
    }

    var diffMillis = fechaFinalizacion.getTime() - hoy.getTime();
    var diffDias = Math.floor(diffMillis / (1000 * 60 * 60 * 24));

    Logger.log("   - Días restantes para vencer: " + diffDias);

    if (diffDias <= diasAnticipacion && diffDias >= 0 && estado !== "Completada") {
      var asunto = "Recordatorio: La tarea '" + tarea + "' está próxima a vencer";
      var cuerpo = "Hola,\n\n" +
                   "La tarea '" + tarea + "' tiene su fecha de finalización próxima (" + 
                   formatearFecha(fechaFinalizacion) + ").\n\n" +
                   "Por favor, revisa y actualiza el estado de la tarea.\n\n" +
                   "Saludos,\nTu sistema de seguimiento de tareas";

      try {
        MailApp.sendEmail({
          to: responsable,
          subject: asunto,
          body: cuerpo
        });
        Logger.log("   ✅ CORREO ENVIADO a " + responsable);
      } catch (e) {
        Logger.log("   ❌ ERROR AL ENVIAR CORREO a " + responsable + ": " + e.toString());
      }
    } else {
      Logger.log("   ⏩ No se envía recordatorio porque la fecha o estado no cumplen los criterios.");
    }
  }
}

function formatearFecha(fecha) {
  var meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
  var dia = fecha.getDate();
  var mes = meses[fecha.getMonth()];
  var año = fecha.getFullYear();
  return dia + "/" + mes + "/" + año;
}
