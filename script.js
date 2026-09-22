// Regla: si envías un correo a alguien @arcadis.com Y el asunto/cuerpo dice "cotización"
// Y ADEMÁS menciona "5444", recuerda copiar a jose.gonzalez@arcadis.com (si no está ya incluido).
// No aplica a correos enviados a otros dominios, ni si falta cualquiera de las dos palabras.
var KEYWORD = "cotizaci"; // cubre "cotización" y "cotizacion" (sin tilde), en minúsculas
var KEYWORD_2 = "5444"; // debe aparecer también, junto con KEYWORD
var TRIGGER_DOMAIN = "@arcadis.com"; // dominio del destinatario que activa la regla
var REMINDER_EMAIL = "jose.gonzalez@arcadis.com"; // contacto que se debe copiar

Office.onReady();

function getAddresses(recipients) {
  if (!recipients) return [];
  return recipients.map(function (r) {
    return (r.emailAddress || "").toLowerCase();
  });
}

function checkForCotizacion(event) {
  var item = Office.context.mailbox.item;

  item.to.getAsync(function (toResult) {
    var toAddrs = toResult.status === Office.AsyncResultStatus.Succeeded
      ? getAddresses(toResult.value)
      : [];

    var sendsToArcadis = toAddrs.some(function (addr) {
      return addr.indexOf(TRIGGER_DOMAIN) !== -1;
    });

    // La regla solo aplica si hay al menos un destinatario @arcadis.com en "Para"
    if (!sendsToArcadis) {
      event.completed({ allowEvent: true });
      return;
    }

    item.subject.getAsync(function (subjectResult) {
      var subject = subjectResult.status === Office.AsyncResultStatus.Succeeded
        ? (subjectResult.value || "").toLowerCase()
        : "";

      item.body.getAsync(Office.CoercionType.Text, function (bodyResult) {
        var body = bodyResult.status === Office.AsyncResultStatus.Succeeded
          ? (bodyResult.value || "").toLowerCase()
          : "";

        var text = subject + " " + body;

        if (text.indexOf(KEYWORD) === -1 || text.indexOf(KEYWORD_2) === -1) {
          event.completed({ allowEvent: true });
          return;
        }

        // Menciona "cotización" y "5444", y va a @arcadis.com: revisar si ya está copiado el contacto
        item.cc.getAsync(function (ccResult) {
          var ccAddrs = ccResult.status === Office.AsyncResultStatus.Succeeded
            ? getAddresses(ccResult.value)
            : [];

          var allAddrs = toAddrs.concat(ccAddrs);
          var alreadyCopied = allAddrs.indexOf(REMINDER_EMAIL.toLowerCase()) !== -1;

          if (alreadyCopied) {
            event.completed({ allowEvent: true });
            return;
          }

          event.completed({
            allowEvent: false,
            errorMessage:
              "Este correo va a " + TRIGGER_DOMAIN + " y menciona 'cotización' y '5444'. ¿Olvidaste copiar a " +
              REMINDER_EMAIL +
              "? Agrégalo en CC y vuelve a enviar, o cierra este aviso y presiona Enviar de nuevo para hacerlo igual.",
            cancelLabel: "Volver a editar",
            sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser
          });
        });
      });
    });
  });
}

Office.actions.associate("checkForCotizacion", checkForCotizacion);
