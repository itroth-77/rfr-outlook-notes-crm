function officeAsync(invoke) {
  return new Promise((resolve, reject) => {
    invoke((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error?.message || "Erreur Outlook inconnue."));
      }
    });
  });
}

function readDate(value) {
  if (value instanceof Date) return Promise.resolve(value);
  if (value && typeof value.getAsync === "function") {
    return officeAsync((done) => value.getAsync(done));
  }
  return Promise.resolve(value ? new Date(value) : null);
}

function readText(value) {
  if (typeof value === "string") return Promise.resolve(value);
  if (value && typeof value.getAsync === "function") {
    return officeAsync((done) => value.getAsync(done));
  }
  return Promise.resolve(value ?? "");
}

export async function getAppointmentContext() {
  await Office.onReady();

  if (Office.context.host !== Office.HostType.Outlook) {
    throw new Error("Ce formulaire doit être ouvert depuis Outlook.");
  }

  const item = Office.context.mailbox.item;
  const [subject, start, end, location] = await Promise.all([
    readText(item.subject),
    readDate(item.start),
    readDate(item.end),
    readText(item.location),
  ]);

  if (!subject || !start || !end) {
    throw new Error("Outlook n’a pas fourni toutes les informations du rendez-vous.");
  }

  return {
    itemId: item.itemId || "",
    subject: String(subject),
    start: new Date(start),
    end: new Date(end),
    location: String(location || ""),
    organizerName: item.organizer?.displayName || "",
    organizerEmail: item.organizer?.emailAddress || "",
  };
}

export async function markAppointmentAsLogged() {
  const item = Office.context.mailbox.item;
  const properties = await officeAsync((done) => item.loadCustomPropertiesAsync(done));
  properties.set("EventLogged", true);
  await officeAsync((done) => properties.saveAsync(done));
}
