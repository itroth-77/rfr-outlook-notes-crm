// Même contrat que l’APK : GUID Dataverse de la séance, pas celui du maître.
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export function appointmentNoteReference(appointmentId) {
  const id = String(appointmentId || '').replace(/[{}]/g, '').toLowerCase();
  if (!GUID.test(id)) throw new Error("La référence Dynamics de la séance n’est pas valide.");
  return `[RFR-RDV:${id}]`;
}
export function withAppointmentReference(text, appointmentId) {
  return `${text.trim()}\n\nRéférence RDV : ${appointmentNoteReference(appointmentId)}`;
}
export function belongsToAppointment(text, appointmentId) {
  const plain = String(text || '').replace(/<[^>]*>/g, ' ').trim().toLowerCase();
  return plain.endsWith(('Référence RDV : ' + appointmentNoteReference(appointmentId)).toLowerCase());
}
export function withoutAppointmentReference(text) {
  return String(text || '').replace(/(?:\s|<br\s*\/?>)*Référence RDV : \[RFR-RDV:[0-9a-f-]{36}\](?=\s*(?:<\/p>)?\s*$)/i, '');
}
