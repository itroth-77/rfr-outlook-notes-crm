/** Le lien ne transporte qu'un identifiant Graph, jamais de jeton ni de contenu CRM. */
export function mobileAppointmentLink(eventId) {
  if (typeof eventId !== "string" || !eventId || eventId.length > 2048 || /[\x00-\x1f\x7f]/.test(eventId)) {
    throw new Error("Identifiant de rendez-vous invalide pour RDV & Notes.");
  }
  return "https://cr-api.roth-france.fr/open?" + new URLSearchParams({ eventId, view: "crm" });
}
