/** Le lien ne transporte qu'un identifiant Graph, jamais de jeton ni de contenu CRM. */
export function mobileAppointmentLink(eventId) {
  if (typeof eventId !== "string" || !eventId || eventId.length > 2048 || /[\x00-\x1f\x7f]/.test(eventId)) {
    throw new Error("Identifiant de rendez-vous invalide pour RDV & Notes.");
  }
  return "https://cr-api.roth-france.fr/open?" + new URLSearchParams({ eventId, view: "crm" });
}

/** User-gesture-only Android intent, fixed package and verified HTTPS route.
 * Chrome supports this syntax; an Outlook WebView may still refuse external apps.
 */
export function androidAppointmentIntent(eventId) {
  const url = mobileAppointmentLink(eventId);
  return url.replace(/^https:/, "intent:") +
    "#Intent;scheme=https;package=fr.roth.compterendu;end";
}

export function mobileLaunchLink(eventId, { isAndroid = false } = {}) {
  return isAndroid ? androidAppointmentIntent(eventId) : mobileAppointmentLink(eventId);
}
