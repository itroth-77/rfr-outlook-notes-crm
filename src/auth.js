import {
  InteractionRequiredAuthError,
  createNestablePublicClientApplication,
} from "@azure/msal-browser";
import { withTimeout } from "./resilience.js";

let msalInstance;

export async function initializeAuthentication(config) {
  if (!globalThis.Office?.context?.requirements?.isSetSupported("NestedAppAuth", "1.1")) {
    throw new Error(
      "Cette version d’Outlook ne prend pas en charge l’authentification intégrée Microsoft (NAA 1.1).",
    );
  }

  if (!msalInstance) {
    msalInstance = await withTimeout(
      createNestablePublicClientApplication({
        auth: {
          clientId: config.clientId,
          authority: `https://login.microsoftonline.com/${config.tenantId}`,
          clientCapabilities: ["CP1"],
        },
        cache: {
          cacheLocation: "localStorage",
        },
      }),
      15_000,
      "La connexion Microsoft n’a pas pu être initialisée.",
      "AUTH_INITIALIZATION_TIMEOUT",
    );
  }
}

export async function getDataverseAccessToken(config) {
  await initializeAuthentication(config);

  const request = {
    scopes: [`${config.dataverseUrl}/user_impersonation`],
  };

  try {
    const response = await withTimeout(
      msalInstance.acquireTokenSilent(request),
      12_000,
      "La connexion Microsoft silencieuse n’a pas répondu.",
      "AUTH_SILENT_TIMEOUT",
    );
    return response.accessToken;
  } catch (error) {
    if (!(error instanceof InteractionRequiredAuthError)) {
      console.warn("Acquisition silencieuse impossible, ouverture de la connexion Microsoft.", error);
    }

    const response = await msalInstance.acquireTokenPopup(request);
    return response.accessToken;
  }
}
