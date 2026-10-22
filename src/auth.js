import {
  InteractionRequiredAuthError,
  createNestablePublicClientApplication,
} from "@azure/msal-browser";

let msalInstance;

export async function initializeAuthentication(config) {
  if (!globalThis.Office?.context?.requirements?.isSetSupported("NestedAppAuth", "1.1")) {
    throw new Error(
      "Cette version d’Outlook ne prend pas en charge l’authentification intégrée Microsoft (NAA 1.1).",
    );
  }

  if (!msalInstance) {
    msalInstance = await createNestablePublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        clientCapabilities: ["CP1"],
      },
      cache: {
        cacheLocation: "localStorage",
      },
    });
  }
}

export async function getDataverseAccessToken(config) {
  await initializeAuthentication(config);

  const request = {
    scopes: [`${config.dataverseUrl}/user_impersonation`],
  };

  try {
    const response = await msalInstance.acquireTokenSilent(request);
    return response.accessToken;
  } catch (error) {
    if (!(error instanceof InteractionRequiredAuthError)) {
      console.warn("Acquisition silencieuse impossible, ouverture de la connexion Microsoft.", error);
    }

    const response = await msalInstance.acquireTokenPopup(request);
    return response.accessToken;
  }
}
