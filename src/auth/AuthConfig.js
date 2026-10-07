export const msalConfig = {
    auth: {
        clientId: "TU_CLIENT_ID_DE_AZURE",
        authority: "https://login.microsoftonline.com/TU_TENANT_ID_DE_AZURE",
        redirectUri: "URL_DE_AWS_PENDIENTE", 
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    },
};

export const loginRequest = {
    scopes: ["User.Read"] 
};