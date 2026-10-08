export const msalConfig = {
    auth: {
<<<<<<< HEAD
        clientId: "TU_CLIENT_ID_DE_AZURE",
        authority: "https://login.microsoftonline.com/TU_TENANT_ID_DE_AZURE",
        redirectUri: "URL_DE_AWS_PENDIENTE", 
=======
        clientId: "1b54f9e6-dc5f-40a3-b612-f17203a45f15",
        authority: "https://login.microsoftonline.com/5d020e86-60f8-47e9-9a84-7abf9906fdff",
        // Reemplazar <IP_EC2_FRONTEND> por la IP pública de la EC2 del Frontend (ej: http://54.210.10.15)
        redirectUri: "http://<IP_EC2_FRONTEND>", 
>>>>>>> 1febf91 (Front)
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    },
};

export const loginRequest = {
<<<<<<< HEAD
    scopes: ["User.Read"] 
=======
    scopes: ["User.Read", "api://1b54f9e6-dc5f-40a3-b612-f17203a45f15/access_as_user"] 
>>>>>>> 1febf91 (Front)
};