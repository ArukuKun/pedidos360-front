export const msalConfig = {
    auth: {
        clientId: "5d7ead55-12ca-4dcd-97e8-5b5e81f77f4b", 
        authority: "https://login.microsoftonline.com/5d020e86-60f8-47e9-9a84-7abf9906fdff",
        redirectUri: "https://j76lk49uq2.execute-api.us-east-1.amazonaws.com/", 
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true,
    },
};

export const loginRequest = {
    scopes: ["User.Read"] 
};