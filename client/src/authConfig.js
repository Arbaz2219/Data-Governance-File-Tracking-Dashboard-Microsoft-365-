import { PublicClientApplication, LogLevel } from "@azure/msal-browser";

export const msalConfig = {
    auth: {
        clientId: "249138ab-21cd-4538-bf4a-7cf81c3594a8",
        authority: "https://login.microsoftonline.com/d0f9ddaa-eec7-4657-9565-aff3f741be7f",
        redirectUri: "http://localhost:5173/",
        navigateToLoginRequestUrl: false,
    },
    cache: {
        cacheLocation: "localStorage", // Better persistence
        storeAuthStateInCookie: true, // Helps with some IE/browser issues
    },
    system: {
        loggerOptions: {
            loggerCallback: (level, message, containsPii) => {
                if (containsPii) return;
                switch (level) {
                    case LogLevel.Error:
                        console.error("[MSAL Error]", message);
                        return;
                    case LogLevel.Info:
                        console.info("[MSAL Info]", message);
                        return;
                    case LogLevel.Verbose:
                        console.debug("[MSAL Verbose]", message);
                        return;
                    case LogLevel.Warning:
                        console.warn("[MSAL Warning]", message);
                        return;
                    default:
                        return;
                }
            }
        }
    }
};

export const loginRequest = {
    scopes: ["User.Read"]
};

export const msalInstance = new PublicClientApplication(msalConfig);
