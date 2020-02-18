import * as msal from 'msal'
import Util from "./Util"

// @ts-ignore
const loggerCallback = (logLevel, message, containsPii) => {
    Util.log(message)
}

export const isIE = (): boolean => {
    const ua = window.navigator.userAgent
    const msie = ua.indexOf("MSIE ") > -1
    const msie11 = ua.indexOf("Trident/") > -1
    const isEdge = ua.indexOf("Edge/") > -1
    return msie || msie11 || isEdge
}

export const msalConfig: msal.Configuration = {
    auth: {
        authority: 'https://login.microsoftonline.com/common',
        clientId: 'dc6e0f7d-2d1b-42aa-acc1-960d35e29617',
        redirectUri: window.location.origin + '/login.html'
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: isIE()
    },
    system: {
        logger: new msal.Logger(loggerCallback, {
            level: msal.LogLevel.Verbose
        })
    }
}

export const authParams: msal.AuthenticationParameters = {
    scopes: ['email', 'offline_access', 'openid', 'profile', 'user.read', 'mail.read'],
}

export enum AuthStatus {
    LOGGED_IN,
    NOT_LOGGED_IN,
    LOGIN_IN_PROGRESS
}
