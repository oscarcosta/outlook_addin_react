import { UserAgentApplication } from 'msal'
import { AuthStatus, authParams } from "./authProvider"
import Util from './Util'

/* LOGIN */

const dialogLoginUrl: string = window.location.origin + '/login.html'
let loginDialog: Office.Dialog

interface AppState {
    authStatus?: AuthStatus
}

/**
 * Open a dialog to execute the msal authentication flow.
 */
export const doLogin = async (setState: (x: AppState) => void, setToken: (x: string) => void, displayError: (x: string) => void) => {
    Util.log("doLogin")

    // Event handler for DialogMessageReceived
    const processLoginMessage = (arg: any) => {
        Util.log("processLoginMessage")
        Util.log(arg)

        let messageFromDialog = JSON.parse(arg.message)
        if (messageFromDialog.status === 'success') {
            Util.log("processLoginMessage: dialog says success")

            // We now have a valid access token.
            loginDialog.close()
            setToken(messageFromDialog.result)
            setState({
                authStatus: AuthStatus.LOGGED_IN
            })
        } else {
            Util.log("processLoginMessage: dialog says error")

            // Something went wrong with authentication or the authorization of the web application.
            loginDialog.close()
            displayError(messageFromDialog.result)
        }
    }

    // Event handler for DialogEventReceived
    const processLoginDialogEvent = (arg: any) => {
        Util.log("processLoginDialogEvent")
        Util.log(arg)

        processDialogEvent(arg, setState, displayError)
    }

    setState({ authStatus: AuthStatus.LOGIN_IN_PROGRESS })

    // Call the dialog
    Office.context.ui.displayDialogAsync(dialogLoginUrl, { height: 40, width: 30 }, (result) => {
        Util.log("displayDialogAsync")

        if (result.status === Office.AsyncResultStatus.Failed) {
            Util.log("displayDialogAsync Failed")
            Util.log(result.error)

            displayError(`${result.error.code} ${result.error.message}`)
        } else {
            Util.log("displayDialogAsync Succeeded")

            loginDialog = result.value
            loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processLoginMessage)
            loginDialog.addEventHandler(Office.EventType.DialogEventReceived, processLoginDialogEvent)
        }
    })
}

export const doLoginHack = async (msal: UserAgentApplication, setState: (x: AppState) => void, displayError: (x: string) => void) => {
    Util.log("doLoginHack")

    //@ts-ignore
    msal.openPopup = () => {
        Util.log("openPopup")

        const dummy = {
            close() {
            },
            location: {
                assign(url) {
                    Util.log("assign")
                    Util.log(url)

                    Office.context.ui.displayDialogAsync(url, { width: 25, height: 50 }, result => {
                        Util.log("displayDialogAsync")
                        
                        if (result.status === Office.AsyncResultStatus.Failed) {
                            Util.log("displayDialogAsync Failed")
                            Util.log(result.error)
                
                            displayError(`${result.error.code} ${result.error.message}`)
                        } else {
                            Util.log("displayDialogAsync Succeeded")

                            dummy.close = result.value.close
                            result.value.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                                Util.log("displayDialogAsync DialogMessageReceived")
                                Util.log(arg)
                                
                                //@ts-ignore
                                dummy.location.href = dummy.location.hash = arg.message
                            })
                            loginDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) =>{
                                Util.log("displayDialogAsync DialogEventReceived")
                                Util.log(arg)

                                processDialogEvent(arg, setState, displayError)
                            })
                        }
                    })
                }
            }
        }
        return dummy
    }

    Util.log('Logging in...')
    await msal.loginPopup(authParams)
}

/* LOGOUT */

const dialogLogoutUrl: string = window.location.origin + '/logout.html'
let logoutDialog: Office.Dialog

export const doLogout = async (setState: (x: AppState) => void, displayError: (x: string) => void) => {
    Util.log("doLogout")

    // Event handler for DialogMessageReceived
    const processLogoutMessage = () => {
        Util.log("processLogoutMessage")

        logoutDialog.close()
        setState({
            authStatus: AuthStatus.NOT_LOGGED_IN
        })
    }

    // Event handler for DialogEventReceived
    const processLogoutDialogEvent = (arg: any) => {
        Util.log("processLogoutDialogEvent")
        Util.log(arg)

        processDialogEvent(arg, setState, displayError)
    }

    // Call the dialog
    Office.context.ui.displayDialogAsync(dialogLogoutUrl, { height: 40, width: 30 }, (result) => {
        Util.log("displayDialogAsync")

        if (result.status === Office.AsyncResultStatus.Failed) {
            Util.log("displayDialogAsync Failed")
            Util.log(result.error)

            displayError(`${result.error.code} ${result.error.message}`)
        } else {
            Util.log("displayDialogAsync Succeeded")

            logoutDialog = result.value
            logoutDialog.addEventHandler(Office.EventType.DialogMessageReceived, processLogoutMessage)
            logoutDialog.addEventHandler(Office.EventType.DialogEventReceived, processLogoutDialogEvent)
        }
    })
}

/* COMMON METHODS */

const processDialogEvent = (arg: any,
    setState: (x: AppState) => void,
    displayError: (x: string) => void) => {
    Util.log("processDialogEvent")

    switch (arg.error) {
        case 12002:
            displayError('The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.')
            break
        case 12003:
            displayError('The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.')
            break
        case 12006:
            // 12006 means that the user closed the dialog instead of waiting for it to close.
            // It is not known if the user completed the login or logout, so assume the user is
            // logged out and revert to the app's starting state. It does no harm for a user to
            // press the login button again even if the user is logged in.
            setState({
                authStatus: AuthStatus.NOT_LOGGED_IN
            })
            break
        default:
            displayError('Unknown error in dialog box.')
            break
    }
}
