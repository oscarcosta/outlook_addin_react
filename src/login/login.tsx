/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as msal from 'msal'
import { msalConfig, authParams } from '../utils/authProvider'
import Util from "../utils/Util"

(() => {
    Util.log('Init login...')

    const msg = (text) => {
        document.getElementById("content").innerText = text
    }

    // The initialize function must be run each time a new page is loaded
    Office.initialize = () => {
        Util.log('Office initialized')

        const msalApp = new msal.UserAgentApplication(msalConfig)

        msalApp.handleRedirectCallback((error: msal.AuthError, response: msal.AuthResponse) => {
            Util.log('handleRedirectCallback')
            Util.log(response)

            if (!error) {
                Util.log('auth OK')

                if (response.tokenType === 'id_token') {
                    localStorage.setItem("loggedIn", "yes")
                } else {
                    // The tokenType is access_token, so send success message and token.
                    Office.context.ui.messageParent(JSON.stringify({ status: 'success', result: response.accessToken }))
                    msg("Please wait, we are redirecting you back to your application.")
                }
            } else {
                Util.log('auth Error')
                Util.log(error)

                const errorData: string = `error: ${error.errorCode} \nmessage: ${error.errorMessage}`
                Office.context.ui.messageParent(JSON.stringify({ status: 'failure', result: errorData }))
            }

            Util.log('account')
            Util.log(msalApp.getAccount())
        })

        Util.log("is login in progress? " + msalApp.getLoginInProgress())

        if (msalApp.getAccount() || localStorage.getItem("loggedIn") === "yes") {
            Util.log('calling acquireTokenRedirect')

            msalApp.acquireTokenRedirect(authParams)
        } else {
            Util.log('calling loginRedirect')
            msalApp.loginRedirect(authParams)
        }
    }
})()
