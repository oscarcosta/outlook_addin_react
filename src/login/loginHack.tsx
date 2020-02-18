//import { UserAgentApplication } from 'msal'
//import { msalConfig } from '../utils/authProvider'
import Util from "../utils/Util"

(() => {
    Util.log('Init login...')

    // This part will be run inside the popup
    if (window.location.hash.includes('id_token=')) {
        Util.log('Id token in hash (is callback)')
        Office.onReady(() => {
            Util.log('Window is callback (wait)')
            if (Office.context.ui) {
                Office.context.ui.messageParent(window.location.hash)
                Util.log('Message sent to parent')
            } else {
                Util.log('Missing Office.context.ui')
            }
        })
        return
    }

    // Only initialize msal if window is not callback (not popup).
    // MSAL will pick up the hash and redirect - we don't want this.
    //const msal = new UserAgentApplication(msalConfig)

})()
