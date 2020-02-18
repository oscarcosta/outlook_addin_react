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
})()
