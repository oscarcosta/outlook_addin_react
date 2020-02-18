/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as msal from 'msal'
import { msalConfig } from '../utils/authProvider'
import Util from "../utils/Util"

(() => {
    Util.log('Init logout...')

    // The initialize function must be run each time a new page is loaded
    Office.initialize = () => {
        Util.log('Office initialized')

        const userAgentApplication = new msal.UserAgentApplication(msalConfig)

        Util.log('calling logout')
        userAgentApplication.logout()
    }
})()