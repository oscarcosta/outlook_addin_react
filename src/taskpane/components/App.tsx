import "office-ui-fabric-react/dist/css/fabric.min.css"
import * as React from "react"
import * as msal from 'msal'
import { Button, ButtonType, DefaultButton, Label, Stack } from "office-ui-fabric-react"
import Util from "../../utils/Util"
import Progress from "./Progress"
import OfficeAddinMessageBar from "./OfficeAddinMessageBar"
import { msalConfig, authParams, AuthStatus } from "../../utils/authProvider"
import { doLogout, doLoginHack } from "../../utils/authDialogHelper"
/* global Button, Header, HeroList, HeroListItem, Progress */

export interface AppProps {
    title: string
}

export interface AppState {
    title: string,
    authStatus?: AuthStatus,
    errorMessage?: string
}

export default class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context)
        this.state = {
            title: this.props.title,
            authStatus: AuthStatus.NOT_LOGGED_IN,
            errorMessage: ''
        }

        // bindings
        this.boundSetState = this.setState.bind(this)
        this.setToken = this.setToken.bind(this)
        this.displayError = this.displayError.bind(this)
        this.login = this.login.bind(this)
    }

    // Auth Properties
    accessToken: string
    msalApp = new msal.UserAgentApplication(msalConfig)

    boundSetState: () => {}

    setToken = (accesstoken: string) => {
        Util.log("setToken: " + accesstoken)
        this.accessToken = accesstoken
    }

    displayError = (error: string) => {
        Util.log("displayError: " + error)
        this.setState({ errorMessage: error })
    }

    errorDismissed = () => {
        this.setState({ errorMessage: '' })
        // If the error occured during a "in process" phase, 
        // the action didn't complete, so return the UI to the preceding state/view.
        this.setState((prevState) => {
            if (prevState.authStatus === AuthStatus.LOGIN_IN_PROGRESS) {
                return { authStatus: AuthStatus.NOT_LOGGED_IN }
            }
            return null
        })
    }

    login = async () => {
        Util.log("login")
        //await doLogin(this.boundSetState, this.setToken, this.displayError)
        await doLoginHack(this.msalApp, this.boundSetState, this.displayError)
        const account = this.msalApp.getAccount()
        if (account) {
            Util.log("account")
            Util.log(account)
            
            const token = await this.msalApp.acquireTokenSilent(authParams)
            this.setToken(token.accessToken)
            this.setState({
                authStatus: AuthStatus.LOGGED_IN
            })
        }
    }

    logout = async () => {
        Util.log("logout")
        await doLogout(this.boundSetState, this.displayError)
    }

    callApi = async () => {
        Util.log('Calling Graph API (/me)...')
        try {
            const token = await this.msalApp.acquireTokenSilent(authParams)
            Util.log(token)
            const me = await fetch('https://graph.microsoft.com/v1.0/me', {
                headers: {
                    authorization: `Bearer ${token.accessToken}`
                }
            }).then(res => res.json())
            Util.log(me)
        } catch (e) {
            Util.log(e)
            throw e
        }
    }

    render() {

        let body

        if (this.state.authStatus === AuthStatus.NOT_LOGGED_IN) {
            body = (
                <Stack tokens={{ childrenGap: 10 }}>
                    <Label>Please click the button below to login!</Label>
                    <DefaultButton text="Login" onClick={this.login} />
                    <DefaultButton text="Logout" onClick={this.logout} />
                </Stack>
            )
        } else if (this.state.authStatus === AuthStatus.LOGIN_IN_PROGRESS) {
            body = (
                <Progress message="Authenticating..." />
            )
        } else {
            body = (
                <Stack tokens={{ childrenGap: 10 }}>
                    <DefaultButton text="Logout" onClick={this.logout} />
                    <Button className="ms-welcome__action"
                        buttonType={ButtonType.hero}
                        iconProps={{ iconName: "ChevronRight" }}
                        onClick={this.callApi}>CallAPI</Button>
                </Stack>
            )
        }

        return (
            <div className='ms-welcome'>
                {this.state.errorMessage ?
                    (<OfficeAddinMessageBar onDismiss={this.errorDismissed} message={this.state.errorMessage + ' '} />)
                    : null}
                {body}
            </div>
        )
    }
}
