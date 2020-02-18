import "office-ui-fabric-react/dist/css/fabric.min.css"
import * as React from "react"
import * as ReactDOM from "react-dom"
import { initializeIcons } from "office-ui-fabric-react"
import App from "./components/App"
import Util from "../utils/Util"

initializeIcons()

const title = "My Addin v1.0.0"

Office.onReady(() => {
    Util.log("Office ready!")

    ReactDOM.render(
        <App title={title} />,
        document.getElementById("container")
    )
})
