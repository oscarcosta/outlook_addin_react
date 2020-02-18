import * as React from "react"
import { Spinner, SpinnerType } from "office-ui-fabric-react"
/* global Spinner */

export interface ProgressProps {
    message: string
}

export default class Progress extends React.Component<ProgressProps> {
    render() {
        const { message } = this.props

        return (
            <section className="ms-welcome__progress ms-u-fadeIn500">
                <Spinner type={SpinnerType.large} label={message} />
            </section>
        )
    }
}
