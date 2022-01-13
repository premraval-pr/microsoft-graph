import React, { Component } from "react";
import { BrowserRouter as Router, Route, Redirect } from "react-router-dom";
import { Container } from "reactstrap";
import NavBar from "./NavBar";
import ErrorMessage from "./ErrorMessage";
import Welcome from "./Welcome";
import "bootstrap/dist/css/bootstrap.css";
import withAuthProvider, { AuthComponentProps } from "./AuthProvider";
import Calendar from "./Calendar";
import NewEvent from "./NewEvent";
import { config } from "./Config";
import { getListMessage, getUserDetails } from "./GraphService";
import { Message, User } from "microsoft-graph";

class App extends Component<AuthComponentProps> {
    formatMessages(messages: Message[], user: User) {
        let formattedMessage: { [key: string]: Message[] } = {};

        messages.forEach((message) => {
            const count = messages.filter(
                (item) => item.conversationId === message.conversationId
            ).length;
            if (count > 1) {
                if (message.conversationId) {
                    if (formattedMessage[message.conversationId]) {
                        formattedMessage[
                            message.conversationId.toString()
                        ].push(message);
                    } else {
                        formattedMessage[message.conversationId.toString()] =
                            [];
                        formattedMessage[
                            message.conversationId.toString()
                        ].push(message);
                    }
                }
            }
        });
        console.log(formattedMessage);

        let output = "";

        Object.keys(formattedMessage).forEach((key) => {
            console.log("CONVERSATION:==" + key);

            for (let i = 0; i < formattedMessage[key].length - 1; i++) {
                console.log(
                    formattedMessage[key][i].sender?.emailAddress?.address,
                    user.mail,
                    formattedMessage[key][i].sender?.emailAddress?.address !==
                        user.mail
                );
                if (
                    formattedMessage[key][i].sender?.emailAddress?.address !==
                    user.mail
                ) {
                    let Q = formattedMessage[key][i].uniqueBody?.content
                        ?.replace(/(\r\n|\n|\r)/gm, " ")
                        .replace(/\s\s+/g, " ")
                        .replace(/(<([^>]+)>)/gi, "")
                        .replaceAll("&nbsp;", "")
                        .replaceAll("&quot;", "'")
                        .replaceAll("&lt;", "<")
                        .replaceAll("&gt;", ">")
                        .replaceAll("&amp;", "&")
                        .trim();
                    let A = formattedMessage[key][i + 1].uniqueBody?.content
                        ?.replace(/(\r\n|\n|\r)/gm, " ")
                        .replace(/\s\s+/g, " ")
                        .replace(/(<([^>]+)>)/gi, "")
                        .replaceAll("&nbsp;", "")
                        .replaceAll("&quot;", "'")
                        .replaceAll("&lt;", "<")
                        .replaceAll("&gt;", ">")
                        .replaceAll("&amp;", "&")
                        .trim();
                    if (Q !== A && Q !== "" && A !== "") {
                        output += "Q: " + Q + "\n";
                        output += "A: " + A + "\n\n";
                    }
                }
            }
        });

        this.downloadTxtFile(output, user.mail ? `${user.mail}.txt` : "");
    }

    downloadTxtFile(output: string, filename: string) {
        const element = document.createElement("a");
        const file = new Blob([output], { type: "text/plain;charset=utf-8" });
        element.href = URL.createObjectURL(file);
        element.download = filename;
        document.body.appendChild(element);
        element.click();
    }

    async componentDidMount() {
        let accessToken = await this.props.getAccessToken(config.scopes);
        let messages = await getListMessage(accessToken);
        let user = await getUserDetails(accessToken);
        this.formatMessages(messages, user);
    }

    render() {
        let error = null;
        if (this.props.error) {
            error = (
                <ErrorMessage
                    message={this.props.error.message}
                    debug={this.props.error.debug}
                />
            );
        }

        return (
            <Router>
                <div>
                    <NavBar
                        isAuthenticated={this.props.isAuthenticated}
                        authButtonMethod={
                            this.props.isAuthenticated
                                ? this.props.logout
                                : this.props.login
                        }
                        user={this.props.user}
                    />
                    <Container>
                        {error}
                        <Route
                            exact
                            path="/"
                            render={(props) => (
                                <Welcome
                                    {...props}
                                    isAuthenticated={this.props.isAuthenticated}
                                    user={this.props.user}
                                    authButtonMethod={this.props.login}
                                />
                            )}
                        />
                        <Route
                            exact
                            path="/calendar"
                            render={(props) =>
                                this.props.isAuthenticated ? (
                                    <Calendar {...props} />
                                ) : (
                                    <Redirect to="/" />
                                )
                            }
                        />
                        <Route
                            exact
                            path="/newevent"
                            render={(props) =>
                                this.props.isAuthenticated ? (
                                    <NewEvent {...props} />
                                ) : (
                                    <Redirect to="/" />
                                )
                            }
                        />
                    </Container>
                </div>
            </Router>
        );
    }
}

export default withAuthProvider(App);
