import * as React from "react";
import styles from "./Template.module.scss";
import type { ITemplateProps } from "./ITemplateProps";
import { escape } from "@microsoft/sp-lodash-subset";
import ListItems from "./lists/ListItems";
import { IState } from "./lists/IList";
import { getListData } from "./lists/ListService"; // Import getListData
import { getMessages } from "./email/GraphService"; // Import getMessages
import MessageList from "./email/MessageList";
import { needsConfiguration } from "./Configuration/ConfigurationChecker";
import ConfigurationDisplay from "./Configuration/ConfigurationDisplay"; // Import the ConfigurationDisplay component

export default class Template extends React.Component<ITemplateProps, IState> {
  constructor(props: ITemplateProps) {
    super(props);

    this.state = {
      listData: [],
      listMessages: [],
      needsConfiguration: false,
    };
  }

  componentDidMount(): void {
    getListData(this.props.context)
      .then((data) => {
        this.setState({ listData: data.value });
      })
      .catch((error) => {
        // Handle any unhandled Promise rejections here
        console.error("Unhandled Promise rejection:", error);
      });
    getMessages(this.props.context)
      .then((messages) => {
        this.setState({ listMessages: messages });
      })
      .catch((error) => {
        // Handle any unhandled Promise rejections here
        console.error("Unhandled Promise rejection:", error);
      });
    if (
      needsConfiguration(
        this.props.preconfiguredListName,
        this.props.order,
        this.props.style
      )
    ) {
      this.setState({ needsConfiguration: true });
    }
  }

  public render(): React.ReactElement<ITemplateProps> {
    const {
      description,
      checkbox,
      toggle,
      multiLineText,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      context,
      Rating,
      listName,
      itemName,
    } = this.props;
    const { listData } = this.state;
    const { listMessages } = this.state;
    const { needsConfiguration } = this.state;

    return (
      <section
        className={`${styles.template} ${hasTeamsContext ? styles.teams : ""}`}
      >
        <div className={styles.welcome}>
          <img
            alt=""
            src={
              isDarkTheme
                ? require("../assets/welcome-dark.png")
                : require("../assets/welcome-light.png")
            }
            className={styles.welcomeImage}
          />
          <h2> User name: {escape(userDisplayName)}!</h2>
          <div> Environment Message :{environmentMessage}</div>
          <div>
            Description : <strong>{escape(description)}</strong>
          </div>
          <div>
            Checkbox :
            <strong>
              {checkbox ? "Checkbox is checked" : "Checkbox is not checked"}
            </strong>
          </div>
          <div>
            Toggle :<strong>{toggle ? "Toggle is on" : "Toggle is off"}</strong>
          </div>
          <div>
            Multi-line : <strong>{escape(multiLineText)}</strong>
          </div>
          <div>
            <div>
              Loading from :{" "}
              <strong>{escape(context.pageContext.web.title)}</strong>
            </div>
          </div>
          <div>
            <div>
              Rating : <strong>{Rating}</strong>
            </div>
          </div>
          <div>
            <div>
              List name: <strong>{escape(listName)}</strong>
            </div>
          </div>
          <div>
            <div>
              Item name: <strong>{escape(itemName)}</strong>
            </div>
          </div>
          <div>
            <div>
              Email:{" "}
              <strong>
                {" "}
                <MessageList messages={listMessages} />{" "}
                {/* Use the MessageList component */}
              </strong>
            </div>
          </div>
        </div>
        <ListItems listData={listData} />

        {needsConfiguration ? (
          <ConfigurationDisplay
            needsConfiguration={needsConfiguration}
            preconfiguredListName={""}
            order={""}
            numberOfItems={0}
            style={""}
          />
        ) : (
          <></>
        )}
      </section>
    );
  }
}
