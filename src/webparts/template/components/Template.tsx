import * as React from "react";
import styles from "./Template.module.scss";
import type { ITemplateProps } from "./ITemplateProps";
import { escape } from "@microsoft/sp-lodash-subset";
import ListItems from "./lists/ListItems";
import { IState } from "./lists/IList";
import { getListData } from "./lists/ListService"; // Import getListData
import { getMessages } from "./email/GraphService"; // Import getMessages
import MessageList from "./email/MessageList";
export default class Template extends React.Component<ITemplateProps, IState> {
  constructor(props: ITemplateProps) {
    super(props);

    this.state = {
      listData: [],
      listMessages: [],
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
      preconfiguredListName,
      order,
      style,
      numberOfItems,
    } = this.props;
    const { listData } = this.state;
    const { listMessages } = this.state;

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
        {this.needsConfiguration() ? (
          <div
            className="ms-Grid"
            style={{
              color: "#666",
              backgroundColor: "#f4f4f4",
              padding: "80px 0",
              alignItems: "center",
              boxAlign: "center",
            }}
          >
            <div className="ms-Grid-row" style={{ color: "#333" }}>
              <div className="ms-Grid-col ms-u-hiddenSm ms-u-md3" />
              <div
                className="ms-Grid-col ms-u-sm12 ms-u-md6"
                style={{
                  height: "100%",
                  whiteSpace: "nowrap",
                  textAlign: "center",
                }}
              >
                <i
                  className="ms-fontSize-su ms-Icon ms-Icon--ThumbnailView"
                  style={{
                    display: "inline-block",
                    verticalAlign: "middle",
                    whiteSpace: "normal",
                  }}
                />
                <span
                  className="ms-fontWeight-light ms-fontSize-xxl"
                  style={{
                    paddingLeft: "20px",
                    display: "inline-block",
                    verticalAlign: "middle",
                    whiteSpace: "normal",
                  }}
                >
                  Gallery
                </span>
              </div>
              <div className="ms-Grid-col ms-u-hiddenSm ms-u-md3" />
            </div>
            <div
              className="ms-Grid-row"
              style={{
                width: "65%",
                verticalAlign: "middle",
                margin: "0 auto",
                textAlign: "center",
              }}
            >
              <span
                style={{
                  color: "#666",
                  fontSize: "17px",
                  display: "inline-block",
                  margin: "24px 0",
                  fontWeight: 100,
                }}
              >
                Show items from the selected list
              </span>
            </div>
            <div className="ms-Grid-row" />
          </div>
        ) : (
          <div>
            <div>
              <div
                className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white `}
              >
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <span className="ms-font-xl ms-fontColor-white">
                    Welcome to SharePoint!
                  </span>
                  <p className="ms-font-l ms-fontColor-white">
                    Customize SharePoint experiences using web parts.
                  </p>
                  <p className="ms-font-l ms-fontColor-white">
                    List: {escape(preconfiguredListName)}
                    <br />
                    Order: {escape(order)}
                    <br />
                    Number of items: {numberOfItems}
                    <br />
                    Style: {escape(style)}
                  </p>
                  <a href="https://aka.ms/spfx">
                    <span>Learn more</span>
                  </a>
                </div>
              </div>
            </div>
          </div>
        )}
      </section>
    );
  }
  private needsConfiguration(): boolean {
    return (
      Template.isEmpty(this.props.preconfiguredListName) ||
      Template.isEmpty(this.props.order) ||
      Template.isEmpty(this.props.style)
    );
  }

  private static isEmpty(value: string): boolean {
    return value === undefined || value === null || value.length === 0;
  }
}
