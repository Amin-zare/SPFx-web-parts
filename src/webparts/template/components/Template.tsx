import * as React from "react";
import styles from "./Template.module.scss";
import type { ITemplateProps } from "./ITemplateProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import ListItems from "./lists/ListItems";
import { ISPLists, IState } from "./lists/IList";
import { MSGraphClientV3 } from "@microsoft/sp-http";

export default class Template extends React.Component<ITemplateProps, IState> {
  constructor(props: ITemplateProps) {
    super(props);

    this.state = {
      listData: [],
      listMessages: [],
    };
  }

  componentDidMount(): void {
    this.getListData().catch((error: string) => {
      // Handle any unhandled Promise rejections here
      console.error("Unhandled Promise rejection:", error);
    });
    this.getMessages().catch((error: string) => {
      // Handle any unhandled Promise rejections here
      console.error("Unhandled Promise rejection:", error);
    });
  }

  getListData = (): Promise<ISPLists> => {
    const { context } = this.props;
    return context.spHttpClient
      .get(
        `${context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((data: ISPLists) => {
        this.setState({ listData: data.value });
      })
      .catch((error: string) => {
        console.error("Something happened:", error);
      });
  };

  getMessages(): Promise<void> {
    const { context } = this.props;

    if (context && context.msGraphClientFactory) {
      return context.msGraphClientFactory
        .getClient("3")
        .then((client: MSGraphClientV3): Promise<void> => {
          return client
            .api("/me/messages")
            .top(5)
            .orderby("receivedDateTime desc")
            .get()
            .then((messages: any) => {
              this.setState({ listMessages: messages.value });
            })
            .catch((error: any) => {
              console.error("Error fetching user information:", error);
            });
        })
        .catch((error: any) => {
          console.error("Error setting up Graph client:", error);
        });
    } else {
      console.error("Context or msGraphClientFactory is not available.");
      return Promise.resolve(); // Return a resolved Promise to handle the missing Promise rejection
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
      preconfiguredListName,
      order,
      style,
      numberOfItems,
    } = this.props;
    const { listData } = this.state;
    const { listMessages } = this.state;

    const listMessage = listMessages.map((item) => (
      <ul key={item.id}>
        <li>
          <span className="ms-font-l">{item.subject}</span>
        </li>
      </ul>
    ));

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
              Email: <strong>{listMessage}</strong>
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
