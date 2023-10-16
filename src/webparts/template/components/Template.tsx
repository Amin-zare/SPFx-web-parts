import { useEffect, useState } from "react";
import * as React from "react";
import styles from "./Template.module.scss";
import { escape } from "@microsoft/sp-lodash-subset";
import ListItems from "./lists/ListItems";
import { ITemplateProps } from "./ITemplateProps";
import { ISPLists } from "./lists/IList";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { getListData } from "./lists/ListService";

const Template: React.FC<ITemplateProps> = (props) => {
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
  } = props;

  interface Message {
    id: string;
    subject: string;
  }

  const [listData, setListData] = useState<ISPLists>({ value: [] });
  const [listMessages, setListMessages] = useState<Message[]>([]);

  const getMessages = (): void => {
    const { context } = props;

    if (context && context.msGraphClientFactory) {
      context.msGraphClientFactory
        .getClient("3")
        .then((client: MSGraphClientV3): Promise<void> => {
          return client
            .api("/me/messages")
            .top(5)
            .orderby("receivedDateTime desc")
            .get((error, messages) => {
              if (error) {
                console.error("Error fetching user information:", error);
              } else {
                setListMessages(messages.value);
              }
            });
        })
        .catch((error: string) => {
          console.error("Error setting up Graph client:", error);
        });
    } else {
      console.error("Context or msGraphClientFactory is not available.");
    }
  };

  useEffect(() => {
    // Fetch list data when the component mounts
    getListData(context)
      .then((data: ISPLists) => {
        setListData(data);
      })
      .catch((error: string) => {
        console.error("Error fetching list data:", error);
      });

    getMessages();
  }, []);

  const listMessage = listMessages.map((item) => (
    <ul key={item.id}>
      <li>
        <span className="ms-font-l">{item.subject}</span>
      </li>
    </ul>
  ));

  const isEmpty = (value: string): boolean => {
    return value === undefined || value === null || value.length === 0;
  };

  const needsConfiguration = (): boolean => {
    return isEmpty(preconfiguredListName) || isEmpty(order) || isEmpty(style);
  };

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
        <div> Environment Message: {environmentMessage}</div>
        <div>
          Description: <strong>{escape(description)}</strong>
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
      <ListItems listData={listData.value} />
      {needsConfiguration() ? (
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
};

export default Template;
