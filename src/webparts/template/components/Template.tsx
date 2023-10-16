import { useEffect, useState } from "react";
import * as React from "react";
import styles from "./Template.module.scss";
import { escape } from "@microsoft/sp-lodash-subset";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import ListItems from "./lists/ListItems";
import { ITemplateProps } from "./ITemplateProps";
import { ISPLists } from "./lists/IList";
import { MSGraphClientV3 } from "@microsoft/sp-http";

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
  } = props;

  interface Message {
    id: string;
    subject: string;
  }

  const [listData, setListData] = useState<ISPLists>({ value: [] });
  const [listMessages, setListMessages] = useState<Message[]>([]);

  const getListData = (): void => {
    context.spHttpClient
      .get(
        `${context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((data: ISPLists) => {
        setListData(data);
      })
      .catch((error: string) => {
        console.error("Something happened:", error);
      });
  };

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
    getListData();
    getMessages();
  }, []);

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
    </section>
  );
};

export default Template;
