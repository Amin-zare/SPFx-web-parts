import { useEffect, useState } from "react";
import * as React from "react";
import styles from "./Template.module.scss";
import { escape } from "@microsoft/sp-lodash-subset";
import ListItems from "./lists/ListItems";
import { ITemplateProps } from "./ITemplateProps";
import { ISPLists } from "./lists/IList";
import { getListData } from "./lists/ListService";
import { getMessages } from "./email/GraphService";
import MessageList from "./email/MessageList";
import ConditionalComponent from "./Configuration/Conditional";

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

  useEffect(() => {
    // Fetch list data when the component mounts
    getListData(context)
      .then((data: ISPLists) => {
        setListData(data);
      })
      .catch((error: string) => {
        console.error("Error fetching list data:", error);
      });

    getMessages(context)
      .then((data) => {
        setListMessages(data);
      })
      .catch((error: string) => {
        console.error("Error fetching messages data:", error);
      });
  }, []);

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
            Email:{" "}
            <strong>
              <MessageList messages={listMessages} />
            </strong>
          </div>
        </div>
      </div>
      <ListItems listData={listData.value} />
      <ConditionalComponent
        preconfiguredListName={preconfiguredListName}
        order={order}
        style={style}
        numberOfItems={numberOfItems}
      />
    </section>
  );
};

export default Template;
