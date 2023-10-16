import { useEffect, useState } from "react";
import * as React from "react";
import styles from "./Template.module.scss";
import { escape } from "@microsoft/sp-lodash-subset";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import ListItems from "./lists/ListItems";
import { ITemplateProps } from "./ITemplateProps";
import { ISPLists } from "./lists/IList";

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
  } = props;

  const [listData, setListData] = useState<ISPLists>({ value: [] });

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
  useEffect(() => {
    // Fetch list data when the component mounts
    getListData();
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
      </div>
      <ListItems listData={listData.value} />
    </section>
  );
};

export default Template;
