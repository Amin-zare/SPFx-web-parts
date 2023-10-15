import * as React from "react";
import styles from "./Template.module.scss";
import type { ITemplateProps } from "./ITemplateProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import ListItems from "./lists/ListItems";
import { ISPLists, IState } from "./lists/IList";

export default class Template extends React.Component<ITemplateProps, IState> {
  constructor(props: ITemplateProps) {
    super(props);

    this.state = {
      listData: [],
    };
  }

  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  componentDidMount() {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this.getListData();
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
        </div>
        <ListItems listData={listData} />
      </section>
    );
  }
}
