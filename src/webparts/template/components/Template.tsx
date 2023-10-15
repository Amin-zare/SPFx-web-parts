import * as React from "react";
import styles from "./Template.module.scss";
import type { ITemplateProps } from "./ITemplateProps";
import { escape } from "@microsoft/sp-lodash-subset";

export default class Template extends React.Component<ITemplateProps, {}> {
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
    } = this.props;

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
        </div>
      </section>
    );
  }
}
