import * as React from "react";
import styles from "./Template.module.scss";
import type { ITemplateProps } from "./ITemplateProps";
import { escape } from "@microsoft/sp-lodash-subset";

export default class Template extends React.Component<ITemplateProps, {}> {
  public render(): React.ReactElement<ITemplateProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
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
        </div>
      </section>
    );
  }
}
