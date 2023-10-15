import { Component } from "react";
import * as React from "react";
import styles from "./List.module.scss";
import { ISPList } from "./IList";

interface IListItemsProps {
  listData: ISPList[];
}

class ListItems extends Component<IListItemsProps> {
  render(): JSX.Element {
    const { listData } = this.props;
    const listItems = listData.map((item) => (
      <ul className={styles.list} key={item.Id}>
        <li className={styles.listItem}>
          <span className="ms-font-l">{item.Title}</span>
        </li>
      </ul>
    ));

    return <div>{listItems}</div>;
  }
}

export default ListItems;
