import * as React from "react";

interface ConfigurationDisplayProps {
  needsConfiguration: boolean;
  preconfiguredListName: string;
  order: string;
  numberOfItems: number;
  style: string;
}

const ConfigurationDisplay: React.FC<ConfigurationDisplayProps> = ({
  needsConfiguration,
  preconfiguredListName,
  order,
  numberOfItems,
  style,
}) => {
  if (needsConfiguration) {
    return (
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
        <div className="ms-Grid-row" />{" "}
      </div>
    );
  } else {
    return (
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
    );
  }
};

export default ConfigurationDisplay;
