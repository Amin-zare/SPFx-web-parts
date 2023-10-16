import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export interface IState {
  listData: ISPList[];
  listMessages: MicrosoftGraph.Message[];
  needsConfiguration: boolean;
}
