import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { ISPLists } from "./IList";

export const getListData = (context: any): Promise<ISPLists> => {
  return context.spHttpClient
    .get(
      `${context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`,
      SPHttpClient.configurations.v1
    )
    .then((response: SPHttpClientResponse) => {
      return response.json();
    })
    .catch((error: string) => {
      console.error("Something happened:", error);
    });
};