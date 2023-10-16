import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { ISPLists } from "./IList";

export const getListData = (context: any): Promise<ISPLists> => {
  return new Promise<ISPLists>((resolve, reject) => {
    context.spHttpClient
      .get(
        `${context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((data: ISPLists) => {
        resolve(data);
      })
      .catch((error: string) => {
        reject(error);
      });
  });
};
