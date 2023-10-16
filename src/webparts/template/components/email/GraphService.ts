import { MSGraphClientV3 } from "@microsoft/sp-http";

export const getMessages = (context: any): Promise<any[]> => {
  if (context && context.msGraphClientFactory) {
    return context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3): Promise<any[]> => {
        return client
          .api("/me/messages")
          .top(5)
          .orderby("receivedDateTime desc")
          .get()
          .then((messages: any) => {
            return messages.value;
          });
      });
  } else {
    console.error("Context or msGraphClientFactory is not available.");
    return Promise.resolve([]); // Return an empty array to handle the missing Promise rejection
  }
};
