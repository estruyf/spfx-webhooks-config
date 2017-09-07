export interface IWebhookSubscriptionWebPartProps {
  listname: string;
  subscriptionUrl: string;
  clientState: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}