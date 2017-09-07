import { IWebPartContext } from "@microsoft/sp-webpart-base";

export interface IWebhookSubscriptionProps {
  listname: string;
  subscriptionUrl: string;
  clientState: string;
  context: IWebPartContext;
}

export interface IWebhookSubscriptionState {
  subscriptions?: ISubscriptionValue[];
  loading?: boolean;
  creating?: boolean;
  subLoading?: string;
  error?: string;
  subTest?: string;
}

export interface ISubscription {
  value: ISubscriptionValue[];
}

export interface ISubscriptionValue {
  id: string;
  clientState: string;
  expirationDateTime: string;
  notificationUrl: string;
  resource: string;
}
