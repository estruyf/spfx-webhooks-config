import * as React from 'react';
import * as moment from 'moment';
import styles from './WebhookSubscription.module.scss';
import { IWebhookSubscriptionProps, IWebhookSubscriptionState, ISubscription, ISubscriptionValue } from './IWebhookSubscriptionProps';
import { SPHttpClient, SPHttpClientResponse, HttpClient, HttpClientResponse } from "@microsoft/sp-http";
import { isEqual } from '@microsoft/sp-lodash-subset';

import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/spinner';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

export default class WebhookSubscription extends React.Component<IWebhookSubscriptionProps, IWebhookSubscriptionState> {
  /**
   * @constructor
   * Constructor
   * @param props
   */
  constructor(props: IWebhookSubscriptionProps) {
    super(props);

    this.state = {
      subscriptions: [],
      loading: true,
      creating: false,
      subLoading: "",
      error: "",
      subTest: ""
    };

    // Bind this to the click events
    this._testSubscriptionUrl = this._testSubscriptionUrl.bind(this);
    this._createSubscription = this._createSubscription.bind(this);
    this._updateSubscription = this._updateSubscription.bind(this);
    this._deleteSubscription = this._deleteSubscription.bind(this);
  }

  /**
   * @function
   * Component is mounted on the page
   */
  public componentDidMount() {
    // Get the list / library subscriptions
    this._getSubscriptions(this.props.listname);
  }

  /**
   * @function
   * Component received property updates
   * @param nextProps
   * @param nextContext
   */
  public componentWillReceiveProps(nextProps: IWebhookSubscriptionProps, nextContext: any) {
    // Check if the listname is configured
    if (nextProps.listname !== this.props.listname) {
      this._getSubscriptions(nextProps.listname);
    }
  }

  /**
   * Specify if the component needs to get updated
   * @param nextProps
   * @param nextState
   */
  public shouldComponentUpdate(nextProps: IWebhookSubscriptionProps, nextState: IWebhookSubscriptionState) {
    if (!isEqual(nextProps, this.props) || !isEqual(nextState, this.state)) {
      return true;
    }

    return false;
  }

  /**
   * @function
   * Retrieving all subscriptions for the specified list
   * @param listName
   */
  private _getSubscriptions(listName: string) {
    const restUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/subscriptions`;
    // Call the subscription API to check all webhooks subs on the list
    this.props.context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => { return response.json(); })
      .then(data => {
        this.setState({
          subscriptions: data.value,
          loading: false,
          creating: false,
          subLoading: "",
          error: ""
        });
      }).catch(error => {
        console.log(`ERROR: ${error}`);
      });
  }

  /**
   * @function
   * Function to test the subscription URL if it is up and running
   */
  private _testSubscriptionUrl() {
    this.setState({
      subTest: ''
    });

    // Test the service URL with a fake token
    // Service URL should respond with the same token
    const serviceUrl = `${this.props.subscriptionUrl}?validationtoken=test-token`;
    this.props.context.httpClient.get(serviceUrl, HttpClient.configurations.v1)
      .then((response: HttpClientResponse) => { return response.text(); })
      .then((value: string) => {
        this.setState({
          subTest: `: Online - ${value}`
        });
      })
      .catch(error => {
        this.setState({
          subTest: ": Failed"
        });
      });
  }

  /**
   * @function
   * Create a new subscription
   */
  private _createSubscription() {
    this.setState({
      creating: true,
      error: ""
    });

    const restUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listname}')/subscriptions`;
    // Do a post request to the subscriptions endpoint
    this.props.context.spHttpClient.post(restUrl, SPHttpClient.configurations.v1, {
      body: JSON.stringify({
        "resource": `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listname}')`,
        "notificationUrl": this.props.subscriptionUrl,
        "expirationDateTime": moment().add(90, 'days'),
        "clientState": this.props.clientState || "SubscriptionFromWebhookWP"
      })
    }).then((response: SPHttpClientResponse) => {
      if (response.status >= 200 && response.status < 300) {
        // Update the subscriptions list
        this._getSubscriptions(this.props.listname);
      } else {
        // Check the error message
        response.json().then(data => {
          if (typeof data.error !== "undefined") {
            console.log('ERROR:', data.error.message);
            this.setState({
              error: data.error.message,
              creating: false
            });
          }
        });
      }
    }).catch(err => {
      console.log('ERROR:', err);
      // Reset the subscription which is loading
      this._setSubLoading();
    });
  }

  /**
   * Update a specific subscription expiration date
   * @param id
   */
  private _updateSubscription(id: string) {
    this._setSubLoading(id);

    // Update the subscription
    const restUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listname}')/subscriptions('${id}')`;
    this.props.context.spHttpClient.fetch(restUrl, SPHttpClient.configurations.v1, {
      method: 'PATCH',
      body: JSON.stringify({
        "expirationDateTime": moment().add(90, 'days')
      })
    }).then((response: SPHttpClientResponse) => {
      // Update the subscriptions list
      this._getSubscriptions(this.props.listname);
    }).catch(err => {
      console.log('ERROR:', err);
      // Reset the subscription which is loading
      this._setSubLoading();
    });
  }

  /**
   * Delete a specific subscription from the list
   * @param id
   */
  private _deleteSubscription(id: string) {
    this._setSubLoading(id);

    // Update the subscription
    const restUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listname}')/subscriptions('${id}')`;
    this.props.context.spHttpClient.fetch(restUrl, SPHttpClient.configurations.v1, {
      method: 'DELETE'
    }).then((response: SPHttpClientResponse) => {
      // Update the subscriptions list
      this._getSubscriptions(this.props.listname);
    }).catch(err => {
      console.log('ERROR:', err);
      // Reset the subscription which is loading
      this._setSubLoading();
    });
  }

  /**
   * Specify if a subscription is loading
   * @param id
   */
  private _setSubLoading(id: string = "") {
    this.setState({
      subLoading: id
    });
  }

  public render(): React.ReactElement<IWebhookSubscriptionProps> {
    if (this.props.listname !== "") {
      if (this.state.loading) {
        return <Spinner size={SpinnerSize.large} label='Loading your data, just wait a sec...' />;
      } else {
        return (
          <div className={`${styles.webhook}`}>
            <h2>Managing subscriptions for: {this.props.listname}</h2>
            <h3>Create a subscription:</h3>
            {
              this.props.subscriptionUrl === "" ?
                <MessageBar messageBarType={MessageBarType.warning}>To create subscriptions you have to configure the subscription URL.</MessageBar> :
                (
                  <div>
                    <p>Subscriptions will be created with the following URL: {this.props.subscriptionUrl}</p>

                    <p><a href="javascript:;" title="Update subscription" onClick={this._testSubscriptionUrl}>Test subscription URL</a><span>{this.state.subTest}</span></p>


                    <p><a href="javascript:;" title="Update subscription" onClick={this._createSubscription}>Create a new subscription</a></p>
                    {
                      (() => {
                        if (this.state.creating) {
                          return <Spinner size={SpinnerSize.medium} label={`Creating a new subscription`} />;
                        }

                        if (this.state.error !== "") {
                          return <MessageBar messageBarType={MessageBarType.error}>{this.state.error}</MessageBar>;
                        }
                      })()
                    }
                  </div>
                )
            }
            <h3>Subscription overview:</h3>
            {
              this.state.subscriptions.length === 0 ? <MessageBar messageBarType={MessageBarType.warning}>No subscriptions created for this list.</MessageBar> : ''
            }
            {
              this.state.subscriptions.map((subscription: ISubscriptionValue) => {
                return (
                  <MessageBar className={`${styles.subscriptions}`}>
                    {
                      this.state.subLoading === subscription.id ?
                        <Spinner size={SpinnerSize.medium} label={`Processing subscription: ${subscription.id}`} /> :
                        (
                          <div>
                            <h4>
                              Subscription with ID: {subscription.id}

                              <a href="javascript:;" title="Update subscription" onClick={() => this._updateSubscription(subscription.id)}><i className="ms-Icon ms-Icon--Refresh" aria-hidden="true"></i></a>
                              <a href="javascript:;" title="Delete subscription" onClick={() => this._deleteSubscription(subscription.id)}><i className="ms-Icon ms-Icon--Delete" aria-hidden="true"></i></a>
                            </h4>

                            <p>Subscription details:</p>
                            <ul key={subscription.id}>
                              <li>clientState: {subscription.clientState}</li>
                              <li>expirationDateTime: {subscription.expirationDateTime}</li>
                              <li>resource: {subscription.resource}</li>
                              <li>notificationUrl: {subscription.notificationUrl}</li>
                            </ul>
                          </div>
                        )
                    }
                  </MessageBar>
                );
              })
            }
          </div>
        );
      }
    } else {
      return (
        <div>
          <h2 className="ms-fontColor-orangeLight">Please configure the web part before you can manage your subscriptions.</h2>
        </div>
      );
    }
  }
}
