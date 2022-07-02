import * as React from 'react';
import styles from './ItemCreator.module.scss';
import { IItemCreatorProps, IItemCreatorState } from './IItemCreatorProps';
import { PeoplePicker } from '@microsoft/mgt-react/dist/es6/spfx';
import { Label, TextField, MessageBar, MessageBarType, Stack, IStackTokens, StackItem, DefaultButton, PrimaryButton } from 'office-ui-fabric-react';
import { AadHttpClient, IHttpClientOptions } from '@microsoft/sp-http';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { IItemAddResult } from "@pnp/sp/items";

const verticalGapStackTokens: IStackTokens = {
  childrenGap: 10,
  padding: 'l2',
};
const stackTokens: IStackTokens = { childrenGap: 40 };

export default class ItemCreator extends React.Component<IItemCreatorProps, IItemCreatorState> {
  private sp: SPFI;
  constructor(props: IItemCreatorProps, state: IItemCreatorState) {
    super(props);
    // set initial state
    this.state = {
      Title: "",
      Description: "",
      Onwers: [],
      Members: [],
      Visitors: [],
      Submitted: false,
    };
  }

  //this.setState({ DataItems: results });

  public render(): React.ReactElement<IItemCreatorProps> {
    this.sp = spfi().using(SPFx({ pageContext: this.props.context.pageContext }));
    const {
      context,
      hasTeamsContext
    } = this.props;
    return (
      <section className={`${styles.itemCreator} ${hasTeamsContext ? styles.teams : ''}`}>
        <Stack tokens={verticalGapStackTokens}>
          <Label className={styles.title}>Add new request</Label>
          <StackItem>
            <TextField label="Title" value={this.state.Title} onChange={(e, newValue) => this.setState({ Title: newValue })}
              description="Project title, It will be applied on project site." required />
          </StackItem>
          <StackItem>
            <TextField label="Description" multiline value={this.state.Description} onChange={(e, newValue) => this.setState({ Description: newValue })}
              description="Project description, It will be applied on project site." required />
          </StackItem>
          <StackItem>
            <Label htmlFor="OwnersPicker">Owners</Label>
            <PeoplePicker id='OwnersPicker' selectionChanged={this.onOwnersChanged} placeholder='Select Owners'></PeoplePicker>
          </StackItem>
          <StackItem>
            <Label htmlFor="MembersPicker">Members</Label>
            <PeoplePicker id='MembersPicker' selectionChanged={this.onMembersChanged} placeholder='Select Members'></PeoplePicker>
          </StackItem>
          <StackItem>
            <Label htmlFor="VisitorsPicker">Visitors</Label>
            <PeoplePicker id='VisitorsPicker' selectionChanged={this.onVisitorsChanged} placeholder='Select Visitors'></PeoplePicker>
          </StackItem>
          {(this.state.Title === '') && this.state.Submitted &&
            <StackItem>
              <MessageBar messageBarType={MessageBarType.error} isMultiline={false}    >
                Title cannot be empty.
              </MessageBar>
            </StackItem>
          }
          {(this.state.Description === '') && this.state.Submitted &&
            <StackItem>
              <MessageBar messageBarType={MessageBarType.error} isMultiline={false}    >
                Description cannot be empty.
              </MessageBar>
            </StackItem>
          }
          {(this.state.Onwers.length === 0) && this.state.Submitted &&
            <StackItem>
              <MessageBar messageBarType={MessageBarType.error} isMultiline={false}    >
                Owners cannot be empty.
              </MessageBar>
            </StackItem>
          }
          {(this.state.Members.length === 0) && this.state.Submitted &&
            <StackItem>
              <MessageBar messageBarType={MessageBarType.error} isMultiline={false}    >
                Members cannot be empty.
              </MessageBar>
            </StackItem>
          }
          <Stack horizontal tokens={stackTokens}>
            <PrimaryButton text="Save" onClick={() => this.saveProjectRequest()} allowDisabledFocus />
            <DefaultButton text="Cancel" onClick={() => console.log('Cancel')} allowDisabledFocus />
          </Stack>
        </Stack>
      </section>
    );
  }

  private onOwnersChanged = async (e: any): Promise<void> => {
    let selusers: number[] = [];
    if (e.detail && e.detail.length > 0) {
      e.detail.map(async user => {
        var Result = await this.sp.web.ensureUser(user.userPrincipalName);
        selusers.push(Result.data.Id);
      });
    }
    this.setState({ Onwers: selusers });
  }
  private onMembersChanged = async (e: any): Promise<void> => {
    let selusers: number[] = [];
    if (e.detail && e.detail.length > 0) {
      e.detail.map(async user => {
        var Result = await this.sp.web.ensureUser(user.userPrincipalName);
        selusers.push(Result.data.Id);
      });
    }
    this.setState({ Members: selusers });
  }
  private onVisitorsChanged = async (e: any): Promise<void> => {
    let selusers: number[] = [];
    if (e.detail && e.detail.length > 0) {
      e.detail.map(async user => {
        var Result = await this.sp.web.ensureUser(user.userPrincipalName);
        selusers.push(Result.data.Id);
      });
    }
    this.setState({ Visitors: selusers });
  }

  private saveProjectRequest = async (): Promise<void> => {
    this.setState({ Submitted: true });
    if (this.state.Title == '' || this.state.Description == '' || this.state.Onwers.length === 0 || this.state.Members.length === 0) {
      return;
    }
    //Loading List 
    const requestList = await this.sp.web.lists.getByTitle(this.props.ListTitle);
    const requestListId = await requestList.select("Id")();

    // add an item to the list
    // *** WARNING ***Append 'Id' on User Field internal Name otherwise api will not work
    const iar: IItemAddResult = await requestList.items.add({
      Title: this.state.Title,
      Description: this.state.Description,
      OwnersId: this.state.Onwers,
      MembersId: this.state.Members,
      VisitorsId: this.state.Visitors
    });

    //Calling Azure function
    const client = await this.props.context.aadHttpClientFactory.getClient(this.props.ClientID);
    const bodyContent: string = JSON.stringify({
      'RequestListItemId': iar.data.Id,
      'RequestListId': requestListId.Id,
      'RequestSPSiteUrl': this.props.context.pageContext.web.absoluteUrl,
      'RequestorId': iar.data.AuthorId
    });
    const httpClientOptions: IHttpClientOptions = {
      body: bodyContent,
    };
    // Redirecting after save or cancel
    const results: any[] = await (await client.post(this.props.apiUrl, AadHttpClient.configurations.v1, httpClientOptions)).json();
    if (this.props.redirectUrl !== '') {
      window.location.href = this.props.redirectUrl;
    }
  }
}
