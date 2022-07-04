import * as React from 'react';
import styles from './ItemCreator.module.scss';
import { IItemCreatorProps, IItemCreatorState } from './IItemCreatorProps';
import { PeoplePicker } from '@microsoft/mgt-react/dist/es6/spfx';
import { Label, TextField, MessageBar, MessageBarType, Stack, IStackTokens, StackItem, DefaultButton, PrimaryButton, Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { AadHttpClient, IHttpClientOptions } from '@microsoft/sp-http';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/sites";
import "@pnp/sp/site-users/web";
import { IItemAddResult } from "@pnp/sp/items";
import { v4 as uuidv4 } from 'uuid';
import "@pnp/sp/site-groups/web";

const verticalGapStackTokens: IStackTokens = {
  childrenGap: 10,
  padding: 'l2',
};
const stackTokens: IStackTokens = { childrenGap: 40 };

export default class ItemCreator extends React.Component<IItemCreatorProps, IItemCreatorState> {
  private _sp: SPFI;
  constructor(props: IItemCreatorProps, state: IItemCreatorState) {
    super(props);
    // set initial state
    this.state = {
      Title: "",
      Description: "",
      OnwersIds: [],
      MembersIds: [],
      VisitorsIds: [],
      OnwersSPNs: [],
      Submitted: false,
      InProgess: false,
      ErrorMessage: ''
    };
  }

  //this.setState({ DataItems: results });

  public render(): React.ReactElement<IItemCreatorProps> {
    this._sp = spfi().using(SPFx({ pageContext: this.props.context.pageContext }));
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
            <PeoplePicker id='OwnersPicker' selectionChanged={this._onOwnersChanged} placeholder='Select Owners'></PeoplePicker>
          </StackItem>
          <StackItem>
            <Label htmlFor="MembersPicker">Members</Label>
            <PeoplePicker id='MembersPicker' selectionChanged={this._onMembersChanged} placeholder='Select Members'></PeoplePicker>
          </StackItem>
          <StackItem>
            <Label htmlFor="VisitorsPicker">Visitors</Label>
            <PeoplePicker id='VisitorsPicker' selectionChanged={this._onVisitorsChanged} placeholder='Select Visitors'></PeoplePicker>
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
          {(this.state.OnwersIds.length === 0) && this.state.Submitted &&
            <StackItem>
              <MessageBar messageBarType={MessageBarType.error} isMultiline={false}    >
                Owners cannot be empty.
              </MessageBar>
            </StackItem>
          }
          {(this.state.MembersIds.length === 0) && this.state.Submitted &&
            <StackItem>
              <MessageBar messageBarType={MessageBarType.error} isMultiline={false}    >
                Members cannot be empty.
              </MessageBar>
            </StackItem>
          }
          {this.state.InProgess &&
            <StackItem>
              <Label>Please wait, we are working on your request. You will be redirecting after success.</Label>
              <Spinner size={SpinnerSize.medium} />
            </StackItem>
          }
          {this.state.ErrorMessage !== '' &&
            <StackItem>
              <MessageBar messageBarType={MessageBarType.error} isMultiline={true}    >
                {this.state.ErrorMessage}
              </MessageBar>
            </StackItem>
          }
          <Stack horizontal tokens={stackTokens}>
            <PrimaryButton text="Save" onClick={() => this._saveProjectRequest()} allowDisabledFocus disabled={this.state.InProgess} />
            <DefaultButton text="Cancel" onClick={() => this._redirectPage()} allowDisabledFocus disabled={this.state.InProgess} />
          </Stack>
        </Stack>
      </section>
    );
  }

  private _onOwnersChanged = async (e: any): Promise<void> => {
    let selusers: number[] = [];
    if (e.detail && e.detail.length > 0) {
      e.detail.map(async user => {
        var Result = await this._sp.web.ensureUser(user.userPrincipalName);
        selusers.push(Result.data.Id);
      });
    }
    this.setState({ OnwersIds: selusers });
  }
  private _onMembersChanged = async (e: any): Promise<void> => {
    let selusers: number[] = [];
    if (e.detail && e.detail.length > 0) {
      e.detail.map(async user => {
        var Result = await this._sp.web.ensureUser(user.userPrincipalName);
        selusers.push(Result.data.Id);
      });
    }
    this.setState({ MembersIds: selusers });
  }
  private _onVisitorsChanged = async (e: any): Promise<void> => {
    let selusers: number[] = [];
    if (e.detail && e.detail.length > 0) {
      e.detail.map(async user => {
        var Result = await this._sp.web.ensureUser(user.userPrincipalName);
        selusers.push(Result.data.Id);
      });
    }
    this.setState({ VisitorsIds: selusers });
  }

  private _saveProjectRequest = async (): Promise<void> => {
    try {
      this.setState({ ErrorMessage: '' });
      this.setState({ Submitted: true });
      if (this.state.Title == '' || this.state.Description == '' || this.state.OnwersIds.length === 0 || this.state.MembersIds.length === 0) {
        return;
      }
      this.setState({ InProgess: true });
      //Loading List 
      const requestList = await this._sp.web.lists.getByTitle(this.props.ListTitle);
      const requestListId = await requestList.select("Id")();

      // add an item to the list
      // *** WARNING ***Append 'Id' on User Field internal Name otherwise api will not work
      const iar: IItemAddResult = await requestList.items.add({
        Title: this.state.Title,
        Description: this.state.Description,
        OwnersId: this.state.OnwersIds,
        MembersId: this.state.MembersIds,
        VisitorsId: this.state.VisitorsIds
      });
      //creating Team Site
      const regEx = /\s+/g
      const newStr = this.state.Title.replace(regEx, "").substring(0, 10);
      var NewSiteUrl = '';
      if (this.props.SiteType === 'GroupWithTeams') {
        var UniueValue = newStr + uuidv4().split('-')[2];
        NewSiteUrl = this.props.context.pageContext.web.absoluteUrl.split("sites")[0] + "sites/" + UniueValue;
        const result = await this._sp.site.createModernTeamSite(
          this.state.Title, //Title
          UniueValue,   //alias
          false, //isPublic
          1033,   //Language ID
          this.state.Description, //Description
          null,   //classification
          [this.props.context.pageContext.user.email], //Owners
          this.props.context.pageContext.legacyPageContext.departmentId, //hubSiteId
          null //siteDesignId
        );
      }
      if (this.props.SiteType === 'GroupWithoutTeams') {
        var UniueValue = uuidv4().split('-')[2];
        NewSiteUrl = this.props.context.pageContext.web.absoluteUrl.split("sites")[0] + "sites/" + UniueValue;
        const result = await this._sp.site.createCommunicationSite(
          this.state.Title, //Title
          1033,             //language id
          true,             //shareByEmailEnabled
          NewSiteUrl,     //Url
          this.state.Description, //Description
          null, //classification
          null, //siteDesignId
          this.props.context.pageContext.legacyPageContext.departmentId,  //hubSiteId
          this.props.context.pageContext.user.email //Owner
        );
      }

      //Calling Azure function
      const client = await this.props.context.aadHttpClientFactory.getClient(this.props.ClientID);
      const bodyContent: string = JSON.stringify({
        'RequestListItemId': iar.data.Id,
        'RequestListId': requestListId.Id,
        'RequestSPSiteUrl': this.props.context.pageContext.web.absoluteUrl,
        'RequestorId': iar.data.AuthorId,
        'NewSiteUrl': NewSiteUrl,
        'ProvisionTemplate': this.props.ProvisionTemplate,
        'SiteType': this.props.SiteType
      });
      const httpClientOptions: IHttpClientOptions = {
        body: bodyContent,
      };
      await (await client.post(this.props.apiUrl, AadHttpClient.configurations.v1, httpClientOptions));

      // Redirecting after save or cancel
      this._redirectPage();
    }

    catch (Error) {
      this.setState({ Submitted: true });
      this.setState({ InProgess: false });
      this.setState({ ErrorMessage: Error.message });
    }
  }
  private _redirectPage(): void {
    // Redirecting after save or cancel
    if (this.props.redirectUrl === '' || this.props.redirectUrl === '#') {
      window.location.reload();
    }
    else {
      window.location.href = this.props.redirectUrl;
    }
  }
}