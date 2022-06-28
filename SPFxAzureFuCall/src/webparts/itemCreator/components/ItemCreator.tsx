import * as React from 'react';
import styles from './ItemCreator.module.scss';
import { IItemCreatorProps, IItemCreatorState } from './IItemCreatorProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PeoplePicker } from '@microsoft/mgt-react/dist/es6/spfx';
import { Label, TextField, CommandBar, ICommandBarItemProps, Stack, IStackTokens, StackItem, DefaultButton, PrimaryButton } from 'office-ui-fabric-react';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

const verticalGapStackTokens: IStackTokens = {
  childrenGap: 10,
  padding: 'l2',
};
const stackTokens: IStackTokens = { childrenGap: 40 };

export default class ItemCreator extends React.Component<IItemCreatorProps, IItemCreatorState> {
  private _items: ICommandBarItemProps[] = [
    { key: 'save', text: 'Save', onClick: () => this.saveProjectRequest(), iconProps: { iconName: 'Save' } },
    { key: 'cancel', text: 'Cancel', onClick: () => console.log('Cancel'), iconProps: { iconName: 'Cancel' } },
  ];
  private LOG_SOURCE = "ProjectRequest";
  private LIBRARY_NAME = "ProjectRequest3";

  constructor(props: IItemCreatorProps, state: IItemCreatorState) {
    super(props);
    // set initial state
    this.state = {
      DataItems: []
    };
  }

  private getData = async (): Promise<void> => {
    console.dir(this.props.context.aadHttpClientFactory);
    const client = await this.props.context.aadHttpClientFactory.getClient('87b09524-2d48-4bd6-bf02-642eccfe5c1b');
    const siteUrl = this.props.context.pageContext.site.absoluteUrl;
    const tenantId = this.props.context.pageContext.aadInfo.tenantId;
    const results: any[] = await (await client.get(`https://samdel-functionapp.azurewebsites.net/api/ProjectRequestAdded?siteUrl=${siteUrl}&tenantId=${tenantId}`, AadHttpClient.configurations.v1)).json();
    this.setState({ DataItems: results });
  }
  public render(): React.ReactElement<IItemCreatorProps> {
    const {
      context,
      hasTeamsContext
    } = this.props;
    if (!this.state.DataItems) {
      return (
        <div>Loading....</div>
      );
    }
  
    return (
      <div >
        <PrimaryButton text="Save" onClick={() => this.getData()} allowDisabledFocus />
        <div>Site lists:</div>
        <ul>
          {this.state.DataItems.map(l => (
            <li>{l.title}</li>
          ))}
        </ul>
      </div>
    );
   /* return (
      <section className={`${styles.itemCreator} ${hasTeamsContext ? styles.teams : ''}`}>
        <CommandBar items={this._items} ariaLabel="Use left and right arrow keys to navigate between commands" />
        <Stack tokens={verticalGapStackTokens}>
          <Label className={styles.title}>Add new request</Label>
          <StackItem>
            <TextField label="Title"
              required description="Project title, It will be applied on project site." />
          </StackItem>
          <StackItem>
            <Label htmlFor="OwnersPicker">Owners</Label>
            <PeoplePicker id='OwnersPicker'></PeoplePicker>
          </StackItem>
          <StackItem>
            <Label htmlFor="MembersPicker">Members</Label>
            <PeoplePicker id='MembersPicker'></PeoplePicker>
          </StackItem>
          <StackItem>
            <Label htmlFor="VisitorsPicker">Visitors</Label>
            <PeoplePicker id='VisitorsPicker'></PeoplePicker>
          </StackItem>
          <StackItem>
            <Label htmlFor="DescriptionField">Description</Label>
            <TextField id="DescriptionField" multiline
              required description="Project description, It will be applied on project site." />
          </StackItem>
          <Stack horizontal tokens={stackTokens}>
            <PrimaryButton text="Save" onClick={() => this.saveProjectRequest()} allowDisabledFocus />
            <DefaultButton text="Cancel" onClick={() => console.log('Cancel')} allowDisabledFocus />
          </Stack>
        </Stack>
      </section>
    );*/
  }

  private saveProjectRequest = async (): Promise<void> => {
  }
}
