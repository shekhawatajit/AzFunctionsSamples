import * as React from 'react';
import styles from './ItemCreator.module.scss';
import { IItemCreatorProps } from './IItemCreatorProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PeoplePicker } from '@microsoft/mgt-react/dist/es6/spfx';
import { Label, TextField, CommandBar, ICommandBarItemProps, Stack, IStackTokens, StackItem, DefaultButton, PrimaryButton } from 'office-ui-fabric-react';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

const verticalGapStackTokens: IStackTokens = {
  childrenGap: 10,
  padding: 'l2',
};
const stackTokens: IStackTokens = { childrenGap: 40 };

export default class ItemCreator extends React.Component<IItemCreatorProps, {}> {
  private _items: ICommandBarItemProps[] = [
    { key: 'save', text: 'Save', onClick: () => this.saveProjectRequest(), iconProps: { iconName: 'Save' } },
    { key: 'cancel', text: 'Cancel', onClick: () => console.log('Cancel'), iconProps: { iconName: 'Cancel' } },
  ];
  private LOG_SOURCE = "ProjectRequest";
  private LIBRARY_NAME = "ProjectRequest3";


  constructor(props: IItemCreatorProps) {
    super(props);
    // set initial state
    this.state = {
      items: [],
      errors: []
    };
  }

  public render(): React.ReactElement<IItemCreatorProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
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
            <PrimaryButton text="Save" onClick={() => this.callExternalapi()} allowDisabledFocus />
            <DefaultButton text="Cancel" onClick={() => console.log('Cancel')} allowDisabledFocus />
          </Stack>
        </Stack>
      </section>
    );
  }
  private callExternalapi = async (): Promise<void> => {
    try {

      const client = await this.props.aadFactory.getClient("https://onrocks.onmicrosoft.com/6c677d39-46c2-4848-93b0-2973cb4d0a72");
      const requestUrl = `https://oip-functionapp.azurewebsites.net/api/ProjectRequestAdded?code=9achtSkEiuZyb0dC5480DdLb4XQtKLz3DzLW6TobXZB3AzFuFys1oA==`;
      const result: any = await (await client.get(requestUrl, AadHttpClient.configurations.v1)).json();
      console.log(result);
      alert("Creation done!");
    } catch (err) {
      console.log(`${this.LOG_SOURCE} (callExternalapi) - ${JSON.stringify(err)} - `);
    }
  }
  private saveProjectRequest = async (): Promise<void> => {

    this.props.aadFactory
      .getClient('6c677d39-46c2-4848-93b0-2973cb4d0a72')
      .then((client: AadHttpClient): void => {
        client
          .get('https://samdel-functionapp.azurewebsites.net/api/ProjectRequestAdded', AadHttpClient.configurations.v1)
          .then((response: HttpClientResponse): Promise<any> => {
            console.log(response.json());
            return response.json();
          })
      });
  }
}
