import * as React from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';

import styles from './AddProjectRequest.module.scss';

export interface IAddProjectRequestProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
}

const LOG_SOURCE: string = 'AddProjectRequest';

export default class AddProjectRequest extends React.Component<IAddProjectRequestProps, {}> {
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: AddProjectRequest mounted');
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: AddProjectRequest unmounted');
  }

  public render(): React.ReactElement<{}> {
    return <div className={styles.addProjectRequest} />;
  }
}
