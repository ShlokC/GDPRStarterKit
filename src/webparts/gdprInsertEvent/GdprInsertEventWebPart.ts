import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'GdprInsertEventWebPartStrings';
import GdprInsertEvent from './components/GdprInsertEvent';
import { IGdprInsertEventProps } from './components/IGdprInsertEventProps';
import { GdprBaseWebPart } from '../../components/GDPRBaseWebPart';
export interface IGdprInsertEventWebPartProps {
  description: string;
}

export default class GdprInsertEventWebPart extends GdprBaseWebPart {

  public render(): void {
    const element: React.ReactElement<IGdprInsertEventProps > = React.createElement(
      GdprInsertEvent,
      {
        context: this.context,
        targetList: this.properties.targetList,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
