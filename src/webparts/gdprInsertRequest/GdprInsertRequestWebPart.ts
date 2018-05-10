import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'GdprInsertRequestWebPartStrings';
import GdprInsertRequest from './components/GdprInsertRequest';
import { IGdprInsertRequestProps } from './components/IGdprInsertRequestProps';
import { GdprBaseWebPart } from '../../components/GDPRBaseWebPart';


export default class GdprInsertRequestWebPart extends GdprBaseWebPart {

  public render(): void {
    const element: React.ReactElement<IGdprInsertRequestProps > = React.createElement(
      GdprInsertRequest,
      {
        context: this.context,
        targetList: this.properties.targetList

      }
    );

    ReactDom.render(element, this.domElement);
  }

}
