import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import * as React from 'react';
import * as ReactDom from 'react-dom';
import DocFlipbook from './components/DocFlipbook';

export interface IDocFlipbookWebPartProps {
  libraryName: string;
}

export default class DocFlipbookWebPart extends BaseClientSideWebPart<IDocFlipbookWebPartProps> {

  public render(): void {
    const element: React.ReactElement = React.createElement(
      DocFlipbook,
      {
        context: this.context,
        libraryName: this.properties.libraryName || "Documents"
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}