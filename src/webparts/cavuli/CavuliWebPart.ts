import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { ICavuliProps, Cavuli } from './components/Cavuli';


export interface ICavuliWebPartProps {
}

export default class CavuliWebPart extends BaseClientSideWebPart<ICavuliWebPartProps> {


  public render(): void {
    const element: React.ReactElement<ICavuliProps> = React.createElement(
      Cavuli,
      {
        displayMode: this.displayMode,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  // protected onInit(): Promise<void> {
  // }



  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{ groups: [] }]
    };
  }
}
