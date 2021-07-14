import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { graph } from "@pnp/graph";
import * as strings from 'PortalWebPartStrings';
import Portal from './components/Portal';
import { IPortalProps } from './components/Portal';
import * as msTeams from "@microsoft/teams-js";
import { initializeIcons } from '@fluentui/react';
import { TeamsTheme, ThemeManager } from './utils/ThemeManager';
import { MSGraphClient } from '@microsoft/sp-http';

export interface IPortalWebPartProps {
  description: string;
}

export default class PortalWebPart extends BaseClientSideWebPart<IPortalWebPartProps> {
  private graphClient: MSGraphClient;

  public async onInit(): Promise<void> {
    this.graphClient = await this.context.msGraphClientFactory.getClient();

    return super.onInit().then(_ => {
      if (this.context.sdks.microsoftTeams) {
        initializeIcons();
        msTeams.initialize();

        let teamsTheme = this.context.sdks.microsoftTeams.context.theme;
        let currentTheme: TeamsTheme = TeamsTheme[teamsTheme];
        ThemeManager.applyTeamsTheme(currentTheme);

        this.context.sdks.microsoftTeams.teamsJs.registerOnThemeChangeHandler((theme) => {
          let t: TeamsTheme = TeamsTheme[theme];
          ThemeManager.applyTeamsTheme(t);
          this.render();
        });
      }
      graph.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<IPortalProps> = React.createElement(
      Portal, {
      graphClient: this.graphClient,
    }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
