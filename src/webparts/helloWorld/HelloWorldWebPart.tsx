import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneHorizontalRule,
  PropertyPaneLabel,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'HelloWorldWebPartStrings';
import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';
import { PropertyPaneWrap } from 'property-pane-wrap';
import { GroupType, PeoplePicker, PersonType, TeamsChannelPicker } from '@microsoft/mgt-react';
import { update } from '@microsoft/sp-lodash-subset';
import { Providers } from '@microsoft/mgt-element';
import { SharePointProvider } from '@microsoft/mgt-sharepoint-provider';

export interface IHelloWorldWebPartProps {
  description: string;
  mgtTheme: string;
  mgtGroupPicker: string;
  mgtGroupMemberPicker: string;
  mgtPeoplePicker: string;
  mgtTeamsChannelPicker: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit().then(_ => {

      // other init code may be present
      Providers.globalProvider = new SharePointProvider(this.context);

    });
  }

  public render(): void {
    const element: React.ReactElement<IHelloWorldProps> = React.createElement(
      HelloWorld,
      {
        properties: this.properties,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  public updateWebPartProperty(property, value, refreshWebPart = true, refreshPropertyPane = true) {

    update(this.properties, property, () => value);
    if (refreshWebPart) this.render();
    if (refreshPropertyPane) this.context.propertyPane.refresh();

  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    let mgtTheme;
    switch (this.properties.mgtTheme) {
      case "inherit": mgtTheme = this._isDarkTheme ? "mgt-dark" : "mgt-light"; break;
      default: mgtTheme = this.properties.mgtTheme;
    }

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "MGT Controls",
              groupFields: [
                PropertyPaneHorizontalRule(),
                PropertyPaneChoiceGroup("mgtTheme", {
                  options: [
                    { key: "inherit", text: "Infer from context" },
                    { key: "mgt-dark", text: "mgt-dark" },
                    { key: "mgt-light", text: "mgt-light" }
                  ],
                  label: "MGT Theme"
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLabel("mgtPeoplePicker", { text: "MGT People Picker" }),
                PropertyPaneWrap("mgtPeoplePicker", {
                  component: PeoplePicker,
                  props: {
                    className: mgtTheme,
                    selectionMode: "single",
                    type: PersonType.person,
                    defaultSelectedUserIds: [this.properties.mgtPeoplePicker],
                    selectionChanged: (e: any) => {
                      let users = [];
                      e.detail.forEach(dtl => users.push(dtl.userPrincipalName));
                      this.updateWebPartProperty("mgtPeoplePicker", users[0]);
                    }
                  }
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLabel("mgtGroupPicker", { text: "MGT Group Picker" }),
                PropertyPaneWrap("mgtGroupPicker", {
                  component: PeoplePicker,
                  props: {
                    className: mgtTheme,
                    selectionMode: "single",
                    type: PersonType.group,
                    groupType: GroupType.unified,
                    defaultSelectedGroupIds: [this.properties.mgtGroupPicker],
                    selectionChanged: (e: any) => {
                      let users = [];
                      e.detail.forEach(dtl => users.push(dtl.id));
                      this.updateWebPartProperty("mgtGroupPicker", users[0]);
                      this.updateWebPartProperty("mgtGroupMemberPicker", "");
                    }
                  }
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLabel("mgtGroupMemberPicker", { text: "MGT Group Member Picker" }),
                PropertyPaneWrap("mgtGroupMemberPicker", {
                  component: PeoplePicker,
                  props: {
                    className: mgtTheme,
                    groupId: this.properties.mgtGroupPicker,
                    selectionMode: "single",
                    type: PersonType.person,
                    defaultSelectedUserIds: [this.properties.mgtGroupMemberPicker],
                    selectionChanged: (e: any) => {
                      let users = [];
                      e.detail.forEach(dtl => users.push(dtl.userPrincipalName));
                      this.updateWebPartProperty("mgtGroupMemberPicker", users[0]);
                    }
                  }
                }),
                // PropertyPaneHorizontalRule(),
                // PropertyPaneLabel("mgtTeamsChannelPicker", { text: "MGT Teams Channel Picker" }),
                // PropertyPaneWrap("mgtTeamsChannelPicker", {
                //   component: TeamsChannelPicker,
                //   props: {
                //     className: mgtTheme,
                //     // selectedItem: this.properties.mgtTeamsChannelPicker,
                //     // selectedChannel: this.properties.mgtTeamsChannelPicker,                    
                //     selectionChanged: (e: any) => {
                //       let slctns = [];
                //       console.log(e);
                //       e.detail.forEach(dtl => slctns.push(dtl.channel.id));
                //       this.updateWebPartProperty("mgtTeamsChannelPicker", slctns[0]);
                //     }
                //   }
                // }),
                PropertyPaneHorizontalRule()
              ]
            },
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneHorizontalRule(),
                PropertyPaneTextField('description', {
                  label: strings.PropertyPaneTextFieldLabel
                }),
                PropertyPaneHorizontalRule()
              ]
            }
          ]
        }
      ]
    };
  }
}
