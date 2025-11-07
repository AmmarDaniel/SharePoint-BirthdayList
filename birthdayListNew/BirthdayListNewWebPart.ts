import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import * as React from "react";
import * as ReactDom from "react-dom";

import BirthdayListNew from "./components/BirthdayListNew";
import { IBirthdayListNewProps } from "./components/IBirthdayListNewProps";

export interface IBirthdayListNewWebPartProps {
  listTitle: string;
  personFieldInternalName: string;
  showDepartment: boolean;
  showAge: boolean;
  maxItems: number;
  headerIconUrl: string;
}

export default class BirthdayListNewWebPart extends BaseClientSideWebPart<IBirthdayListNewWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IBirthdayListNewProps> =
      React.createElement(BirthdayListNew, {
        context: this.context,
        listTitle: this.properties.listTitle?.trim() ?? "EmployeeBirthdays",
        personFieldInternalName:
          this.properties.personFieldInternalName?.trim() ?? "Employee",
        showDepartment: this.properties.showDepartment ?? true,
        showAge: this.properties.showAge ?? false,
        maxItems: this.properties.maxItems ?? 5,
        headerIconUrl:
          this.properties.headerIconUrl?.trim() ??
          "https://www.iconarchive.com/download/i136009/microsoft/fluentui-emoji-3d/Birthday-Cake-3d.1024.png",
      });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Configure the Birthday List (New)" },
          groups: [
            {
              groupName: "Data",
              groupFields: [
                PropertyPaneTextField("listTitle", {
                  label: "SharePoint List Title",
                  placeholder: "EmployeeBirthdays",
                }),
                PropertyPaneTextField("personFieldInternalName", {
                  label: "Person field internal name",
                  placeholder: "Employee",
                }),
                PropertyPaneSlider("maxItems", {
                  label: "Visible rows (scroll to see more)",
                  min: 1,
                  max: 15,
                  step: 1,
                  showValue: true,
                }),
              ],
            },
            {
              groupName: "Display",
              groupFields: [
                PropertyPaneToggle("showDepartment", {
                  label: "Show department",
                  onText: "On",
                  offText: "Off",
                }),
                PropertyPaneToggle("showAge", {
                  label: "Show age (this year)",
                  onText: "On",
                  offText: "Off",
                }),
                PropertyPaneTextField("headerIconUrl", {
                  label: "Header icon URL (optional)",
                  placeholder: "https://.../cake.png",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
