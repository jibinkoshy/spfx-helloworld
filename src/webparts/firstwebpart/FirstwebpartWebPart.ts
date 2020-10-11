import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./FirstwebpartWebPart.module.scss";
import * as strings from "FirstwebpartWebPartStrings";

export interface IFirstwebpartWebPartProps {
  ListTitle: string;
  ListUrl: string;
  PercentCompleted: string;
  ValidationRequired: boolean;
  ListName: string;
}

export default class FirstwebpartWebPart extends BaseClientSideWebPart<
  IFirstwebpartWebPartProps
> {
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.firstwebpart}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.title}">Welcome to Jibin's World!!</span>
              <p class="${
                styles.subTitle
              }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${styles.description}">${escape(
      this.properties.ListTitle
    )}</p>
    <p>${this.properties.ListUrl}</p>
    </p>
    <p>${this.properties.PercentCompleted}</p>
            
            </div>
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  public ValidateListUrl(value: string): string {
    if (value.length > 256) return "URL should be less than 256 character";
    if (value.length === 0) return "Enter the list URL";
    return "";
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("ListTitle", {
                  label: strings.ListFieldLabel,
                }),
                PropertyPaneTextField("ListUrl", {
                  label: strings.ListUrl,
                  onGetErrorMessage: this.ValidateListUrl.bind(this),
                }),
                PropertyPaneSlider("PercentCompleted", {
                  label: strings.PercentCompleted,
                  min: 0,
                  max: 10,
                  value: 0,
                }),
                PropertyPaneCheckbox("ValidationRequired", {
                  checked: false,
                  text: "Validation Required",
                }),
                PropertyPaneDropdown("ListName", {
                  label: "Select your list",
                  selectedKey: "--Select YOur List--",
                  options: [
                    {
                      key: "--Select Your List--",
                      text: "--Select Your List--",
                    },
                    {
                      key: "Documents",
                      text: "Documents",
                    },
                    {
                      key: "Test",
                      text: "Test",
                    },
                    {
                      key: "Hands-On",
                      text: "Hands-On",
                    },
                  ],
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
