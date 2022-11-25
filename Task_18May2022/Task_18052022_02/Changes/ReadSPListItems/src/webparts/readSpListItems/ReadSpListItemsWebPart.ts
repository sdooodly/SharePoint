//import
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape, isEmpty } from "@microsoft/sp-lodash-subset";
import styles from "./ReadSpListItemsWebPart.module.scss";
import * as strings from "ReadSpListItemsWebPartStrings";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";

//Property config, change here first ???
export interface IReadSpListItemsWebPartProps {
  listName: string;
  listField: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Id: string;
  Title: string;
  Address: string;
  Number: string;
}

export default class ReadSpListItemsWebPart extends BaseClientSideWebPart<IReadSpListItemsWebPartProps> {
  private _getListData(): Promise<ISPLists> {
    //Change, 1. How to filter using REST API? 2. How to filter and select fields at the same time?
    //return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('${this.properties.listName}')?$select=Title,Id,Address,Number",SPHttpClient.configurations.v1)
    //Added if-else to display the entire list if listField(ID) is empty

    if (this.properties.listField === "") {
      return this.context.spHttpClient
        .get(
          `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items?$select=Title,Id,Address,Number`,
          //cross-check
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=nometadata",
              "odata-version": "",
            },
          }
        )

        .then((response: SPHttpClientResponse) => {
          return response.json();
        });
    } else {
      return this.context.spHttpClient
        .get(
          `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items?$filter=ID eq '${this.properties.listField}'`,

          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=nometadata",
              "odata-version": "",
            },
          }
        )

        .then((response: SPHttpClientResponse) => {
          return response.json();
        });
    }
  }

  private _renderListAsync(): void {
    if (
      Environment.type == EnvironmentType.SharePoint ||
      Environment.type == EnvironmentType.ClassicSharePoint
    ) {
      this._getListData().then((response) => {
        this._renderList(response.value);
      });
    }
  }

  private _renderList(items: ISPList[]): void {
    let html: string =
      '<table border=1 width=100% style="border-collapse: collapse;">';

    html += "<th>ID</th> <th>Title</th><th>Address</th><th>Number</th>";

    items.forEach((item: ISPList) => {
      html += `
    <tr>            
        <td>${item.Id}</td>
        <td>${item.Title}</td>
        <td>${item.Address}</td>
        <td>${item.Number}</td>
        </tr>
        `;
    });

    html += "</table>";

    const listContainer: Element =
      this.domElement.querySelector("#spListContainer");

    listContainer.innerHTML = html;
  }

  public render(): void {
    this.domElement.innerHTML = `

      <div class="${styles.readSpListItems}">

          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${
            styles.row
          }">
          <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
          <p class="ms-font-l ms-fontColor-white">Loading from ${
            this.context.pageContext.web.title
          }</p>
          <h1><i>List name:</i> ${escape(this.properties.listName)}</h1>
          <div><i>Purpose:</i> Sample application to read data from a SharePoint list.</div>
        </div>
      </div> 

          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${
            styles.row
          }">
          <br>
           <div id="spListContainer" />
        </div>

      </div>`;

    this._renderListAsync();
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
  //Property config
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Enter the list name and field below.",
          },

          groups: [
            {
              groupName: strings.BasicGroupName,

              groupFields: [
                PropertyPaneTextField("listName", {
                  label: strings.ListNameFieldLabel,
                }),
                PropertyPaneTextField("listField", {
                  label: "Enter ID of the required row",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
