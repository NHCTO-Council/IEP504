/*
 * Apache License, Version 2.0
 *
 * Copyright (c) 2019 Axis Business Solutions
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at:
 *     http://www.apache.org/licenses/LICENSE-2.0
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { Version } from "@microsoft/sp-core-library";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneLink,
  PropertyPaneTextField,
  PropertyPaneToggle
} from "@microsoft/sp-webpart-base";
import * as strings from "Iep504DocumentUiWebPartStrings";
import * as $ from "jquery";
import * as Tabulator from "tabulator-tables";

import * as models from "../../shared/models";
import { SPData } from "../../shared/providers/SPData";
import { sp, Items } from "@pnp/sp";


export interface IIep504DocumentUiWebPartProps {
  docLibraryName: string;
  auditListName: string;
  uiMode: string;
  treatPreviewAsRead: boolean;
  allowExport: boolean;
  debugMode: boolean;
  authorLink: string;
  //multilineTextboxField: string;
  //checkboxField: boolean;
  //sliderField: number;
}
const loading: any = require("./assets/loading.gif");

export default class Iep504DocumentUiWebPart extends BaseClientSideWebPart<
  IIep504DocumentUiWebPartProps
  > {
  public onInit(): Promise<void> {

    return super.onInit().then(_ => {

      // other init code may be present

      sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    //TODO - Style elements need to be moved out of application
    this.domElement.innerHTML = `
    <img class="loading" src="${loading}" alt="Please wait" style="display:block; margin-left: auto; margin-right:auto; position:relative; z-index:999;" />
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.6.0/css/all.css"
    integrity="sha384-aOkxzJ5uQz7WBObEZcHvV5JvRW3TUc2rNPA7pe3AwnsUohiw1Vj2Rgx2KSOkF5+h"
    crossorigin="anonymous">
    <link href="https://unpkg.com/tabulator-tables@4.1.3/dist/css/bootstrap/tabulator_bootstrap4.min.css" rel="stylesheet">
    <div id="document-table"></div><pre id="debug" style="padding:5px; border:2px solid Black;display:none;"><h2>Debug Info:</pre>
    <style>
        /*Tabulator Styles */
    .tabulator-group-level-0.tabulator-group.tabulator-group-visible{
      background-color:beige!important;
    }
    .tabulator-group-level-1{
      background-color:#c9c9c9!important;
    }
    .tabulator-table {min-width:100%!important;}
    </style>
    `;
    //TODO : This query logic to be replaced with PNP and moved to data provider.
    var siteUrl = this.context.pageContext.web.absoluteUrl;
    var docLibUrl =
      siteUrl +
      `/_api/web/lists/getbytitle('${escape(
        this.properties.docLibraryName
      )}')/items?`;

    var baseQuery =
      "$Select=*,File/Name,Teachers/Id,Teachers/Title,Teachers/EMail,CaseManager/Id,CaseManager/Title,CaseManager/EMail,StudentWebId&$expand=File,Teachers,CaseManager";
    var finalQuery;
    switch ("" + escape(this.properties.uiMode)) {
      case "Admin":
        finalQuery = baseQuery;
        break;
      case "Manager":
        finalQuery =
          baseQuery +
          "&$filter=(CaseManager/Id eq " +
          this.context.pageContext.legacyPageContext["userId"] +
          ")";
        break;
      case "Teacher":
        finalQuery =
          baseQuery +
          "&$filter=(Teachers/Id eq " +
          this.context.pageContext.legacyPageContext["userId"] +
          ")";
        break;
      default:
        finalQuery = "";
    }
    var documentData;
    $(document).ready(() => {
      $.ajax({
        url: docLibUrl + finalQuery + '&$Top=5000',
        type: "GET",
        headers: { Accept: "application/json;odata=nometadata" },
        cache: false
      })
        .done(data => {
          documentData = [];
          var docItem: models.IEP504Document, teacher: models.IEP504Teacher;
          data.value.forEach(item => {
            docItem = new models.IEP504Document();
            docItem.fileName = item.File.Name;
            docItem.modified = item.Modified;
            docItem.student = new models.IEP504Student();
            docItem.student.firstName = item.StudentFirstName;
            docItem.student.lastName = item.StudentLastName;
            docItem.student.id = item.StudentWebId;

            try {
              docItem.student.caseManager = new models.IEP504CaseManager(
                item.CaseManager.Id
              );
              docItem.student.caseManager.displayName = item.CaseManager.Title;
              docItem.student.caseManager.email = item.CaseManager.EMail;
            } catch (e) {
              docItem.student.caseManager = new models.IEP504CaseManager(null);
              docItem.student.caseManager.displayName = "Unknown";
              console.warn(
                "Warning: Check Case Manager Assignments for student '" +
                docItem.student.id +
                "'"
              );
            }
            docItem.teachers = [];
            try {
              item.Teachers.forEach(teacherItem => {
                teacher = new models.IEP504Teacher(teacherItem.Id);
                teacher.displayName = teacherItem.Title;
                teacher.email = teacherItem.EMail;
                docItem.teachers.push(teacher);
              });
            } catch (e) {
              var teacherError = new models.IEP504Teacher(null);
              teacherError.displayName =
                "Error: Could not associate document to a teacher.";
              docItem.teachers.push(teacherError);
              //console.log(e);
              console.warn(
                "Warning: Check Teacher Assignments for document '" +
                docItem.fileName +
                "'"
              );
            }

            docItem.SetHeaderRow();
            documentData.push(docItem);
          });
        })
        .fail((jqXHR, textStatus) => {
          console.log("Request failed: " + textStatus);
        })
        .always(() => {
          documentData.forEach((doc: models.IEP504Document) => {
            doc.teachers.forEach((teacher: models.IEP504Teacher) => {
              let spData = new SPData(this.properties, siteUrl);
              // sp.GetUserById(teacher.id)
              //   .then(function (doc_data) { teacher.displayName = doc_data.Title });
              spData.GetFirstAccessedDate(
                doc.fileName,
                doc.modified,
                teacher.id
              ).then(doc_access_data => {
                teacher.documentAccessed = doc_access_data;
              });
            });
          });
        });
    });
    $(document).ajaxStop(() => {
      if (this.properties.debugMode) {
        $("#debug").show();
        $("#debug").append(JSON.stringify(documentData, null, 4));
      }
      $(".loading").hide();

      //create Tabulator on DOM element with id "document-table"
      var table = new Tabulator("#document-table", {
        columnMinWidth: 80,
        data: documentData,
        layout: "fitDataFill", //fit columns to width of table (optional)
        placeholder: "No data is available to display in this view.",
        groupHeader: (value, count, data, group) => {
          var countUnopenedDocs = 0;
          var docCount = 0;
          var infoString = "";
          try {
            var docGroup = group.getSubGroups();
            docCount = docGroup.length;
            docGroup.forEach(docItem => {
              countUnopenedDocs += docItem
                .getRows()[0]
                .getData()
                .teachers.filter(x => x.documentAccessed == null).length;
            });
          } catch (err) {
            console.log(err);
          }

          if (countUnopenedDocs > 0) {
            infoString +=
              " <span style='color:orange; margin-left:10px;'><i class='fas fa-exclamation-triangle'></i></span> " +
              countUnopenedDocs +
              " warning(s).";
          }
          // uncomment the next line to include a document count warning.
          // if (docCount > 1) { infoString += " <span style='color:orange; margin-left:10px;'><i class='fas fa-exclamation-triangle'></i></span> This student has multiple documents. (" + docCount + ") "; }
          if (infoString.length > 0) {
            return value + infoString;
          } else {
            return value;
          }
        },
        groupToggleElement: "header",
        groupBy: [
          data => {
            return data.headerRow;
          },
          data => {
            return data.fileName;
          }
        ],
        groupStartOpen: [
          false,
          (value, count, data, group) => {
            return group.getParentGroup().getSubGroups().length < 2;
          }
        ],
        columns: [
          /*No Columns Specified, since details are shown in group headers.*/
        ],
        rowFormatter: row => {
          //create and style holder elements
          var holderEl = document.createElement("div");
          var tableEl = document.createElement("div");

          holderEl.style.boxSizing = "border-box";
          holderEl.style.padding = "10px 40px 10px 40px";
          holderEl.style.borderTop = "1px solid #333";
          holderEl.style.borderBottom = "1px solid #333";
          holderEl.style.background = "#ddd";

          tableEl.style.border = "1px solid #333";

          holderEl.appendChild(tableEl);

          row.getElement().appendChild(holderEl);

          var subTable = new Tabulator(tableEl, {
            columnMinWidth: 200,
            layout: "fitDataFill",
            initialSort: [
              { column: "documentAccessed", dir: "asc" } //then sort by this second
            ],
            data: row.getData().teachers,
            columns: [
              { title: "Teacher", field: "displayName" },
              {
                title: "Opened On",
                field: "documentAccessed",
                formatter: cell => {
                  if (!cell.getValue()) {
                    return "Never";
                  } else {
                    return new Date(cell.getValue()).toLocaleString();
                  }
                }
              }
            ],
            rowFormatter: thisRow => {
              if (!thisRow.getData().documentAccessed) {
                thisRow.getElement().style.backgroundColor = "Red";
                thisRow.getElement().style.color = "White";
                thisRow.getElement().style.fontWeight = "Bold";
              }
            }
          });
          if (this.properties.uiMode == "Teacher") {
            subTable.setFilter(
              "id",
              "=",
              this.context.pageContext.legacyPageContext["userId"]
            );
          }
        }
      });
      if (this.properties.uiMode == "Admin") {
        // if the UIMode is set to administrator, add a top-level grouping by CaseManager.

        table.setGroupHeader((value, count, data, group) => {
          var countUnopenedDocs = 0;
          var docCount = 0;
          var infoString = "";
          try {
            var docGroup = group.getSubGroups();
            docCount = docGroup.length;
            docGroup.forEach(docItem => {
              countUnopenedDocs += docItem
                .getRows()[0]
                .getData()
                .teachers.filter(x => {
                  return x.documentAccessed == null;
                }).length;
            });
          } catch (err) {
            //bury the exception.
          }

          if (countUnopenedDocs > 0) {
            infoString +=
              " <span style='color:orange; margin-left:10px;'><i class='fas fa-exclamation-triangle'></i></span> " +
              countUnopenedDocs +
              " warning(s).";
          }
          // uncomment the next line to include a document count warning.
          // if (docCount > 1) { infoString += " <span style='color:orange; margin-left:10px;'><i class='fas fa-exclamation-triangle'></i></span> This student has multiple documents. (" + docCount + ") "; }
          if (infoString.length > 0) {
            return value + infoString;
          } else {
            return value;
          }
        });

        table.setGroupBy([
          data => {
            return "Case Manager: " + data.student.caseManager.displayName;
          },
          data => {
            return data.headerRow;
          },
          data => {
            return data.fileName;
          }
        ]);

        table.setGroupStartOpen([
          false,
          false,
          (value, count, data, group) => {
            return group.getParentGroup().getSubGroups().length < 2;
          }
        ]);
      }
    });
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
                PropertyPaneTextField("docLibraryName", {
                  label: strings.docLibraryNameFieldLabel
                }),
                PropertyPaneTextField("auditListName", {
                  label: strings.auditListNameFieldLabel
                }),
                PropertyPaneDropdown("uiMode", {
                  label: strings.uiModeLabel,
                  options: [
                    { key: "Teacher", text: "Teacher's Console" },
                    { key: "Manager", text: "Case Manager Console" },
                    { key: "Admin", text: "Administrator Console" }
                  ]
                }),
                PropertyPaneToggle("treatPreviewAsRead", {
                  label: strings.treatPreviewAsReadLabel,
                  onText: "Yes",
                  offText: "No"
                }),
                PropertyPaneToggle("allowExport", {
                  label: strings.allowExportLabel,
                  onText: "Yes",
                  offText: "No"
                }),
                PropertyPaneToggle("debugMode", {
                  label: strings.debugModeLabel,
                  onText: "On",
                  offText: "Off"
                }),
                PropertyPaneLabel("blankRow", {
                  text: ""
                }),
                PropertyPaneLabel("blankRow", {
                  text: ""
                }),
                PropertyPaneLabel("blankRow", {
                  text: ""
                }),
                PropertyPaneLabel("blankRow", {
                  text: "Version: " + this.manifest.version
                }),
                PropertyPaneLink("authorLink", {
                  text: strings.copyright,
                  href: "https://axisbusiness.com",
                  target: "_blank"
                })
                // PropertyPaneTextField('multilineTextboxField', {
                //   label: 'Multi-line Textbox label',
                //   multiline: true
                // }),
                // PropertyPaneCheckbox('checkboxField', {
                //   text: 'Checkbox text'
                // }),

                // PropertyPaneSlider('sliderField', {
                //   label: 'Slider label',
                //   min: 0,
                //   max: 100
                // })
              ]
            }
          ]
        }
      ]
    };
  }
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }
}
