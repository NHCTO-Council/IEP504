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
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import { Version } from "@microsoft/sp-core-library";
import { escape } from "@microsoft/sp-lodash-subset";
import { BaseClientSideWebPart, PropertyPaneDropdown, PropertyPaneLabel, PropertyPaneLink, PropertyPaneTextField, PropertyPaneToggle } from "@microsoft/sp-webpart-base";
import * as strings from "Iep504DocumentUiWebPartStrings";
import * as $ from "jquery";
import * as Tabulator from "tabulator-tables";
import * as models from "../../shared/models";
import { SPData } from "../../shared/providers/SPData";
var loading = require("./assets/loading.gif");
var Iep504DocumentUiWebPart = (function (_super) {
    __extends(Iep504DocumentUiWebPart, _super);
    function Iep504DocumentUiWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Iep504DocumentUiWebPart.prototype.render = function () {
        var _this = this;
        this.domElement.innerHTML = "\n    <img class=\"loading\" src=\"" + loading + "\" alt=\"Please wait\" style=\"display:block; margin-left: auto; margin-right:auto; position:relative; z-index:999;\" />\n    <link rel=\"stylesheet\" href=\"https://use.fontawesome.com/releases/v5.6.0/css/all.css\" \n    integrity=\"sha384-aOkxzJ5uQz7WBObEZcHvV5JvRW3TUc2rNPA7pe3AwnsUohiw1Vj2Rgx2KSOkF5+h\"\n    crossorigin=\"anonymous\">\n    <link href=\"https://unpkg.com/tabulator-tables@4.1.3/dist/css/bootstrap/tabulator_bootstrap4.min.css\" rel=\"stylesheet\">\n    <div id=\"document-table\"></div><pre id=\"debug\" style=\"padding:5px; border:2px solid Black;display:none;\"><h2>Debug Info:</pre>\n    <style>\n        /*Tabulator Styles */\n    .tabulator-group-level-0.tabulator-group.tabulator-group-visible{\n      background-color:beige!important;\n    }\n    .tabulator-group-level-1{\n      background-color:#c9c9c9!important;\n    }\n    .tabulator-table {min-width:100%!important;}\n    </style>\n    ";
        var wpProperties = this.properties;
        var siteUrl = this.context.pageContext.web.absoluteUrl;
        var docLibUrl = siteUrl +
            ("/_api/web/lists/getbytitle('" + escape(this.properties.docLibraryName) + "')/items?");
        var baseQuery = "$Select=*,File/Name,Teachers/Id,Teachers/Title,Teachers/EMail,CaseManager/Id,CaseManager/Title,CaseManager/EMail,StudentWebId&$expand=File,Teachers,CaseManager";
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
        $(document).ready(function () {
            $.ajax({
                url: docLibUrl + finalQuery,
                type: "GET",
                headers: { Accept: "application/json;odata=nometadata" },
                cache: false
            })
                .done(function (data) {
                documentData = [];
                var docItem, teacher;
                data.value.forEach(function (item) {
                    docItem = new models.IEP504Document();
                    docItem.fileName = item.File.Name;
                    docItem.modified = item.Modified;
                    docItem.student = new models.IEP504Student();
                    docItem.student.firstName = item.StudentFirstName;
                    docItem.student.lastName = item.StudentLastName;
                    docItem.student.id = item.StudentWebId;
                    try {
                        docItem.student.caseManager = new models.IEP504CaseManager(item.CaseManager.Id);
                        docItem.student.caseManager.displayName = item.CaseManager.Title;
                        docItem.student.caseManager.email = item.CaseManager.EMail;
                    }
                    catch (e) {
                        docItem.student.caseManager = new models.IEP504CaseManager(null);
                        docItem.student.caseManager.displayName = "Unknown";
                        console.warn("Warning: Check Case Manager Assignments for student '" +
                            docItem.student.id +
                            "'");
                    }
                    docItem.teachers = [];
                    try {
                        item.Teachers.forEach(function (teacherItem) {
                            teacher = new models.IEP504Teacher(teacherItem.Id);
                            teacher.displayName = teacherItem.Title;
                            teacher.email = teacherItem.EMail;
                            docItem.teachers.push(teacher);
                        });
                    }
                    catch (e) {
                        var teacherError = new models.IEP504Teacher(null);
                        teacherError.displayName =
                            "Error: Could not associate document to a teacher.";
                        docItem.teachers.push(teacherError);
                        //console.log(e);
                        console.warn("Warning: Check Teacher Assignments for document '" +
                            docItem.fileName +
                            "'");
                    }
                    docItem.SetHeaderRow();
                    documentData.push(docItem);
                });
            })
                .fail(function (jqXHR, textStatus) {
                console.log("Request failed: " + textStatus);
            })
                .always(function () {
                documentData.forEach(function (doc) {
                    doc.teachers.forEach(function (teacher) {
                        var sp = new SPData(wpProperties, siteUrl);
                        // sp.GetUserById(teacher.id)
                        //   .then(function (doc_data) { teacher.displayName = doc_data.Title });
                        sp.GetFirstAccessedDate(doc.fileName, doc.modified, teacher.id).then(function (doc_access_data) {
                            teacher.documentAccessed = doc_access_data;
                        });
                    });
                });
            });
        });
        $(document).ajaxStop(function () {
            if (wpProperties.debugMode) {
                $("#debug").show();
                $("#debug").append(JSON.stringify(documentData, null, 4));
            }
            $(".loading").hide();
            //create Tabulator on DOM element with id "document-table"
            var table = new Tabulator("#document-table", {
                columnMinWidth: 80,
                data: documentData,
                layout: "fitDataFill",
                placeholder: "No data is available to display in this view.",
                groupHeader: function (value, count, data, group) {
                    var countUnopenedDocs = 0;
                    var docCount = 0;
                    var infoString = "";
                    try {
                        var docGroup = group.getSubGroups();
                        docCount = docGroup.length;
                        docGroup.forEach(function (docItem) {
                            countUnopenedDocs += docItem
                                .getRows()[0]
                                .getData()
                                .teachers.filter(function (x) { return x.documentAccessed == null; }).length;
                        });
                    }
                    catch (err) {
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
                    }
                    else {
                        return value;
                    }
                },
                groupToggleElement: "header",
                groupBy: [
                    function (data) {
                        return data.headerRow;
                    },
                    function (data) {
                        return data.fileName;
                    }
                ],
                groupStartOpen: [
                    false,
                    function (value, count, data, group) {
                        return group.getParentGroup().getSubGroups().length < 2;
                    }
                ],
                columns: [],
                rowFormatter: function (row) {
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
                                formatter: function (cell) {
                                    if (!cell.getValue()) {
                                        return "Never";
                                    }
                                    else {
                                        return new Date(cell.getValue()).toLocaleString();
                                    }
                                }
                            }
                        ],
                        rowFormatter: function (thisRow) {
                            if (!thisRow.getData().documentAccessed) {
                                thisRow.getElement().style.backgroundColor = "Red";
                                thisRow.getElement().style.color = "White";
                                thisRow.getElement().style.fontWeight = "Bold";
                            }
                        }
                    });
                    if (wpProperties.uiMode == "Teacher") {
                        subTable.setFilter("id", "=", _this.context.pageContext.legacyPageContext["userId"]);
                    }
                }
            });
            if (wpProperties.uiMode == "Admin") {
                // if the UIMode is set to administrator, add a top-level grouping by CaseManager.
                table.setGroupHeader(function (value, count, data, group) {
                    var countUnopenedDocs = 0;
                    var docCount = 0;
                    var infoString = "";
                    try {
                        var docGroup = group.getSubGroups();
                        docCount = docGroup.length;
                        docGroup.forEach(function (docItem) {
                            countUnopenedDocs += docItem
                                .getRows()[0]
                                .getData()
                                .teachers.filter(function (x) {
                                return x.documentAccessed == null;
                            }).length;
                        });
                    }
                    catch (err) {
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
                    }
                    else {
                        return value;
                    }
                });
                table.setGroupBy([
                    function (data) {
                        return "Case Manager: " + data.student.caseManager.displayName;
                    },
                    function (data) {
                        return data.headerRow;
                    },
                    function (data) {
                        return data.fileName;
                    }
                ]);
                table.setGroupStartOpen([
                    false,
                    false,
                    function (value, count, data, group) {
                        return group.getParentGroup().getSubGroups().length < 2;
                    }
                ]);
            }
        });
    };
    Object.defineProperty(Iep504DocumentUiWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse("1.0");
        },
        enumerable: true,
        configurable: true
    });
    Iep504DocumentUiWebPart.prototype.getPropertyPaneConfiguration = function () {
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
    };
    Object.defineProperty(Iep504DocumentUiWebPart.prototype, "disableReactivePropertyChanges", {
        get: function () {
            return true;
        },
        enumerable: true,
        configurable: true
    });
    return Iep504DocumentUiWebPart;
}(BaseClientSideWebPart));
export default Iep504DocumentUiWebPart;
//# sourceMappingURL=Iep504DocumentUiWebPart.js.map