import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './NewsWidgetWebPart.module.scss';
import * as strings from 'NewsWidgetWebPartStrings';

import { SPComponentLoader } from '@microsoft/sp-loader';
import * as pnp from "sp-pnp-js";

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IODataList } from "@microsoft/sp-odata-types";
import { SPPermission } from '@microsoft/sp-page-context';
import { PermissionKind } from 'sp-pnp-js';



export interface INewsWidgetWebPartProps {
  description: string;
  listName: string;
  ShowDate: boolean;
  listId: string;
  footerURL : string;
  foundList : boolean;
}

var mythis;
const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
export default class NewsWidgetWebPart extends BaseClientSideWebPart<INewsWidgetWebPartProps> {

  private _dropdownOptions: IPropertyPaneDropdownOption[] = [];
  private listsFetched: boolean;
  private selListID: any;
  private listData = [];
  //private footerURL: string;
  //foundList: boolean = false;


  public onInit(): Promise<void> {
    this.properties.foundList = false;

    return super.onInit().then(_ => {

      pnp.setup({
        spfxContext: this.context
      });
    });
  }

  public constructor() {
    super();
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css');
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css');

    SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jquery/3.1.1/jquery.min.js', { globalExportsName: 'jQuery' }).then((jQuery: any): void => {
      SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/js/bootstrap.min.js', { globalExportsName: 'jQuery' }).then((): void => {
      });
    });
    //Custom CSS
    SPComponentLoader.loadCss('https://msletb.sharepoint.com/sites/FET/_catalogs/masterpage/School/css/style.css');

    //Variable to initiliaze list found or not
    
  }

  public getDataFromList(): void {

    pnp.sp.web.lists.getById(this.properties.listId).items.orderBy('Modified', false).get().then(function (result) {
      //console.log("Got News List Data:" + JSON.stringify(result));
      mythis.displayData(result);
    }, function (er) {
      alert("Oops, Something went wrong, Please try after sometime");
      console.log("Error:" + er);
    });
  }

  //Check for "body" column in selected list
  public getAllCloumns() {
    pnp.sp.web.lists.getById(this.properties.listId).fields.get().then(function (result) {
      result.forEach(function (val) {
        if (val.InternalName == 'body') {
          mythis.properties.foundList = true;
        }
      })

      //check whether list is found or not
      if (mythis.properties.foundList) {
        console.log('Found Column, List Proper');
        mythis.getDataFromList();
        mythis.properties.foundList = false;
      } else {
        var myHtml = '<h3>Please select correct list from Property Panel</h3>'
        var div = document.getElementById("newsList");
        div.innerHTML = myHtml;
      }
    }, function (er) {
      console.log('Error in getting columns:' + er)
    })



  }

  public displayData(data): void {
    try {
      if (this.properties.listName) {
        //var myHtml = "";

        //check whether column is there or not

        /*if (data[0].body === undefined) {
          myHtml = '<h3>Please select correct list from Property Panel</h3>'
          var div = document.getElementById("newsList");
          div.innerHTML = myHtml;
        } 
        else {*/

        var myHtml="";
        data.forEach(function (val) {
          var exDate = val.Expires ? val.Expires : "";
          var expireDate = new Date(exDate);
          var now = new Date();
          now.setHours(0, 0, 0, 0);
          //if (expireDate) {
          if (expireDate > now) {
            //console.log("Not Expired");
            var _Title = val.Title ? val.Title : "";
            var _ID = val.ID;
            var _Description = val.body ? val.body : "";
            var tmp = new Date(val.Created);
            var _CalDate = tmp.getDate();
            var _CalMonth = monthNames[tmp.getMonth() + 1];
            var tmp1 = tmp.getFullYear();
            var _CalYr = tmp1.toString().slice(-2);

            var regex = /(<([^>]+)>)/ig;

            var _Title_Long = _Title.replace(regex, "");
            var announcementDesc = _Description.replace(regex, "");

            var _Title_Display = _Title_Long;
            var _BodySubString = announcementDesc;

            //var newsURL = mythis.context.pageContext.web.absoluteUrl + "/Lists/" + mythis.properties.listName + "/DispForm.aspx?ID=" + _ID + "&Source=" + mythis.context.pageContext.web.absoluteUrl;
            var newsURL = mythis.context.pageContext.web.absoluteUrl + "/SitePages/News-Details.aspx?lid=" + mythis.properties.listName + "&nid=" + _ID;
            if (_Title_Display.length > 40) _Title_Display = _Title_Long.substring(0, 40) + '...'; _Title_Display;
            if (_BodySubString.length > 60) _BodySubString = announcementDesc.substring(0, 60) + '...'; _BodySubString;
            myHtml += "<li style='margin-top:15px'>" +
              "<div class='media-left'>" +
              "<div class='dt'>" + _CalDate + "<small>" + _CalMonth + " - " + _CalYr + "</small></div> " +
              "</div>" +
              "<div class='media-body' style='padding-left:15px'>" +
              "<h4 class='media-heading'><a href='" + newsURL + "' target='_blank'>" + _Title_Display + "</a></h4>" +
              "<p>" + _BodySubString + "</p>" +
              "</div>" +
              "</li>";

          }
          //}
        });
        //console.log(mythis.properties.ShowDate);

        var newsHTML = document.getElementById("newsList");
        newsHTML.innerHTML = myHtml;


        if (!mythis.properties.ShowDate) {

          // Loop through each element with the class name of slidePanel
          var elementList = document.querySelectorAll('.media-left');

          // Iterate through each element in the array
          for (var i = 0; i < elementList.length; i++) {
            //Create the MouseDown, MouseUp, and MouseMove events for the element
            var ele = elementList[i];
            ele.setAttribute('style', 'display:none');
          }
        }

        //Check Whether user has EDIT permission in site or not
        pnp.sp.web.currentUserHasPermissions(PermissionKind.EditListItems).then(perms => {

          console.log(perms);
          if (perms) {           

            var footHTML = "<div class='section-footer box-shadow'>" +
              "<span><a id='footerLink' href='" +mythis.properties.footerURL + "' target='_blank'> View All </span>" +
              "</div>";

            var footDIV = document.getElementById('footer');
            footDIV.innerHTML = footHTML;
          }
        });

        //}
      } else {
        myHtml = '<h3>Please select a list from Property Panel</h3>'
        var div = document.getElementById("newsList");
        div.innerHTML = myHtml;
      }
    }
    catch (err) {
      //console.log('Please Select A List');
      myHtml = '<h3>Please select a correct list from Property Panel</h3>'
      var div = document.getElementById("newsList");
      div.innerHTML = myHtml;
      console.log("Error Occured:" + err);
    }
  }
  public render(): void {


    mythis = this;


    if (!this.listsFetched) {
      this.fetchOptions().then((response) => {
        this._dropdownOptions = response;
        this.listsFetched = true;
        // now refresh the property pane, now that the promise has been resolved..
        //this.onDispose();
      });
    }
    var title = this.properties.listName ? this.properties.listName : "Please select a list"
    this.domElement.innerHTML = "<div class='section box-shadow' style='height:auto;'>" +
      "<div class='section-header'>" +
      "<h2 id='listTitle'>" + title + "</h2>" +
      "</div>" +
      "<div class='whats_new section-body'>" +
      "<ul class='news_list' id='newsList'>" +
      "</div>" +
      "</div>" +
      "<div id='footer'></div>";

    if (!this.properties.listName) {
      var myHtml = '<h3>Please select a list from Property Panel</h3>'
      var div = document.getElementById("newsList");
      div.innerHTML = myHtml;
    } else {
      //this.getDataFromList();
      this.getAllCloumns();
    }


  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    // ...
    if (this.properties.listName) {

      return;
    }

    this.fetchOptions().then((response) => {
      this._dropdownOptions = response;
      this.listsFetched = true;
      // now refresh the property pane, now that the promise has been resolved..
      //this.onDispose();
    });
    // ...
  }


  private fetchLists(url: string): Promise<any> {
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      if (response.ok) {
        return response.json();
      } else {
        console.log("WARNING - failed to hit URL " + url + ". Error = " + response.statusText);
        return null;
      }
    });
  }


  private fetchOptions(): Promise<IPropertyPaneDropdownOption[]> {
    var url = mythis.context.pageContext.web.absoluteUrl + '/_api/web/lists?$select=Title,RootFolder/ServerRelativeUrl,*&$expand=RootFolder&$filter=Hidden%20eq%20false';

    return this.fetchLists(url).then((response) => {
      var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
      var alllistdata = response.value;
      response.value.map((list: IODataList) => {
        //console.log("Found list with title = " + list.Title);
        options.push({ text: list.Title, key: list.Id + "|" + list.Title });
      });
      alllistdata.forEach(function (val) {
        mythis.listData.push({
          Title: val.Title ? val.Title : "",
          URL: val.RootFolder.ServerRelativeUrl ? val.RootFolder.ServerRelativeUrl : "#",
          ID: val.Id
        })
      })

      return options;
    });
  }

  protected onPropertyPaneFieldChanged(propertypath: string, oldval: any, newval: any) {

    //get list name
    if (propertypath == "displaylist") {
      var Ltitle;
      Ltitle = newval ? newval.split('|')[1] : "";
      this.properties.listId = newval ? newval.split('|')[0] : "";
      this.properties.listName = Ltitle;
      mythis.listData.forEach(function (val) {
        if (val.ID == mythis.properties.listId) {
          mythis.properties.footerURL = val.URL ? val.URL : "#";
        }
      })
      //this.getDataFromList();
      //this.getAllCloumns();
    }


    //show date or not
    if (propertypath == "displaydate") {
      this.properties.ShowDate = newval;
      console.log("Show Date:" + this.properties.ShowDate);


    }
  }


  /*protected get disableReactivePropertyChanges(): boolean { 
   return true; 
 }*/



  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Select list for News Webpart"
          },
          groups: [
            {
              groupFields: [
                PropertyPaneDropdown('displaylist', {
                  label: 'Select list for News Webpart',
                  options: this._dropdownOptions
                }),

              ]
            },
            {
              groupFields: [
                PropertyPaneToggle('displaydate', {
                  label: strings.ShowDate,
                  onText: 'Show Date',
                  offText: 'Hide Date',
                  checked: this.properties.ShowDate
                })
              ]
            }
          ]
        }
      ]
    }
  }
}
