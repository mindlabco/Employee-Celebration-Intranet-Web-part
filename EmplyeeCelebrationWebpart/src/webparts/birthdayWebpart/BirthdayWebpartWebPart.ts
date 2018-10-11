import { SPComponentLoader } from '@microsoft/sp-loader';
import * as pnp from "sp-pnp-js";
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './BirthdayWebpartWebPart.module.scss';
import * as strings from 'BirthdayWebpartWebPartStrings';

require('./app/style.css');

export interface IBirthdayWebpartWebPartProps {
  description: string;
}
export default class BirthdayWebpartWebPart extends BaseClientSideWebPart<IBirthdayWebpartWebPartProps> {


  public constructor() {
    super();
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css');
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css');

    SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jquery/3.1.1/jquery.min.js', { globalExportsName: 'jQuery' }).then((jQuery: any): void => {
      SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/js/bootstrap.min.js',  { globalExportsName: 'jQuery' }).then((): void => {        
      });
    });
  }

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      pnp.setup({
        spfxContext: this.context
      });
      
    });
  }

  public getDataFromList():void {
    var mythis =this;
    pnp.sp.web.lists.getByTitle('EmployeeCelebrations').items.get().then(function(result){
      console.log("Got List Data:"+JSON.stringify(result));
      mythis.displayData(result);
    },function(er){
      alert("Oops, Something went wrong, Please try after sometime");
      console.log("Error:"+er);
    });


  }

  public displayData(data):void{
    data.forEach(function(val){
      var img = val.Employee_Img?val.Employee_Img.Url:"https://dtgxwmigmg3gc.cloudfront.net/files/58a11f43777a4239b21f5475-icon-256x256.png";
      var myHtml = 
      '<img src="'+img+'" alt="Image" class="img-responsive img-circle contactimg resp"/>'+
      '<p class="name">'+val.Title+'</p>'+
      '<p class="desc">'+val.qpwe+'</p>'+
      '<hr>';
        var div = document.getElementById("bdayCelebration");
        div.innerHTML+=myHtml;
    });
    
  }

  public render(): void {
    this.domElement.innerHTML = `<div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
    <div class="card card-stats events-news">
        <div class="card-header"  style="background-color: #da3b01!important;">
            Events
        </div>
        <div class="card-content panel-body rowtop">
            <section class="people slider" id="bdayCelebration">


            </section>
        </div>
        <div class="panel-footer" style="text-align:center color:#337ab7;">
            <a href="/sites/Intranet/SPFX/Lists/EmployeeCelebrations/AllItems.aspx" target="_blank">Read more</a>
        </div>
    </div>

</div>`;

  this.getDataFromList();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
