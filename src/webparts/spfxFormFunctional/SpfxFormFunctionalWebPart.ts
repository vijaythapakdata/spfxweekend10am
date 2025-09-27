import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'SpfxFormFunctionalWebPartStrings';
import SpfxFormFunctional from './components/SpfxFormFunctional';
import { ISpfxFormFunctionalProps } from './components/ISpfxFormFunctionalProps';

export interface ISpfxFormFunctionalWebPartProps {
 ListName: string;
 CityOptions:any;
}

export default class SpfxFormFunctionalWebPart extends BaseClientSideWebPart<ISpfxFormFunctionalWebPartProps> {



  public async render():Promise<void> {
    const cityopt=await this.getLookupValue();
    const element: React.ReactElement<ISpfxFormFunctionalProps> = React.createElement(
      SpfxFormFunctional,
      {
        ListName:this.properties.ListName,
        siteurl:this.context.pageContext.web.absoluteUrl,
        context:this.context,
        DepartmentOptions:await this.getChoiceFields(this.context.pageContext.web.absoluteUrl,"Department",this.properties.ListName),
        GenderOptions:await this.getChoiceFields(this.context.pageContext.web.absoluteUrl,"Gender",this.properties.ListName),
        SkillsOptions:await this.getChoiceFields(this.context.pageContext.web.absoluteUrl,"Skills",this.properties.ListName),
        CityOptions:cityopt

      }
    );

    ReactDom.render(element, this.domElement);
  }
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
  //get choice fields
  private async getChoiceFields(siteurl:string,fieldName:string,ListName:string):Promise<any>{
    try{
const response=await fetch(`${siteurl}/_api/web/lists/getbytitle('${ListName}')/fields?$filter=EntityPropertyName eq '${fieldName}'`,{
  method:'GET',
  headers:{
    'Accept':'application/json;odata=nometadata'
  }
});
if(!response.ok){
  throw new Error(`Error while fetching the choice fields : ${response.status}`);
}
const data=await response.json();
const choices=data.value[0].Choices;
return choices.map((choice:any)=>({
  key:choice,
  text:choice
}))
    }
    catch(err){
console.error(err)
return[];
    }
  }

  //get lookup
  private async getLookupValue():Promise<any[]>{
    try{
const response=await fetch(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Cities')/items?$select=Title,ID`,{
  method:'GET',
  headers:{
    'Accept':'application/json;odata=nometadata'
  }
});
if(!response.ok){
  throw new Error(`Error while fetching the lookup fields : ${response.status}`);
}
const data=await response.json();
return data.value.map((city:{ID:string,Title:string})=>({
  key:city.ID,
  text:city.Title
}))
    }
    catch(err){
console.error(err)
return[];
    }
  }
}
