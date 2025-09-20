import * as React from 'react';
// import styles from './SpfxFormFunctional.module.scss';
import type { ISpfxFormFunctionalProps } from './ISpfxFormFunctionalProps';
import { ISpfxFormFunctionalState } from './ISpfxFormFunctionalState';
import {Web} from "@pnp/sp/presets/all";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import { Dialog } from '@microsoft/sp-dialog';
import { TextField,Slider,PrimaryButton } from '@fluentui/react';
import {  PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
const SpfxFormFunctional:React.FC<ISpfxFormFunctionalProps>=(props)=>{
  const [formdata,setFormData]=React.useState<ISpfxFormFunctionalState>({
    Name:"",
    Age:"",
    FullAddress:"",
    Email:"",
    Score:1,
    Salary:"",
    Admin:"",
    AdminId:"",
    Manager:[],
    ManagerId:[]
  });
  //Create form function
  const createForm=async()=>{
    try{
// we are going to call site url
const web=Web(props.siteurl);
const list=web.lists.getByTitle(props.ListName); //it will store the list name
const item=await list.items.add({
  Title:formdata.Name,
  Age:parseInt(formdata.Age),
  Address:formdata.FullAddress,
  EmailAddress:formdata.Email,
  Score:formdata.Score,
  Salary:parseFloat(formdata.Salary),
  AdminId:formdata.AdminId,
});
Dialog.alert(`Item created successfully with ID ${item.data.Id}`);
console.log(item);
setFormData({
   Name:"",
    Age:"",
    FullAddress:"",
    Email:"",
    Score:1,
    Salary:"",
     Admin:"",
    AdminId:"",
    Manager:[],
    ManagerId:[]
})
    }
    catch(err){
console.log(err);
Dialog.alert(`Error while creating item`);
    }
  }
//form event
const handleChange=(fieldValue:keyof ISpfxFormFunctionalState,value:boolean|number|string)=>{
  setFormData(prev=>({...prev,[fieldValue]:value}));//[1,2,3,4],[5,6,7]=>...a,b=>[1,2,3],[2,3]

}
//get admins 
const _getAdmins=(items: any[]) =>{
 if(items.length>0){
  setFormData(a=>({...a,Admin:items[0].text,AdminId:items[0].id}))
 }
}
  return(
    <>
    <TextField
    label='Name'
    value={formdata.Name}
    onChange={(_,val)=>handleChange("Name",val||"")}
    placeholder='enter your name'
    iconProps={{iconName:'person'}}
    />
    <TextField
    label='Age'
    value={formdata.Age}
    onChange={(_,val)=>handleChange("Age",val||"")}
    // placeholder='enter your name'
    // iconProps={{iconName:'person'}}
    />
    <TextField
    label='Email Address'
    value={formdata.Email}
    onChange={(_,val)=>handleChange("Email",val||"")}
    placeholder='enter your email Address'
    iconProps={{iconName:'mail'}}
    />
    <TextField
    label='Salary'
    value={formdata.Salary}
    onChange={(_,val)=>handleChange("Salary",val||"")}
    // placeholder='enter your name'
    // iconProps={{iconName:'person'}}
    prefix='₹'
    suffix='INR'
    />
    <Slider
    label='Score'
    value={formdata.Score}
    onChange={(val)=>handleChange("Score",val)}
    min={1}
    max={100}
    step={1}
    />
    <PeoplePicker
    context={props.context as any}
    titleText="Admin"
    personSelectionLimit={1}
  
    showtooltip={true}
    
  
    onChange={_getAdmins}
  
    principalTypes={[PrincipalType.User]}
    resolveDelay={1000} 
    ensureUser={true}
    defaultSelectedUsers={[formdata.Admin? formdata.Admin:""]} // to show the selected user on edit form
    webAbsoluteUrl={props.siteurl}
    />
    <TextField
    label='Full Address'
    value={formdata.FullAddress}
    onChange={(_,val)=>handleChange("FullAddress",val||"")}
    placeholder='enter your address'
    iconProps={{iconName:'home'}}
    multiline
    rows={5}
    />
    <br/>
    <PrimaryButton text='Save' onClick={createForm} iconProps={{iconName:'save'}}/>
    </>
  )
}
export default SpfxFormFunctional;
