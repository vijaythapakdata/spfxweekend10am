import * as React from 'react';
// import styles from './SpfxFormFunctional.module.scss';
import type { ISpfxFormFunctionalProps } from './ISpfxFormFunctionalProps';
import { ISpfxFormFunctionalState } from './ISpfxFormFunctionalState';
import {Web} from "@pnp/sp/presets/all";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import { Dialog } from '@microsoft/sp-dialog';
import { TextField,Slider,PrimaryButton, IDatePickerStrings, DatePicker, Dropdown, ChoiceGroup, IDropdownOption, Label } from '@fluentui/react';
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
    ManagerId:[],
    DOB:null,
    City:"",
    Department:"",
    Gender:"",
    Skills:[],
    Attachments:[]
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
  ManagerId:{results:formdata.ManagerId},
  DOB:new Date(formdata.DOB),
  CityId:formdata.City,
  Department:formdata.Department,
  Gender:formdata.Gender,
  Skills:{results:formdata.Skills}

});
const itemId=item.data.Id;
//upload multiple
for(const file of formdata.Attachments){
  const arrayBuffer=await file.arrayBuffer();
  await list.items.getById(itemId).attachmentFiles.add(file.name,arrayBuffer);
}
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
    ManagerId:[],
    DOB:null,
     City:"",
    Department:"",
    Gender:"",
    Skills:[],
    Attachments:[]
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
//get Managers
const getManagers=(items:any)=>{
  const managerName=items.map((i:any)=>i.text);
   const managerNameId=items.map((i:any)=>i.id);
   setFormData(a=>({...a,Manager:managerName,ManagerId:managerNameId}))
}

//skils 

const onSkillsChange=(event:React.FormEvent<HTMLInputElement>,options:IDropdownOption):void=>{
  const selectedKey=options.selected?[...formdata.Skills,options.key as string]:formdata.Skills.filter((key)=>key!==options.key);
  setFormData(a=>({...a,Skills:selectedKey}))
}
//file upload

const handleUploadFile=(event:React.ChangeEvent<HTMLInputElement>):void=>{
  const files=event.target.files;
  if(files){
    setFormData(prev=>({...prev,Attachments:Array.from(files)}));
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
    <PeoplePicker
    context={props.context as any}
    titleText="Managers"
    personSelectionLimit={3}
  
    showtooltip={true}
    
  
    onChange={getManagers}
  
    principalTypes={[PrincipalType.User]}
    resolveDelay={1000} 
    ensureUser={true}
    // defaultSelectedUsers={[formdata.Admin? formdata.Admin:""]} // to show the selected user on edit form
    defaultSelectedUsers={formdata.Manager}
    webAbsoluteUrl={props.siteurl}
    />
    <DatePicker
    label='Date of Birth'
    strings={DatePickerStrings}
    value={formdata.DOB}
    formatDate={FormateDate}
    onSelectDate={(date)=>setFormData(a=>({...a,DOB:date}))}
    />
    {/* City from Lookup Field */}
    <Dropdown
    label='City'
    options={props.CityOptions}
    selectedKey={formdata.City}
    onChange={(_,options)=>handleChange("City",options?.key as string)}
    />
    {/* Department single selected dropdown */}
    <Dropdown
    label='Department'
    options={props.DepartmentOptions}
    selectedKey={formdata.Department}
    onChange={(_,options)=>handleChange("Department",options?.key as string)}
    />
    {/* Gender as Radio button */}
    <ChoiceGroup
    label='Gender'
    options={props.GenderOptions}
    selectedKey={formdata.Gender}
    onChange={(_,options)=>handleChange("Gender",options?.key as string)}
    />
    {/* Skills multi select dropdown */}
    <Dropdown
    label='Skills'
    options={props.SkillsOptions}
    // selectedKey={formdata.City}
    defaultSelectedKeys={formdata.Skills}
    // onChange={(_,options)=>handleChange("City",options?.key as string)}
    onChange={onSkillsChange}
    multiSelect
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
    <Label>File Upload</Label>
    <input type='File' multiple onChange={handleUploadFile}/>
    <br/>
    <br/>
    <PrimaryButton text='Save' onClick={createForm} iconProps={{iconName:'save'}}/>
    </>
  )
}
export default SpfxFormFunctional;

//DatePicket Strings

export const DatePickerStrings:IDatePickerStrings={
  months:["January","February","March","April","May","June","July","August","September","October","November","December"],
  shortMonths:["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],
  days:["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"],
  shortDays:["Sun","Mon","Tue","Wed","Thu","Fri","Sat"],
  goToToday:"Go to today",
  prevMonthAriaLabel:"Go to previous month",
  nextMonthAriaLabel:"Go to next month",
  prevYearAriaLabel:"Go to previous year",
  nextYearAriaLabel:"Go to next year",
  closeButtonAriaLabel:"Close date picker"

}
export const FormateDate=(date:any):string=>{
  var date1=new Date(date);
  //get year
  var year=date1.getFullYear();
  //get month
  var month=(1+date1.getMonth()).toString();
  month=month.length>1?month:"0"+month;
  //get day
  var day=date1.getDate().toString();
  day=day.length>1?day:"0"+day;
  return month+"/"+day+"/"+year;
}