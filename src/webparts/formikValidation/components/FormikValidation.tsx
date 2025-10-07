import * as React from 'react';
import styles from './FormikValidation.module.scss';
import type { IFormikValidationProps } from './IFormikValidationProps';
import { Service } from '../../../FormikUtilities/FormikService';
import { sp } from '@pnp/sp/presets/all';
import *as Yup from 'yup';
import { Formik,FormikProps } from 'formik';
import { Dialog } from '@microsoft/sp-dialog';
import { Dropdown, Label, Stack, TextField } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
const stackTokens={childrenGap:20};

const FormikValidation:React.FC<IFormikValidationProps>=(props)=>{
  const [service,setService]=React.useState<Service|null>(null);

  React.useEffect(()=>{
    sp.setup({
      spfxContext:props.context as any
    });
    setService(new Service(props.siteurl));

  },[props.context,props.siteurl]);

  const validationForm=Yup.object().shape({
    name:Yup.string().required("Task name is required"),
    details:Yup.string().min(15,"Minimum 15 characters are required ").required("Task details are required"),
    startDate:Yup.date().required("Start date is required"),
    endDate:Yup.date().required("End date is required"),
    phoneNumber:Yup.string().required("Phone number is required").matches(/^[0-9]{10}$/,"Phone number must be 10 digits"),
    emailAddres:Yup.string().email("Email address is not valid").required("Email is required"),
    projectName:Yup.string().required("Project name is required")
  })

  const getFieldProps=(formik:FormikProps<any>,field:string)=>({
    ...formik.getFieldHelpers(field),errorMessage:formik.errors[field] as string
  });

  const createRecord=async(record:any)=>{
  try{
const item=await service?.createItems(props.ListName,{
  Title:record.name,
  TaskDetails:record.details,
  StartDate:record.startDate,
  EndDate:record.endDate,
  ProjectName:record.projectName,
  PhoneNumber:record.phoneNumber,
  EmailAddress:record.emailAddress
});
console.log(item);
Dialog.alert("Saved successfullly");
  }
  catch(err){
console.log("error in creating record ",err);
  }
  }
  return(
    <>
      <Formik
       initialValues={{
        name:"",
        details:"",
        startDate:null,
        endDate:null,
        phoneNumber:"",
        emailAddress:"",
        projectnName:""
       }}
       validationSchema={validationForm}
       onSubmit={(values,helpers)=>{
        createRecord(values).then(()=>helpers.resetForm())
       }}
     >
{(formik:FormikProps<any>)=>(
  <form onSubmit={formik.handleSubmit}>
    <div className={styles.form}>
      <Stack tokens={stackTokens}>
        <Label className={styles.lbl}>User Name</Label>
 <PeoplePicker
    context={props.context as any}
    personSelectionLimit={1}
    showtooltip={true}
    principalTypes={[PrincipalType.User]}
 
    ensureUser={true}
    defaultSelectedUsers={[props.context.pageContext.user.displayName as any]} // to show the selected user on edit form
    webAbsoluteUrl={props.siteurl}
    disabled={true}
    />
    <Label className={styles.lbl}>Task Name</Label>
    <TextField
    {...getFieldProps(formik,"name")}
    />
       <Label className={styles.lbl}>Email Address</Label>
    <TextField
    {...getFieldProps(formik,"emailAddress")}
    />
       <Label className={styles.lbl}>Phone Number</Label>
    <TextField
    {...getFieldProps(formik,"phoneNumber")}
    />
      <Label className={styles.lbl}>Project Name</Label>
      <Dropdown
      options={[
        {key:'IT',text:'IT'},
        {key:'HR',text:'HR'}
      ]}
      selectedKey={formik.values.projectName}
      onChange={(_,options)=>formik.setFieldValue("ProjectName",options?.key as string)}
      errorMessage={formik.errors.projectName as string}
      />
      </Stack>
    </div>

  </form>
)}

     </Formik>
    </>
  )
}
export default FormikValidation;
