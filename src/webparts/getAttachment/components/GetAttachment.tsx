import * as React from 'react';
// import styles from './GetAttachment.module.scss';
import type { IGetAttachmentProps } from './IGetAttachmentProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import {Web} from "@pnp/sp/presets/all";
import { IGetAttachmentState } from './IGetAttachmentState';
const GetAttachment:React.FC<IGetAttachmentProps>=(props)=>{
const [stateAttachment,setAttachments]=React.useState<IGetAttachmentState>({
  Attachments:[]
});


//handle file

const handleFileChange=(event:React.ChangeEvent<HTMLInputElement>)=>{
  const files=event.target.files;
  if(files){
    setAttachments({Attachments:Array.from(files)});
  }
}
//upload file
const uploadDocuments=async()=>{
  try{
const web=Web(props.siteurl);
const list=web.lists.getByTitle(props.ListName);
//add an empty item first
const item=await list.items.add({});
const itemId=item.data.id;

//upload multiple docs
for(const file of stateAttachment.Attachments){
  const arrayBuffer=await file.arrayBuffer();
  await list.items.getById(itemId).attachmentFiles.add(file.name,arrayBuffer);
}

console.log("Files saved successfully");
  }
  catch(err){
console.error("Error while uploading the files");
  }
}
  return(
    <>
    <input
    type='file' onChange={handleFileChange}
    multiple
    
    />
    <button
    onClick={uploadDocuments}
    >Upload file</button>
    </>
  )
}
export default GetAttachment;