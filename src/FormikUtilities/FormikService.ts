import {Web} from "@pnp/sp/presets/all";

export class Service{
    private web:any;
    constructor(siteurl:string){
        this.web=Web(siteurl);
    }

public async createItems(ListName:string,body:any){
    try{
let createItem=await this.web.lists.getByTitle(ListName).items.add(body);
return createItem;
    }
    catch(err){
console.log("Error in creating item:",err);
throw err;
    }
    finally{
        console.log("Create item function executed");
    }
}
}