import * as React from 'react';
// import styles from './Crud.module.scss';
import type { ICrudProps } from './ICrudProps';
import { spfi,SPFx } from '@pnp/sp/presets/all';


interface ICrudState{
  id:number;
  name:string;
  email:string;
  admin:any;
  joiningDate:any
}
const Crud:React.FC<ICrudProps>=(props)=>{
  const _sp=spfi().using(SPFx(props.context));
  const [items,setItems]=React.useState<ICrudState[]>([]);
  const [loading,setLoading]=React.useState<boolean>(true);
  const [addHidden,setIsAddHidden]=React.useState<boolean>(true);
  const[newName,setNewName]=React.useState<string>('');
  const[newEmail,setNewEmail]=React.useState<string>('');
  const [newAdmin,setNewAdmin]=React.useState<{id:number,text:string}>();
  const[newJoiningDate,setNewJoiningDate]=React.useState<string|any>("");

  React.useEffect(()=>{
    fetchItems();
  },[loading]);
  //Read items
  const fetchItems=async()=>{
    try{
const data:any[]=await _sp.web.lists.getByTitle(props.ListName).items.select("Id","Title","EmailAddress","Admin/Id","Admin/Title","JoiningDate")
.expand("Admin")();
setItems(data.map((items:any)=>({
  id:items.Id,
  name:items.Title,
  email:items.EmailAddress,
  admin:items.Admin?{id:items.Admin.Id,text:items.Admin.Title}:undefined,
  joiningDate:items.DOB
})))
    }
    catch(err){
console.error("Error fetching items:", err);
    }
  }
  //add items
  const addItems=async()=>{
    try{
await _sp.web.lists.getByTitle(props.ListName).items.add({
  Title:newName,
  EmailAddress:newEmail,
  AdminId:newAdmin?parseInt(newAdmin.id.toString()):null,
  JoiningDate:newJoiningDate
});
setLoading(!loading);
    }
    catch(err){
console.error("Error adding item:", err);
    }
    finally{
      setIsAddHidden(true);
    }
  }
  return(
    <></>
  )
}
export default Crud;