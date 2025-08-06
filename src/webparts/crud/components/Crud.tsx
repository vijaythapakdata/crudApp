import * as React from 'react';
// import styles from './Crud.module.scss';
import type { ICrudProps } from './ICrudProps';
import { spfi,SPFx } from '@pnp/sp/presets/all';
import { DayOfWeek, DetailsList, Dialog, DialogFooter, DialogType, IconButton, PrimaryButton, SelectionMode } from '@fluentui/react';
import {PeoplePicker,PrincipalType} from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { TextField,DatePicker } from '@fluentui/react';


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
  const [currentId,setCurrentId]=React.useState<number>(0);
  const[editHidden,setIsEditHidden]=React.useState<boolean>(true);
  const[editName,setEditName]=React.useState<string>('');
  const[editEmail,setEditEmail]=React.useState<string>('');
  const[editAdmin,setEditAdmin]=React.useState<{id:number,text:string}>();
  const[editJoiningDate,setEditJoiningDate]=React.useState<string|any>("");

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
  //edit Dialog
  const openEditDialog=async(id:number)=>{
    const item:any=items.find((item)=>item.id===id);
    if(!items) return;
    setCurrentId(id);
    setEditName(item.name);
    setEditEmail(item.email);
    setEditAdmin(item.admin?{id:item.admin.id,text:item.admin.text}:undefined);
    setEditJoiningDate(item.joiningDate?new Date(item.joiningDate):undefined);
    setIsEditHidden(false);
  }
  //update items
  const updateItems=async()=>{
    try{
await _sp.web.lists.getByTitle(props.ListName).items.getById(currentId).update({
  Title:editName,
  EmailAddress:editEmail,
  AdminId:editAdmin?parseInt(editAdmin.id.toString()):null,
  JoiningDate:editJoiningDate?new Date(editJoiningDate):null
});
setLoading(!loading);
    }
    catch(err){
console.error("Error updating item:", err);
    }
    finally{
setIsEditHidden(true)
    }
  }
  //delete items
  const deleteItems=async(id:number)=>{
    try{
await _sp.web.lists.getByTitle(props.ListName).items.getById(id).delete();
    }
    catch(err){
console.error("Error deleting item:", err);
    }
    finally{
      setLoading(!loading)
    }
  }
  return(
    <>
    <DetailsList
    items={items}
    columns={[
      {
        key:'name',
        name:'Title',
        fieldName:'Title',
        minWidth:100,
        isResizable:true,
        onRender:(item:ICrudState)=><span>{item.name}</span>
      },
      {
        key:'Email Address',
        name:'EmailAddress',
        fieldName:'EmailAddress',
        minWidth:100,
        isResizable:true,
        onRender:(item:ICrudState)=><span>{item.email}</span>
      },
      {
       key:'Admin',
       name:'Admin',
       fieldName:'Admin',
       minWidth:100,
       isResizable:true,
       onRender:(item:ICrudState)=><span>{item.admin?.text}</span>
      },
      {
        key:'Joining Date',
        name:'JoiningDate',
        fieldName:'JoiningDate',
        minWidth:100,
        isResizable:true,
        onRender:(item:ICrudState)=><span>{item.joiningDate}</span>
      },
      {
        key:'Action',
        name:'Action',
        fieldName:'Action',
        minWidth:100,
        isResizable:true,
        onRender:(item:ICrudState)=>(
          <>
          <IconButton
          iconProps={{iconName:'Edit'}}
          aria-label='Edit'
          onClick={()=>openEditDialog(item.id)}
          />
           <IconButton
          iconProps={{iconName:'delete'}}
          aria-label='Delete'
          onClick={()=>deleteItems(item.id)}
          />
          </>
        )
      },
    ]}
    selectionMode={SelectionMode.none}
    />
    <Dialog
    hidden={addHidden}
    onDismiss={()=>setIsAddHidden(true)}
    dialogContentProps={{
      title:'Add Item',
      type:DialogType.largeHeader
    }}
    >
<TextField
label='Name'
value={newName}
onChange={(e, newValue)=>setNewName(newValue||'')}
/>
<TextField
label='Email Address'
value={newEmail}
onChange={(e, newValue)=>setNewEmail(newValue||'')}
/>
<PeoplePicker
context={props.context as any}
titleText='Admin'
personSelectionLimit={1}
showtooltip={true}
required={false}
principalTypes={[PrincipalType.User]}
webAbsoluteUrl={props.siteurl}
defaultSelectedUsers={newAdmin?.text?[newAdmin.text]:[]}
onChange={(items:any[])=>{
  if(items.length>0){
    setNewAdmin({id:items[0].id,text:items[0].text});
  }else{
    setNewAdmin(undefined);
  }
}}
/>
<DatePicker
label='Joining Date'
placeholder='Select a date'
value={newJoiningDate}
onSelectDate={(e)=>setNewJoiningDate(e)}
firstDayOfWeek={DayOfWeek.Sunday}
/>
<DialogFooter>
  <PrimaryButton
  onClick={addItems}
  text='Save'
  iconProps={{iconName:'Save'}}
  />
   <PrimaryButton
  onClick={()=>setIsAddHidden(true)}
  text='Cancel'
  iconProps={{iconName:'Cancel'}}
  />
</DialogFooter>
    </Dialog>
    {/* ----- */}
     <Dialog
    hidden={editHidden}
    onDismiss={()=>setIsEditHidden(true)}
    dialogContentProps={{
      title:'Edit Item',
      type:DialogType.largeHeader
    }}
    >
<TextField
label='Name'
value={editName}
onChange={(e, newValue)=>setEditName(newValue||'')}
/>
<TextField
label='Email Address'
value={editEmail}
onChange={(e, newValue)=>setEditEmail(newValue||'')}
/>
<PeoplePicker
context={props.context as any}
titleText='Admin'
personSelectionLimit={1}
showtooltip={true}
required={false}
principalTypes={[PrincipalType.User]}
webAbsoluteUrl={props.siteurl}
defaultSelectedUsers={editAdmin?.text?[editAdmin.text]:[]}
onChange={(items:any[])=>{
  if(items.length>0){
    setEditAdmin({id:items[0].id,text:items[0].text});
  }else{
    setEditAdmin(undefined);
  }
}}
/>
<DatePicker
label='Joining Date'
placeholder='Select a date'
value={editJoiningDate}
onSelectDate={(e)=>setEditJoiningDate(e)}
firstDayOfWeek={DayOfWeek.Sunday}
/>
<DialogFooter>
  <PrimaryButton
  onClick={updateItems}
  text='Save'
  iconProps={{iconName:'Save'}}
  />
   <PrimaryButton
  onClick={()=>setIsEditHidden(true)}
  text='Cancel'
  iconProps={{iconName:'Cancel'}}
  />
</DialogFooter>
    </Dialog>
<PrimaryButton
text='Add Item'
onClick={()=>setIsAddHidden(false)}
iconProps={{iconName:'Add'}}
/>
    </>
  )
}
export default Crud;