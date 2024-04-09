import * as React from 'react';

import type { IEmployeeDemoProps } from './IEmployeeDemoProps';
import { DetailsList,PrimaryButton,TextField,DefaultButton } from 'office-ui-fabric-react';
//import { escape } from '@microsoft/sp-lodash-subset';
import{ MSGraphClientV3 } from "@microsoft/sp-http";

export interface Iuser{
  displayName:string;
  mail:string;
  jobTitle:string;
}
export interface IuserState{ // to pass the array data in state.
  userstate:Iuser[];
}

export default class EmployeeDemo extends React.Component<IEmployeeDemoProps,IuserState, {}> {
 constructor(props:IEmployeeDemoProps){
  super(props);
  this.state={userstate:[]};
 }

  public alluser:Iuser[]=[]; //blank array to store value form we get from api.
/**
 * GetUser=
 =>*/
public GetUser=():void=> {
  
   this.props.context.msGraphClientFactory.getClient("3")
    .then((msGraphClient: MSGraphClientV3) =>{
      msGraphClient.api("users")
      .version("v1.0").select("displayName,mail,jobTitle")
      .get((err, res)=>{ 
        if(err){
          console.log("error occured",err);
      }
      console.log("error occured",err);
      console.log("Response",res);
      res.value.map((result:any)=>{
        this.alluser.push({displayName: result.displayName,
          mail:result.mail,jobTitle:result.jobTitle
        });
      });
      this.setState({userstate:this.alluser});
      });
    });
};
/**
 * searchResult
 */
public searchResult=(searchTerm:string):void => {
 
   const filteredUsers = this.alluser.filter((user=>
    user.displayName.toLowerCase().includes(searchTerm.toLowerCase())||
    user.mail.toLowerCase().includes(searchTerm.toLowerCase())||
    user.jobTitle.toLowerCase().includes(searchTerm.toLowerCase())
   ));
   this.setState({userstate:filteredUsers});
}
  public render(): React.ReactElement<IEmployeeDemoProps> {
    

    return (
      <div>
        <PrimaryButton text='Search User' 
        onClick={this.GetUser}
        ></PrimaryButton>
        <DetailsList items={this.state.userstate}></DetailsList>

    <TextField     
    required={true}     
    name="txtSearchText"     
    placeholder="Search..." 
    onChange={(e, newValue) => this.searchResult(newValue || '')}   
    ></TextField>
    <DetailsList items={this.state.userstate}></DetailsList>

    <DefaultButton    
     data-automation-id="search"    
     target="_blank"    
     title="Search"
     >
     </DefaultButton>

      </div>
    );
  }
}
