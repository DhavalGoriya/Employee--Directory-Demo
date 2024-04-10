import * as React from 'react';
import styles from './EmployeeDemo.module.scss';
import type { IEmployeeDemoProps } from './IEmployeeDemoProps';
import { DetailsList, PrimaryButton, TextField, } from 'office-ui-fabric-react';
//import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClientV3 } from "@microsoft/sp-http";


export interface Iuser {
  displayName: string;
  mail: string;
  surname: string;


}
export interface IData {
  userID: string;
  managerID: string;
}
export interface IuserState { // to pass the array data in state.
  userstate: Iuser[];
  searchTerm: string;
  managerstate: IData[];
}
userManagaer: { };
export default class EmployeeDemo extends React.Component<IEmployeeDemoProps, IuserState, {}> {
  constructor(props: IEmployeeDemoProps) {
    super(props);
    this.state = {
      userstate: [],
      searchTerm: '',
      managerstate:[]
    };
  }

  componentDidMount(): void {
    this.GetUser();
  }

  public alluser: Iuser[] = []; //blank array to store value form we get from api.
  public userData: IData[] = [];





  /**
   * GetUser=
   =>*/
  public GetUser = (): void => {

    this.props.context.msGraphClientFactory.getClient("3")
      .then((msGraphClient: MSGraphClientV3) => {
        msGraphClient.api("users")
          .version("v1.0").select("displayName,mail,surname")
          .get((err, res) => {
            if (err) {
              console.log("error occured", err);
            }
            console.log("error occured", err);
            console.log("Response", res);
            res.value.map((result: any) => {
              this.alluser.push({
                displayName: result.displayName,
                mail: result.mail, surname: result.surname
              });
            });
            this.setState({ userstate: this.alluser });
          });
      });
  };
  
  /**
   * searchResult
   */
  public searchResult = (searchTerm: string): void => {

            const filteredUsers = this.alluser.filter((user =>
              user.displayName?.toLowerCase().includes(searchTerm.toLowerCase()) ||
              user.mail?.toLowerCase().includes(searchTerm.toLowerCase()) ||
              user.surname?.toLowerCase().includes(searchTerm.toLowerCase())
            ));
            this.setState({ userstate: filteredUsers });
          } 

/**
   * mapUserManager=
   */


public mapUserManager = (): void => {
  this.props.context.msGraphClientFactory.getClient("3")
    .then((Client: MSGraphClientV3) => {

      Client.api("users/id | userPrincipalName?$expand=directReports")
        .version("v1.0").select("UserID,mangerID")
        .get((err
          , res) => {
          if (err) {
            console.log("error occurred", err);
          }
          console.log("error occured", err);
          console.log("Response", res);
          res.value.map((results: any) => {
            this.userData.push({
              userID: results.userID, managerID: results.managerID
            });
          });
            this.setState({managerstate:this.userData});
        });
      });
    }
          

  public render(): React.ReactElement < IEmployeeDemoProps > {
            return(
      <div>
        <TextField
          className={styles.row}
          value={this.state.searchTerm}
          name="txtSearchText"
          placeholder="Search..."
          // onChange={(e, newValue) => this.searchResult(newValue || '')}
          onChange={(e: React.ChangeEvent<HTMLInputElement>) => this.setState({ searchTerm: e.target.value })}
        ></TextField>
        <PrimaryButton text='Search'
          onClick={() => this.searchResult(this.state.searchTerm)}
        ></PrimaryButton>
        <DetailsList items={this.state.userstate}></DetailsList>

      </div >
    );
      }
}
