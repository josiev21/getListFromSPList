import { Version } from '@microsoft/sp-core-library';
import { spfi, SPFx } from '@pnp/sp'
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

// import {
//   type IPropertyPaneConfiguration,
//   PropertyPaneTextField
// } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IGetListWebPartProps {
  description: string;
}

//class definition =>type
interface IWorkdayEmployee {
  Title: string;
  Leave_Balance: number;
  Sub_TimeOff: number;
  Sub_Birthday: number;
  Inbox_Count: number;
}

export default class GetListWebPart extends BaseClientSideWebPart<IGetListWebPartProps> {


  private username: String = "";
  private jobTitle: String ="";
  private Sub_Birthday: any = "" ;
  private Leave_Balance: any = "";
  private error: String ="";
  // private id: String = "";
  private employeeId: String ="";

  public render(): void {
    this.domElement.innerHTML = `
    <div class="card" style="width:380px; border: solid 3px #D7DADF; border-radius: 8px; padding:18px; ">
    <div class="card-header" style="display: flex; width: 100%; align-items: center; height:26px;">
    <div class="col-sm-6" style="display:flex; justify-content: flex-start; width:50%;">
      <h1 style="font-weight:350; font-size:20px;line-height:26px;">Team <span style="display:none;">${this.username}${this.error}</h1>
    </div>
    <div class="col-sm-6" style="display:flex; justify-content: flex-end; width:50%;">
    </div>
    </div>
      <div class="card-body" style="">
        <div class="container" style="display:flex;">
          <div class = "col-sm-6" style="width:100%; display:flex; flex-direction:column; align-items:flex-start; border-right: solid 2px #EDEBE9; margin-top:-10px;"> 
            <div class="row" style="display: flex; flex-direction: row; justify-content: space-between; margin-left: 10px; align-items: center;">    
              <div class="col-sm-6">
                <p style="font-family:'Segoe UI'; font-weight:600; font-size:32px; line-height:48px; color:#161819;">${this.Sub_Birthday}</p>
              </div>
              <div class="col-sm-6" style="display: flex; flex-direction: column; margin-left: 10px;">
              <h1 style="font-family:'Segoe UI'; font-size:12px; line-height:0;">Employees</h1>
              <p style="font-family:'Segoe UI'; font-size:12px; line-height:0; color:#848993;">(Within 2 weeks)</p>
              </div> 
              </div> 
              <p style="font-family:'Segoe UI'; font-weight:600; font-size:16px; color: #161819; margin-left:10px; margin-top:-10px;">Birthday</p>
              <button class="seeEmployee" style="border: solid 3px #DE0E13; border-radius: 6px; background: #fff; margin-left: 10px; width: 90px; height: 24px; font-size: 10px; color: #DE0E13; font-weight: 400;">See Employee</button>
          </div>
          <div class = "col-sm-6" style="width:100%; display:flex; flex-direction:column; align-items:flex-start; margin-left:10px; margin-top:-10px;">
            <div class="row" style="display: flex; flex-direction: row; justify-content: space-between; margin-left: 10px; align-items: center;">    
              <div class="col-sm-6">
                <p style="font-family:'Segoe UI'; font-weight:600; font-size:32px; line-height:48px; color:#161819;">${this.Leave_Balance}</p>
              </div>
              <div class="col-sm-6" style="display: flex; flex-direction: column; margin-left: 10px;">
              <h1 style="font-family:'Segoe UI'; font-size:12px; line-height:0;">Employees</h1>
              <p style="font-family:'Segoe UI'; font-size:12px; line-height:0; color:#848993;">(In the next 2 weeks)</p>
              </div> 
              </div> 
              <p style="font-family:'Segoe UI'; font-weight:600; font-size:16px; color: #161819; margin-left:10px; margin-top:-10px;">Leave Plan</p>
              <button class="seeCalendar" style="border: solid 3px #DE0E13; border-radius: 6px; background: #fff; margin-left: 10px; width: 90px; height: 24px; font-size: 10px; color: #DE0E13; font-weight: 400;">See Calendar</button>
          </div>
        </div>
      </div>
    </div>
      `;
  }

  private async getList(employeeId: String) {
    
    const sp = spfi().using(SPFx(this.context))
    // const profileCardDataList = sp.web.lists.getByTitle('ProfileCardsData');
    const employeeIdNumb: IWorkdayEmployee[] = await sp.web.lists.getByTitle('ProfileCardsData').items.top(1).filter(`Title eq '${employeeId}'`)();
    this.Leave_Balance = employeeIdNumb[0].Leave_Balance;
    this.Sub_Birthday = employeeIdNumb[0].Sub_Birthday;
    console.log(employeeIdNumb[0].Leave_Balance)
    this.render()
    // console.log(employeeIdNumb[0].Leave_Balance);
    // // @ts-ignore
    // for await (const items of profileCardDataList.items.top(1000)) {

    //   console.log("Items length " + items.length + ", items size " + items.size + ". First one is " + items[0]);

    //   items.forEach((employee: IWorkdayEmployee) => {
    //     console.log(employee);
    //   });

    // }

    
  }



  protected async onInit(): Promise<void> {
    await super.onInit()   
    // username and job title
    this.context.msGraphClientFactory
    .getClient('3')
    .then((client: MSGraphClientV3) => {
        // eslint-disable-next-line no-void
      void client
          .api('/me')
          .get((error: { message: String; }, response: any) => {
            if (error) {
              this.error = error.message;
            } else {
                this.username = response.displayName;            
                this.jobTitle = response.jobTitle;
                console.log(this.jobTitle);
                // this.id = response.id
              }

            this.render()
          });
    })
    .catch((error) => {
      this.error = error.message
      this.render()
    });

    // profile picture
    this.context.msGraphClientFactory
    .getClient('3')
    .then((client: MSGraphClientV3) => {
        // eslint-disable-next-line no-void
      void client
          .api('/me/photo/$value')
          .get((error: any, response: any) => {
            // this.profpicURL = URL.createObjectURL(response)
            this.render()
          });
    })
    .catch((error) => {
      this.error = error.message
      this.render()
    });

    // Employee Id
    this.context.msGraphClientFactory
    .getClient('3')
    .then((client: MSGraphClientV3) => {
        // eslint-disable-next-line no-void
      void client
          .api('/me/?$select=EmployeeId')
          .get((error: any, response: any) => {
            this.employeeId = response.employeeId;
            this.render();
            this.getList(this.employeeId);

          });
    })
    .catch((error) => {
      this.error = error.message
      this.render()
    });

    // //Get Data from List
    // const siteUrl = 'https://wqcmg.sharepoint.com/sites/JoTest';
    // const listTitle = 'ProfileCardsData';

    // // Step 1: Get the current user’s ID
    // fetch(`${siteUrl}/_api/web/currentuser`, {
    //     method: 'GET',
    //     headers: {
    //         'Accept': 'application/json;odata=verbose'
    //     }
    // })
    // .then(response => response.json())
    // .then(userData => {
    //     const userId = userData.d.Id;

    //     // Step 2: Query the SharePoint list using the logged-in user's ID
    //     const queryUrl = `${siteUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$filter=User/Id eq ${userId}`;

    //     return fetch(queryUrl, {
    //         method: 'GET',
    //         headers: {
    //             'Accept': 'application/json;odata=verbose'
    //         }
    //     });
    // })
    // .then(response => response.json())
    // .then(data => {
    //     console.log(data.d.results); // Process the filtered data here
    // })
    // .catch(error => console.error('Error:', error));
  }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}

