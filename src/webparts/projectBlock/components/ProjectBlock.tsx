import * as React from 'react';
import styles from './ProjectBlock.module.scss';
import { IProjectBlockProps } from './IProjectBlockProps';
import { IProjectBlockState } from './IProjectBlockState';
import { escape } from '@microsoft/sp-lodash-subset';
import { Image } from 'office-ui-fabric-react/lib/Image';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { ServiceScope } from '@microsoft/sp-core-library';
import { CamlQuery } from '@pnp/sp';
import * as pnp from '@pnp/sp';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { DefaultButton, IDropdownOption, Dropdown } from 'office-ui-fabric-react';
import * as $ from 'jquery';

let listItem: any = [];

let ReportOption: IDropdownOption[] = [];
export default class ProjectBlock extends React.Component<IProjectBlockProps, IProjectBlockState> {


  public constructor(props: IProjectBlockProps) {
    super(props);

    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css');
    this.state = {
      ID: null,
      Title: "",
      ProjectLead: "",
      Description: "",
      ProjectTeam: "",
      StartDate: "",
      length: "",
      Size: "",
      BusinessIdea: null,
      showModal: false
    };

    let serviceScope: ServiceScope;
    serviceScope = this.props.serviceScope;

  }

  public async componentDidMount() {
    await this.getlistItem(this.props.site);
    await this.getFields(this.props.site);
    this.GetLinkData(this.props.site, "Project Blocks", this.props.currentUser).then((response: any) => {
      this.setState({
        BusinessIdea: response.results
      });
    });

  }

  public async getlistItem(siteUrl) {
    let web = new pnp.Web(siteUrl);

    const result = await web.lists.getByTitle("Project Blocks").items.get().then(result => {
      let Fields = result;
      listItem = result;
    });
  }

  public async getFields(siteUrl) {
    let web = new pnp.Web(siteUrl);

    const result = await web.lists.getByTitle("Project Blocks").fields.get().then(result => {
      let Fields = result;

      for (var resultVal of Fields) {
        ReportOption.push({
          key: resultVal.Title,
          text: resultVal.InternalName,
        });
      }
    });
  }

  /***************************/
  public GetLinkData(webUrl: string, listId: string, currentUser: string): Promise<any> {
    let p = new Promise<any>(async (resolve) => {

      let camlQuery: string = '';
      camlQuery = `<View Scope='Recursive'>
                    <Query>
                          <OrderBy>
                              <FieldRef Name="ID" 'Ascending="FALSE"'} /> 
                          </OrderBy>  
                        </Query>
                      <ViewFields>
                                    <FieldRef Name="ID" />
                                    <FieldRef Name="Title" />
                                    <FieldRef Name="ProjectLead" />		
                                    <FieldRef Name="Description" />
                                    <FieldRef Name="ProjectTeam" />
                                    <FieldRef Name="StartDate" />
                                    <FieldRef Name="length" />
                                    <FieldRef Name="Size" />
                      </ViewFields>`;

      const query: CamlQuery = {
        ViewXml: `${camlQuery}<RowLimit>10000</RowLimit></View>`,
        ListItemCollectionPosition: {
          "PagingInfo": "Paged=TRUE&p_ID=0"
        },
        FolderServerRelativeUrl: ''
      };

      const countQuery: CamlQuery = {
        ViewXml: `${camlQuery}</View>`,
      };

      let response = this.BindWorkItems(webUrl, listId, query, currentUser);
      resolve(response);
    });
    return p;
  }

  /***************Bind Ideas************/
  public async BindWorkItems(siteUrl: string, listId: string, query: CamlQuery, currentUser: string) {
    let web = new pnp.Web(siteUrl);

    const result = await web.lists.getByTitle(listId).getItemsByCAMLQuery(query, 'FieldValuesAsText');

    var response: any = {};
    let IdeasObj: any = [];
    result.forEach((item: any) => {
      var userEmail;
      web.siteUsers.getById(item.ProjectLeadId).get().then((result) => {
        var userInfo = "";
        userEmail = result.Email;
      });
      var imgPath = this.props.site + "/_layouts/15/userphoto.aspx?size=L&accountname=" + userEmail;
      var date = this.getFormattedDate(new Date(item.StartDate));
      IdeasObj.push({
        ID: item.Id,
        Title: item.Title,
        ProjectLead: item.FieldValuesAsText.ProjectLead,
        Description: item.Description,
        ProjectTeam: item.FieldValuesAsText.ProjectTeam,
        StartDate: date,
        length: item.length,
        Size: item.Size,
      });
    });

    response.results = IdeasObj;
    return response;
  }

  private _showModal = (): void => {
    this.setState({ showModal: true });
  }

  /***************************/

  public getFormattedDate(date) {
    var year = date.getFullYear();

    var month = (1 + date.getMonth()).toString();
    month = month.length > 1 ? month : '0' + month;

    var day = date.getDate().toString();
    day = day.length > 1 ? day : '0' + day;

    return month + '/' + day + '/' + year;
  }

  private handleReportComapnyDropDownOnChange = async (
    ev: any,
    selectedOption: any | undefined
  ): Promise<void> => {
    const selectedKey: string = selectedOption
      ? (selectedOption.text as string)
      : "";

    $('#dynamicField').append('<p className={styles.DetailPara}>' + selectedOption.text + '</p>');
  }

  public render(): React.ReactElement<IProjectBlockProps> {

    const Ideas: JSX.Element = this.state.BusinessIdea ?
      <div>
        {this.state.BusinessIdea.map((Ideas) => {

          var itemURl = this.props.site + "/Lists/Project%20Blocks/dispform.aspx?ID=" + Ideas.ID;
          return (
            <div className={styles.projectBlock}>
              <div className={styles.container}>
                <div className={styles.dummyInline} style={{ width: "45%" }}>
                  <Label className={styles.ttle}>{Ideas.Title}</Label>
                  <Link href={itemURl} target="_blank">View Other Fields...</Link>
                  <p className={styles.DetailPara}>{Ideas.Description}</p>

                  <div className={styles.dummyInline}>
                    Project Lead
              <div style={{ display: "flex" }}>
                      <div style={{ display: "block" }}>
                        <Image
                          src="http://placehold.it/50x50"
                          alt="Example implementation of the property image fit using the none value on an image smaller than the frame."
                          style={{ borderRadius: "50%" }}
                        />
                        <p style={{ margin: "1%" }}>{Ideas.ProjectLead}</p>
                      </div>
                    </div>
                  </div>
                  <div className={styles.dummyInline}>
                    Project Team
              <div style={{ display: "flex" }}>
                      <div style={{ display: "block" }}>
                        <Image
                          src="http://placehold.it/50x50"
                          alt="Example implementation of the property image fit using the none value on an image smaller than the frame."
                          style={{ borderRadius: "50%" }}
                        />
                        <p style={{ margin: "1%" }}>{Ideas.ProjectTeam}</p>
                      </div>

                    </div>

                  </div>
                </div>
                <div className={styles.dummyInline} style={{ width: "45%" }}>
                
                  {/* <Dropdown placeholder="Select an option" options={ReportOption} multiSelect={true} onChange={this.handleReportComapnyDropDownOnChange} /> */}

                  <div style={{ borderBottom: "1px solid" }}><p className={styles.DetailPara} style={{ fontWeight: 600 }}>At a Glance</p></div>
                  <ul></ul>
                  <div>
                    <div className={styles.inline}>
                      <div>Start Date</div></div>
                    <div className={styles.inline}>
                      <div>{Ideas.StartDate}</div></div>
                  </div>
                  <div>
                    <div className={styles.inline}>
                      <div>Contract length</div></div>
                    <div className={styles.inline}>
                      <div>{Ideas.length} year's</div></div>
                  </div>
                  <div>
                    <div className={styles.inline}>
                      <div>Contract Size</div></div>
                    <div className={styles.inline}>
                      <div>{Ideas.Size}</div></div>
                  </div>
                  <div className={styles.divInline}>

                  </div>
                </div>
              </div>
            </div >

          );
        })
        }
      </div >

      : <div />;
    return (
      <div className={styles.projectBlock}>
        {Ideas}
      </div>
    );
  }

}
