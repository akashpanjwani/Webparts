import * as React from 'react';
import styles from './BusinessAreaBlock.module.scss';
import { IBusinessAreaBlockProps } from './IBusinessAreaBlockProps';
import { IBusinessAreaBlockState } from './IBusinessAreaBlockState';
import { escape } from '@microsoft/sp-lodash-subset';
import { Image } from 'office-ui-fabric-react/lib/Image';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { FontSizes } from '@uifabric/styling';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { ServiceScope } from '@microsoft/sp-core-library';
import { CamlQuery, Item } from '@pnp/sp';
import * as pnp from '@pnp/sp';
import * as Newpnp from 'sp-pnp-js'; 

let userEmail;
let userPhone;
export default class BusinessAreaBlock extends React.Component<IBusinessAreaBlockProps, IBusinessAreaBlockState> {

  public constructor(props: IBusinessAreaBlockProps) {
    super(props);

    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css');
    this.state = {
      Title: "",
      BusinessIdea: null,
      URL: "",
      Description: "",
      fileName: "",
      fileURL: "",
      user: "",
      userEmail: "",
      userPhone: ""
    };

    let serviceScope: ServiceScope;
    serviceScope = this.props.serviceScope;

  }

  public componentDidMount() {
    this.GetLinkData(this.props.site, "BusinessBlock", this.props.currentUser).then((response: any) => {
      this.setState({
        BusinessIdea: response.results
      });
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
                                    <FieldRef Name="URL" />		
                                    <FieldRef Name="Description" />
                                    <FieldRef Name="user" />
                                    <FieldRef Name="UserSpecialist" />
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

    const result = await web.lists.getByTitle(listId).getItemsByCAMLQuery(query, 'FieldValuesAsText', 'fileLeafRef', 'FileRef');

    var response: any = {};
    let IdeasObj: any = [];


    for (let item of result){
    //result.forEach((item: any) => {

      await web.siteUsers.getById(item.userId).get().then(async (userResult) => {
        var userInfo = "";
        userEmail = userResult.Email;
        let loginName = "i:0#.f|membership|" + userEmail;
      //   await Newpnp.sp.profiles.getPropertiesFor(loginName).then(resp => {
      //     let props = {};
      //     resp.UserProfileProperties.map((val) => {
      //       if (val.Key == "WorkPhone") {
      //         console.log(val.Value);
      //         userPhone = val.Value;
      //       }
      //     });
          
      // });
    });

    var imgPath = this.props.site + "/_layouts/15/userphoto.aspx?size=L&accountname=" + item.FieldValuesAsText.user;
          IdeasObj.push({
            Title: item.FieldValuesAsText.Title,
            URL: item.URL,
            Description: item.Description,
            fileName: item.FileLeafRef,
            fileURL: item.FileRef,
            user: item.FieldValuesAsText.user,
            imgPath: imgPath,
            userEmail: userEmail,
            userPhone: userPhone,
          });
        }

        response.results = IdeasObj;
        return response;

  }

  /***************************/

  public render(): React.ReactElement<IBusinessAreaBlockProps> {

    const Ideas: JSX.Element = this.state.BusinessIdea ?
      <div>
        {this.state.BusinessIdea.map((Ideas) => {
          var fileUrl = this.props.site + "/" + Ideas.fileURL;
          var emailURL = "mailto:" + Ideas.userEmail;
          var Tel = "tel:" + userPhone;
          return (
            <div className={styles.businessAreaBlock}>
              <div className={styles.container}>
                <Image
                  src={fileUrl}
                  alt="Example implementation with no image fit property and only width is specified."
                  width={1210}
                  height={300}
                />
                <div><Label className={styles.ttle}>{Ideas.Title}</Label></div>
                <div><p className={styles.DetailPara}>{Ideas.Description}</p></div>
                <div className={styles.dummyInline}>
                  <Image
                    src={Ideas.imgPath}
                    alt="Example implementation of the property image fit using the none value on an image smaller than the frame."
                    style={{ borderRadius: "50%", height: "70px" }}
                  />
                </div>
                <div className={styles.dummyInline} style={{ width: "15%" }}>
                  <p>Eve Compton</p>
                  <p>{Ideas.user}</p>
                </div>
                <div className={styles.dummyInlineicon}>
                  <a href={emailURL} target="_blank"><i className="fa fa-envelope" id={styles.envelope}></i></a>
                  <a href={Tel} target="_blank"><i className="fa fa-phone" id={styles.envelope}></i></a>
                  <i className="fa fa-file" id={styles.envelope}></i>
                </div>
                <div>
                </div>
              </div>
            </div>
          );
        })}
      </div>
      : <div />;
    return (
      <div className={styles.businessAreaBlock}>
        {Ideas}
      </div>
    );
  }
}
