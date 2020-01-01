import * as React from 'react';
import styles from './LinkImage.module.scss';
import { ILinkImageProps } from './ILinkImageProps';
import { ILinkImageState } from './ILinkImageState';
import { escape } from '@microsoft/sp-lodash-subset';
import { Image } from 'office-ui-fabric-react/lib/Image';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { ServiceScope } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { CamlQuery } from '@pnp/sp';
import * as pnp from '@pnp/sp';

export default class LinkImage extends React.Component<ILinkImageProps, ILinkImageState> {

  public constructor(props: ILinkImageProps) {
    super(props);

    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css');
    this.state = {
      Title: "",
      Idea: null,
      URL: "",
      Description: "",
      fileName: "",
      fileURL: "",
    };

    let serviceScope: ServiceScope;
    serviceScope = this.props.serviceScope;

  }

  public componentDidMount() {
    this.GetLinkData(this.props.site, "LinkData", this.props.currentUser).then((response: any) => {
      this.setState({
        Idea: response.results
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
    result.forEach((item: any) => {

      IdeasObj.push({
        Title: item.Title,
        URL: item.URL,
        Description: item.Description,
        fileName: item.FileLeafRef,
        fileURL: item.FileRef,
      });

    });

    response.results = IdeasObj;
    return response;

  }

  /***************************/
  public render(): React.ReactElement<ILinkImageProps> {

    var siteURL=this.props.site.split('/');
    var webURl=siteURL.splice(0,3).join('/');

    const Ideas: JSX.Element = this.state.Idea ?
      <div>
        {this.state.Idea.map((Ideas) => {
          var fileUrl = webURl+"/"+ Ideas.fileURL;
          return (
            <div className={styles.linkImage}>
              <div className={styles.container}>
                <div className={styles.divImg} style={{ width: "8%" }}>
                  <Image
                    src={fileUrl}
                    alt="Example implementation of the property image fit using the none value on an image smaller than the frame."
                    style={{ borderRadius: "50%" }}
                    height="10"
                    width="10"
                  />
                </div>
                <div className={styles.divText} style={{ width: "65%" }}>
                  <a href={Ideas.URL} target="_blank" style={{cursor:"pointer"}}><Label style={{cursor:"pointer"}} className={styles.ttle}>{Ideas.Title}</Label></a>
                  <p className={styles.DetailPara}>{Ideas.Description}</p>
                </div>
              </div>
            </div>
          );
        })}
      </div>
      : <div />;
    return (
      <div className={styles.linkImage}>
        {Ideas}
      </div>
    );
  }
}
