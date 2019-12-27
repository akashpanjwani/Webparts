import { ServiceScope } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ILinkImageProps {
  title: string;
  site: string;
  currentUser: string;
  serviceScope: ServiceScope;
  context: WebPartContext;
}
