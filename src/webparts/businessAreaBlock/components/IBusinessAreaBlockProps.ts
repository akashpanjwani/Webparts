import { ServiceScope } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBusinessAreaBlockProps {
  title: string;
  site: string;
  currentUser: string;
  serviceScope: ServiceScope;
  context: WebPartContext;
}
