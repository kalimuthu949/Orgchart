import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IPhoneguideProps {
  description?: string;
  propertyToggle?:string;
  isDarkTheme?: boolean;
  environmentMessage?: string;
  hasTeamsContext?: boolean;
  userDisplayName?: string;
  context?:WebPartContext
}
