import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IBannerProps {
  displayMode: DisplayMode;
  listName: string;
  newListName:string;
  createNewList: boolean;
  layout: string;
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  height: string;
  autoplay: boolean;
  autoplaySpeed: number;
  navStyle: string;
  slideEffect: string;
  pauseOnHover: boolean;
  captionPosition: string;
  displayStyle: string;
  hideUpload: boolean;
  context: WebPartContext;
  backgroundColor: string;
  borderRadius: number;
  captionFontSize: number;
  captionWeight: string;
  captionColor: string;
}
