import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  IPropertyPaneGroup,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPFx, spfi } from "@pnp/sp";
import "@pnp/sp/fields";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/views";
import "@pnp/sp/webs";
import { PropertyFieldColorPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import * as strings from 'BannerWebPartStrings';
import { IList } from '../../services/IList';
import { ListService } from '../../services/ListService';
import { IBannerProps } from './components/IBannerProps';
import Banner from './components/Banner';

export interface IBannerWebPartProps {
  displayMode: DisplayMode;
  listName: string;
  newListName:string;
  createNewList: boolean;
  layout: string;
  description: string;
  height: string;
  autoplay: boolean;
  autoplaySpeed: number;
  navStyle: string;
  slideEffect: string;
  pauseOnHover: boolean;
  captionPosition: string;
  displayStyle: string;
  hideUpload: boolean;
  backgroundColor: string;
  borderRadius: number;
  captionFontSize: number;
  captionWeight: string;
  captionColor: string;
}

export default class BannerWebPart extends BaseClientSideWebPart<IBannerWebPartProps> {
    private _isDarkTheme: boolean = false;
    private _environmentMessage: string = "";
    private lists: IPropertyPaneDropdownOption[];
    private listsDropdownDisabled: boolean = true;
    private showCreateList: boolean = false;
    private createListBoolean: boolean = false;
    private listData: any[] = [];

    protected async onPropertyPaneConfigurationStart(): Promise<void> {
        this.listsDropdownDisabled = !this.lists;
        this.properties.createNewList = false;
        if (this.lists) {
            return;
        }

        this.context.statusRenderer.displayLoadingIndicator(this.domElement, "lists");
        if (this.properties.captionFontSize === undefined) this.properties.captionFontSize = 15;

        await this.loadLists().then((listOptions: IPropertyPaneDropdownOption[]): void => {
            this.lists = listOptions;
            this.listsDropdownDisabled = false;
            let listExisted = false;
            this.lists.map((item) => {
                if (item.key === this.properties.listName) listExisted = true;
            });
            listExisted ? (this.properties.hideUpload = false) : (this.properties.hideUpload = true);
            this.context.propertyPane.refresh();
            this.context.statusRenderer.clearLoadingIndicator(this.domElement);
            this.render();
        });
    }

    private async loadLists(): Promise<IPropertyPaneDropdownOption[]> {
        const dataService = new ListService(this.context);
        const response = await dataService.getDocumentLibrary();
        const options: IPropertyPaneDropdownOption[] = [];
        const data: any[] = [];

        response.forEach((item: any) => {

            data.push({
                id: item.Id,
                title: item.Title,
                InternalName: item.EntityTypeName
              })

            options.push({ key: item.Title, text: item.Title });
        });

        this.listData = [...data];

        return options;
    }

    private async createDocumentLibrary(listName: string): Promise<void> {
        this.listsDropdownDisabled = true;
        const sp = spfi(this.context.pageContext.web.absoluteUrl).using(SPFx(this.context));
        //const listEnsureResult = await sp.web.lists.ensure(listName);
        //const listEnsureResult = await sp.web.folders.addUsingPath(listName);
        const listExisted = await (await sp.web.getFolderByServerRelativePath(listName).select("Exists")()).Exists;
        if (listExisted) {
            alert(`Carousel Library "${listName}" already exist!`);
            return;
        }

        await sp.web.lists
            .add(listName, "", 101, true, { AllowContentTypes: true, ContentTypesEnabled: true })
            .then(async (doc) => {
                await doc.list.fields.addNumber("Sequence");
                await doc.list.fields.addText("Caption");
                const list = await sp.web.lists.getByTitle(listName);
                const view = await list.defaultView;
                await view.fields.add("Sequence");
                await view.fields.add("Caption");

                alert(`Carousel library "${listName}" created!`);
                this.properties.createNewList = false;
                this.properties.listName = listName;
                this.properties.hideUpload = false;
                await this.loadLists().then((listOptions: IPropertyPaneDropdownOption[]): void => {
                    this.lists = listOptions;
                    this.listsDropdownDisabled = false;
                    this.context.propertyPane.refresh();
                    this.context.statusRenderer.clearLoadingIndicator(this.domElement);
                    this.showCreateList = !this.showCreateList;
                    this.render();
                });
            });
    }

    public render(): void {
        const element: React.ReactElement<IBannerProps> = React.createElement(Banner, {
            displayMode: this.displayMode,
            listName: this.properties.listName,
            newListName: this.properties.newListName,
            createNewList: this.properties.createNewList,
            layout: this.properties.layout,
            description: this.properties.description,
            isDarkTheme: this._isDarkTheme,
            environmentMessage: this._environmentMessage,
            hasTeamsContext: !!this.context.sdks.microsoftTeams,
            userDisplayName: this.context.pageContext.user.displayName,
            height: this.properties.height,
            autoplay: this.properties.autoplay,
            autoplaySpeed: this.properties.autoplaySpeed,
            navStyle: this.properties.navStyle,
            slideEffect: this.properties.slideEffect,
            pauseOnHover: this.properties.pauseOnHover,
            captionPosition: this.properties.captionPosition,
            displayStyle: this.properties.displayStyle,
            hideUpload: this.properties.hideUpload,
            context: this.context,
            backgroundColor: this.properties.backgroundColor,
            borderRadius: this.properties.borderRadius,
            captionFontSize: this.properties.captionFontSize,
            captionWeight: this.properties.captionWeight,
            captionColor: this.properties.captionColor
        });

        ReactDom.render(element, this.domElement);
    }

    protected onInit(): Promise<void> {
        return this._getEnvironmentMessage().then((message) => {
            this._environmentMessage = message;
        });
    }

    private async handleCreateList(): Promise<void>  {
        this.showCreateList = !this.showCreateList;
        this.render();
    }

    private _getEnvironmentMessage(): Promise<string> {
        if (!!this.context.sdks.microsoftTeams) {
            // running in Teams, office.com or Outlook
            return this.context.sdks.microsoftTeams.teamsJs.app.getContext().then((context) => {
                let environmentMessage: string = "";
                switch (context.app.host.name) {
                    case "Office": // running in Office
                        environmentMessage = this.context.isServedFromLocalhost
                            ? strings.AppLocalEnvironmentOffice
                            : strings.AppOfficeEnvironment;
                        break;
                    case "Outlook": // running in Outlook
                        environmentMessage = this.context.isServedFromLocalhost
                            ? strings.AppLocalEnvironmentOutlook
                            : strings.AppOutlookEnvironment;
                        break;
                    case "Teams": // running in Teams
                    case "TeamsModern":
                        environmentMessage = this.context.isServedFromLocalhost
                            ? strings.AppLocalEnvironmentTeams
                            : strings.AppTeamsTabEnvironment;
                        break;
                    default:
                        environmentMessage = strings.UnknownEnvironment;
                }

                return environmentMessage;
            });
        }

        return Promise.resolve(
            this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentSharePoint
                : strings.AppSharePointEnvironment
        );
    }

    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
        if (!currentTheme) {
            return;
        }

        this._isDarkTheme = !!currentTheme.isInverted;
        const { semanticColors } = currentTheme;

        if (semanticColors) {
            this.domElement.style.setProperty("--bodyText", semanticColors.bodyText || null);
            this.domElement.style.setProperty("--link", semanticColors.link || null);
            this.domElement.style.setProperty("--linkHovered", semanticColors.linkHovered || null);
        }
    }

    private async updateEditPanelView(): Promise<void>  {
        this.createListBoolean = !this.createListBoolean;
        this.showCreateList = !this.showCreateList;
        this.render();
    }

    private async goToBackend(): Promise<void>  {
        if (this.listData && this.listData.length > 0) {
            const listName = this.listData.find(ld => ld.title === this.properties.listName);
            window.open(this.context.pageContext.web.absoluteUrl+ '/' + listName.InternalName.replace("_x0020_", " ") + '/Forms/AllItems.aspx', "_blank");
        }
        else {
            window.open(this.context.pageContext.web.absoluteUrl+ '/' + this.properties.listName + '/Forms/AllItems.aspx', "_blank");
        }
        // window.open(
        //     this.context.pageContext.web.absoluteUrl + "/" + this.properties.listName + "/Forms/AllItems.aspx",
        //     "_blank"
        // );
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse("1.0");
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        const selectCreateListGroup: IPropertyPaneGroup = this.showCreateList
            ? {
                  groupName: "Create Carousel Library",
                  groupFields: [
                      PropertyPaneButton("back", {
                          text: "Back",
                          buttonType: PropertyPaneButtonType.Primary,
                          icon: "Back",
                          onClick: () => this.handleCreateList(),
                      }),
                      PropertyPaneTextField("newListName", {
                          label: "Carousel Libaray Name",
                      }),
                      PropertyPaneButton("createList", {
                          text: "Create",
                          buttonType: PropertyPaneButtonType.Primary,
                          icon: "Add",
                          onClick: () => this.createDocumentLibrary(this.properties.newListName || "Carousel Library"),
                      }),
                    //   PropertyPaneDropdown("navStyle", {
                    //       label: "Navigation Styling",
                    //       options: [
                    //           { key: "dots", text: "Dots" },
                    //           { key: "thumb", text: "Thumbnail" },
                    //       ],
                    //       selectedKey: "dots",
                    //   }),
                    //   PropertyPaneDropdown("slideEffect", {
                    //       label: "Slide Effect",
                    //       options: [
                    //           { key: "slide", text: "Slide" },
                    //           { key: "fade", text: "Fade" },
                    //       ],
                    //       selectedKey: "slide",
                    //   }),
                  ],
              }
            // : this.properties.layout === "Template2"
            // ? {
            //       groupName: "Template2 settings",
            //       groupFields: [
            //           PropertyPaneCheckbox("autoplay", {
            //               text: "Autoplay2",
            //               checked: false,
            //           }),
            //       ],
            //   }
            : {
                  groupName: "Document Library Configuration",
                  groupFields: [
                    PropertyPaneButton("createList", {
                        text: "Create List",
                        buttonType: PropertyPaneButtonType.Primary,
                        icon: "Add",
                        onClick: () => this.updateEditPanelView(),
                    }),
                    PropertyPaneDropdown("listName", {
                        label: "List Name",
                        options: this.lists,
                        disabled: this.listsDropdownDisabled,
                    }),
                    this.properties.listName !== undefined &&
                        this.properties.listName !== "" &&
                        PropertyPaneButton("Backend", {
                            text: "Edit your carousel",
                            buttonType: PropertyPaneButtonType.Normal,
                            onClick: () => this.goToBackend(),
                        }),
                  ],
              };
        
        const CaptionGroup: IPropertyPaneGroup = {
            groupName: "Caption Settings",
            groupFields: [
                PropertyPaneDropdown("captionPosition", {
                    label: "Caption Position",
                    options: [
                        { key: "bottom", text: "Bottom" },
                        { key: "top", text: "Top" },
                        { key: "tleft", text: "Top Left" },
                        { key: "tright", text: "Top Right" },
                        { key: "bleft", text: "Bottom Left" },
                        { key: "bright", text: "Bottom Right" },
                    ],
                    selectedKey: "bottom",
                }),
                PropertyFieldNumber("captionFontSize", {
                    key: "captionFontSize",
                    label: "Caption Font Size",
                    value: this.properties.captionFontSize,
                    maxValue: 30,
                    minValue: 15,
                }),
                PropertyPaneDropdown('captionWeight', {
                    label: 'Caption Weight',
                    options: [
                      { key: 'normal', text: 'normal' },
                      { key: 'bold', text: 'bold' },
                      { key: 'lighter', text: 'lighter' }
                    ],
                    selectedKey: 'normal'
                }),
                PropertyFieldColorPicker("captionColor", {
                    label: "Caption Color",
                    selectedColor: this.properties.captionColor,
                    onPropertyChange: this.onPropertyPaneFieldChanged,
                    properties: this.properties,
                    key: "captionColor",
                }),
            ],
        };

        return {
            pages: [
                {
                    /*header: {
                        description: strings.PropertyPaneDescription,
                    },*/
                    groups: [
                        selectCreateListGroup,
                        CaptionGroup,
                        {
                            groupName: "Carousel settings",
                            groupFields: [
                                PropertyPaneDropdown("layout", {
                                    label: "Layout",
                                    options: [
                                        { key: "1", text: "Large Image (1:1)" },
                                        { key: "2", text: "Small Image (1:3)" },
                                    ],
                                    selectedKey: "1",
                                }),
                                
                                this.properties.layout === "1" && 
                                PropertyPaneDropdown("displayStyle", {
                                    label: "Display Style (Image)",
                                    options: [
                                        { key: "cover", text: "Cover" },
                                        { key: "contain", text: "Contain" },
                                        { key: "fill", text: "Fill" },
                                    ],
                                    selectedKey: "cover",
                                }),
                                PropertyPaneToggle("autoplay", {
                                    label: "Auto Play",
                                    checked: this.properties.autoplay,
                                    offText: "Off",
                                    onText: "On",
                                }),
                                PropertyFieldNumber("autoplaySpeed", {
                                    key: "autoplaySpeed",
                                    label: "Auto Play Speed (in milliseconds)",
                                    value: this.properties.autoplaySpeed,
                                    maxValue: 9999,
                                    minValue: 3000,
                                    disabled: !this.properties.autoplay,
                                }),
                                PropertyPaneToggle("pauseOnHover", {
                                    label: "Pause Slide Show On Hover",
                                    checked: this.properties.pauseOnHover,
                                    offText: "Off",
                                    onText: "On",
                                }),
                                PropertyPaneTextField("borderRadius", {
                                    label: "Border Radius (px)",
                                }),
                                PropertyPaneTextField("height", {
                                    label: "Height (px)",
                                }),
                                PropertyFieldColorPicker("backgroundColor", {
                                    label: "Background Color",
                                    selectedColor: this.properties.backgroundColor,
                                    onPropertyChange: this.onPropertyPaneFieldChanged,
                                    properties: this.properties,
                                    key: "backgroundColor",
                                }),
                            ],
                        },
                        //templateGroup
                    ],
                },
            ],
        };
    }
}
