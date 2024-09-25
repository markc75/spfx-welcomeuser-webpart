import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLabel,
  PropertyPaneToggle,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';

export interface IWelcomeUserWebPartProps {
  description: string;
  showCustomMessage: boolean;
  customMessage: string;
  buttonText: string;
  buttonUrl: string;
  showButton: boolean;
  buttonColor: string;
  buttonTextColor: string;
  backgroundType: string;
  backgroundValue: string;
  customMessageColor: string;
  welcomeUserTextColor: string;
}



export default class WelcomeUserWebPart extends BaseClientSideWebPart<IWelcomeUserWebPartProps> {

  private sp: any;

  protected onInit(): Promise<void> {
    this.sp = spfi().using(SPFx(this.context));
    return super.onInit();
  }

  public render(): void {
    this._getUserProperties().then(userName => {
      const customMessage = this.properties.customMessage || "Have a great day!";
      const customMessageColor = this.properties.customMessageColor || "#000000"; // Default to black for the custom message
  
      const welcomeUserTextColor = this.properties.welcomeUserTextColor || "#000000"; // Default to black for welcome user text
  
      const buttonText = this.properties.buttonText || "Call to Action";
      const buttonUrl = this.properties.buttonUrl || "#";
      const showButton = this.properties.showButton !== false;
      const buttonColor = this.properties.buttonColor && /^#[0-9A-F]{6}$/i.test(this.properties.buttonColor)
        ? this.properties.buttonColor
        : "#0078d4"; // Default to blue if invalid color
        const buttonTextColor = this.properties.buttonTextColor && /^#[0-9A-F]{6}$/i.test(this.properties.buttonTextColor)
        ? this.properties.buttonTextColor
        : "#FFFFFF"; // Default to white for button text if no valid color is provided
  
      const backgroundType = this.properties.backgroundType || 'color';
      const backgroundValue = this.properties.backgroundValue || "#ffffff"; // Default to white
      const backgroundStyle = backgroundType === 'gradient' ? backgroundValue : backgroundValue;
  
      const customMessageHtml = this.properties.showCustomMessage 
        ? `<span style="flex-grow: 1; text-align: center; font-size: 24px; font-weight: bold; margin: 0 20px; color: ${escape(customMessageColor)};">${escape(customMessage)}</span>`
        : ''; 
  
      this.domElement.innerHTML = `
        <div class="welcomeUser" style="display: flex; justify-content: space-between; align-items: center; background: ${escape(backgroundStyle)}; padding: 10px;">
          <div style="flex: 1; margin-left: 20px; white-space: nowrap; font-size: 24px; font-weight: bold; color: ${escape(welcomeUserTextColor)};">
            Welcome, ${escape(userName)}!
          </div>
          ${customMessageHtml}
          ${showButton ? `
            <div style="margin-right: 20px;">
              <a href="${escape(buttonUrl)}" target="_blank" rel="noopener noreferrer">
                <button class="callToActionBtn" style="padding: 10px 20px; background-color: ${escape(buttonColor)}; color: ${escape(buttonTextColor)}; border: none; border-radius: 5px; cursor: pointer;">
                  ${escape(buttonText)}
                </button>
              </a>
            </div>` : ''}
        </div>`;
    });
  }
  

  

  private async _getUserProperties(): Promise<string> {
    try {
      const user = await this.sp.web.currentUser();
      return user.Title || user.LoginName;
    } catch (error) {
      console.error('Error fetching user properties', error);
      return 'User';
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Customize your web part"
          },
          groups: [
            {
              groupName: "Settings",
              groupFields: [
                PropertyPaneLabel('developerLabel', {
                  text: "Developer: Mark Colbert"
                }),
                PropertyPaneTextField('welcomeUserTextColor', { 
                  label: "Welcome User Text Color (Hex Code)",
                  value: "#000000"  // Default to black for welcome user text
                }),
                PropertyPaneToggle('showCustomMessage', {
                  label: "Show Custom Message",
                  onText: "Show",
                  offText: "Hide",
                  checked: true // Default to showing the message
                }),
                PropertyPaneTextField('customMessage', {
                  label: "Custom Message",
                  value: "Have a great day!"
                }),
                PropertyPaneTextField('customMessageColor', {
                  label: "Custom Message Color (Hex Code)",
                  value: "#000000"  // Default to black for custom message color
                }),
                PropertyPaneToggle('showButton', {
                  label: "Show Call to Action Button",
                  onText: "Show",
                  offText: "Hide",
                  checked: true // Default to show the button
                }),
                PropertyPaneTextField('buttonText', {
                  label: "Button Text",
                  value: "Get Started" // Default button text
                }),
                PropertyPaneTextField('buttonUrl', {
                  label: "Button URL",
                  value: "https://www.bing.com" // Default URL
                }),
                PropertyPaneTextField('buttonColor', {
                  label: "Button Color (Hex Code)",
                  value: "#AFAFAF"  // Default color
                }),
                PropertyPaneTextField('buttonTextColor', {
                  label: "Button Text Color (Hex Code)",
                  value: "#FFFFFF" // Default to white
                }),
                PropertyPaneDropdown('backgroundType', {
                  label: "Background Type",
                  options: [
                    { key: 'color', text: 'Solid Color' },
                    { key: 'gradient', text: 'Gradient' }
                  ],
                  selectedKey: 'color' // Default to solid color
                }),
                PropertyPaneTextField('backgroundValue', {
                  label: "Background Color/Gradient",
                  description: "Enter hex color (e.g., #ff0000) or gradient (e.g., linear-gradient(to right, #ff7e5f, #feb47b))",
                  value: "#ffffff" // Default to white for solid background
                })
              ]
            }
          ]
        }
      ]
    };
  }
  
  
  
  
}
