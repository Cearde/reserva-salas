import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SrscWebPartStrings';
import Srsc from './components/Srsc';
//import { ISrscProps } from './components/ISrscProps';
import { APP_VERSION } from './utils/utils';
import { AuthProvider } from './contexts/AuthContext';
import { SPService } from './services/sp';
//import { Web } from  "@pnp/sp/webs";
export interface ISrscWebPartProps {
  description: string;
}

export default class SrscWebPart extends BaseClientSideWebPart<ISrscWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _usuarioIDLista: number = 0; // Nueva propiedad para almacenar el userId
  private _usuarioIDDivision: number = 0; // Nueva propiedad para almacenar el userId
  public render(): void {
    /*const element: React.ReactElement<ISrscProps> = React.createElement(
      Srsc,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context
      }
    );
    React.useEffect(() => {
    const webBase = Web(this.context.pageContext.web.absoluteUrl);
    webBase.ensureUser(this.context.pageContext.user.loginName).then(result => {
      this._userId = result.data.Id; // Guardamos el userId en la propiedad del WebPart
    }).catch(error => {});
    }, []);*/


    const element: React.ReactElement = React.createElement(
    AuthProvider,
    {
      context: this.context,
      // Pasamos explícitamente el hijo dentro de las props para satisfacer a TS
      children: React.createElement(Srsc, {
      description: this.properties.description,
      isDarkTheme: this._isDarkTheme,
      environmentMessage: this._environmentMessage,
      hasTeamsContext: !!this.context.sdks.microsoftTeams,
      userDisplayName: this.context.pageContext.user.displayName,
      loginName: this.context.pageContext.user.loginName,
      usuarioIDLista: this._usuarioIDLista,//result.data.Id, // Usamos el ID del usuario obtenido de ensureUser
      usuarioIDDivision: this._usuarioIDDivision, // Puedes obtener esta información adicionalmente si la necesitas
      context: this.context
      })
    }
  );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    this.hideSharePointChrome();
    console.log(`SRSC WebPart Version: ${APP_VERSION}`);
    const link = document.createElement('link');
    link.rel = 'stylesheet';
    link.href = 'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css';
    document.head.appendChild(link);
    this._environmentMessage = await this._getEnvironmentMessage();
    try {
      const spService = new SPService(this.context);
      const usuarios = await spService.getUsuarios(true);
      const emailActual = this.context.pageContext.user.email.toLowerCase();
      
      const usuarioEncontrado = usuarios.find(u => u.email?.toLowerCase() === emailActual);
      this._usuarioIDLista = (usuarioEncontrado && usuarioEncontrado.Id) ?? 0;// usuarioEncontrado ? usuarioEncontrado.Id : 0;
      this._usuarioIDDivision = usuarioEncontrado?.divisionId ?? 0; // Si necesitas la división, la puedes obtener aquí

    } catch (error) {
      console.error("Error al obtener el ID del usuario en onInit:", error);
      this._usuarioIDLista = 0;
      this._usuarioIDDivision = 0;
    }

  return super.onInit();
    /*return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });*/
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  // En tu WebPart.ts o componente raíz
private hideSharePointChrome(): void {
  const style = document.createElement('style');
  style.innerHTML = `
    #spLeftNav, #SuiteNavWrapper, #spSiteHeader, #sp-appBar, #spCommandBar {
      display: none !important;
    }

    #SuiteNavWrapper, #spSiteHeader, #sp-appBar, #spCommandBar

    #contentBox {
      margin-left: 0 !important;
    }
  `;
  document.head.appendChild(style);
}

}
