// Importa React y ReactDOM para renderizar componentes
import * as React from 'react';
import * as ReactDom from 'react-dom';

// Importa SPFx: Version, property pane y base de WebPart
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

// Importa cadenas de traducción/localización
import * as strings from 'BuscadorWebPartStrings';

// Importa el componente principal React y la interfaz de props
import WebpartBuscador from './components/Buscador';
import { IWebpartBuscadorProps } from './components/IBuscadorProps';

// Importa PnPjs y presets para SPFx
import { spfi, SPFI } from '@pnp/sp';
import { SPFx } from '@pnp/sp/presets/all';

// Interfaz de props de la webpart (editable desde property pane)
export interface IWebpartBuscadorWebPartProps {
  description: string; // Texto de descripción configurable por el usuario
}

// Clase principal de la WebPart que hereda de BaseClientSideWebPart
export default class WebpartBuscadorWebPart extends BaseClientSideWebPart<IWebpartBuscadorWebPartProps> {

  private _sp!: SPFI; // Instancia de SPFI que usaremos en el componente React

  // =========================
  // Inicialización de la WebPart
  // =========================
  protected async onInit(): Promise<void> {
    await super.onInit(); // Llamada al init del padre

    // Creamos la instancia de SPFI usando el contexto de SPFx
    // Esto permite que todas las llamadas a SharePoint se realicen correctamente con contexto actual
    this._sp = spfi(`https://wslg4.sharepoint.com/sites/WebpartBuscador`).using(SPFx(this.context));
  }

  // =========================
  // Renderizamos el componente React
  // =========================
  public render(): void {
    // Creamos el elemento React con las props que necesita
    const element: React.ReactElement<IWebpartBuscadorProps> = React.createElement(
      WebpartBuscador,
      {
        description: this.properties.description, // Pasamos la descripción
        sp: this._sp,                             // Pasamos la instancia SPFI
        context: this.context                     // Pasamos el contexto SPFx
      }
    );

    // Renderizamos el elemento React en el contenedor de la WebPart
    ReactDom.render(element, this.domElement);
  }

  // =========================
  // Limpieza cuando se elimina la WebPart
  // =========================
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement); // Desmonta React
  }

  // =========================
  // Versión de datos de la WebPart
  // =========================
  protected get dataVersion(): Version {
    return Version.parse('1.0'); // Version inicial de datos
  }

  // =========================
  // Configuración del Property Pane (panel de propiedades editable)
  // =========================
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription }, // Texto de cabecera
          groups: [
            {
              groupName: strings.BasicGroupName, // Nombre del grupo
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel // Campo editable de descripción
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

// =========================
// CONFIGURACIÓN DE THEME FLUENT UI
// =========================
import { loadTheme } from '@fluentui/react';

// Carga un theme personalizado para Fluent UI (colores corporativos Andbank)
loadTheme({
  palette: {
    themePrimary: '#c8102e',     // Color principal
    themeLighterAlt: '#fff5f6',
    themeLighter: '#f6d6da',
    themeLight: '#ecaab2',
    themeTertiary: '#d35a6b',
    themeSecondary: '#bf0f2c',
    themeDarkAlt: '#a70d26',
    themeDark: '#8c0b20',
    themeDarker: '#660816',

    neutralLighterAlt: '#fafafa',
    neutralLighter: '#f4f4f4',
    neutralLight: '#eaeaea',
    neutralQuaternaryAlt: '#d6d6d6',
    neutralQuaternary: '#cfcfcf',
    neutralTertiaryAlt: '#c8c8c8',
    neutralTertiary: '#6f6f6f',

    neutralSecondary: '#4a4a4a',
    neutralPrimaryAlt: '#3a3a3a',
    neutralPrimary: '#2b2b2b',
    neutralDark: '#1c1c1c',

    black: '#000000',
    white: '#ffffff',
  }
});