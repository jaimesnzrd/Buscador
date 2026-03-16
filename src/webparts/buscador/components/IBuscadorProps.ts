// Importa SPFI de PnPjs, que se usa para interactuar con SharePoint de forma tipada
import { SPFI } from '@pnp/sp';

// Importa el contexto de la WebPart de SPFx, necesario para acceder a site, lista, usuario, etc
import { WebPartContext } from '@microsoft/sp-webpart-base';

// Definimos la interfaz de props que recibe la WebPart Buscador
export interface IWebpartBuscadorProps {
  description: string;     // Texto descriptivo que puede mostrarse en la webpart
  sp: SPFI;                // Instancia de PnP SPFI para realizar llamadas a SharePoint
  context: WebPartContext; // Contexto de la webpart, incluye info del sitio, usuario y entorno SPFx
}