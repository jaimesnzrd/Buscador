import { SPFI } from "@pnp/sp";
import "@pnp/sp/search";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

// =====================
// ===== INTERFACES FILTROS =====
// =====================

// Interfaz para filtros de documentos
export interface IDocumentsFilters {
  texto?: string[];       // Palabras o frases a buscar en el documento
  tipoArchivo?: string[]; // Tipos de archivo permitidos (docx, pdf, etc.)
  autor?: string[];       // Autores permitidos
  carpeta?: string;       // Carpeta específica dentro del site
  fechaDesde?: Date;      // Fecha mínima de creación
  fechaHasta?: Date;      // Fecha máxima de creación
  titulo?: string[];      // Títulos específicos
}

// Interfaz para cada resultado de búsqueda
export interface ISearchResultItem {
  Title?: string;         // Título del documento
  Path?: string;          // Ruta al documento
  Author?: string;        // Autor
  Created?: string;       // Fecha de creación
  Tipo?: string;          // Tipo de documento (FileExtension)
}

// =====================
// ===== SERVICIO DE BÚSQUEDA =====
// =====================
export class SearchService {
  private _sp: SPFI;

  constructor(sp: SPFI) {
    this._sp = sp;
  }

  // =====================
  // ===== BUSCAR DOCUMENTOS =====
  // =====================
  public async buscarDocumentos(
    filtros: IDocumentsFilters,
    startRow = 0
  ): Promise<{ resultados: ISearchResultItem[]; total: number }> {
    try {
      // Base de la query KQL apuntando a la librería de documentos
      let kql = `Path:"https://wslg4.sharepoint.com/sites/WebpartBuscador/DocsBuscador"`;

      // Agregar filtros KQL
      if (filtros.texto?.length) kql += ` AND (${filtros.texto.join(" OR ")})`;
      if (filtros.titulo?.length) kql += ` AND (${filtros.titulo.join(" OR ")})`;
      if (filtros.tipoArchivo?.length) kql += ` AND FileExtension:(${filtros.tipoArchivo.join(" OR ")})`;
      if (filtros.autor?.length) kql += ` AND Author:(${filtros.autor.join(" OR ")})`;
      if (filtros.carpeta) kql += ` AND Path:${filtros.carpeta}`;
      if (filtros.fechaDesde) kql += ` AND Created>=${filtros.fechaDesde.toISOString()}`;
      if (filtros.fechaHasta) kql += ` AND Created<=${filtros.fechaHasta.toISOString()}`;

      // Ejecutar búsqueda con PnPJS Search
      const res = await this._sp.search({
        Querytext: kql,
        RowLimit: 5,
        StartRow: startRow,
        SelectProperties: ["Title", "Path", "Created", "Author", "FileExtension"],
      });

      // Mapear resultados a ISearchResultItem
      const resultados = res.PrimarySearchResults.map((r: any) => ({
        Title: r["Title"],
        Path: r["Path"],
        Created: r["Created"],
        Author: r["Author"],
        Tipo: r["FileExtension"],
      }));

      return { resultados, total: res.TotalRows };
    } catch (err) {
      console.error("Error buscarDocumentos:", err);
      throw err;
    }
  }
  // =====================
  // ===== OBTENER TIPOS DE ARCHIVO REALES =====
  // =====================
  public async obtenerTiposArchivo(): Promise<string[]> {
    try {
      const res = await this._sp.search({
        Querytext: `Path:"https://wslg4.sharepoint.com/sites/WebpartBuscador/DocsBuscador"`,
        RowLimit: 1,           // No nos interesan los resultados en sí
        SelectProperties: [],  // No necesitamos columnas
        Refiners: "FileExtension", // Pide a SharePoint que agregue TODOS los tipos
      });

      const extensions = new Set<string>();

      // Los refiners vienen en res.RawSearchResults.PrimaryQueryResult.RefinementResults
      const refinementResults = (res as any).RawSearchResults?.PrimaryQueryResult?.RefinementResults?.Refiners;

      if (refinementResults) {
        refinementResults.forEach((refiner: any) => {
          if (refiner.Name === "FileExtension") {
            refiner.Entries.forEach((entry: any) => {
              const ext = entry.RefinementName?.toLowerCase();
              if (ext) extensions.add(ext);
            });
          }
        });
      }

      return Array.from(extensions).sort();
    } catch (err) {
      console.error("Error obtenerTiposArchivo:", err);
      return [];
    }
  }
}