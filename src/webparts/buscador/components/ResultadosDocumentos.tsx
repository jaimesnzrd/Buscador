// Importa React completo
import * as React from 'react';

// Importa Stack de Fluent UI para layouts
import { Stack } from '@fluentui/react/lib/Stack';

// Importa Icon de Fluent UI para mostrar iconos de tipo de documento
import { Icon } from '@fluentui/react/lib/Icon';

// Importa SPFI para consultas a SharePoint
import { SPFI } from "@pnp/sp";

// Define la interfaz de props del componente
interface IResultadosDocumentosProps {
  resultados: any[]; // Array de resultados a mostrar
  onOpenPreview: (url: string, tipo: string, title: string) => void; // Función para abrir vista previa
  sp: SPFI; // Instancia de SPFI para consultas
}

// Estilo de las cajitas de metadatos
const cajitaStyle: React.CSSProperties = {
  borderRadius: 30,
  padding: '4px 12px',
  backgroundColor: '#f0f0f0',
  fontSize: 12
};

// Componente funcional ResultadosDocumentos
const ResultadosDocumentos: React.FC<IResultadosDocumentosProps> = ({ resultados, onOpenPreview, sp }) => {

  // Estado para guardar metadatos de cada documento
  const [metadatosMap, setMetadatosMap] = React.useState<Record<string, any>>({});

  // Función para devolver el icono según el tipo de documento
  const getIcon = (tipo?: string) => {
    switch(tipo) {
      case 'pdf': 
        return <><Icon iconName="PDF" styles={{ root: { color: '#d13438' } }} /> </>;
      case 'xlsx': 
        return <><Icon iconName="ExcelDocument" styles={{ root: { color: '#217346' } }} /> </>;
      case 'docx': 
        return <><Icon iconName="WordDocument" styles={{ root: { color: '#2B579A' } }} /> </>;
      case 'pptx': 
        return <><Icon iconName="PowerPointDocument" styles={{ root: { color: '#d26e26' } }} /> </>;
      default: 
        return tipo;
    }
  };

  // Función para cargar metadatos de un documento
  const cargarMetadatosFila = async (r: any) => {
    if (!r.Path || metadatosMap[r.Path]) return; // ya cargado
    try {
      const serverPath = r.Path.replace("https://wslg4.sharepoint.com", "");
      const file: any = await sp.web.getFileByServerRelativePath(serverPath)();
      const item: any = await sp.web.getFileByServerRelativePath(serverPath).getItem();

      let authorTitle = '-', editorTitle = '-';
      if (item?.AuthorId) {
        try { const author = await sp.web.getUserById(item.AuthorId)(); authorTitle = author.Title; } catch {}
      }
      if (item?.EditorId) {
        try { const editor = await sp.web.getUserById(item.EditorId)(); editorTitle = editor.Title; } catch {}
      }

      setMetadatosMap(prev => ({
        ...prev,
        [r.Path]: {
          Name: file.Name || file.Title || '-',
          Created: file.TimeCreated || '-',
          Modified: file.TimeLastModified || '-',
          Author: authorTitle,
          Editor: editorTitle,
          ServerRelativeUrl: file.ServerRelativeUrl || '-',
          TimeLastModified: file.TimeLastModified || '-',
          Length: file.Length || '-',
          UniqueId: file.UniqueId || '-'
        }
      }));
    } catch (err) {
      console.error("Error cargando metadatos fila", err);
    }
  };

  // Render principal del componente
  return (
    <>
      {resultados.map((r, idx) => {
        // Llamamos a cargar metadatos al montar la fila
        React.useEffect(() => { cargarMetadatosFila(r); }, [r.Path]);

        return (
          <Stack 
            key={idx} 
            style={{ 
              border: '1px solid #dcdcdc', 
              borderRadius: 6, 
              padding: 12, 
              marginBottom: 8, 
              background: '#ffffff' 
            }}
          >
            {/* Título del documento */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              {r.Tipo && 
                <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                  {getIcon(r.Tipo)}
                </span>
              }
              <strong style={{ color: '#ed1e40' }}>{r.Title?.replace(/_/g, ' ')}</strong>
            </div>

            {/* Ruta por carpetas */}
            {r.Path && (
              <div style={{ fontSize: 12, color: '#333333', marginTop: 4 }}>
                {r.Path
                  .split('/')
                  .slice(r.Path.split('/').indexOf('DocsBuscador') + 1, r.Path.split('/').length - 1)
                  .slice(0, 3)
                  .join(' > ')}
              </div>
            )}

            {/* Metadatos debajo de la ruta */}
            {metadatosMap[r.Path] && (
              <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: 6 }}>
                <div style={cajitaStyle}>Nombre: {metadatosMap[r.Path].Name}</div>
                <div style={cajitaStyle}>Creado: {metadatosMap[r.Path].Created}</div>
                <div style={cajitaStyle}>Modificado: {metadatosMap[r.Path].Modified}</div>
                <div style={cajitaStyle}>Autor: {metadatosMap[r.Path].Author}</div>
                <div style={cajitaStyle}>Editor: {metadatosMap[r.Path].Editor}</div>
                <div style={cajitaStyle}>Ruta: {metadatosMap[r.Path].ServerRelativeUrl}</div>
                <div style={cajitaStyle}>Última modificación: {metadatosMap[r.Path].TimeLastModified}</div>
                <div style={cajitaStyle}>Tamaño: {metadatosMap[r.Path].Length}</div>
                <div style={cajitaStyle}>GUID: {metadatosMap[r.Path].UniqueId}</div>
              </Stack>
            )}

            {/* Vista previa */}
            {r.Path && (
              <span
                onClick={() => onOpenPreview(r.Path, r.Tipo || '', r.Title || '')}
                style={{ 
                  color: 'red', 
                  fontWeight: 600, 
                  textDecoration: 'underline', 
                  cursor: 'pointer', 
                  marginTop: 8, 
                  display: 'inline-block' 
                }}
              >
                Vista previa
              </span>
            )}
          </Stack>
        );
      })}
    </>
  );
};

// Exporta el componente para usarlo en otros archivos
export default ResultadosDocumentos;