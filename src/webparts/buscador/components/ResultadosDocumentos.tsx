// Importa React completo
import * as React from 'react';

// Importa Stack de Fluent UI para layouts
import { Stack } from '@fluentui/react/lib/Stack';

// Importa Icon de Fluent UI para mostrar iconos de tipo de documento
import { Icon } from '@fluentui/react/lib/Icon';

// Importa SPFI para consultas a SharePoint
import { SPFI } from "@pnp/sp";

import { TooltipHost } from '@fluentui/react/lib/Tooltip';

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

const scrollContainerStyle: React.CSSProperties = {
  marginTop: 6,
  marginRight: 110,
  display: 'flex',
  gap: 10,
  overflowX: 'auto',
  paddingBottom: 4,
  cursor: 'grab',
  scrollbarWidth: 'none',        // Firefox
  msOverflowStyle: 'none',       // IE 10+
};

// Componente funcional ResultadosDocumentos
const ResultadosDocumentos: React.FC<IResultadosDocumentosProps> = ({ resultados, onOpenPreview, sp }) => {

  // Estado para guardar metadatos de cada documento
  const [metadatosMap, setMetadatosMap] = React.useState<Record<string, any>>({});

  // Función para devolver el icono según el tipo de documento
  const getIcon = (tipo?: string) => {
    const iconStyle = { root: { color: '#d13438', fontSize: 24, width: 24, height: 24 } };

    switch(tipo?.toLowerCase()) {
      case 'pdf': 
        return <Icon iconName="PDF" styles={iconStyle} />;

      case 'xlsx': 
        return <Icon iconName="ExcelDocument" styles={iconStyle} />;

      case 'docx': 
        return <Icon iconName="WordDocument" styles={iconStyle} />;

      case 'pptx': 
        return <Icon iconName="PowerPointDocument" styles={iconStyle} />;

      case 'one':
      case 'onenote':
        return <Icon iconName="OneNoteLogo" styles={iconStyle} />;

      case 'loop':
        return <Icon iconName="Loop" styles={iconStyle} />;

      case 'jpg':
      case 'jpeg':
      case 'png':
      case 'gif':
      case 'photo':
        return <Icon iconName="Photo2" styles={iconStyle} />;

      case 'mp4':
      case 'avi':
      case 'mov':
      case 'video':
        return <Icon iconName="Video" styles={iconStyle} />;

      case 'html':
      case 'htm':
      case 'web':
      case 'url':
        return <Icon iconName="Globe" styles={iconStyle} />;

      default: 
        return <Icon iconName="Page" styles={iconStyle} />; // fallback bonito
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

  // ----------------------------
  // Aquí corregimos el hook
  // ----------------------------
  React.useEffect(() => {
    resultados.forEach(r => {
      if (r.Path && !metadatosMap[r.Path]) {
        void cargarMetadatosFila(r);
      }
    });
  }, [resultados]);

  // Render principal del componente
  return (
    <>
      {resultados.map((r, idx) => (
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
            <TooltipHost content={r.Title?.replace(/_/g, ' ')}>
              <strong
                style={{
                  color: '#D52B1E',
                  fontWeight: 600,
                  fontStyle: 'normal',
                  fontSize: 16,
                  lineHeight: '140%',
                  letterSpacing: '1%',
                  whiteSpace: 'nowrap',
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                  maxWidth: 300,
                  display: 'inline-block'
                }}
              >
                {r.Title?.replace(/_/g, ' ')}
              </strong>
            </TooltipHost>
          </div>

          {/* Ruta por carpetas */}
          {r.Path && (
            <span 
              style={{ fontSize: 12, color: '#333333', marginTop: 4, display: 'inline-block' }}
              onClick={(e) => e.stopPropagation()} 
            >
              {r.Path
                .split('/')
                .slice(r.Path.split('/').indexOf('DocsBuscador') + 1, r.Path.split('/').length - 1)
                .slice(0, 3)
                .join(' > ')}
            </span>
          )}

          {/* Metadatos debajo de la ruta */}
          {metadatosMap[r.Path] && (
            <div
              style={scrollContainerStyle}
              onMouseDown={(e) => {
                const container = e.currentTarget;
                let startX = e.pageX - container.offsetLeft;
                let scrollLeft = container.scrollLeft;

                container.style.cursor = 'grabbing';

                const onMouseMove = (ev: MouseEvent) => {
                  const x = ev.pageX - container.offsetLeft;
                  container.scrollLeft = scrollLeft - (x - startX);
                };

                const onMouseUp = () => {
                  container.style.cursor = 'grab';
                  document.removeEventListener('mousemove', onMouseMove);
                  document.removeEventListener('mouseup', onMouseUp);
                };

                document.addEventListener('mousemove', onMouseMove);
                document.addEventListener('mouseup', onMouseUp);
              }}
            >
              {['Name','Created','Modified','Author','Editor','ServerRelativeUrl','TimeLastModified','Length','UniqueId'].map(key => (
                <div
                  key={key}
                  style={{
                    ...cajitaStyle,
                    height: 27,
                    display: 'flex',
                    alignItems: 'center',
                    whiteSpace: 'nowrap',
                    flexShrink: 0,
                    fontWeight: 400,  
                    fontSize: 14      
                  }}
                  onClick={(e) => e.stopPropagation()}
                >
                  {`${key}: ${metadatosMap[r.Path][key]}`}
                </div>
              ))}
            </div>
          )}

          {/* Vista previa: separada en un div independiente */}
          {r.Path && (
            <div style={{ marginTop: 8 }}>
              <span
                onClick={(e) => {
                  e.stopPropagation();
                  onOpenPreview(r.Path, r.Tipo || '', r.Title || '');
                }}
                style={{ 
                  color: 'red', 
                  fontWeight: 400, 
                  textDecoration: 'underline', 
                  cursor: 'pointer', 
                  display: 'inline-block'
                }}
              >
                Vista previa
              </span>
            </div>
          )}
        </Stack>
      ))}
    </>
  );
};

// Exporta el componente para usarlo en otros archivos
export default ResultadosDocumentos;