// Importa React completo
import * as React from 'react';

// Importa Stack de Fluent UI para layouts
import { Stack } from '@fluentui/react/lib/Stack';

// Importa Icon de Fluent UI para mostrar iconos de tipo de documento
import { Icon } from '@fluentui/react/lib/Icon';

// Define la interfaz de props del componente
interface IResultadosDocumentosProps {
  resultados: any[]; // Array de resultados a mostrar
  onOpenPreview: (url: string, tipo: string, title: string) => void; // Función para abrir vista previa
}

// Componente funcional ResultadosDocumentos
const ResultadosDocumentos: React.FC<IResultadosDocumentosProps> = ({ resultados, onOpenPreview }) => {

  // Función para devolver el icono según el tipo de documento
  const getIcon = (tipo?: string) => {
    switch(tipo) {
      case 'pdf': 
        // Icono PDF con color rojo
        return <><Icon iconName="PDF" styles={{ root: { color: '#d13438' } }} /> </>;
      case 'xlsx': 
        // Icono Excel con color rojo
        return <><Icon iconName="ExcelDocument" styles={{ root: { color: '#217346' } }} /> </>;
      case 'docx': 
        // Icono Word con color rojo
        return <><Icon iconName="WordDocument" styles={{ root: { color: '#2B579A' } }} /> </>;
      case 'pptx': 
        // Icono PowerPoint con color rojo
        return <><Icon iconName="PowerPointDocument" styles={{ root: { color: '#d26e26' } }} /> </>;
      default: 
        // Si no es ninguno de los anteriores, muestra el tipo tal cual
        return tipo;
    }
  };

  // Render principal del componente
  return (
    <>
      {resultados.map((r, idx) => (
        // Cada resultado va en un Stack con borde y padding
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
          {/* Título del documento, reemplaza _ por espacios */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            {r.Tipo && 
              <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                {getIcon(r.Tipo)}
              </span>
            }
            <strong>{r.Title?.replace(/_/g, ' ')}</strong>
          </div>

          {/* Fecha y tipo de documento en horizontal */}
          <Stack horizontal horizontalAlign="space-between">
            {r.Created && 
              // Mostrar fecha de creación si existe
              <span><strong>Fecha:</strong> {new Date(r.Created).toLocaleDateString()}</span>
            }
            
          </Stack>

          {/* Vista previa clicable si hay path */}
          {r.Path && (
            <span
              onClick={() => onOpenPreview(r.Path, r.Tipo || '', r.Title || '')} // Llama a la función de vista previa
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
      ))}
    </>
  );
};

// Exporta el componente para usarlo en otros archivos
export default ResultadosDocumentos;