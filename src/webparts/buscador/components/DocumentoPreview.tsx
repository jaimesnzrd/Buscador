// Importa React completo
import * as React from 'react';

// Importa useState para manejar estados internos del componente
import { useState } from 'react';

// Importa componentes de Fluent UI para UI: Modal, IconButton, Stack, PrimaryButton y Spinner
import { Modal, IconButton, Stack, PrimaryButton, Spinner, SpinnerSize } from '@fluentui/react';

// Importa SPFI de PnPjs para interactuar con SharePoint
import { SPFI } from "@pnp/sp";

// Define las props que recibe el componente DocumentoPreview
interface DocumentoPreviewProps {
  sp: SPFI;            // Instancia de SPFI para hacer consultas a SharePoint
  isOpen: boolean;     // Si el modal principal está abierto
  onDismiss: () => void; // Función para cerrar el modal
  url: string;         // URL del documento a previsualizar
  tipo: string;        // Tipo de archivo (pdf, docx, xlsx, pptx)
  title: string;       // Título del documento
}

// Componente funcional DocumentoPreview
const DocumentoPreview: React.FC<DocumentoPreviewProps> = ({ sp, isOpen, onDismiss, url, tipo, title }) => {

  // Estado para almacenar metadatos del documento
  const [metadatos, setMetadatos] = useState<any>(null);

  // Estado que indica si los metadatos se están cargando
  const [loadingMeta, setLoadingMeta] = useState(false);

  // Estado que controla el modal de metadatos
  const [metaModalOpen, setMetaModalOpen] = useState(false);

  // URLs base de Office para previsualización online según tipo de archivo
  const officeUrls: Record<string, string> = {
    pptx: 'https://wslg4.sharepoint.com/:p:/r/sites/WebpartBuscador/_layouts/15/Doc.aspx',
    xlsx: 'https://wslg4.sharepoint.com/:x:/r/sites/WebpartBuscador/_layouts/15/Doc.aspx',
    docx: 'https://wslg4.sharepoint.com/:w:/r/sites/WebpartBuscador/_layouts/15/Doc.aspx'
  };

  // =========================
  // FUNCION PARA CARGAR METADATOS DEL DOCUMENTO
  // =========================
  const cargarMetadatos = async () => {
    if (!url) return; // Si no hay URL, no hace nada
    try {
      setLoadingMeta(true); // Activa spinner

      // Convierte URL completa a path relativo del sitio
      const serverPath = url.replace("https://wslg4.sharepoint.com", "");

      // 1. Obtenemos el archivo con info básica: Name, Created, Modified
      const file: any = await sp.web.getFileByServerRelativePath(serverPath)();

      // 2. Obtenemos el item de lista asociado para tener AuthorId y EditorId
      const item: any = await sp.web.getFileByServerRelativePath(serverPath).getItem();

      // 3. Inicializa nombres de Author y Editor
      let authorTitle = '-', editorTitle = '-';

      // Si hay AuthorId, obtenemos su título
      if (item?.AuthorId) {
        try {
          const author = await sp.web.getUserById(item.AuthorId)();
          authorTitle = author.Title;
        } catch {}
      }

      // Si hay EditorId, obtenemos su título
      if (item?.EditorId) {
        try {
          const editor = await sp.web.getUserById(item.EditorId)();
          editorTitle = editor.Title;
        } catch {}
      }

      // 4. Guardamos todos los metadatos en estado
      setMetadatos({
        Name: file.Name || file.Title || '-', // Nombre del archivo
        Created: file.TimeCreated || '-', // Fecha de creación
        Modified: file.TimeLastModified || '-', // Fecha de última modificación
        Author: authorTitle, // Autor
        Editor: editorTitle, // Editor
        ServerRelativeUrl: file.ServerRelativeUrl || '-', // Path relativo
        TimeLastModified: file.TimeLastModified || '-', // Última modificación
        Length: file.Length || '-', // Tamaño en bytes
        UniqueId: file.UniqueId || '-', // GUID del archivo
      });

      setMetaModalOpen(true); // Abre modal de metadatos
    } catch (err) {
      console.error("Error cargando metadatos", err);
      // Si falla, setea valores por defecto
      setMetadatos({
        Name: '-',
        Created: '-',
        Modified: '-',
        Author: '-',
        Editor: '-'
      });
      setMetaModalOpen(true);
    } finally {
      setLoadingMeta(false); // Apaga spinner
    }
  };

  // =========================
  // FUNCION PARA RENDERIZAR EL CONTENIDO DEL DOCUMENTO
  // =========================
  const renderContenido = () => {
    if (!url) return <p>Documento no disponible</p>; // Si no hay URL, mensaje

    const ext = tipo?.toLowerCase(); // Tipo de archivo en minúsculas

    // Si es PDF, iframe directo
    if (ext === 'pdf') {
      return <iframe src={url} width="100%" height="600px" title={title} style={{ border: 'none' }} />;
    }

    // Si es Office (docx, xlsx, pptx) usamos la URL de Office Online
    if (ext && officeUrls[ext]) {
      const urlFinal = `${officeUrls[ext]}?sourcedoc=${encodeURIComponent(url)}&file=${encodeURIComponent(title || '')}&action=default&mobileredirect=true`;
      return <iframe src={urlFinal} width="100%" height="600px" title={title} style={{ border: 'none' }} />;
    }

    // Otros tipos no previsualizables
    return <p>No se puede previsualizar este tipo de archivo</p>;
  };

  // =========================
  // FUNCION PARA RENDERIZAR METADATOS EN FORMATO VERTICAL
  // =========================
  const renderMetadatos = () => {
    if (!metadatos) return <p>No hay metadatos disponibles.</p>;

    const campos = [
      { key: 'Name', label: 'Nombre' },
      { key: 'Created', label: 'Creado' },
      { key: 'Modified', label: 'Modificado' },
      { key: 'Author', label: 'Creado por' },
      { key: 'Editor', label: 'Modificado por' },
      { key: 'ServerRelativeUrl', label: 'Ruta relativa' },
      { key: 'TimeLastModified', label: 'Última modificación (TimeLastModified)' },
      { key: 'Length', label: 'Tamaño del archivo en bytes' },
      { key: 'UniqueId', label: 'GUID' }
    ];

    // Dividir campos en 2 columnas, máximo 7 filas por columna
    const columnas = [campos.slice(0, 7), campos.slice(7)];

    return (
      <div style={{ display: 'flex', gap: 40, justifyContent: 'center', marginTop: 10 }}>
        {columnas.map((col, i) => (
          <div key={i} style={{ display: 'flex', flexDirection: 'column', gap: 15 }}>
            {col.map(c => (
              <div key={c.key} style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
                <span style={{ fontWeight: 600 }}>{c.label}</span>
                <span style={{ wordBreak: 'break-all', color: '#333' }}>{metadatos[c.key] || '-'}</span>
              </div>
            ))}
          </div>
        ))}
      </div>
    );
  };

  // =========================
  // RETURN DEL COMPONENTE
  // =========================
  return (
    <>
      {/* Modal principal de documento */}
      <Modal
        isOpen={isOpen}          // Controla si se muestra
        onDismiss={onDismiss}    // Función para cerrar
        isBlocking={false}       // No bloquea fondo
        styles={{ main: { maxWidth: '90%', minWidth: 600 } }}
      >
        <Stack tokens={{ childrenGap: 10 }} style={{ padding: 20 }}>
          {/* Header con título y botón cerrar */}
          <Stack horizontal horizontalAlign="space-between">
            <h3>{title}</h3>
            <IconButton iconProps={{ iconName: 'Cancel' }} onClick={onDismiss} />
          </Stack>

          {/* Contenido del documento */}
          {renderContenido()}

          {/* Botón y spinner para metadatos */}
          <Stack style={{ marginTop: 10 }}>
            <PrimaryButton text="Mostrar metadatos" onClick={cargarMetadatos} />
            {loadingMeta && <Spinner label="Cargando metadatos..." size={SpinnerSize.medium} />}
          </Stack>
        </Stack>
      </Modal>

      {/* Modal de metadatos */}
      <Modal
        isOpen={metaModalOpen}                   // Controla si se muestra
        onDismiss={() => setMetaModalOpen(false)} // Cierra modal de metadatos
        isBlocking={false}                       // No bloquea fondo
        styles={{ 
          main: { 
            width: 700,           // ancho fijo
            height: 740,          // alto fijo
            maxWidth: 700,        // asegura que no supere 700px
            maxHeight: 740,       // asegura que no supere 740px
            minWidth: 700,        // mínimo ancho
            minHeight: 740,       // mínimo alto
            overflowY: 'auto',    // scroll si el contenido excede
            borderTopLeftRadius: 8,
            borderBottomLeftRadius: 8,
            opacity: 1            // opacidad completa
          }
        }}
      >
        <Stack tokens={{ childrenGap: 10 }} style={{ padding: 20 }}>
          {/* Header del modal de metadatos */}
          <Stack horizontal horizontalAlign="space-between">
            <h3>Metadatos de {title}</h3>
            <IconButton iconProps={{ iconName: 'Cancel' }} onClick={() => setMetaModalOpen(false)} />
          </Stack>

          {/* Render de los metadatos */}
          {renderMetadatos()}
        </Stack>
      </Modal>
    </>
  );
};

// Exporta componente para usarlo en otros lugares
export default DocumentoPreview;