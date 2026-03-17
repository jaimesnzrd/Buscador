import * as React from 'react';
import { useState, useEffect } from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Label } from '@fluentui/react/lib/Label';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { SPFI } from '@pnp/sp';
import { IWebpartBuscadorProps } from './IBuscadorProps';
import { SearchService } from '../services/SearchService';
import FiltrosDocumentos from './FiltrosDocumentos';
import ResultadosDocumentos from './ResultadosDocumentos';
import DocumentoPreview from './DocumentoPreview';

interface IBuscadorPropsExtended extends IWebpartBuscadorProps {
  sp: SPFI;
}

const BuscadorDocumentos: React.FC<IBuscadorPropsExtended> = ({ description, sp }) => {
  const searchService = React.useMemo(() => new SearchService(sp), [sp]);

  // Estados de búsqueda
  const [tipoBusqueda, setTipoBusqueda] = useState<'documentos' | null>('documentos');
  const [filtroTitulo, setFiltroTitulo] = useState('');
  const [filtroTipoArchivo, setFiltroTipoArchivo] = useState<string[]>([]);
  const [filtroCarpeta, setFiltroCarpeta] = useState('');
  const [filtroFechaDocsDesde, setFiltroFechaDocsDesde] = useState<Date | undefined>();
  const [filtroFechaDocsHasta, setFiltroFechaDocsHasta] = useState<Date | undefined>();
  const [filtroTexto, setFiltroTexto] = useState<string[]>([]);

  // Opciones de dropdown
  const [opcionesTipoArchivo, setOpcionesTipoArchivo] = useState<string[]>([]);

  // Resultados
  const [resultados, setResultados] = useState<any[]>([]);
  const [totalResultados, setTotalResultados] = useState(0);
  const [paginaActual, setPaginaActual] = useState(1);
  const resultadosPorPagina = 5;
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Modal de preview
  const [modalOpen, setModalOpen] = useState(false);
  const [modalUrl, setModalUrl] = useState('');
  const [modalTipo, setModalTipo] = useState('');
  const [modalTitle, setModalTitle] = useState('');

  // Función para borrar filtros
  const borrarFiltros = () => {
    setFiltroTitulo('');
    setFiltroTipoArchivo([]);
    setFiltroCarpeta('');
    setFiltroFechaDocsDesde(undefined);
    setFiltroFechaDocsHasta(undefined);
    setFiltroTexto([]);
    setError(null);
  };

  // Cargar opciones de usuarios al iniciar
  useEffect(() => {
    const cargarOpciones = async () => {
      try {
        const tipos = await searchService.obtenerTiposArchivo();
        setOpcionesTipoArchivo(tipos);

      } catch (err) {
        console.error("Error cargando opciones:", err);
      }
    };
    void cargarOpciones();
  }, [sp]);

  // Función de búsqueda
  const buscar = async (pagina: number = 1) => {
    if (!tipoBusqueda) return;
    setLoading(true);
    setError(null);
    setPaginaActual(pagina);

    try {
      const { resultados: docRes, total } = await searchService.buscarDocumentos(
        {
          texto: filtroTexto,
          tipoArchivo: filtroTipoArchivo,
          carpeta: filtroCarpeta,
          titulo: filtroTitulo ? [filtroTitulo] : [],
          fechaDesde: filtroFechaDocsDesde,
          fechaHasta: filtroFechaDocsHasta
        },
        (pagina - 1) * resultadosPorPagina
      );
      setResultados(docRes);
      setTotalResultados(total);
    } catch (err: any) {
      setError(err.message || 'Error en la búsqueda');
    } finally {
      setLoading(false);
    }
  };

  // Paginación
  const paginasTotales = Math.ceil(totalResultados / resultadosPorPagina);
  const siguientePagina = () => { if (paginaActual < paginasTotales) void buscar(paginaActual + 1); };
  const anteriorPagina = () => { if (paginaActual > 1) void buscar(paginaActual - 1); };

  // Abrir modal de preview
  const handleOpenPreview = (url: string, tipo: string, title: string) => {
    setModalUrl(url);
    setModalTipo(tipo);
    setModalTitle(title);
    setModalOpen(true);
  };

  return (
    <div style={{ padding: 20 }}>
      <p>{description}</p>
      <h2>Buscador de Documentos</h2>

      <Stack horizontal tokens={{ childrenGap: 10 }}>
        {/* Columna de filtros */}
        <Stack tokens={{ childrenGap: 10 }} styles={{ root: { width: '30%', border: '1px solid #cd0000', padding: 10 } }}>
          <Label>Filtros</Label>

          <FiltrosDocumentos
            titulo={filtroTitulo}
            setTitulo={setFiltroTitulo}
            fechaDesde={filtroFechaDocsDesde}
            setFechaDesde={setFiltroFechaDocsDesde}
            fechaHasta={filtroFechaDocsHasta}
            setFechaHasta={setFiltroFechaDocsHasta}
            opcionesTipoArchivo={opcionesTipoArchivo}
            filtroTipoArchivo={filtroTipoArchivo}
            setFiltroTipoArchivo={setFiltroTipoArchivo}
          />

          <Stack horizontal tokens={{ childrenGap: 10 }}>
            <PrimaryButton text="Buscar" onClick={() => void buscar(1)} />
            <DefaultButton text="Borrar filtros" onClick={borrarFiltros} />
          </Stack>
        </Stack>

        {/* Columna de resultados */}
        <Stack tokens={{ childrenGap: 10 }} styles={{ root: { width: '70%', padding: 10 } }}>
          {error && <p style={{ color: 'red' }}>{error}</p>}

          {resultados.length > 0 && (
            <div>
              <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: '10px', fontWeight: 200 }}>
                Total resultados: {totalResultados}
              </div>
              <ResultadosDocumentos resultados={resultados} onOpenPreview={handleOpenPreview} />
            </div>
          )}

          {totalResultados > 0 && (
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center" tokens={{ childrenGap: 10 }}>
              <Stack horizontal tokens={{ childrenGap: 10 }}>
                <DefaultButton text="Anterior" onClick={anteriorPagina} disabled={paginaActual === 1} />
                <DefaultButton text="Siguiente" onClick={siguientePagina} disabled={paginaActual === paginasTotales} />
              </Stack>
              <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: '10px', fontWeight: 200 }}>
                <span>Página {paginaActual} de {paginasTotales}</span>
              </div>
            </Stack>
          )}
        </Stack>
      </Stack>

      <DocumentoPreview
        sp={sp}
        isOpen={modalOpen}
        onDismiss={() => setModalOpen(false)}
        url={modalUrl}
        tipo={modalTipo}
        title={modalTitle}
      />
    </div>
  );
};

export default BuscadorDocumentos;