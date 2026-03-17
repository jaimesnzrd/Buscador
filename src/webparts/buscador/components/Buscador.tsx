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

const opcionesTipoArchivo = ['docx', 'xlsx', 'pdf', 'pptx'];

const BuscadorDocumentos: React.FC<IBuscadorPropsExtended> = ({ description, sp }) => {
  const searchService = React.useMemo(() => new SearchService(sp), [sp]);

  // Estados de búsqueda
  const [tipoBusqueda, setTipoBusqueda] = useState<'documentos' | null>('documentos');
  const [filtroTitulo, setFiltroTitulo] = useState('');
  const [filtroCreatedBy, setFiltroCreatedBy] = useState<string[]>([]);
  const [filtroTipoArchivo, setFiltroTipoArchivo] = useState<string[]>([]);
  const [filtroAutor, setFiltroAutor] = useState<string[]>([]);
  const [filtroCarpeta, setFiltroCarpeta] = useState('');
  const [filtroFechaDocsDesde, setFiltroFechaDocsDesde] = useState<Date | undefined>();
  const [filtroFechaDocsHasta, setFiltroFechaDocsHasta] = useState<Date | undefined>();
  const [filtroTexto, setFiltroTexto] = useState<string[]>([]);

  // Opciones de dropdown
  const [opcionesUsuarios, setOpcionesUsuarios] = useState<string[]>([]);

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

  // Función para limpiar filtros
  const limpiarFiltros = () => {
    setFiltroTitulo('');
    setFiltroCreatedBy([]);
    setFiltroTipoArchivo([]);
    setFiltroAutor([]);
    setFiltroCarpeta('');
    setFiltroFechaDocsDesde(undefined);
    setFiltroFechaDocsHasta(undefined);
    setFiltroTexto([]);
    setResultados([]);
    setTotalResultados(0);
    setPaginaActual(1);
    setError(null);
  };

  // Cargar opciones de usuarios al iniciar
  useEffect(() => {
    const cargarOpciones = async () => {
      try {
        // Solo documentos
        const items: any[] = await sp.web.lists.getByTitle("DocsBuscador")
          .items.top(5000)
          .select("Author/Title", "Title", "Created")
          .expand("Author")();
        setOpcionesUsuarios(Array.from(new Set(items.map(i => i.Author?.Title || i.Author).filter(Boolean))));
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
          autor: filtroAutor,
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
            opcionesUsuarios={opcionesUsuarios}
            filtroCreatedBy={filtroCreatedBy}
            setFiltroCreatedBy={setFiltroCreatedBy}
            opcionesTipoArchivo={opcionesTipoArchivo}
            filtroTipoArchivo={filtroTipoArchivo}
            setFiltroTipoArchivo={setFiltroTipoArchivo}
          />

          <Stack horizontal tokens={{ childrenGap: 10 }}>
            <PrimaryButton text="Buscar" onClick={() => void buscar(1)} />
            <DefaultButton text="Limpiar" onClick={limpiarFiltros} />
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
              <ResultadosDocumentos resultados={resultados} onOpenPreview={handleOpenPreview} sp={sp}/>
            </div>
          )}

          {totalResultados > 0 && (
          <div style={{ display: 'flex', justifyContent: 'center', marginTop: 10 }}>
            <div
              style={{
                width: 184,
                height: 35,
                display: 'flex',
                border: '1px solid #CFD8DC',
                borderRadius: 4,
                overflow: 'hidden',
                fontFamily: 'Roboto',
                fontSize: 16,
                fontWeight: 400,
                lineHeight: '100%',
                letterSpacing: '1%',
                textAlign: 'center'
              }}
            >
              {/* Botón Anterior */}
              <div
                onClick={paginaActual > 1 ? anteriorPagina : undefined}
                style={{
                  width: 76,
                  height: 35,
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  borderRight: '1px solid #CFD8DC',
                  cursor: paginaActual > 1 ? 'pointer' : 'default',
                  backgroundColor: '#fff',
                  borderTopRightRadius: 4,
                  borderBottomRightRadius: 4
                }}
                onMouseEnter={e => { (e.currentTarget.querySelector('span') as HTMLSpanElement).style.color = '#D52B1E'; }}
                onMouseLeave={e => { (e.currentTarget.querySelector('span') as HTMLSpanElement).style.color = paginaActual > 1 ? '#000' : '#999'; }}
              >
                <span style={{ color: paginaActual > 1 ? '#000' : '#999' }}>Anterior</span>
              </div>

              {/* Número de página */}
              <div
                style={{
                  width: 35,
                  height: 35,
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  backgroundColor: '#D52B1E',
                  color: '#fff',
                  borderRight: '1px solid #CFD8DC'
                }}
              >
                {paginaActual}
              </div>

              {/* Botón Siguiente */}
              <div
                onClick={paginaActual < paginasTotales ? siguientePagina : undefined}
                style={{
                  width: 76,
                  height: 35,
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  cursor: paginaActual < paginasTotales ? 'pointer' : 'default',
                  backgroundColor: '#fff',
                  borderTopRightRadius: 4,
                  borderBottomRightRadius: 4
                }}
                onMouseEnter={e => { (e.currentTarget.querySelector('span') as HTMLSpanElement).style.color = '#D52B1E'; }}
                onMouseLeave={e => { (e.currentTarget.querySelector('span') as HTMLSpanElement).style.color = paginaActual < paginasTotales ? '#000' : '#999'; }}
              >
                <span style={{ color: paginaActual < paginasTotales ? '#000' : '#999' }}>Siguiente</span>
              </div>
            </div>
          </div>
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