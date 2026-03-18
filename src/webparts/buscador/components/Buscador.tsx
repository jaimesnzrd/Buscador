import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Label } from '@fluentui/react/lib/Label';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { SPFI } from '@pnp/sp';
import { IWebpartBuscadorProps } from './IBuscadorProps';
import { SearchService, ICarpetaInfo } from '../services/SearchService';
import FiltrosDocumentos from './FiltrosDocumentos';
import ResultadosDocumentos from './ResultadosDocumentos';
import DocumentoPreview from './DocumentoPreview';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';

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
  const [buscarContenido, setBuscarContenido] = useState(false);

  // Opciones de dropdown
  const [opcionesTipoArchivo, setOpcionesTipoArchivo] = useState<string[]>([]);

  // Resultados
  const [resultados, setResultados] = useState<any[]>([]);
  const [totalResultados, setTotalResultados] = useState(0);
  const [paginaActual, setPaginaActual] = useState(1);
  const resultadosPorPagina = 5;
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Filtros jerárquicos
  const [bloques, setBloques] = useState<ICarpetaInfo[]>([]);
  const [selectedBloques, setSelectedBloques] = useState<string[]>([]);
  const [loadingBloques, setLoadingBloques] = useState(false);
  const [secciones, setSecciones] = useState<ICarpetaInfo[]>([]);
  const [selectedSecciones, setSelectedSecciones] = useState<string[]>([]);
  const [loadingSecciones, setLoadingSecciones] = useState(false);
  const [subSecciones, setSubSecciones] = useState<ICarpetaInfo[]>([]);
  const [selectedSubSecciones, setSelectedSubSecciones] = useState<string[]>([]);
  const [loadingSubSecciones, setLoadingSubSecciones] = useState(false);

  // Selector de biblioteca
  const [bibliotecas, setBibliotecas] = useState<{ nombre: string; path: string }[]>([]);
  const [selectedLibrary, setSelectedLibrary] = useState<string>('');
  const [loadingLibraries, setLoadingLibraries] = useState(false);

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
    setBuscarContenido(false);
    setSelectedBloques([]);
    setSelectedSecciones([]);
    setSelectedSubSecciones([]);
    setSecciones([]);
    setSubSecciones([]);
    setError(null);
  };

  // Cargar bibliotecas al inicio
  useEffect(() => {
    const cargarBibliotecas = async (): Promise<void> => {
      try {
        setLoadingLibraries(true);
        const libs = await searchService.obtenerBibliotecas();
        setBibliotecas(libs);
      } catch (err) {
        console.error("Error cargando bibliotecas:", err);
      } finally {
        setLoadingLibraries(false);
      }
    };
    void cargarBibliotecas();
  }, [sp]);

  // Cuando se selecciona una library, cargar sus datos
  useEffect(() => {
    if (!selectedLibrary) return;
    const cargarDatosLibrary = async (): Promise<void> => {
      try {
        searchService.setLibraryPath(selectedLibrary);

        // Cargar tipos de archivo
        const tipos = await searchService.obtenerTiposArchivo();
        setOpcionesTipoArchivo(tipos);

        // Cargar bloques (carpetas de primer nivel)
        setLoadingBloques(true);
        const bloquesData = await searchService.obtenerCarpetas(selectedLibrary);
        setBloques(bloquesData);
        setLoadingBloques(false);

        // Limpiar resultados y filtros anteriores
        setResultados([]);
        setTotalResultados(0);
        borrarFiltros();
      } catch (err) {
        console.error("Error cargando datos de library:", err);
      }
    };
    void cargarDatosLibrary();
  }, [selectedLibrary]);

  // Cargar Secciones al cambiar Bloques
  useEffect(() => {
    const cargar = async (): Promise<void> => {
      if (selectedBloques.length === 0) {
        setSecciones([]); setSelectedSecciones([]);
        setSubSecciones([]); setSelectedSubSecciones([]);
        return;
      }
      setLoadingSecciones(true);
      const todas: ICarpetaInfo[] = [];
      for (const bp of selectedBloques) {
        const s = await searchService.obtenerCarpetas(bp);
        todas.push(...s);
      }
      setSecciones(todas);
      setSelectedSecciones(prev => prev.filter(s => todas.some(t => t.path === s)));
      setLoadingSecciones(false);
    };
    void cargar();
  }, [selectedBloques]);

  // Cargar Sub-Secciones al cambiar Secciones
  useEffect(() => {
    const cargar = async (): Promise<void> => {
      if (selectedSecciones.length === 0) {
        setSubSecciones([]); setSelectedSubSecciones([]);
        return;
      }
      setLoadingSubSecciones(true);
      const todas: ICarpetaInfo[] = [];
      for (const sp2 of selectedSecciones) {
        const ss = await searchService.obtenerCarpetas(sp2);
        todas.push(...ss);
      }
      setSubSecciones(todas);
      setSelectedSubSecciones(prev => prev.filter(s => todas.some(t => t.path === s)));
      setLoadingSubSecciones(false);
    };
    void cargar();
  }, [selectedSecciones]);

  // Toggles
  const handleToggleBloque = useCallback((path: string): void => {
    setSelectedBloques(prev => prev.includes(path) ? prev.filter(p => p !== path) : [...prev, path]);
  }, []);
  const handleToggleSeccion = useCallback((path: string): void => {
    setSelectedSecciones(prev => prev.includes(path) ? prev.filter(p => p !== path) : [...prev, path]);
  }, []);
  const handleToggleSubSeccion = useCallback((path: string): void => {
    setSelectedSubSecciones(prev => prev.includes(path) ? prev.filter(p => p !== path) : [...prev, path]);
  }, []);

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
          carpetas: selectedSubSecciones.length > 0 ? selectedSubSecciones
            : selectedSecciones.length > 0 ? selectedSecciones
            : selectedBloques.length > 0 ? selectedBloques
            : [],
          fechaDesde: filtroFechaDocsDesde,
          fechaHasta: filtroFechaDocsHasta,
          buscarContenido: buscarContenido
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

      <Dropdown
        label="Biblioteca de documentos"
        placeholder={loadingLibraries ? "Cargando bibliotecas..." : "Selecciona una biblioteca"}
        options={bibliotecas.map(b => ({ key: b.path, text: b.nombre }))}
        selectedKey={selectedLibrary}
        onChange={(_, option) => {
          if (option) setSelectedLibrary(option.key as string);
        }}
        disabled={loadingLibraries}
        styles={{ root: { maxWidth: 400, marginBottom: 20 } }}
      />

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
            buscarContenido={buscarContenido}
            setBuscarContenido={setBuscarContenido}
            bloques={bloques}
            selectedBloques={selectedBloques}
            onToggleBloque={handleToggleBloque}
            loadingBloques={loadingBloques}
            secciones={secciones}
            selectedSecciones={selectedSecciones}
            onToggleSeccion={handleToggleSeccion}
            loadingSecciones={loadingSecciones}
            subSecciones={subSecciones}
            selectedSubSecciones={selectedSubSecciones}
            onToggleSubSeccion={handleToggleSubSeccion}
            loadingSubSecciones={loadingSubSecciones}
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