// Importa React completo
import * as React from 'react';

// Fluent UI
import { Stack } from '@fluentui/react/lib/Stack';
import { Label } from '@fluentui/react/lib/Label';
import { TextField } from '@fluentui/react/lib/TextField';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { DayOfWeek } from '@fluentui/react/lib/Calendar';
import { Dropdown, IDropdownOption, IDropdownStyles } from '@fluentui/react/lib/Dropdown';
import { Checkbox } from '@fluentui/react/lib/Checkbox';

// Carpetas filtro
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Link } from '@fluentui/react/lib/Link';
import { ICarpetaInfo } from '../services/SearchService';

// Props del componente
interface IFiltrosDocumentosProps {
  titulo: string;
  setTitulo: (v: string) => void;
  fechaDesde?: Date;
  setFechaDesde: (d?: Date) => void;
  fechaHasta?: Date;
  setFechaHasta: (d?: Date) => void;
  opcionesTipoArchivo: string[];
  filtroTipoArchivo: string[];
  setFiltroTipoArchivo: (v: string[]) => void;
  buscarContenido: boolean;
  setBuscarContenido: (v: boolean) => void;
  bloques: ICarpetaInfo[];
  selectedBloques: string[];
  onToggleBloque: (path: string) => void;
  loadingBloques: boolean;
  secciones: ICarpetaInfo[];
  selectedSecciones: string[];
  onToggleSeccion: (path: string) => void;
  loadingSecciones: boolean;
  subSecciones: ICarpetaInfo[];
  selectedSubSecciones: string[];
  onToggleSubSeccion: (path: string) => void;
  loadingSubSecciones: boolean;
}

const FiltrosDocumentos: React.FC<IFiltrosDocumentosProps> = ({
  titulo,
  setTitulo,
  fechaDesde,
  setFechaDesde,
  fechaHasta,
  setFechaHasta,
  opcionesTipoArchivo,
  filtroTipoArchivo,
  setFiltroTipoArchivo,
  buscarContenido,
  setBuscarContenido,
  bloques, selectedBloques, onToggleBloque, loadingBloques,
  secciones, selectedSecciones, onToggleSeccion, loadingSecciones,
  subSecciones, selectedSubSecciones, onToggleSubSeccion, loadingSubSecciones,
}) => {

  // Conversión de opciones a formato Dropdown
  const opcionesTipoDropdown: IDropdownOption[] = opcionesTipoArchivo.map(o => ({ key: o, text: o }));

  // Estilo para que los Dropdowns tengan el mismo ancho que TextField
  const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 170 } };

  // Formatea fecha a DD/MM/AAAA
  const formatDate = (date?: Date): string => {
    if (!date) return '';
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
  };

  // Strings en español para el calendario interno del DatePicker
  const spanishCalendarStrings = {
    months: ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'],
    shortMonths: ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'],
    days: ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'],
    shortDays: ['D','L','M','X','J','V','S'],
    goToToday: 'Ir a hoy',
    prevMonthAriaLabel: 'Mes anterior',
    nextMonthAriaLabel: 'Mes siguiente',
    prevYearAriaLabel: 'Año anterior',
    nextYearAriaLabel: 'Año siguiente',
    closeButtonAriaLabel: 'Cerrar',
    monthPickerHeaderAriaLabel: '{0}, selecciona para cambiar el año',
    yearPickerHeaderAriaLabel: '{0}, selecciona para cambiar el mes',
  };

  const hoy = new Date();

  return (
    <Stack tokens={{ childrenGap: 10 }}>

      {/* Campo de texto para título */}
      <TextField 
        label="Nombre del documento" 
        value={titulo} 
        onChange={(_, val) => setTitulo(val || '')} 
      />

      {/* Checkbox para buscar dentro del contenido */}
      <Checkbox
        label="Buscar dentro de los documentos"
        checked={buscarContenido}
        onChange={(_, checked) => setBuscarContenido(!!checked)}
      />

      {/* Contenedor horizontal para los DatePickers */}
      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <DatePicker 
          label="Fecha desde" 
          value={fechaDesde} 
          onSelectDate={date => setFechaDesde(date ?? undefined)}
          maxDate={fechaHasta || hoy}
          formatDate={formatDate}
          firstDayOfWeek={DayOfWeek.Monday}
          strings={spanishCalendarStrings}
        />
        <DatePicker 
          label="Fecha hasta" 
          value={fechaHasta} 
          onSelectDate={date => setFechaHasta(date ?? undefined)}
          minDate={fechaDesde}
          maxDate={hoy}
          formatDate={formatDate}
          firstDayOfWeek={DayOfWeek.Monday}
          strings={spanishCalendarStrings}
        />
      </Stack>

      {/* Dropdown multi-select para "Formato de archivo" */}
      <Label>Formato de archivo</Label>
      <Dropdown
        placeholder="Selecciona tipos"
        multiSelect
        selectedKeys={filtroTipoArchivo}
        options={opcionesTipoDropdown}
        onChange={(_, option) => {
          if (!option) return;
          const key = option.key as string;
          setFiltroTipoArchivo(
            option.selected
              ? [...filtroTipoArchivo, key]
              : filtroTipoArchivo.filter(k => k !== key)
          );
        }}
        styles={dropdownStyles}
      />

      {/* Bloques */}
      <div>
        <Label styles={{ root: { fontWeight: 700 } }}>Bloques</Label>
        {loadingBloques ? <Spinner size={SpinnerSize.small} /> : bloques.map(b => (
          <Checkbox
            key={b.path}
            label={`${b.nombre} (${b.count})`}
            checked={selectedBloques.indexOf(b.path) !== -1}
            onChange={() => onToggleBloque(b.path)}
            styles={{ root: { marginTop: 4 } }}
          />
        ))}
      </div>

      {/* Secciones (solo si hay bloque seleccionado) */}
      {selectedBloques.length > 0 && (
        <div>
          <Label styles={{ root: { fontWeight: 700 } }}>Secciones</Label>
          {loadingSecciones ? <Spinner size={SpinnerSize.small} /> : secciones.map(s => (
            <Checkbox
              key={s.path}
              label={`${s.nombre} (${s.count})`}
              checked={selectedSecciones.indexOf(s.path) !== -1}
              onChange={() => onToggleSeccion(s.path)}
              styles={{ root: { marginTop: 4 } }}
            />
          ))}
        </div>
      )}

      {/* Sub-Secciones (solo si hay sección seleccionada) */}
      {selectedSecciones.length > 0 && (
        <div>
          <Label styles={{ root: { fontWeight: 700 } }}>Sub-Secciones</Label>
          {loadingSubSecciones ? <Spinner size={SpinnerSize.small} /> : subSecciones.map(ss => (
            <Checkbox
              key={ss.path}
              label={`${ss.nombre} (${ss.count})`}
              checked={selectedSubSecciones.indexOf(ss.path) !== -1}
              onChange={() => onToggleSubSeccion(ss.path)}
              styles={{ root: { marginTop: 4 } }}
            />
          ))}
        </div>
      )}

    </Stack>
  );
};

export default FiltrosDocumentos;