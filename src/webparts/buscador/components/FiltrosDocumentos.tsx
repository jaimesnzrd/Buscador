// Importa React completo
import * as React from 'react';

// Fluent UI
import { Stack } from '@fluentui/react/lib/Stack';
import { Label } from '@fluentui/react/lib/Label';
import { TextField } from '@fluentui/react/lib/TextField';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { Dropdown, IDropdownOption, IDropdownStyles } from '@fluentui/react/lib/Dropdown';

// Props del componente
interface IFiltrosDocumentosProps {
  titulo: string;
  setTitulo: (v: string) => void;
  fechaDesde?: Date;
  setFechaDesde: (d?: Date) => void;
  fechaHasta?: Date;
  setFechaHasta: (d?: Date) => void;
  opcionesUsuarios: string[];
  filtroCreatedBy: string[];
  setFiltroCreatedBy: (v: string[]) => void;
  opcionesTipoArchivo: string[];
  filtroTipoArchivo: string[];
  setFiltroTipoArchivo: (v: string[]) => void;
}

const FiltrosDocumentos: React.FC<IFiltrosDocumentosProps> = ({
  titulo,
  setTitulo,
  fechaDesde,
  setFechaDesde,
  fechaHasta,
  setFechaHasta,
  opcionesUsuarios,
  filtroCreatedBy,
  setFiltroCreatedBy,
  opcionesTipoArchivo,
  filtroTipoArchivo,
  setFiltroTipoArchivo,
}) => {

  // Conversión de opciones a formato Dropdown
  const opcionesUsuariosDropdown: IDropdownOption[] = opcionesUsuarios.map(o => ({ key: o, text: o }));
  const opcionesTipoDropdown: IDropdownOption[] = opcionesTipoArchivo.map(o => ({ key: o, text: o }));

  // Estilo para que los Dropdowns tengan el mismo ancho que TextField
  const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 170 } };

  return (
    <Stack tokens={{ childrenGap: 10 }}>

      {/* Campo de texto para título */}
      <TextField 
        label="Título" 
        value={titulo} 
        onChange={(_, val) => setTitulo(val || '')} 
      />

      {/* Contenedor horizontal para los DatePickers */}
      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <DatePicker 
          label="Fecha Desde" 
          value={fechaDesde} 
          onSelectDate={date => setFechaDesde(date ?? undefined)} 
        />
        <DatePicker 
          label="Fecha Hasta" 
          value={fechaHasta} 
          onSelectDate={date => setFechaHasta(date ?? undefined)} 
        />
      </Stack>

      {/* Dropdown multi-select para "Creado por" */}
      <Label>Creado por</Label>
      <Dropdown
        placeholder="Selecciona usuarios"
        multiSelect
        selectedKeys={filtroCreatedBy}
        options={opcionesUsuariosDropdown}
        onChange={(_, option) => {
          if (!option) return;
          const key = option.key as string;
          setFiltroCreatedBy(
            option.selected
              ? [...filtroCreatedBy, key]
              : filtroCreatedBy.filter(k => k !== key)
          );
        }}
        styles={dropdownStyles}
      />

      {/* Dropdown multi-select para "Tipo Archivo" */}
      <Label>Tipo Archivo</Label>
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

    </Stack>
  );
};

export default FiltrosDocumentos;