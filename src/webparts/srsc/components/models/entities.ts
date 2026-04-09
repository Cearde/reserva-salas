//import { IPersonaProps } from '@fluentui/react/lib/Persona';





export interface ISPItem {
  Id: number;
  Title: string;
}

// New interface for a more specific list item type
export interface ISPListItem {
  Id: number;
  Title: string;
  activo: boolean;
  IMAGEN?: string;
}

// Interface for the raw SharePoint item from LM_PISOS
export interface ISPPisoItem {
  Id: number;
  Title: string;
  activo: boolean;
  IMAGEN?: string;
  PLANTA?: {
    Id: number;
    Title: string; // Added Title
  };
}

// New interface for Piso items used in the component
export interface IPisoItem {
  Id?: number; // Optional for new items
  Title: string;
  activo: boolean;
  IMAGEN?: string;
  idImgPiso?: number; // New field for image ID
  PlantaId?: number; // Lookup ID for LM_PLANTAS, can be undefined for new items
  PlantaTitle?: string; // Display title for PLANTA
  Horarios?: { Id: number; Title: string; }[]; // Selected hours from LM_PlanificacionHoras
}

export interface ISPUser {
  Id: number;
  Title: string;
  Email: string;
}

// Interface for the item to be sent to LO_PUESTORESERVADO
export interface IReservationSPItem {
  PISOId: number;
  puestoid: number; // Added for SP item creation
  PUESTOId: number;
  USUARIOId: number;
  FECHAINICIORESERVA: string;
  FECHATERMINORESERVA: string;
  ESTADO: string;
  HORARESERVADA?: string; // Make optional as it might not be present for blocks
  FECHACOMIENZORESERVA?: string; // Make optional
  FECHALIMITECHECK?: string; // Make optional
  ASISTENTESId?:  number[] ;//{ results: string[] };
  COMENTARIOBLOQUEO?: string; // Added for blocking functionality
  USUARIORESERVAId: number; // Para GUARDAR (POST/PATCH)
  USUARIORESERVA?: ISPUser;  // Para LEER (GET con $expand)
}

export interface IReservationData {
  pisoId: number;
  usuarioId: number;
  fechaReserva: Date;
}

export interface IHotspot {
  Id: number;
  Title: string;
  COORDENADA: string;
  PISOId: number;
  CAPACIDAD: number;
}

export interface IHorarioItem {
  Id: number;
  Title: string; // Day of the week, e.g., "Lunes"
  HORAS: { ID: number; Title: string; }[]; // Lookup field, array of objects
  PISOId: number; // Lookup to LM_PISOS
  PlantaId: number; // Lookup to LM_PLANTAS
}

export interface IExistingReservation {
  HORARESERVADA: string;
  USUARIO: {
    Title: string;
    Id: number;
  };
}

export interface IAvailableRoom {
  pisoName: string;
  salaName: string;
  CAPACIDAD: number;
  availableHours: string;
  salaId: number;
  pisoId: number;
}

export interface IReportItem {
  Id: number;
  Planta: string; // This will be derived from Piso.Title
  Piso: string;
  IMAGEN: string; // Added for the image URL
  COORDENADA: string; // Added for room coordinates
  Sala: string;
  Usuario: string;
  FechaReserva: string; // FECHAINICIORESERVA
  FechaTerminoReserva: string; // FECHATERMINORESERVA
  BloqueHorario: string; // HORARESERVADA
  Estado: string;
  FechaCheckIn: string; // FECHAENTRADA
  FechaCheckOut: string; // FECHASALIDA
  KPI: string; // Empty for now
  Ver: string; // Empty for now (for the icon)
}

export interface IReservationReportItem {
  Id: number;
  FECHAINICIORESERVA: string;
  FECHATERMINORESERVA: string;
  PISO: { Title: string; ID: number; IMAGEN: string; }; // Added ID and IMAGEN
  ESTADO: string;
  HORARESERVADA: string;
  FECHASALIDA: string;
  FECHAENTRADA: string;
  PUESTO: { Title: string; ID: number; COORDENADA: string; }; // Added ID and COORDENADA
  USUARIO: { Title: string; ID: number; }; // Added ID
}

// Interface for a SharePoint Sala Item (LM_PUESTOSPISO)
export interface ISPSalaItem {
  Id?: number; // Optional for creation
  Title: string; // Room Name
  COORDENADA: string; // X,Y coordinates
  PISOId?: number; // Lookup to LM_PISOS (made optional)
  CAPACIDAD?: number; // Capacity of the room
  activo: boolean; // Active status
}

// New interface for LM_VICEPRESIDENCIAS
export interface IVicepresidenciaItem {
  Id?: number; // Optional for new items
  Title: string;
  activo: boolean;
}
export interface IPlanificacionHoraItem {
  Id?: number; // Optional for new items
  Title: string;
}

export interface IDivisionItem {
    Id?: number; // Optional for new items
    Title: string;
    activo: boolean;
  }

export interface IGerenciaItem {
  Id?: number; // Optional for new items
  Title: string;
  activo: boolean;
  VicepresidenciaId: number; // Lookup ID for LM_VICEPRESIDENCIAS
  VicepresidenciaTitle: string; // Display title for Vicepresidencia
}

export interface ISPGerenciaItem {
  Id: number;
  Title: string;
  activo: boolean;
  VICEPRESIDENCIA: {
    Id: number;
    Title: string;
  };
}

export interface IHorarioSala {
  ID: string; // Changed to string for composite ID
  IDSala: number;
  HORA: string;
  Disponibilidad: boolean;
  ReservadoPara?: string;
  statusText: string;
  reservadoPorText: string;
  statusColor: 'green' | 'red' | 'black';
}

export interface IReservaSalaScheduleProps {
  isOpen: boolean;
  pisoId: number;
  selectedPisoName: string; // New prop
  roomName: string;
  selectedDate: Date | undefined; // New prop
  usuarioId: number; // New prop
  horarios: IHorarioSala[];
  onClose: () => void;
  onSelectSlot: (horarios: IHorarioSala[], attendees: any[]) => void;
  isLoading?: boolean;
  CAPACIDAD: number;
}

export interface ISala {
  ID: number;
  Nombre: string;
  PisoID: number;
  PosicionX: number; // Position as a percentage
  PosicionY: number; // Position as a percentage
  CAPACIDAD: number;
  Disponibilidad: string; // "Disponible" or "No Disponible"
}

export interface IUploadResult {
  Id: number;
  Url: string;
}

// Raw item from LM_USUARIOS
export interface ISPUsuarioItem {
  Id: number;
  activo: boolean;
  USUARIO: {
    Id: number;
    Name: string;
    Title: string;
    EMail: string;
  };
 /* VICEPRESIDENCIA: {
    Id: number;
    Title: string;
  };*/
   division: {
    Id: number;
    Title: string;
  };
  GERENCIA: {
    Id: number;
    Title: string;
  };
  esAdmin: boolean;
}

// Item used in the component for LM_USUARIOS
export interface IUsuarioItem {
  Id?: number; // Optional for new items
  usuarioId: number;
  LoginName: string;
  text: string;
  email: string;
  //secondaryText: string;
  //VicepresidenciaId: number;
  //VicepresidenciaTitle?: string;
  GerenciaId: number;
  GerenciaTitle?: string;
  activo: boolean;
  esAdmin: boolean;
  divisionId?: number;
  divisionTitle?: string;
}

export interface Iperson { 
id: number;
text: string;
secondaryText: string;
loginName: string;
imageUrl?: string;
imageInitials: string;
}
