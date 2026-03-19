import * as React from 'react';
import { IViewProps } from './IViewProps';
import { IHorarioSala,ISala } from '../models/entities';
import { SPService } from '../../services/sp';
import { Dropdown, DatePicker, PrimaryButton, Stack, Spinner, SpinnerSize, DayOfWeek, MessageBar, MessageBarType } from '@fluentui/react';
import styles from '../Srsc.module.scss';
import ReservaSalaSchedule from './ReservaSalaSchedule';
import { useSPFxContext } from '../../contexts/SPFxContext';
import * as strings from 'SrscWebPartStrings';
import { useFormDropdownOptions, useSalasByPiso } from '../../hooks/useReservaSalaData';
import { getRestrictedDates, getStatusColor,onFormatDate  } from '../../utils/utils';

// --- State Management with useReducer ---

interface ReservaSalaState {
  loading: boolean; // Initial loading for the whole component
  isScheduleLoading: boolean;
  message?: string;
  messageType: MessageBarType;
  selectedPlanta?: number;
  selectedPiso?: number;
  selectedUsuario?: number;
  selectedDate?: Date;
  selectedPisoImageUrl?: string;
  selectedSala?: ISala;
  horariosSala: IHorarioSala[];
  showSchedule: boolean;
}

type ReservaSalaAction =
  | { type: 'SET_LOADING'; payload: boolean }
  | { type: 'SET_SCHEDULE_LOADING'; payload: boolean }
  | { type: 'SET_MESSAGE'; payload: { message?: string; type: MessageBarType } }
  | { type: 'SET_SELECTED_PLANTA'; payload?: number }
  | { type: 'SET_SELECTED_PISO'; payload?: number; imageUrl?: string, capacidad?: number }
  | { type: 'SET_SELECTED_USUARIO'; payload?: number }
  | { type: 'SET_SELECTED_DATE'; payload?: Date }
  | { type: 'SET_SELECTED_SALA'; payload?: ISala }
  | { type: 'SET_HORARIOS_SALA'; payload: IHorarioSala[] }
  | { type: 'SET_SHOW_SCHEDULE'; payload: boolean }
  | { type: 'RESET_FORM' }
  | { type: 'RESET_RESERVATION_FIELDS' };

const initialState: ReservaSalaState = {
  loading: true,
  isScheduleLoading: false,
  message: undefined,
  messageType: MessageBarType.info,
  selectedPlanta: undefined,
  selectedPiso: undefined,
  selectedUsuario: undefined,
  selectedDate: undefined,
  selectedPisoImageUrl: undefined,
  selectedSala: undefined,
  horariosSala: [],
  showSchedule: false,
};

function reservaSalaReducer(state: ReservaSalaState, action: ReservaSalaAction): ReservaSalaState {
  switch (action.type) {
    case 'SET_LOADING':
      return { ...state, loading: action.payload };
    case 'SET_SCHEDULE_LOADING':
      return { ...state, isScheduleLoading: action.payload };
    case 'SET_MESSAGE':
      return { ...state, message: action.payload.message, messageType: action.payload.type };
    case 'SET_SELECTED_PLANTA':
      return { ...state, selectedPlanta: action.payload, 
        //selectedPiso: undefined, 
        selectedPiso: 0,
        selectedPisoImageUrl: undefined, 
        selectedSala: undefined, 
        horariosSala: [], 
        showSchedule: false };
    case 'SET_SELECTED_PISO':
      return { ...state, selectedPiso: action.payload, selectedPisoImageUrl: action.imageUrl, selectedSala: undefined, horariosSala: [], showSchedule: false };
    case 'SET_SELECTED_USUARIO':
      return { ...state, selectedUsuario: action.payload };
    case 'SET_SELECTED_DATE':
      return { ...state, selectedDate: action.payload, selectedSala: undefined, horariosSala: [], showSchedule: false };
    case 'SET_SELECTED_SALA':
      return { ...state, selectedSala: action.payload };
    case 'SET_HORARIOS_SALA':
      return { ...state, horariosSala: action.payload };
    case 'SET_SHOW_SCHEDULE':
      return { ...state, showSchedule: action.payload };
    case 'RESET_FORM':
      return { ...initialState, loading: false }; // Reset to initial state, but keep loading false
    case 'RESET_RESERVATION_FIELDS':
      return {
        ...state,
        selectedSala: undefined,
        selectedDate: undefined,
        horariosSala: [],
        showSchedule: false,
        isScheduleLoading: false,
        message: undefined, // Clear messages related to the previous reservation
      };
    default:
      return state;
  }
}


const ReservaSala: React.FC<IViewProps> = () => {
  const context = useSPFxContext();
  const spService = React.useMemo(() => new SPService(context), [context]);

  const [state, dispatch] = React.useReducer(reservaSalaReducer, initialState);
  const [imageDimensions, setImageDimensions] = React.useState<{ width: number; height: number } | undefined>(undefined);

  const {
    loading,
    isScheduleLoading,
    message,
    messageType,
    selectedPlanta,
    selectedPiso,
    selectedUsuario,
    selectedDate,
    selectedPisoImageUrl,
    selectedSala,
    horariosSala,
    showSchedule,
  } = state;

  // --- Custom Hooks for Data Fetching ---
  const { data: dropdownOptions, loading: dropdownLoading, error: dropdownError } = useFormDropdownOptions(spService);
  const { data: salas, loading: salasLoading, error: salasError } = useSalasByPiso(spService, selectedPiso, selectedDate,selectedPlanta);

  // Update loading state based on hooks
  React.useEffect(() => {
    dispatch({ type: 'SET_LOADING', payload: dropdownLoading || salasLoading });
  }, [dropdownLoading, salasLoading]);

  


  // Filtramos los pisos basándonos en la planta seleccionada
  const filteredPisoOptions = React.useMemo(() => {
    if (!selectedPlanta) return [];
    
    // Asumiendo que en 'data' de la opción del piso guardas el ID de la planta
    // o que el hook useFormDropdownOptions ya trae esa relación.
    return dropdownOptions.pisos.filter(piso => piso.data?.PLANTAId === selectedPlanta);
  }, [dropdownOptions.pisos, selectedPlanta]);



  // Handle errors from hooks
  React.useEffect(() => {
    if (dropdownError) {
      dispatch({ type: 'SET_MESSAGE', payload: { message: dropdownError.message, type: dropdownError.type } });
    } else if (salasError) {
      dispatch({ type: 'SET_MESSAGE', payload: { message: salasError.message, type: salasError.type } });
    } else {
      dispatch({ type: 'SET_MESSAGE', payload: { message: undefined, type: MessageBarType.info } });
    }
  }, [dropdownError, salasError]);

  // Define work week days for the calendar
  const workWeekDays = [
    DayOfWeek.Monday,
    DayOfWeek.Tuesday,
    DayOfWeek.Wednesday,
    DayOfWeek.Thursday,
    DayOfWeek.Friday,
  ];

 

  // --- Event Handlers ---
  const handleSalaClick = async (sala: ISala): Promise<void> => {
    if (!selectedPiso || !selectedPlanta || !selectedDate) {
      dispatch({ type: 'SET_MESSAGE', payload: { message: strings.SelectPlantaPisoFechaSalaWarning, type: MessageBarType.warning } });
      return;
    }
    dispatch({ type: 'SET_MESSAGE', payload: { message: undefined, type: MessageBarType.info } }); // Clear previous messages

    dispatch({ type: 'SET_SELECTED_SALA', payload: sala });
    dispatch({ type: 'SET_SCHEDULE_LOADING', payload: true });
    dispatch({ type: 'SET_SHOW_SCHEDULE', payload: true });

    try {
      const mappedHorarios = await spService.getScheduleForSala(sala, selectedPiso, selectedPlanta, selectedDate);
      dispatch({ type: 'SET_HORARIOS_SALA', payload: mappedHorarios });
    } catch (error) {
      console.error(`Error fetching schedules for sala ${sala.ID}:`, error);
      dispatch({ type: 'SET_HORARIOS_SALA', payload: [] });
      dispatch({ type: 'SET_MESSAGE', payload: { message: `${strings.ErrorLoadingHorariosForSala} ${sala.Nombre}.`, type: MessageBarType.error } });
    } finally {
      dispatch({ type: 'SET_SCHEDULE_LOADING', payload: false });
    }
  };

  

  // Get selected piso name for display in the schedule modal
  const selectedPisoName = React.useMemo(() => {
    const pisoOption = dropdownOptions.pisos.find(piso => piso.key === selectedPiso);
    return pisoOption ? String(pisoOption.text) : 'Desconocido';
  }, [dropdownOptions.pisos, selectedPiso]);

  return (
    <div>
      <h2>{strings.ReservaSalaTitle}</h2>
      {loading ? (
        <Spinner size={SpinnerSize.large} label={strings.LoadingConfig} />
      ) : (
        <div className={styles.viewContent}>
          {message && (
            <MessageBar
              messageBarType={messageType}
              isMultiline={false}
              onDismiss={() => dispatch({ type: 'SET_MESSAGE', payload: { message: undefined, type: MessageBarType.info } })}
              dismissButtonAriaLabel={strings.CloseButton}
            >
              {message}
            </MessageBar>
          )}
          <form onSubmit={(e) => e.preventDefault()}>
            <Stack horizontal wrap tokens={{ childrenGap: 15 }}>
              <div className={`${styles.formGroup} ${styles.formControl}`}>
                <Dropdown
                  label={strings.PlantaLabel}
                  placeholder={strings.SelectPlantaPlaceholder}
                  options={dropdownOptions.plantas}
                  onChange={(e, option) => dispatch({ type: 'SET_SELECTED_PLANTA', payload: option ? Number(option.key) : undefined })}
                  selectedKey={selectedPlanta}
                  required={true}
                />
              </div>
              <div className={`${styles.formGroup} ${styles.formControl}`}>
                <Dropdown
                  label={strings.PisoLabel}
                  placeholder={strings.SelectPisoPlaceholder}
                  options={filteredPisoOptions} //{dropdownOptions.pisos}
                  onChange={(e, option) => {
                    dispatch({ type: 'SET_SELECTED_PISO', payload: 
                                                            option ? Number(option.key) : undefined, 
                                                            imageUrl: option?.data?.IMAGEN || undefined,
                                                            capacidad: option?.data?.capacidad || 0 });
                  }}
                  selectedKey={selectedPiso}
                  required={true}
                />
              </div>
              <div className={`${styles.formGroup} ${styles.formControl}`}>
                <Dropdown
                  label={strings.UsuarioLabel}
                  placeholder={strings.SelectUsuarioPlaceholder}
                  options={dropdownOptions.usuarios}
                  onChange={(e, option) => dispatch({ type: 'SET_SELECTED_USUARIO', payload: option ? Number(option.key) : undefined })}
                  selectedKey={selectedUsuario}
                  required={true}
                />
              </div>
              <div className={`${styles.formGroup} ${styles.formControl}`}>
                <DatePicker
                  label={strings.FechaReservaLabel}
                  placeholder={strings.SelectFechaReservaPlaceholder}
                  minDate={new Date()}
                  maxDate={new Date(new Date().setFullYear(new Date().getFullYear() + 1))}
                  onSelectDate={(date) => dispatch({ type: 'SET_SELECTED_DATE', payload: date || undefined })}
                  value={selectedDate}
                  isRequired
                  allowTextInput
                  formatDate={onFormatDate}
                  textField={{ errorMessage: "Este campo es obligatorio" }}
                  calendarProps={{
                    workWeekDays: workWeekDays,
                    restrictedDates: getRestrictedDates()
                  }}
                />
              </div>
            </Stack>
          </form>

          {selectedPisoImageUrl  && selectedPlanta && selectedPiso && selectedUsuario && selectedDate && (
            <div style={{ marginTop: '20px' }}>
              <h3>{strings.PlanoPisoTitle}</h3>
              <div style={{ position: 'relative', display: 'inline-block' }}>
                <img
                  src={selectedPisoImageUrl}
                  alt={strings.PlanoPisoTitle}
                  style={{ maxWidth: '100%', height: 'auto', border: '1px solid #ccc' }}
                  onLoad={(e) => {
                    const img = e.currentTarget;
                    setImageDimensions({ width: img.naturalWidth, height: img.naturalHeight });
                  }}
                />
                {salas.map(sala => (
                  <PrimaryButton
                    key={sala.ID}
                    text={sala.Nombre}
                    onClick={() => handleSalaClick(sala)}
                    disabled={!selectedDate || !selectedPlanta || !selectedPiso || !selectedUsuario}
                    title={!selectedDate || !selectedPlanta ? strings.SelectPlantaPisoFechaAvailability : `${strings.ViewHorariosFor} ${sala.Nombre}`}
                    style={{
                      position: 'absolute',
                      left: imageDimensions ? `${(sala.PosicionX / imageDimensions.width) * 100}%` : '0',
                      top: imageDimensions ? `${(sala.PosicionY / imageDimensions.height) * 100}%` : '0',
                      transform: 'translate(-50%, -50%)',
                      visibility: imageDimensions ? 'visible' : 'hidden',
                      backgroundColor: getStatusColor(sala.Disponibilidad)
                    }}
                  />
                ))}
              </div>
            </div>
          )}

          {showSchedule && selectedPiso && selectedSala && selectedUsuario && (
            <ReservaSalaSchedule
              isOpen={showSchedule}
              pisoId={selectedPiso}
              selectedPisoName={selectedPisoName}
              roomName={selectedSala.Nombre}
              selectedDate={selectedDate}
              usuarioId={selectedUsuario}
              horarios={horariosSala}
              CAPACIDAD= {selectedSala.CAPACIDAD}
              onClose={() => dispatch({ type: 'SET_SHOW_SCHEDULE', payload: false })}
              onSelectSlot={async (selectedHorarios, attendees) => {
                if (!selectedPiso || !selectedSala || !selectedDate || !selectedUsuario || selectedHorarios.length === 0) {
                  dispatch({ type: 'SET_MESSAGE', payload: { message: strings.MissingReservationData, type: MessageBarType.error } });
                  return;
                }
                dispatch({ type: 'SET_MESSAGE', payload: { message: undefined, type: MessageBarType.info } }); // Clear previous messages

                dispatch({ type: 'SET_SCHEDULE_LOADING', payload: true });

                try {
                  await spService.createReservation({
                    pisoId: selectedPiso,
                    salaId: selectedSala.ID,
                    usuarioId: selectedUsuario,
                    selectedDate: selectedDate,
                    selectedHorarios: selectedHorarios,
                    attendees: attendees
                  });

                  dispatch({ type: 'SET_MESSAGE', payload: { message: strings.ReservationConfirmedSuccess, type: MessageBarType.success } });
                  dispatch({ type: 'SET_SHOW_SCHEDULE', payload: false });
                  dispatch({ type: 'RESET_RESERVATION_FIELDS' }); // Reset only reservation-specific fields
                } catch (error) {
                  console.error("Error al confirmar la reserva:", error);
                  dispatch({ type: 'SET_MESSAGE', payload: { message: strings.ReservationConfirmError, type: MessageBarType.error } });
                } finally {
                  dispatch({ type: 'SET_SCHEDULE_LOADING', payload: false });
                }
              }}
              isLoading={isScheduleLoading}
            />
          )}
        </div>
      )}
    </div>
  );
};

export default ReservaSala;