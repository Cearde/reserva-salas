import * as React from 'react';
import { IViewProps } from './IViewProps';
//import { IHorarioSala } from '../interfaces/IHorarioSala'; // Although not directly used for blocking, it's part of the common interfaces
import { SPService } from '../../services/sp'; // IReservationSPItem will be adapted or a new one created for blocking
import { ISala,IReservationSPItem } from '../models/entities';
import { Dropdown, IDropdownOption, DatePicker, PrimaryButton, Stack, Spinner, SpinnerSize, DayOfWeek, TextField, DefaultButton, MessageBar, MessageBarType } from '@fluentui/react';
import styles from '../Srsc.module.scss';
import { useSPFxContext } from '../../contexts/SPFxContext';
import * as strings from 'SrscWebPartStrings';
import { getRestrictedDates,getStatusColor,onFormatDate } from '../../utils/utils';



const BloqueoSala: React.FC<IViewProps> = () => {
  const context = useSPFxContext();
  const spService = React.useMemo(() => new SPService(context), [context]);

  // --- State Management (similar to ReservaSala) ---
  const [plantas, setPlantas] = React.useState<IDropdownOption[]>([]);
  const [pisos, setPisos] = React.useState<IDropdownOption[]>([]);
  const [usuarios, setUsuarios] = React.useState<IDropdownOption[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [message, setMessage] = React.useState<string | undefined>(undefined);
  const [messageType, setMessageType] = React.useState<MessageBarType>(MessageBarType.info);

  const [selectedPlanta, setSelectedPlanta] = React.useState<number | undefined>();
 // const [selectedPlanta, setSelectedPlanta] = React.useState<number>();
  const [selectedPiso, setSelectedPiso] = React.useState<number | undefined>();
  const [selectedUsuario, setSelectedUsuario] = React.useState<number | undefined>();
  const [selectedDate, setSelectedDate] = React.useState<Date | undefined>();
  const [selectedPisoImageUrl, setSelectedPisoImageUrl] = React.useState<string | undefined>();

  const [filteredPisos, setFilteredPisos] = React.useState<IDropdownOption[]>([]);


  const [salas, setSalas] = React.useState<ISala[]>([]);
  const [selectedSala, setSelectedSala] = React.useState<ISala | undefined>();

  // --- State for Blocking Modal ---
  const [showBloqueoModal, setShowBloqueoModal] = React.useState<boolean>(false);
  const [comentarioBloqueo, setComentarioBloqueo] = React.useState<string>('');
  const [isBlocking, setIsBlocking] = React.useState<boolean>(false);
  const [refreshCounter, setRefreshCounter] = React.useState<number>(0);
  
  // --- Data Fetching Effects (similar to ReservaSala) ---
  React.useEffect(() => {
    setLoading(true);
    spService.getFormDropdownOptions()
      .then(options => {
        setPlantas(options.plantas);
        setPisos(options.pisos);
        setUsuarios(options.usuarios);
      })
      .catch(error => {
        console.error("Error fetching dropdown options for BloqueoSala", error);
        setMessage(strings.ErrorLoadingFilterOptions);
        setMessageType(MessageBarType.error);
      })
      .finally(() => setLoading(false));
  }, [spService]);

   React.useEffect(() => {
      if (selectedPlanta) {
          setMessage(undefined); 
          const newFilteredPisos = pisos.filter(p => p.data.PLANTAId === selectedPlanta);
          setFilteredPisos(newFilteredPisos);
      } else {
          setFilteredPisos([]);
      }
      setSelectedPiso(0); // Reset piso selection
      setSelectedPisoImageUrl(undefined);
  }, [selectedPlanta]);

  React.useEffect(() => {
    if (selectedPiso && selectedDate && selectedPlanta) {
      //spService.getSalasByPiso(selectedPiso)
      spService.getDisponibilidadSalas(selectedPiso, selectedDate || new Date(), selectedPlanta)
        .then(setSalas)
        .catch(error => {
          console.error(`Error fetching salas for piso ${selectedPiso} in BloqueoSala:`, error);
          setSalas([]);
          setMessage(`${strings.ErrorLoadingSalasForPiso} ${selectedPiso}.`);
          setMessageType(MessageBarType.error);
        });
    } else {
      setSalas([]);
    }
  }, [selectedPiso,selectedDate,refreshCounter]);//[selectedPiso, spService,selectedDate,selectedPlanta]);

  // Define work week days for the calendar
  const workWeekDays = [
    DayOfWeek.Monday,
    DayOfWeek.Tuesday,
    DayOfWeek.Wednesday,
    DayOfWeek.Thursday,
    DayOfWeek.Friday,
  ];

  // Function to format date to dd-MM-yyyy
  

  // --- Event Handlers ---
  const handleSalaClick = (sala: ISala): void => {
    if (!selectedPiso || !selectedPlanta || !selectedDate || !selectedUsuario) {
      setMessage(strings.SelectPlantaPisoUsuarioFechaBlockWarning);
      setMessageType(MessageBarType.warning);
      return;
    }
    setMessage(undefined); // Clear previous messages

    setSelectedSala(sala);
    setShowBloqueoModal(true);
  };

  const handleConfirmBloqueo = async (): Promise<void> => {
    if (!selectedSala || !selectedPiso || !selectedUsuario || !selectedDate || !comentarioBloqueo.trim()) {
      setMessage(strings.CompleteAllFieldsBlockWarning);
      setMessageType(MessageBarType.error);
      return;
    }
    setMessage(undefined); // Clear previous messages

    setIsBlocking(true);
    try {

      const horas = await spService.getScheduleForSala(selectedSala, selectedPiso, selectedPlanta! , selectedDate);
      const reservedHours = new Set<string>();
      horas.filter(res => res.Disponibilidad === true)
        .forEach(res => {
          if (res.HORA) {
            reservedHours.add(res.HORA);
          }
        });

      const horasSeparadas = Array.from(reservedHours).join(',');

      const newItem: IReservationSPItem = {
        PISOId: selectedPiso,
        PUESTOId: selectedSala.ID,
        USUARIOId: selectedUsuario,
        FECHAINICIORESERVA: selectedDate.toISOString(),
        FECHATERMINORESERVA: selectedDate.toISOString(), // For a single day block, start and end are the same
        COMENTARIOBLOQUEO: comentarioBloqueo.trim(),
        ESTADO: "BLOQUEADO",
        USUARIORESERVAId: context.pageContext.legacyPageContext.userId,
        puestoid: selectedSala.ID,
        HORARESERVADA: horasSeparadas
         // FECHACOMIENZORESERVA, FECHALIMITECHECK, ASISTENTESId are optional for blocks
      };

      await spService.blockRoom(newItem);

      setMessage(strings.RoomBlockedSuccess);
      setMessageType(MessageBarType.success);
      setRefreshCounter(prev => prev + 1);
      setShowBloqueoModal(false);
      setComentarioBloqueo('');
      setSelectedSala(undefined); // Clear selected sala
    } catch (error) {
      console.error("Error al bloquear la sala:", error);
      setMessage(strings.RoomBlockError);
      setMessageType(MessageBarType.error);
    } finally {
      setIsBlocking(false);
    }
  };

  const handleCancelBloqueo = (): void => {
    setShowBloqueoModal(false);
    setComentarioBloqueo('');
    setSelectedSala(undefined);
    setMessage(undefined); // Clear messages when modal closes
  };
/*
  const getStatusColor = (disponibilidad: string): string => {
  switch (disponibilidad.toLowerCase()) {
    case 'full':
      return '#E74C3C';//'red';
    case 'empty':
      return '#27AE60';//'green';
    case 'partial':
      return '#F1C40F';'yellow';
    default:
      return 'grey'; // Color de respaldo por si el dato viene mal
  }
};*/

  // Get selected piso name for display in the modal
  const selectedPisoName = React.useMemo(() => {
    const pisoOption = pisos.find(piso => piso.key === selectedPiso);
    return pisoOption ? String(pisoOption.text) : 'Desconocido';
  }, [pisos, selectedPiso]);

  return (
    <div>
      <h2>{strings.BloqueoSalaTitle}</h2>
      {loading ? (
        <Spinner size={SpinnerSize.large} label={strings.LoadingConfig} />
      ) : (
        <div className={styles.viewContent}>
          {message && (
            <MessageBar
              messageBarType={messageType}
              isMultiline={false}
              onDismiss={() => setMessage(undefined)}
              dismissButtonAriaLabel={strings.CloseButton}
            >
              {message}
            </MessageBar>
          )}
          <form>
            <Stack horizontal wrap tokens={{ childrenGap: 15 }}>
              <div className={`${styles.formGroup} ${styles.formControl}`}>
                <Dropdown
                  label={strings.PlantaLabel}
                  placeholder={strings.SelectPlantaPlaceholder}
                  options={plantas}
                  onChange={(e, option) => setSelectedPlanta(option ? Number(option.key) : undefined)}
                  selectedKey={selectedPlanta}
                  required={true}
                />
              </div>
              <div className={`${styles.formGroup} ${styles.formControl}`}>
                <Dropdown
                  label={strings.PisoLabel}
                  placeholder={strings.SelectPisoPlaceholder}
                  options={filteredPisos}//{pisos}
                  onChange={(e, option) => {
                    setSelectedPiso(option ? Number(option.key) : undefined);
                    setSelectedPisoImageUrl(option?.data?.IMAGEN || undefined);
                  }}
                  selectedKey={selectedPiso}
                  required={true}
                />
              </div>
              <div className={`${styles.formGroup} ${styles.formControl}`}>
                <Dropdown
                  label={strings.UsuarioLabel}
                  placeholder={strings.SelectUsuarioPlaceholder}
                  options={usuarios}
                  onChange={(e, option) => setSelectedUsuario(option ? Number(option.key) : undefined)}
                  selectedKey={selectedUsuario}
                  required={true}
                />
              </div>
              <div className={`${styles.formGroup} ${styles.formControl}`}>
                <DatePicker
                  label={strings.FechaBloqueoLabel}
                  placeholder={strings.SelectFechaBloqueoPlaceholder}
                  minDate={new Date()}
                  maxDate={new Date(new Date().setFullYear(new Date().getFullYear() + 1))}
                  onSelectDate={(date) => setSelectedDate(date || undefined)}
                  value={selectedDate}
                  isRequired
                  allowTextInput
                  textField={{ errorMessage: "Este campo es obligatorio" }}
                  formatDate={onFormatDate}
                  calendarProps={{
                    workWeekDays: workWeekDays,
                    restrictedDates: getRestrictedDates() // Example of restricted date}
                  }}
                />
              </div>
            </Stack>
          </form>

          {selectedPisoImageUrl && (
            <div style={{ marginTop: '20px' }}>
              <h3>{strings.PlanoPisoTitle}</h3>
              <div style={{ position: 'relative', display: 'inline-block' }}>
                <img src={selectedPisoImageUrl} alt={strings.PlanoPisoTitle} style={{ maxWidth: '100%', height: 'auto', border: '1px solid #ccc' }} />
                {salas.map(sala => (
                  <PrimaryButton
                    key={sala.ID}
                    text={sala.Nombre}
                    onClick={() => handleSalaClick(sala)}
                    disabled={!selectedDate || !selectedPlanta || !selectedPiso || !selectedUsuario || sala.Disponibilidad.toLowerCase() === 'full'} // Disable if no date/planta/piso/usuario or if sala is fully booked
                    title={!selectedDate || !selectedPlanta || !selectedPiso || !selectedUsuario ? strings.SelectPlantaPisoUsuarioFechaBlockWarning : `${strings.BlockRoom} ${sala.Nombre}`}
                    style={{
                      position: 'absolute',
                      top: `${sala.PosicionY}px`,
                      left: `${sala.PosicionX}px`,
                      transform: 'translate(-50%, -50%)',
                      backgroundColor: getStatusColor(sala.Disponibilidad)
                      // getStatusColor(sala.Disponibilidad),
                    }}
                  />
                ))}
              </div>
            </div>
          )}

          {/* Blocking Modal */}
          {showBloqueoModal && selectedSala && selectedPiso && selectedUsuario && selectedDate && (
            <div style={{
              position: 'fixed',
              top: 0,
              left: 0,
              right: 0,
              bottom: 0,
              backgroundColor: 'rgba(0, 0, 0, 0.7)',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              zIndex: 1000,
            }}>
              <div style={{
                backgroundColor: '#fff',
                padding: '20px',
                borderRadius: '5px',
                minWidth: '700px',
                maxWidth: '90%',
                boxShadow: '0 5px 15px rgba(0, 0, 0, 0.3)',
                display: 'flex',
                flexDirection: 'column',
                alignItems: 'center',
              }}>
                <h3>{strings.ConfirmBloqueoSalaTitle}</h3>
                <div style={{ marginBottom: '15px', width: '100%' }}>
                  <p><strong>{strings.PisoLabel}:</strong> {selectedPisoName}</p>
                  <p><strong>{strings.SalaColumn}:</strong> {selectedSala.Nombre}</p>
                  <p><strong>{strings.FechaBloqueoLabel}:</strong> {onFormatDate(selectedDate)}</p>
                </div>
                <TextField
                  label={strings.ComentarioBloqueoLabel}
                  multiline
                  rows={6}
                  value={comentarioBloqueo}
                  onChange={(e, newValue) => setComentarioBloqueo(newValue || '')}
                  placeholder={strings.ComentarioBloqueoPlaceholder}
                  required
                  style={{ width: '500px' }}
                />
                <Stack horizontal tokens={{ childrenGap: 10 }} style={{  marginTop: '20px' }}>
                  <PrimaryButton onClick={handleConfirmBloqueo} disabled={isBlocking || !comentarioBloqueo.trim()}>
                    {isBlocking ? <Spinner size={SpinnerSize.xSmall} /> : strings.ConfirmBloqueoButton}
                  </PrimaryButton>
                  <DefaultButton onClick={handleCancelBloqueo} disabled={isBlocking}>{strings.CancelButton}</DefaultButton>
                </Stack>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default BloqueoSala;
//};

//export default BloqueoSala;

