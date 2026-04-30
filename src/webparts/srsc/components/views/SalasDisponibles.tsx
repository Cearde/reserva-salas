import * as React from 'react';
import { IViewProps } from './IViewProps';
import { SPService } from '../../services/sp';
import {IAvailableRoom,IHorarioSala,ISala} from '../models/entities';

import { Dropdown, IDropdownOption, DatePicker, Stack, Spinner, SpinnerSize, DayOfWeek, DetailsList, IColumn, IconButton, TooltipHost, MessageBar, MessageBarType } from '@fluentui/react';
import styles from '../Srsc.module.scss';
import ReservaSalaSchedule from './ReservaSalaSchedule';
import { useSPFxContext } from '../../contexts/SPFxContext';
import * as strings from 'SrscWebPartStrings';
import { getRestrictedDates } from '../../utils/utils';
import {useAuth} from '../../contexts/AuthContext';



const SalasDisponibles: React.FC<IViewProps> = (props) => {
  const context = useSPFxContext();
  const { isAdmin } = useAuth();
  const  { usuarioIDLista }  = props;
  const  { usuarioIDDivision }  = props;
  const spService = React.useMemo(() => new SPService(context), [context]);

  // --- State Management ---
  const [plantas, setPlantas] = React.useState<IDropdownOption[]>([]);
  const [pisos, setPisos] = React.useState<IDropdownOption[]>([]);
  const [usuarios, setUsuarios] = React.useState<IDropdownOption[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [message, setMessage] = React.useState<string | undefined>(undefined);
  const [messageType, setMessageType] = React.useState<MessageBarType>(MessageBarType.info);

  const [selectedPlanta, setSelectedPlanta] = React.useState<number | undefined>();
  const [selectedUsuario, setSelectedUsuario] = React.useState<number | undefined>();
  const [selectedDate, setSelectedDate] = React.useState<Date | undefined>(new Date());

  // Modal state for ReservaSalaSchedule
  const [isModalOpen, setIsModalOpen] = React.useState<boolean>(false);
  const [selectedSalaForModal, setSelectedSalaForModal] = React.useState<IAvailableRoom | undefined>(undefined);
  const [horariosSala, setHorariosSala] = React.useState<IHorarioSala[]>([]);
  const [modalIsLoading, setModalIsLoading] = React.useState<boolean>(false);

  // Accordion and DetailsList state
  const [filteredPisos, setFilteredPisos] = React.useState<IDropdownOption[]>([]);
  const [expandedPisoId, setExpandedPisoId] = React.useState<number | undefined>();
  const [pisoSchedules, setPisoSchedules] = React.useState<{ [pisoId: number]: IAvailableRoom[] }>({});
  const [pisoIsLoading, setPisoIsLoading] = React.useState<{ [pisoId: number]: boolean }>({});

  // --- Data Fetching Effects ---
  React.useEffect(() => {
    setLoading(true);
    Promise.all([
      spService.fetchActiveListItems('LM_PLANTAS'),
      spService.getPisosWithPlantaId(),
      spService.fetchActiveListItems('LM_USUARIOS')
    ]).then(([plantasData, pisosData, usuariosData]) => {
      setPlantas(plantasData);
      setPisos(pisosData);
      setUsuarios(usuariosData);
    }).catch(error => {
      console.error("Error fetching dropdown options for SalasDisponibles", error);
      setMessage(strings.ErrorLoadingFilterOptions);
      setMessageType(MessageBarType.error);
    }).finally(() => {
      setLoading(false);
    });
  }, [spService]);

  // Effect to filter pisos when planta changes
  React.useEffect(() => {
    if (selectedPlanta && pisos.length > 0) {
      const newFilteredPisos = pisos.filter(p => p.data.plantaId === selectedPlanta);
      setFilteredPisos(newFilteredPisos);
      setExpandedPisoId(undefined);
      setPisoSchedules({});
    } else {
      setFilteredPisos([]);
    }
  }, [selectedPlanta, pisos]);

  React.useEffect( () => {
      // 1. Supongamos que ya cargaste tus usuarios en dropdownOptions.usuarios
      if (usuarios.length > 0 && !selectedUsuario) {
        // 3. Verificamos si el usuario actual existe en la lista del dropdown
        const currentUserOption = usuarios.filter(u => Number(u.key) === usuarioIDLista);
  
        if (currentUserOption) {
          //setSelectedUsuario(Number(currentUserOption.key));
          if(isAdmin) {
            setUsuarios(usuarios);
            setSelectedUsuario(Number(currentUserOption[0].key));
          }else {
            setUsuarios(currentUserOption);
            setSelectedUsuario(Number(currentUserOption[0].key));
          }
        }
      }
  }, [usuarios]); 

  //usuarioIDDivision
  React.useEffect( () => {
    
      // 1. Supongamos que ya cargaste tus usuarios en dropdownOptions.usuarios
      if (plantas.length > 0 && !selectedPlanta) {
        // 3. Verificamos si el usuario actual existe en la lista del dropdown
        const currentUserOption = plantas.filter(u => Number(u.key) === usuarioIDDivision);
  
        if (currentUserOption) {
          if(isAdmin) {
            setPlantas(plantas);
            setSelectedPlanta(Number(currentUserOption[0].key));
          }else {
            setPlantas(currentUserOption);
            setSelectedPlanta(Number(currentUserOption[0].key));
          }
        }
      }
      
  }, [plantas]); 

  // Effect to fetch schedules when a piso is expanded
  React.useEffect(() => {
    if (expandedPisoId === undefined || !selectedPlanta || !selectedDate) {
      return;
    }

    const piso = filteredPisos.find(p => Number(p.key) === expandedPisoId);
    if (!piso) return;

    const fetchSchedules = async (): Promise<void> => {
      setPisoIsLoading(prev => ({ ...prev, [expandedPisoId]: true }));
      try {
        const results = await spService.getAvailableRoomsByPiso(expandedPisoId, piso.text, selectedPlanta, selectedDate);
        setPisoSchedules(prev => ({ ...prev, [expandedPisoId]: results }));
      } catch (error) {
        console.error(`Error fetching schedule for piso ${expandedPisoId}:`, error);
        setPisoSchedules(prev => ({ ...prev, [expandedPisoId]: [] }));
        setMessage(`${strings.ErrorLoadingAvailableRoomsForPiso} ${piso.text}.`);
        setMessageType(MessageBarType.error);
      } finally {
        setPisoIsLoading(prev => ({ ...prev, [expandedPisoId]: false }));
      }
    };

    void fetchSchedules();
  }, [expandedPisoId, selectedPlanta, selectedDate, spService, filteredPisos]);

  const handlePisoExpand = (piso: IDropdownOption): void => {
    const pisoId = Number(piso.key);
    const newExpandedPisoId = expandedPisoId === pisoId ? undefined : pisoId;
    setExpandedPisoId(newExpandedPisoId);
  };

  const handleReserveClick = async (item: IAvailableRoom): Promise<void> => {
    if (!selectedPlanta || !selectedDate || !selectedUsuario) {
      setMessage(strings.SelectPlantaUsuarioFechaReserveWarning);
      setMessageType(MessageBarType.warning);
      return;
    }
    setMessage(undefined); // Clear previous messages

    setSelectedSalaForModal(item);
    setIsModalOpen(true);
    setModalIsLoading(true);

    try {
      // We need an ISala object, but we have an IAvailableRoom. We can construct it.
      const sala: ISala = { ID: item.salaId, 
                            Nombre: item.salaName, 
                            PisoID: item.pisoId, 
                            PosicionX: 0, 
                            PosicionY: 0, 
                            CAPACIDAD: item.CAPACIDAD,
                            Disponibilidad: ''
                          }; // This property is not used in ReservaSalaSchedule, so we can default it};
      const mappedHorarios = await spService.getScheduleForSala(sala, item.pisoId, selectedPlanta, selectedDate);
      setHorariosSala(mappedHorarios);
    } catch (error) {
      console.error(`Error fetching schedule for sala ${item.salaId}:`, error);
      setHorariosSala([]);
      setMessage(`${strings.ErrorLoadingHorariosForSala} ${item.salaName}.`);
      setMessageType(MessageBarType.error);
    } finally {
      setModalIsLoading(false);
    }
  };

  const handleModalClose = (): void => {
    setIsModalOpen(false);
    setSelectedSalaForModal(undefined);
    setHorariosSala([]);
    setMessage(undefined); // Clear messages when modal closes
  };

  const handleConfirmReservation = async (selectedHorarios: IHorarioSala[], attendees: { id?: string; }[]): Promise<void> => {
    if (!selectedSalaForModal || !selectedDate || !selectedUsuario || selectedHorarios.length === 0) {
      setMessage(strings.MissingReservationData);
      setMessageType(MessageBarType.error);
      return;
    }
    setMessage(undefined); // Clear previous messages

    setModalIsLoading(true);

    try {
      await spService.createReservation({
        pisoId: selectedSalaForModal.pisoId,
        salaId: selectedSalaForModal.salaId,
        usuarioId: selectedUsuario,
        selectedDate: selectedDate,
        selectedHorarios: selectedHorarios,
        attendees: attendees
      });

      setMessage(strings.ReservationConfirmedSuccess);
      setMessageType(MessageBarType.success);
      handleModalClose();
      // Optionally, refresh the list for the current floor
      setExpandedPisoId(undefined); // Collapse and force re-expand to refresh
    } catch (error) {
      console.error("Error al confirmar la reserva:", error);
      setMessage(strings.ReservationConfirmError);
      setMessageType(MessageBarType.error);
    } finally {
      setModalIsLoading(false);
    }
  };

  // --- DetailsList Columns ---
  const columns: IColumn[] = [
    { key: 'column1', name: strings.PisoColumn, fieldName: 'pisoName', minWidth: 100, maxWidth: 150, isResizable: true },
    { key: 'column2', name: strings.SalaColumn, fieldName: 'salaName', minWidth: 100, maxWidth: 150, isResizable: true },
    { key: 'column3', name: strings.CapacidadColumn, fieldName: 'capacidad', minWidth: 70, maxWidth: 100, isResizable: true },
    { key: 'column4', name: strings.HorasDisponiblesColumn, fieldName: 'availableHours', minWidth: 200, isResizable: true, isMultiline: true },
    {
      key: 'column5', name: strings.AccionesColumn, minWidth: 50,
      onRender: (item: IAvailableRoom) => (
        <TooltipHost content={strings.ViewDetailAndReserveRoom}>
          <IconButton
            iconProps={{ iconName: 'Calendar' }}
            title={strings.ReserveButton}
            ariaLabel={strings.ReserveButton}
            onClick={() => handleReserveClick(item)}
            disabled={!selectedPlanta || !selectedDate || !selectedUsuario}
          />
        </TooltipHost>
      ),
    },
  ];

  // --- DatePicker Configuration ---
  const onFormatDate = (date?: Date): string => {
    if (!date) return '';
    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const year = date.getFullYear();
    return `${day}-${month}-${year}`;
  };

  const workWeekDays = [ DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday ];

  console.log(`[SalasDisponibles.tsx] Render. isModalOpen: ${isModalOpen}, selectedSalaForModal: ${selectedSalaForModal?.salaName}, horariosSala.length: ${horariosSala.length}, modalIsLoading: ${modalIsLoading}`);

  return (
    <div>
      <h2>{strings.SalasDisponiblesTitle}</h2>
      {loading ? <Spinner size={SpinnerSize.large} label={strings.LoadingFilters} /> : (
        <div className={styles.viewContent}>
          <span className={styles.camposObligatoriosTitle}>{strings.camposObligatoriosTitle}</span>
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
          {/* Filter controls */}
          <Stack horizontal wrap tokens={{ childrenGap: 15 }} style={{ marginBottom: '20px' }}>
            <div className={`${styles.formGroup} ${styles.formControl}`}>
              <Dropdown 
              label={strings.PlantaLabel} 
              placeholder={strings.SelectPlantaPlaceholder} 
              options={plantas} 
              onChange={(e, option) => setSelectedPlanta(option ? Number(option.key) : undefined)} 
              selectedKey={selectedPlanta} 
              required />
            </div>
            <div className={`${styles.formGroup} ${styles.formControl}`}>
              <Dropdown 
                label={strings.UsuarioLabel} 
                placeholder={strings.SelectUsuarioPlaceholder} 
                options={usuarios} 
                onChange={(e, option) => setSelectedUsuario(option ? Number(option.key) : undefined)} 
                selectedKey={selectedUsuario} 
                required 
              />
            </div>
            <div className={`${styles.formGroup} ${styles.formControl}`}>
              <DatePicker 
                label={strings.FechaReservaLabel} 
                placeholder={strings.SelectFechaReservaPlaceholder} 
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
                  restrictedDates: getRestrictedDates()
                }} />
              </div>
          </Stack>

          {/* Accordion */}
          <div className={styles.accordion}>
            {filteredPisos.length > 0 ? (
              filteredPisos.map(piso => (
                <div key={piso.key} className={styles.accordionItem}>
                  <button className={styles.accordionHeader} onClick={() => handlePisoExpand(piso)}>
                    <span>{piso.text}</span>
                    <IconButton iconProps={{ iconName: expandedPisoId === Number(piso.key) ? 'ChevronUp' : 'ChevronDown' }} />
                  </button>
                  {expandedPisoId === Number(piso.key) && (
                    <div className={styles.accordionContent}>
                      {pisoIsLoading[Number(piso.key)] ? <Spinner label={`${strings.LoadingSalasFor} ${piso.text}...`} /> : (
                        pisoSchedules[Number(piso.key)] && pisoSchedules[Number(piso.key)].length > 0 ? (
                          <DetailsList items={pisoSchedules[Number(piso.key)]} columns={columns} selectionMode={0} layoutMode={1} isHeaderVisible={true} />
                        ) : <p>{strings.NoAvailableRoomsForDate}</p>
                      )}
                    </div>
                  )}
                </div>
              ))
            ) : <p>{selectedPlanta ? strings.NoPisosConfiguredForPlanta : strings.SelectPlantaToViewPisos}</p>}
          </div>
        </div>
      )}

      {isModalOpen && selectedSalaForModal && selectedDate && selectedUsuario && (
        <ReservaSalaSchedule
          isOpen={isModalOpen}
          pisoId={selectedSalaForModal.pisoId}
          selectedPisoName={selectedSalaForModal.pisoName}
          CAPACIDAD={selectedSalaForModal.CAPACIDAD}
          roomName={selectedSalaForModal.salaName}
          selectedDate={selectedDate}
          usuarioId={selectedUsuario}
          horarios={horariosSala}
          onClose={handleModalClose}
          onSelectSlot={handleConfirmReservation}
          isLoading={modalIsLoading}
        />
      )}
    </div>
  );
};

export default SalasDisponibles;