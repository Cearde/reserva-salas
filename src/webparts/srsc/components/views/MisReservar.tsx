import * as React from 'react';
import { IViewProps } from './IViewProps';
import { SPService } from '../../services/sp';
import {IReportItem } from '../models/entities';
import {   Dialog,DefaultButton,
  DialogType,
  DialogFooter,DatePicker, PrimaryButton, Stack, Spinner, SpinnerSize, DetailsList, IColumn, SelectionMode, MessageBar, MessageBarType, Dropdown, IDropdownOption, IconButton,TooltipHost } from '@fluentui/react';
import styles from '../Srsc.module.scss';
import { useSPFxContext } from '../../contexts/SPFxContext';
import * as strings from 'SrscWebPartStrings';
import { getRestrictedDates,getReportColor } from '../../utils/utils';
import {useAuth} from '../../contexts/AuthContext';

const MisReservas: React.FC<IViewProps> = (props) => {
    const context = useSPFxContext();
    const { isAdmin } = useAuth();
    const  { usuarioIDLista }  = props;
    const  { usuarioIDDivision }  = props;
    const spService = React.useMemo(() => new SPService(context), [context]);

    // Filter state
    const [plantas, setPlantas] = React.useState<IDropdownOption[]>([]);
    const [pisos, setPisos] = React.useState<IDropdownOption[]>([]);
    const [usuarios, setUsuarios] = React.useState<IDropdownOption[]>([]);
    const [filteredPisos, setFilteredPisos] = React.useState<IDropdownOption[]>([]);
    const [selectedPlanta, setSelectedPlanta] = React.useState<number | undefined>();
    const [selectedPiso, setSelectedPiso] = React.useState<number | undefined>();
    const [startDate, setStartDate] = React.useState<Date | undefined>();
    const [endDate, setEndDate] = React.useState<Date | undefined>();

    // Component state
    const [reportData, setReportData] = React.useState<IReportItem[]>([]);
    const [loading, setLoading] = React.useState<boolean>(false);
    const [loadingFilters, setLoadingFilters] = React.useState<boolean>(true);
    const [error, setError] = React.useState<string | undefined>();
    const [hasSearched, setHasSearched] = React.useState<boolean>(false);

    const [showDeleteConfirm, setShowDeleteConfirm] = React.useState<boolean>(false);
    const [reservaToDelete, setReservaToDelete] = React.useState<IReportItem | undefined>();
    const [message, setMessage] = React.useState<{ type: MessageBarType, text: string } | undefined>(undefined);

    const [selectedUsuario, setSelectedUsuario] = React.useState<number | undefined>();

    const [reservaToShow, setReservaToShow] = React.useState<IReportItem | undefined>();
    const [showSalaReserva, setShowSalaReserva] = React.useState<boolean>(false);
    const [imageDimensions, setImageDimensions] = React.useState<{ width: number; height: number } | undefined>(undefined);

    // Fetch filter options
    React.useEffect(() => {
        setLoadingFilters(true);
        Promise.all([
            spService.fetchActiveListItems('LM_PLANTAS'),
            spService.fetchActiveListItems('LM_USUARIOS'),
            spService.getPisosWithPlantaId()
        ]).then(([plantasData, usuariosData, pisosData]) => {
            setPlantas(plantasData);
            setUsuarios(usuariosData);
            setPisos(pisosData);
        }).catch(err => {
            setError(strings.ErrorLoadingFilterOptions);
            console.error(err);
        }).finally(() => {
            setLoadingFilters(false);
        });
    }, [spService]);

    

    const handleEliminarReserva = (item: IReportItem): void => {
        setReservaToDelete(item);
        setShowDeleteConfirm(true);
    };

    const handleVerSalaReserva = (item: IReportItem): void => {
        setReservaToShow(item);
        setShowSalaReserva(true);
    };
    

    const confirmDelete = async () => {
          if (reservaToDelete?.Id) {
              try {
                  await spService.deleteReservasParaSala(reservaToDelete.Id);
                  setMessage({ type: MessageBarType.success, text: strings.ReservaEliminadaSuccess });
                  handleSearch();
                  //fetchGerenciasAndVicepresidencias();
              } catch (err) {
                  //setError(strings.ErrorDeletingGerencia + " " + err.message);
                  setMessage({ type: MessageBarType.error, text: strings.ReservaEliminadaError });
                  const msg = err instanceof Error ? err.message : String(err);
                  console.error("Error eliminando la reserva:", msg);
              } finally {
                  setShowDeleteConfirm(false);
                  setReservaToDelete(undefined);
              }
          } else {
              setMessage({ type: MessageBarType.error, text: strings.CannotDeleteGerenciaWithoutId });
              setShowDeleteConfirm(false);
              setReservaToDelete(undefined);
          }
    };
    // Filter pisos when planta changes
    React.useEffect(() => {
        if (selectedPlanta) {
            const newFilteredPisos = pisos.filter(p => p.data.plantaId === selectedPlanta);
            setFilteredPisos(newFilteredPisos);
        } else {
            setFilteredPisos([]);
        }
        setSelectedPiso(undefined); // Reset piso selection
    }, [selectedPlanta, pisos]);

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

    React.useEffect( () => {
          
            // 1. Supongamos que ya cargaste tus usuarios en dropdownOptions.usuarios
            if (usuarios.length > 0 && !selectedUsuario) {
                // 3. Verificamos si el usuario actual existe en la lista del dropdown
                const currentUserOption = usuarios.filter(u => Number(u.key) === usuarioIDLista);
        
                if (currentUserOption) {
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

    const columns: IColumn[] = [
        { key: 'column1', name: strings.PlantaLabel, fieldName: 'Planta', minWidth: 100, maxWidth: 120, isResizable: true, isMultiline: true }, // Added isMultiline
        { key: 'column2', name: strings.PisoColumn, fieldName: 'Piso', minWidth: 100, maxWidth: 120, isResizable: true, isMultiline: true }, // Added isMultiline
       // { key: 'column3', name: strings.UsuarioLabel, fieldName: 'Usuario', minWidth: 150, isResizable: true, isMultiline: true }, // Added isMultiline
        { key: 'column4', name: strings.SalaReservadaColumn, fieldName: 'Sala', minWidth: 100, maxWidth: 150, isResizable: true, isMultiline: true }, // Added isMultiline
        { key: 'column5', name: strings.FechaReservaColumn, fieldName: 'FechaReserva', minWidth: 100, maxWidth: 120, isResizable: true, isMultiline: true }, // Added isMultiline
        { key: 'column6', name: strings.BloqueHorarioColumn, fieldName: 'BloqueHorario', minWidth: 150, isResizable: true, isMultiline: true },
        { key: 'column7', name: strings.EstadoColumn, fieldName: 'Estado', minWidth: 80, maxWidth: 100, isResizable: true, isMultiline: true }, // Added isMultiline
        { key: 'column8', name: strings.FechaCheckInColumn, fieldName: 'FechaCheckIn', minWidth: 120, maxWidth: 150, isResizable: true, isMultiline: true }, // Added isMultiline
        { key: 'column9', name: strings.FechaCheckOutColumn, fieldName: 'FechaCheckOut', minWidth: 120, maxWidth: 150, isResizable: true, isMultiline: true }, // Added isMultiline
        { 
                    key: 'column10', 
                    name: strings.KPIColumn, 
                    fieldName: 'KPI', 
                    minWidth: 50, 
                    maxWidth: 50, 
                    isResizable: true, 
                    isMultiline: true,
                    onRender:(item: any) => (
                        <TooltipHost content={`Estado: ${item.Estado}`}>
                            <div style={{
                            width: '14px',
                            height: '14px',
                            borderRadius: '50%',
                            backgroundColor: getReportColor(item.Estado), // Una función auxiliar que retorne el Hex
                            //margin: '0 auto'
                            }} />
                        </TooltipHost>
                    )
            
                },
        {
            key: 'column11',
            name: strings.VerColumn,
            fieldName: 'Ver',
            minWidth: 50,
            maxWidth: 50,
            isResizable: false,
            onRender: (item: IReportItem) => (
                <IconButton
                    iconProps={{ iconName: 'View' }}
                    title={strings.ViewDetailsButton}
                    ariaLabel={strings.ViewDetailsButton}
                    // onClick={() => handleViewDetails(item)} // Implement this later if needed
                    //disabled={true} // Disabled for now as per requirement
                    onClick={() => handleVerSalaReserva(item)}
                />
            ),
        },
        {
            key: 'column12',
            name: strings.EliminarColumn,
            fieldName: 'Eliminar',
            minWidth: 50,
            maxWidth: 50,
            isResizable: false,
            onRender: (item: IReportItem) => (
                <IconButton
                    iconProps={{ iconName: 'Delete' }}
                    title={strings.EliminarColumn}
                    ariaLabel={strings.EliminarColumn}
                    onClick={() => handleEliminarReserva(item)} // Implement this later if needed 
                    hidden={item.Estado.toLowerCase() !== 'reservado'} // Only show delete button if status is "Reservado"
                />
            ),
        }
    ];

    const handleSearch = async (): Promise<void> => {
        setLoading(true);
        setError(undefined);
        setReportData([]);
        setHasSearched(true);

        try {
            const rawData = await spService.getReportData({ startDate, endDate, pisoId: selectedPiso, usuarioId: selectedUsuario });
            //const rawData = await spService.getMisReservas({ startDate, endDate, pisoId: selectedPiso, usuarioId: selectedUsuario });
            const pisos = await spService.getPisos();
            const mappedData: IReportItem[] = rawData.map(item => {
                //const pisoOption = pisos.find(p => p.key === item.PISO?.ID);
                const plantaName = pisos.find(p => p.Id === selectedPiso)?.PlantaTitle;

                return {
                    Id: item.Id,
                    Planta: plantaName ? plantaName : strings.NAStatus,
                    Piso: item.PISO?.Title || strings.NAStatus,
                    IMAGEN: item.PISO?.IMAGEN || '', 
                    COORDENADA: item.PUESTO?.COORDENADA || '',
                    Sala: item.PUESTO?.Title || strings.NAStatus,
                    Usuario: item.USUARIO?.Title || strings.NAStatus,
                    FechaReserva: item.FECHAINICIORESERVA ? new Date(item.FECHAINICIORESERVA).toLocaleDateString() : strings.NAStatus,
                    FechaTerminoReserva: item.FECHATERMINORESERVA ? new Date(item.FECHATERMINORESERVA).toLocaleDateString() : strings.NAStatus,
                    BloqueHorario: item.HORARESERVADA || strings.NAStatus,
                    Estado: item.ESTADO || strings.NAStatus,
                    FechaCheckIn: item.FECHAENTRADA ? new Date(item.FECHAENTRADA).toLocaleString() : strings.NAStatus,
                    FechaCheckOut: item.FECHASALIDA ? new Date(item.FECHASALIDA).toLocaleString() : strings.NAStatus,
                    KPI: '', // Empty for now
                    Ver: '', // Empty for now (for the icon)
                };
            });
            setReportData(mappedData);
        } catch (err) {
            setError(strings.ErrorLoadingReport);
            console.error(err);
        } finally {
            setLoading(false);
        }
    };


    if (loadingFilters) {
        return <Spinner size={SpinnerSize.large} label={strings.LoadingFilters} />;
    }

    return (
        <div>
            <h2>{strings.MisReservasTitle}</h2>
            {message && (
                <MessageBar messageBarType={message.type} isMultiline={false} onDismiss={() => setMessage(undefined)} dismissButtonAriaLabel={strings.CloseButton}>
                {message.text}
                </MessageBar>
            )}

            <div className={styles.viewContent}>
                <span className={styles.camposObligatoriosTitle}>{strings.camposObligatoriosTitle}</span>
                
                {/*******************  elimnar reserva *******************************/}
                <Dialog
                    hidden={!showDeleteConfirm}
                    onDismiss={() => setShowDeleteConfirm(false)}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: "Confimra que desa eliminar la reserva",
                    }}
                    modalProps={{
                        isBlocking: true,
                        styles: { main: { maxWidth: 450 } }
                    }}
                >
                    <DialogFooter>
                        <PrimaryButton onClick={confirmDelete} text={strings.DeleteReservationButton} />
                        <DefaultButton onClick={() => setShowDeleteConfirm(false)} text={strings.CancelButton} />
                    </DialogFooter>
                </Dialog>
                {/*******************  ver sala reserva *******************************/}
                {showSalaReserva && (
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
                        <div style={{ position: 'relative', display: 'inline-block' }}>
                            <img 
                            src={reservaToShow?.IMAGEN || 'https://via.placeholder.com/600x400?text=No+Image'} 
                            alt={strings.PlanoPisoTitle} 
                            style={{ maxWidth: '100%', height: 'auto', border: '1px solid #ccc' }}
                            onLoad={(e) => {
                                const img = e.currentTarget;
                                setImageDimensions({ width: img.naturalWidth, height: img.naturalHeight });
                            }}
                            />
                            
                            <PrimaryButton
                                key={reservaToShow?.Id}
                                text={reservaToShow?.Sala}
                                //onClick={() => handleSalaClick(sala)}
                                //disabled={!selectedDate || !selectedPlanta || !selectedPiso || !selectedUsuario || sala.Disponibilidad.toLowerCase() === 'full'} // Disable if no date/planta/piso/usuario or if sala is fully booked
                                title={reservaToShow?.Sala}
                                /*style={{
                                position: 'absolute',
                                top: `${sala.PosicionY}px`,
                                left: `${sala.PosicionX}px`,
                                transform: 'translate(-50%, -50%)',
                                backgroundColor: getStatusColor(sala.Disponibilidad)
                                // getStatusColor(sala.Disponibilidad),
                                }}
                                const [x, y] = coordenadaString.split(',').map(coord => parseInt(coord.trim(), 10));
                                */
                            style={{
                                position: 'absolute',
                                left: imageDimensions ? `${(parseInt(reservaToShow?.COORDENADA.split(',')[0] || '0') / imageDimensions.width) * 100}%` : '0',
                                top: imageDimensions ? `${(parseInt(reservaToShow?.COORDENADA.split(',')[1] || '0') / imageDimensions.height) * 100}%` : '0',
                                transform: 'translate(-50%, -50%)',
                                visibility: imageDimensions ? 'visible' : 'hidden',
                                //backgroundColor: getStatusColor(reservaToShow?.Sala?.Disponibilidad)
                                }}
                            />
                        </div>
                        <Stack horizontal tokens={{ childrenGap: 10 }} style={{  marginTop: '20px' }}>
                            <DefaultButton onClick={() => setShowSalaReserva(false)} >{strings.CancelButton}</DefaultButton>
                        </Stack>
                    </div>
                    </div>
                )}



                <Stack horizontal  tokens={{ childrenGap: 15 }} style={{ marginBottom: '20px' }}>
                    <div className={`${styles.formGroup} ${styles.formControl}`}>
                        <Dropdown
                            label={strings.PlantaLabel}
                            placeholder={strings.AllPlaceholder}
                            options={plantas}
                            onChange={(e, option) => setSelectedPlanta(option ? Number(option.key) : undefined)}
                            selectedKey={selectedPlanta}
                            required
                        />
                    </div>
                    <div className={`${styles.formGroup} ${styles.formControl}`}>
                        <Dropdown
                            label={strings.PisoLabel}
                            placeholder={strings.AllPlaceholder}
                            options={filteredPisos}
                            onChange={(e, option) => setSelectedPiso(option ? Number(option.key) : undefined)}
                            selectedKey={selectedPiso}
                            required
                            disabled={!selectedPlanta}
                        />
                    </div>
                    <div className={`${styles.formGroup} ${styles.formControl}`}>
                        <Dropdown
                            label={strings.UsuarioLabel}
                            placeholder={strings.AllPlaceholder}
                            options={usuarios}
                            required
                            onChange={(e, option) => setSelectedUsuario(option ? Number(option.key) : undefined)}
                            selectedKey={selectedUsuario} 
                        />
                    </div>
                    <div className={`${styles.formGroup} ${styles.formControl}`}>
                        <DatePicker
                            label={strings.FechaDesdeLabel}
                            placeholder={strings.SelectFechaReservaPlaceholder}
                            onSelectDate={(date) => setStartDate(date || undefined)}
                            value={startDate}
                            formatDate={(d) => d ? d.toLocaleDateString() : ''}
                            maxDate={endDate}
                            isRequired
                            allowTextInput
                            calendarProps={{ 
                                restrictedDates: getRestrictedDates()
                            }}
                            textField={{ errorMessage: "Este campo es obligatorio" }}
                        />
                    </div>
                    <div className={`${styles.formGroup} ${styles.formControl}`}>
                        <DatePicker
                            label={strings.FechaHastaLabel}
                            placeholder={strings.SelectFechaReservaPlaceholder}
                            onSelectDate={(date) => setEndDate(date || undefined)}
                            value={endDate}
                            formatDate={(d) => d ? d.toLocaleDateString() : ''}
                            minDate={startDate}
                            isRequired
                            allowTextInput
                            calendarProps={{ 
                                restrictedDates: getRestrictedDates()
                            }} 
                            textField={{ errorMessage: "Este campo es obligatorio" }}
                        />
                    </div>
                    <Stack.Item align="end">
                        <PrimaryButton onClick={handleSearch} disabled={loading}>
                            {loading ? <Spinner size={SpinnerSize.xSmall} /> : strings.SearchButton}
                        </PrimaryButton>
                    </Stack.Item>
                </Stack>

                {error && <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>{error}</MessageBar>}

                {loading ? (
                    <Spinner size={SpinnerSize.large} label={strings.LoadingReport} />
                ) : (
                    hasSearched && (
                        reportData.length > 0 ? (
                            <DetailsList
                                items={reportData}
                                columns={columns}
                                selectionMode={SelectionMode.none}
                                layoutMode={1} // 1 is for fixed layout
                                isHeaderVisible={true}
                                compact={true}
                            />
                        ) : (
                            <MessageBar>{strings.NoResultsFound}</MessageBar>
                        )
                    )
                )}
                
            </div>
            
        </div>
        
    );
};

export default MisReservas;