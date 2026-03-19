import * as React from 'react';
import { IViewProps } from './IViewProps';
import { SPService } from '../../services/sp';
import {IReportItem } from '../models/entities';
import { DatePicker, PrimaryButton, Stack, Spinner, SpinnerSize, DetailsList, IColumn, SelectionMode, MessageBar, MessageBarType, Dropdown, IDropdownOption, IconButton } from '@fluentui/react';
import styles from '../Srsc.module.scss';
import { useSPFxContext } from '../../contexts/SPFxContext';
import * as strings from 'SrscWebPartStrings';
import { getRestrictedDates } from '../../utils/utils';

const Reportes: React.FC<IViewProps> = () => {
    const context = useSPFxContext();
    const spService = React.useMemo(() => new SPService(context), [context]);

    // Filter state
    const [plantas, setPlantas] = React.useState<IDropdownOption[]>([]);
    const [pisos, setPisos] = React.useState<IDropdownOption[]>([]);
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

    // Fetch filter options
    React.useEffect(() => {
        setLoadingFilters(true);
        Promise.all([
            spService.fetchActiveListItems('LM_PLANTAS'),
            spService.getPisosWithPlantaId(),
        ]).then(([plantasData, pisosData]) => {
            setPlantas(plantasData);
            setPisos(pisosData);
        }).catch(err => {
            setError(strings.ErrorLoadingFilterOptions);
            console.error(err);
        }).finally(() => {
            setLoadingFilters(false);
        });
    }, [spService]);

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


    const columns: IColumn[] = [
        { key: 'column1', name: strings.PlantaLabel, fieldName: 'Planta', minWidth: 100, maxWidth: 120, isResizable: true, isMultiline: true }, // Added isMultiline
        { key: 'column2', name: strings.PisoColumn, fieldName: 'Piso', minWidth: 100, maxWidth: 120, isResizable: true, isMultiline: true }, // Added isMultiline
        { key: 'column3', name: strings.UsuarioLabel, fieldName: 'Usuario', minWidth: 150, isResizable: true, isMultiline: true }, // Added isMultiline
        { key: 'column4', name: strings.SalaReservadaColumn, fieldName: 'Sala', minWidth: 100, maxWidth: 150, isResizable: true, isMultiline: true }, // Added isMultiline
        { key: 'column5', name: strings.FechaReservaColumn, fieldName: 'FechaReserva', minWidth: 100, maxWidth: 120, isResizable: true, isMultiline: true }, // Added isMultiline
        { key: 'column6', name: strings.BloqueHorarioColumn, fieldName: 'BloqueHorario', minWidth: 150, isResizable: true, isMultiline: true },
        { key: 'column7', name: strings.EstadoColumn, fieldName: 'Estado', minWidth: 80, maxWidth: 100, isResizable: true, isMultiline: true }, // Added isMultiline
        { key: 'column8', name: strings.FechaCheckInColumn, fieldName: 'FechaCheckIn', minWidth: 120, maxWidth: 150, isResizable: true, isMultiline: true }, // Added isMultiline
        { key: 'column9', name: strings.FechaCheckOutColumn, fieldName: 'FechaCheckOut', minWidth: 120, maxWidth: 150, isResizable: true, isMultiline: true }, // Added isMultiline
        { key: 'column10', name: strings.KPIColumn, fieldName: 'KPI', minWidth: 50, maxWidth: 80, isResizable: true, isMultiline: true }, // Added isMultiline
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
                    disabled={true} // Disabled for now as per requirement
                />
            ),
        },
    ];

    const handleSearch = async (): Promise<void> => {
        setLoading(true);
        setError(undefined);
        setReportData([]);
        setHasSearched(true);

        try {
            const rawData = await spService.getReportData({ startDate, endDate, pisoId: selectedPiso });
            
            const mappedData: IReportItem[] = rawData.map(item => {
                const pisoOption = pisos.find(p => p.key === item.PISO?.ID);
                const plantaName = pisoOption?.data?.plantaTitle || strings.NAStatus;

                return {
                    Id: item.Id,
                    Planta: plantaName,
                    Piso: item.PISO?.Title || strings.NAStatus,
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

    const exportToCSV = (): void => {
        if (reportData.length === 0) {
            return;
        }

        const csvHeader = columns.map(c => `"${c.name}"`).join(',');
        const csvRows = reportData.map(row => {
            const values = [
                `"${row.Planta}"`,
                `"${row.Piso}"`,
                `"${row.Usuario}"`,
                `"${row.Sala}"`,
                `"${row.FechaReserva}"`,
                `"${row.BloqueHorario}"`,
                `"${row.Estado}"`,
                `"${row.FechaCheckIn}"`,
                `"${row.FechaCheckOut}"`,
                `"${row.KPI}"`,
                `"${row.Ver}"` // This will be empty string for now
            ];
            return values.join(',');
        }).join('\n');

        const csvContent = `\uFEFF${csvHeader}\n${csvRows}`; // Add BOM for Excel
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        link.setAttribute('href', url);
        link.setAttribute('download', 'reporte_reservas.csv');
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };

    if (loadingFilters) {
        return <Spinner size={SpinnerSize.large} label={strings.LoadingFilters} />;
    }

    return (
        <div>
            <h2>{strings.ReportesTitle}</h2>
            <div className={styles.viewContent}>
                <Stack horizontal wrap tokens={{ childrenGap: 15 }} style={{ marginBottom: '20px' }}>
                    <div className={`${styles.formGroup} ${styles.formControl}`}>
                        <Dropdown
                            label={strings.PlantaLabel}
                            placeholder={strings.AllPlaceholder}
                            options={plantas}
                            onChange={(e, option) => setSelectedPlanta(option ? Number(option.key) : undefined)}
                            selectedKey={selectedPlanta}
                        />
                    </div>
                    <div className={`${styles.formGroup} ${styles.formControl}`}>
                        <Dropdown
                            label={strings.PisoLabel}
                            placeholder={strings.AllPlaceholder}
                            options={filteredPisos}
                            onChange={(e, option) => setSelectedPiso(option ? Number(option.key) : undefined)}
                            selectedKey={selectedPiso}
                            disabled={!selectedPlanta}
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
                    <Stack.Item align="end">
                        <PrimaryButton onClick={exportToCSV} disabled={reportData.length === 0 || loading} iconProps={{ iconName: 'Download' }}>
                            {strings.ExportToCSVButton}
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

export default Reportes;