import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { IViewProps } from './IViewProps';
import {
  PrimaryButton,
  DefaultButton,
  Dialog,
  DialogType,
  DialogFooter,
 // TextField,
  Stack,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  IconButton,
  TooltipHost,
  Dropdown,
  IDropdownOption
} from '@fluentui/react';
import { useSPFxContext } from '../../contexts/SPFxContext';
import { SPService  } from '../../services/sp';
//import {IHorarioItem, IPlanificacionHoraItem } from '../models/entities';
import { IPlanificacionHoraItem } from '../models/entities';
import * as strings from 'SrscWebPartStrings';

// Helper function to parse HH:MM AM/PM string into hour, minute, and ampm
const parseTime = (timeString: string): { hour: number; minute: number; ampm: string } | null => {
  const match = timeString.match(/(\d{1,2}):(\d{2})\s(AM|PM)/i);
  if (!match) return null;

  let hour = parseInt(match[1], 10);
  const minute = parseInt(match[2], 10);
  const ampm = match[3].toUpperCase();

  return { hour, minute, ampm };
};

// Helper function to format hour, minute, ampm into HH:MM AM/PM string
const formatTime = (hour: number, minute: number, ampm: string): string => {
  const formattedHour = hour.toString().padStart(2, '0');
  const formattedMinute = minute.toString().padStart(2, '0');
  return `${formattedHour}:${formattedMinute} ${ampm}`;
};

const MantenedorHorarios: React.FC<IViewProps> = () => {
  const spfxContext = useSPFxContext();
  const spService = React.useMemo(() => new SPService(spfxContext), [spfxContext]);

  const [horas, setHoras] = useState<IPlanificacionHoraItem[]>([]);
  const [isModalOpen, setIsModalOpen] = useState<boolean>(false);
  const [currentHora, setCurrentHora] = useState<IPlanificacionHoraItem | undefined>(undefined);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [titleError, setTitleError] = useState<string | undefined>(undefined);
  const [message, setMessage] = React.useState<{ type: MessageBarType, text: string } | undefined>(undefined);//useState<string | undefined>(undefined);
  //const [messageType, setMessageType] = useState<MessageBarType>(MessageBarType.info);

  const [showDeleteConfirm, setShowDeleteConfirm] = React.useState<boolean>(false);
  const [planificacionHoraToDelete, setPlanificacionHoraDelete] = React.useState<IPlanificacionHoraItem | undefined>(undefined);

  // Time input states for the modal
  const [selectedHour, setSelectedHour] = useState<number | undefined>(undefined);
  const [selectedMinute, setSelectedMinute] = useState<number | undefined>(undefined);
  const [selectedAMPM, setSelectedAMPM] = useState<string | undefined>(undefined);
  //const [timeInputError, setTimeInputError] = useState<string | undefined>(undefined);

  const fetchHoras = useCallback(async () => {
    setIsLoading(true);
    setTitleError(undefined);
    try {
      const fetchedHoras = await spService.getPlanificacionHorasItems();
      // Sort fetched hours by time
      fetchedHoras.sort((a, b) => {
        const timeA = parseTime(a.Title);
        const timeB = parseTime(b.Title);
        if (!timeA || !timeB) return 0; // Handle invalid formats by not sorting them
        
        const dateA = new Date(0, 0, 0, timeA.ampm === 'PM' && timeA.hour !== 12 ? timeA.hour + 12 : timeA.hour === 12 && timeA.ampm === 'AM' ? 0 : timeA.hour, timeA.minute);
        const dateB = new Date(0, 0, 0, timeB.ampm === 'PM' && timeB.hour !== 12 ? timeB.hour + 12 : timeB.hour === 12 && timeB.ampm === 'AM' ? 0 : timeB.hour, timeB.minute);
        return dateA.getTime() - dateB.getTime();
      });
      setHoras(fetchedHoras);
    } catch (err) {
      setMessage({ type: MessageBarType.error, text: strings.ErrorFetchingHoras });
      console.error(err);
    } finally {
      setIsLoading(false);
    }
  }, [spService]);

  useEffect(() => {
    void fetchHoras();
  }, [fetchHoras]);

  const hourOptions: IDropdownOption[] = Array.from({ length: 12 }, (_, i) => ({ key: i + 1, text: (i + 1).toString().padStart(2, '0') }));
  const minuteOptions: IDropdownOption[] = [{ key: 0, text: '00' }, { key: 30, text: '30' }];//Array.from({ length: 12 }, (_, i) => ({ key: i * 5, text: (i * 5).toString().padStart(2, '0') }));
  const ampmOptions: IDropdownOption[] = [{ key: 'AM', text: strings.AMLabel }, { key: 'PM', text: strings.PMLabel }];

  const validateForm = (): boolean => {
    if (selectedHour === undefined || selectedMinute === undefined || selectedAMPM === undefined) {
      setTitleError(strings.RequiredField);
      return false;
    }
    setTitleError(undefined);
    return true;
  };

  const onAddHora = () => {
    setCurrentHora(undefined); // Clear current hora
    setSelectedHour(undefined);
    setSelectedMinute(undefined);
    setSelectedAMPM(undefined);
    setIsModalOpen(true);
    setTitleError(undefined);
    setMessage(undefined); 
  };

  const onEditHora = (item: IPlanificacionHoraItem) => {
    setCurrentHora({ ...item }); // Create a copy to edit
    const parsed = parseTime(item.Title);
    if (parsed) {
      setSelectedHour(parsed.hour);
      setSelectedMinute(parsed.minute);
      setSelectedAMPM(parsed.ampm);
    } else {
      // Handle invalid format in existing item
      setSelectedHour(undefined);
      setSelectedMinute(undefined);
      setSelectedAMPM(undefined);
      setTitleError(strings.InvalidTimeFormat);
    }
    setIsModalOpen(true);
    //setError(undefined);
    //setMessage(undefined);
  };

  const handleDelete = (item: IPlanificacionHoraItem) => {
          setPlanificacionHoraDelete(item);
          setShowDeleteConfirm(true);
  };
  const onDeleteHora = async() => {// (id: number, title: string): Promise<void> => {
    if (planificacionHoraToDelete?.Id) {
            try {
                await spService.deletePlanificacionHora(planificacionHoraToDelete.Id);
                setMessage({ type: MessageBarType.success, text: strings.HoraDeletedSuccess });
                fetchHoras();
            } catch (err) {
                //setError(strings.ErrorDeletingHora + " " + err.message);
                setMessage({ type: MessageBarType.error, text: strings.ErrorDeletingHora });
                console.error("Error Eliminando horario:", err);
            } finally {
                setShowDeleteConfirm(false);
                setPlanificacionHoraDelete(undefined);
            }
        } else {
            setMessage({ type: MessageBarType.error, text: strings.CannotDeleteHoraWithoutId });
            setShowDeleteConfirm(false);
            setPlanificacionHoraDelete(undefined);
        }
  };

  const onSaveHora = async () => {
    if (!validateForm()) {
      //setMessageType(MessageBarType.error);
      //setMessage({ type: MessageBarType.warning, text: strings.FormErrorsWarning });
      setTitleError(strings.FormErrorsWarning);
      return;
    }

    setIsLoading(true);
    setTitleError(undefined);
    //setMessage(undefined);

    const formattedTime = formatTime(selectedHour!, selectedMinute!, selectedAMPM!);
    const horaToSave: IPlanificacionHoraItem = {
      ...currentHora,
      Title: formattedTime,
    };

    try {
      if (horaToSave.Id) {
        // Update existing hora
        await spService.updatePlanificacionHora(horaToSave);
        setMessage({ type: MessageBarType.success, text: strings.HoraUpdatedSuccess });
        //setMessageType(MessageBarType.success);
      } else {
        // Create new hora
        await spService.createPlanificacionHora(horaToSave);
        setMessage({ type: MessageBarType.success, text: strings.HoraAddedSuccess });
        //setMessageType(MessageBarType.success);
      }
      setIsModalOpen(false);
      setCurrentHora(undefined);
      await fetchHoras(); // Refresh the list
    } catch (err) {
      //setError(horaToSave.Id ? strings.ErrorUpdatingHora : strings.ErrorAddingHora);
      setMessage({ type: MessageBarType.error, text: horaToSave.Id ? strings.ErrorUpdatingHora : strings.ErrorAddingHora });
      //setMessageType(MessageBarType.error);
      console.error(err);
    } finally {
      setIsLoading(false);
    }
  };

  const onCancel = (): void => {
    setIsModalOpen(false);
    setShowDeleteConfirm(false);
    setCurrentHora(undefined);
    setSelectedHour(undefined);
    setSelectedMinute(undefined);
    setSelectedAMPM(undefined);
    setTitleError(undefined);
    setMessage(undefined); 
  };

  const columns: IColumn[] = [
   /* {
      key: 'idColumn',
      name: 'ID',
      fieldName: 'Id',
      minWidth: 40,
      maxWidth: 60, // Keep ID small
      isResizable: true,
    },*/
    {
      key: 'titleColumn',
      name: strings.HoraNameLabel,
      fieldName: 'Title',
      minWidth: 150, // Main column
      isResizable: true,
    },
    {
      key: 'actionsColumn',
      name: strings.AccionesColumn,
      minWidth: 100,
      isResizable: true,
      onRender: (item: IPlanificacionHoraItem) => (
        <Stack horizontal tokens={{ childrenGap: 5 }} wrap>
          <TooltipHost content={strings.EditHoraButton}>
            <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => onEditHora(item)} />
          </TooltipHost>
          <TooltipHost content={strings.DeleteHoraButton}>
            <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => {
              if (item.Id) {
                void handleDelete(item)// onDeleteHora(item.Id, item.Title);
              } else {
                //setError(strings.CannotDeleteHoraWithoutId);
                setMessage({ type: MessageBarType.error, text: strings.CannotDeleteHoraWithoutId });
              }
            }} />
          </TooltipHost>
        </Stack>
      ),
    },
  ];

  return (
    <div style={{ padding: 20 }}>
      <h2>{strings.MantenedorHorariosTitle}</h2>

      
      {message && (
        <MessageBar messageBarType={message.type} isMultiline={false} onDismiss={() => setMessage(undefined)} dismissButtonAriaLabel={strings.CloseButton}>
          {message.text}
        </MessageBar>
      )}

      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 10 }}>
        <PrimaryButton text={strings.AddHoraButton} onClick={onAddHora} iconProps={{ iconName: 'Add' }} />
      </Stack>

      {isLoading ? (
        <Spinner size={SpinnerSize.large} label={strings.LoadingHoras} />
      ) : (
        horas.length > 0 ? (
          <DetailsList
            items={horas}
            columns={columns}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
            isHeaderVisible={true}
          />
        ) : (
          <MessageBar>{strings.NoHorasFound}</MessageBar>
        )
      )}

      <Dialog
        hidden={!isModalOpen}
        onDismiss={onCancel}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: currentHora?.Id ? strings.EditHoraButton : strings.AddHoraButton,
        }}
        modalProps={{
          isBlocking: isLoading,
        }}
      >
        {titleError && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false} dismissButtonAriaLabel={strings.CloseButton}>
          {titleError}
        </MessageBar>
      )}
        <Stack tokens={{ childrenGap: 15 }}>
          <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="end">
            <Dropdown
              label={strings.HourLabel}
              options={hourOptions}
              selectedKey={selectedHour}
              onChange={(e, option) => setSelectedHour(option ? Number(option.key) : undefined)}
              placeholder={strings.SelectHourPlaceholder}
              required
              errorMessage={titleError && selectedHour === undefined ? strings.RequiredField : undefined}
              styles={{ root: { width: 100 } }}
            />
            <Dropdown
              label={strings.MinuteLabel}
              options={minuteOptions}
              selectedKey={selectedMinute}
              onChange={(e, option) => setSelectedMinute(option ? Number(option.key) : undefined)}
              placeholder={strings.SelectMinutePlaceholder}
              required
              errorMessage={titleError && selectedMinute === undefined ? strings.RequiredField : undefined}
              styles={{ root: { width: 100 } }}
            />
            <Dropdown
              label="" // Label is handled by the overall time input
              options={ampmOptions}
              selectedKey={selectedAMPM}
              onChange={(e, option) => setSelectedAMPM(option ? String(option.key) : undefined)}
              placeholder={strings.SelectAMPMPlaceholder}
              required
              errorMessage={titleError && selectedAMPM === undefined ? strings.RequiredField : undefined}
              styles={{ root: { width: 100 } }}
            />
          </Stack>
        </Stack>

        <DialogFooter>
          <PrimaryButton onClick={onSaveHora} text={strings.SaveButton} disabled={isLoading} />
          <DefaultButton onClick={onCancel} text={strings.CancelButton} disabled={isLoading} />
        </DialogFooter>
      </Dialog>

      <Dialog
            hidden={!showDeleteConfirm}
            onDismiss={() => setShowDeleteConfirm(false)}
            dialogContentProps={{
                type:  DialogType.normal,
                title:  "¿Está seguro que desea eliminar el Horario " + planificacionHoraToDelete?.Title,
            }}
            modalProps={{
                isBlocking: true,
                styles: { main: { maxWidth: 450 } }
            }}
        >
            <DialogFooter>
                <PrimaryButton onClick={onDeleteHora} text={strings.DeleteHoraButton} />
                <DefaultButton onClick={onCancel} text={strings.CancelButton} />
            </DialogFooter>
        </Dialog>
    </div>
  );
};

export default MantenedorHorarios;
