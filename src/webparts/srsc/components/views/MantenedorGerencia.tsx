import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { IViewProps } from './IViewProps';
import {
  PrimaryButton,
  DefaultButton,
  Dialog,
  DialogType,
  DialogFooter,
  TextField,
  Toggle,
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
import { SPService } from '../../services/sp';
import { IGerenciaItem } from '../models/entities';
import * as strings from 'SrscWebPartStrings';

const MantenedorGerencia: React.FC<IViewProps> = () => {
  const spfxContext = useSPFxContext();
  const spService = React.useMemo(() => new SPService(spfxContext), [spfxContext]);

  const [gerencias, setGerencias] = useState<IGerenciaItem[]>([]);
  const [vicepresidencias, setVicepresidencias] = useState<IDropdownOption[]>([]);
  const [isModalOpen, setIsModalOpen] = useState<boolean>(false);
  const [currentGerencia, setCurrentGerencia] = useState<IGerenciaItem | undefined>(undefined);
  const [isLoading, setIsLoading] = useState<boolean>(true);
 //const [error, setError] = useState<string | undefined>(undefined);
  //const [message, setMessage] = useState<string | undefined>(undefined);
  //const [messageType, setMessageType] = useState<MessageBarType>(MessageBarType.info);



  const [showDeleteConfirm, setShowDeleteConfirm] = React.useState<boolean>(false);
  const [gerenciaToDelete, setGerenciaToDelete] = React.useState<IGerenciaItem | undefined>(undefined);
  const [message, setMessage] = React.useState<{ type: MessageBarType, text: string } | undefined>(undefined);
  const [gerenciaErrors, setGerenciaErrors] = useState<string | undefined>(undefined);

  // Form validation state
  const [titleError, setTitleError] = useState<string | undefined>(undefined);
  const [vicepresidenciaError, setVicepresidenciaError] = useState<string | undefined>(undefined);

  const fetchGerenciasAndVicepresidencias = useCallback(async () => {
    setIsLoading(true);
    setTitleError(undefined);
    try {
      const [fetchedGerencias, fetchedVicepresidencias] = await Promise.all([
        spService.getGerencias(true), // Fetch all gerencias, including inactive
        spService.getVicepresidencias(false) // Fetch only active vicepresidencias for dropdown
      ]);
      setGerencias(fetchedGerencias);
      setVicepresidencias(fetchedVicepresidencias.map(vp => ({ key: vp.Id!, text: vp.Title })));
    } catch (err) {
      //setError(strings.ErrorFetchingGerencias);
      setMessage({ type: MessageBarType.error, text: strings.ErrorFetchingGerencias });
      console.error(err);
    } finally {
      setIsLoading(false);
    }
  }, [spService]);

  useEffect(() => {
    void fetchGerenciasAndVicepresidencias();
  }, [fetchGerenciasAndVicepresidencias]);

  const validateForm = (): boolean => {
    let isValid = true; 
  
    if (!currentGerencia?.Title || currentGerencia.Title.trim() === '') {
     
      setGerenciaErrors(strings.RequiredField);
      isValid = false;
    } else {
      setGerenciaErrors(undefined);
    }

    if (!currentGerencia?.VicepresidenciaId) { 
      setVicepresidenciaError(strings.RequiredField);
      isValid = false;
    } else {
      setVicepresidenciaError(undefined);
    }
    return isValid;
  };

  const onAddGerencia = () => {
    setCurrentGerencia({
      Title: '',
      activo: true,
      VicepresidenciaId: 0,
      VicepresidenciaTitle: '',
    } as IGerenciaItem);
    setIsModalOpen(true);
    //setError(undefined);
    setMessage(undefined);
    setTitleError(undefined);
    setVicepresidenciaError(undefined);
  };

  const onEditGerencia = (item: IGerenciaItem) => {
    setCurrentGerencia({ ...item }); // Create a copy to edit
    setIsModalOpen(true);
    //setError(undefined);
    setMessage(undefined);
    setTitleError(undefined);
    setVicepresidenciaError(undefined);
  };
/*
  const onDeleteGerencia = async (id: number, title: string): Promise<void> => {
    if (window.confirm(strings.ConfirmDeleteGerencia.replace('{0}', title))) {
      setIsLoading(true);
      setError(undefined);
      setMessage(undefined);
      try {
        await spService.deleteGerencia(id);
        setMessage({ type: MessageBarType.success, text: strings.GerenciaDeletedSuccess });
       // setMessageType(MessageBarType.success);
        await fetchGerenciasAndVicepresidencias();
      } catch (err) {
        setError(strings.ErrorDeletingGerencia);
        //setMessageType(MessageBarType.error);
        console.error(err);
      } finally {
        setIsLoading(false);
      }
    }
  };*/
  const handleDelete = (item: IGerenciaItem) => {
          setGerenciaToDelete(item);
          setShowDeleteConfirm(true);
      };
  
  const confirmDelete = async () => {
      if (gerenciaToDelete?.Id) {
          try {
              await spService.deleteGerencia(gerenciaToDelete.Id);
              setMessage({ type: MessageBarType.success, text: strings.GerenciaDeletedSuccess });
              fetchGerenciasAndVicepresidencias();
          } catch (err) {
              //setError(strings.ErrorDeletingGerencia + " " + err.message);
              setMessage({ type: MessageBarType.error, text: strings.ErrorDeletingGerencia });
              console.error("Error eliminando la gerencia:", err.message);
          } finally {
              setShowDeleteConfirm(false);
              setGerenciaToDelete(undefined);
          }
      } else {
          setMessage({ type: MessageBarType.error, text: strings.CannotDeleteGerenciaWithoutId });
          setShowDeleteConfirm(false);
          setGerenciaToDelete(undefined);
      }
  };

  const onSaveGerencia = async () => {
    if (!validateForm() || !currentGerencia) {
      //setMessageType(MessageBarType.error);
      //setMessage({ type: MessageBarType.error, text: strings.FormErrorsWarning });
      setTitleError(strings.FormErrorsWarning)
      return;
    }

    setIsLoading(true);
    //setError(undefined);
    setMessage(undefined);

    try {
      if (currentGerencia.Id) {
        // Update existing gerencia
        await spService.updateGerencia(currentGerencia);
        setMessage({ type: MessageBarType.success, text: strings.GerenciaUpdatedSuccess });
        //setMessageType(MessageBarType.success);
      } else {
        // Create new gerencia
        await spService.createGerencia(currentGerencia);
        setMessage({ type: MessageBarType.success, text: strings.GerenciaAddedSuccess });
        //setMessageType(MessageBarType.success);
      }
      setIsModalOpen(false);
      setCurrentGerencia(undefined);
      await fetchGerenciasAndVicepresidencias(); // Refresh the list
    } catch (err) {
      //setError(currentGerencia.Id ? strings.ErrorUpdatingGerencia : strings.ErrorAddingGerencia);
      setMessage({ type: MessageBarType.error, text: currentGerencia.Id ? strings.ErrorUpdatingGerencia : strings.ErrorAddingGerencia });
      //setMessageType(MessageBarType.error);
      console.error(err);
    } finally {
      setIsLoading(false);
    }
  };

  const onCancel = (): void => {
    setIsModalOpen(false);
    setShowDeleteConfirm(false);
    setCurrentGerencia(undefined);
    //setError(undefined);
    setMessage(undefined);
    setTitleError(undefined);
    setVicepresidenciaError(undefined);
  };

  const columns: IColumn[] = [
    {
      key: 'idColumn',
      name: 'ID',
      fieldName: 'Id',
      minWidth: 40,
      maxWidth: 60, // Keep ID small
      isResizable: true,
    },
    {
      key: 'titleColumn',
      name: strings.GerenciaNameLabel,
      fieldName: 'Title',
      minWidth: 150, // Main column
      isResizable: true,
    },
    {
      key: 'vicepresidenciaColumn',
      name: strings.GerenciaVicepresidenciaLabel,
      fieldName: 'VicepresidenciaTitle',
      minWidth: 150,
      isResizable: true,
      onRender: (item: IGerenciaItem) => item.VicepresidenciaTitle || strings.NAStatus,
    },
    {
      key: 'activoColumn',
      name: strings.GerenciaActiveLabel,
      fieldName: 'activo',
      minWidth: 80,
      isResizable: true,
      onRender: (item: IGerenciaItem) => (item.activo ? strings.YesLabel : strings.NoLabel),
    },
    {
      key: 'actionsColumn',
      name: strings.AccionesColumn,
      minWidth: 100,
      isResizable: true,
      onRender: (item: IGerenciaItem) => (
        <Stack horizontal tokens={{ childrenGap: 5 }} wrap>
          <TooltipHost content={strings.EditGerenciaButton}>
            <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => onEditGerencia(item)} />
          </TooltipHost>
          <TooltipHost content={strings.DeleteGerenciaButton}>
            <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => {
              if (item.Id) {
                void handleDelete(item)// onDeleteGerencia(item.Id, item.Title);
              } else {
                //setError(strings.CannotDeleteGerenciaWithoutId);
                setMessage({ type: MessageBarType.error, text: strings.CannotDeleteGerenciaWithoutId });
              }
            }} />
          </TooltipHost>
        </Stack>
      ),
    },
  ];

  return (
    <div style={{ padding: 20 }}>
      <h2>{strings.MantenedorGerenciaView}</h2>

     
      {message && (
        <MessageBar messageBarType={message.type} isMultiline={false} onDismiss={() => setMessage(undefined)} dismissButtonAriaLabel={strings.CloseButton}>
          {message.text}
        </MessageBar>
      )}

      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 10 }}>
        <PrimaryButton text={strings.AddGerenciaButton} onClick={onAddGerencia} iconProps={{ iconName: 'Add' }} />
      </Stack>

      {isLoading ? (
        <Spinner size={SpinnerSize.large} label={strings.LoadingGerencias} />
      ) : (
        gerencias.length > 0 ? (
          <DetailsList
            items={gerencias}
            columns={columns}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
            isHeaderVisible={true}
          />
        ) : (
          <MessageBar>{strings.NoGerenciasFound}</MessageBar>
        )
      )}

      <Dialog
        hidden={!isModalOpen}
        onDismiss={onCancel}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: currentGerencia?.Id ? strings.EditGerenciaButton : strings.AddGerenciaButton,
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
          <TextField
            label={strings.GerenciaNameLabel}
            required
            value={currentGerencia?.Title || ''}
            onChange={(e, newValue) =>
              setCurrentGerencia((prev: IGerenciaItem | undefined) => ({ ...prev, Title: newValue || '' } as IGerenciaItem))
            }
            errorMessage={gerenciaErrors}
          />
          <Dropdown
            label={strings.GerenciaVicepresidenciaLabel}
            required
            options={vicepresidencias}
            selectedKey={currentGerencia?.VicepresidenciaId || null}
            onChange={(e, option) =>
              setCurrentGerencia((prev: IGerenciaItem | undefined) => ({ ...prev, VicepresidenciaId: option?.key as number } as IGerenciaItem))
            }
            placeholder={strings.SelectVicepresidenciaPlaceholder}
            errorMessage={vicepresidenciaError}
          />
          <Toggle
            label={strings.GerenciaActiveLabel}
            onText={strings.YesLabel}
            offText={strings.NoLabel}
            checked={currentGerencia?.activo || false}
            onChange={(e, checked) =>
              setCurrentGerencia((prev: IGerenciaItem | undefined) => ({ ...prev, activo: checked || false } as IGerenciaItem))
            }
          />
        </Stack>

        <DialogFooter>
          <PrimaryButton onClick={onSaveGerencia} text={strings.SaveButton} disabled={isLoading} />
          <DefaultButton onClick={onCancel} text={strings.CancelButton} disabled={isLoading} />
        </DialogFooter>
      </Dialog>

    <Dialog
        hidden={!showDeleteConfirm}
        onDismiss={() => setShowDeleteConfirm(false)}
        dialogContentProps={{
            type: DialogType.normal,
            title: strings.ConfirmDeleteGerencia.replace('{0}', gerenciaToDelete?.Title || ''),
        }}
        modalProps={{
            isBlocking: true,
            styles: { main: { maxWidth: 450 } }
        }}
    >
        <DialogFooter>
            <PrimaryButton onClick={confirmDelete} text={strings.DeleteGerenciaButton} />
            <DefaultButton onClick={onCancel} text={strings.CancelButton} />
        </DialogFooter>
    </Dialog>
    </div>
  );
};

export default MantenedorGerencia;
