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
  TooltipHost
} from '@fluentui/react';
import { useSPFxContext } from '../../contexts/SPFxContext';
import { SPService } from '../../services/sp';
import { IVicepresidenciaItem } from '../models/entities';
import * as strings from 'SrscWebPartStrings';
//import { set } from 'date-fns';

const MantenedorVicepresidencias: React.FC<IViewProps> = () => {
  const spfxContext = useSPFxContext();
  const spService = React.useMemo(() => new SPService(spfxContext), [spfxContext]);

  const [vicepresidencias, setVicepresidencias] = useState<IVicepresidenciaItem[]>([]);
  const [isModalOpen, setIsModalOpen] = useState<boolean>(false);
  const [currentVicepresidencia, setCurrentVicepresidencia] = useState<IVicepresidenciaItem | undefined>(undefined);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  //const [error, setError] = useState<string | undefined>(undefined);
  const [message, setMessage] = React.useState<{ type: MessageBarType, text: string } | undefined>(undefined);// useState<string | undefined>(undefined);
  //const [messageType, setMessageType] = useState<MessageBarType>(MessageBarType.info);
  //const [message, setMessage] =  useState<string | undefined>(undefined);
 
  const [showDeleteConfirm, setShowDeleteConfirm] = React.useState<boolean>(false);
  const [vicepresidenciaToDelete, setVicepresidenciaToDelete] = React.useState<IVicepresidenciaItem | undefined>(undefined);

  // Form validation state
  const [titleError, setTitleError] = useState<string | undefined>(undefined);

  const fetchVicepresidencias = useCallback(async () => {
    setIsLoading(true);
    setTitleError(undefined);
    //setMessage(undefined);
    try {
      const fetchedVicepresidencias = await spService.getVicepresidencias(true); // Fetch all, including inactive
      setVicepresidencias(fetchedVicepresidencias);
    } catch (err) {
      //setError( strings.ErrorFetchingVicepresidencias );
      setMessage({ type: MessageBarType.error, text: strings.ErrorFetchingVicepresidencias });
      console.error(err);
    } finally {
      setIsLoading(false);
    }
  }, [spService]);

  useEffect(() => {
    void fetchVicepresidencias();
  }, [fetchVicepresidencias]);

  const validateForm = (): boolean => {
    let isValid = true;
    if (!currentVicepresidencia?.Title || currentVicepresidencia.Title.trim() === '') {
      setTitleError(strings.FormErrorsWarning);
      isValid = false;
    } else {
      setTitleError(undefined);
    }
    return isValid;
  };

  const onAddVicepresidencia = () => {
    setCurrentVicepresidencia({
      Title: '',
      activo: true,
    } as IVicepresidenciaItem);
    setIsModalOpen(true);
    //setMessage(undefined);
    //setError(undefined);
  };

  const onEditVicepresidencia = (item: IVicepresidenciaItem) => {
    setCurrentVicepresidencia({ ...item }); // Create a copy to edit
    setIsModalOpen(true);
    //setMessage(undefined);
    //setTitleError(undefined);
  };

  const handleDelete = (item: IVicepresidenciaItem) => {
          setVicepresidenciaToDelete(item);
          setShowDeleteConfirm(true);
    };

    const confirmDelete = async () => {
        if (vicepresidenciaToDelete?.Id) {
            try {
                await spService.deleteVicepresidencia(vicepresidenciaToDelete.Id);
                setMessage({ type: MessageBarType.success, text: strings.VicepresidenciaDeletedSuccess });
                fetchVicepresidencias();
            } catch (err) {
                //setError(strings.ErrorDeletingVicepresidencia + " " + err.message);
                const msg = err instanceof Error ? err.message : String(err);
                setMessage({ type: MessageBarType.error, text: strings.ErrorDeletingVicepresidencia + ": " + msg });
                console.error("Error eliminando vicepresidencia:", err);
            } finally {
                setShowDeleteConfirm(false);
                setVicepresidenciaToDelete(undefined);
            }
        } else {
            //setMessage( strings.CannotDeleteWithoutId );
            setMessage({ type: MessageBarType.error, text: strings.CannotDeleteWithoutId });
            setShowDeleteConfirm(false);
            setVicepresidenciaToDelete(undefined);
        }
    };

    
/*
  const onDeleteVicepresidencia = async (id: number, title: string): Promise<void> => {
    if (window.confirm(strings.ConfirmDeleteVicepresidencia.replace('{0}', title))) {
      setIsLoading(true);
      setError(undefined);
      setMessage(undefined);
      try {
        await spService.deleteVicepresidencia(id);
        setMessage(strings.VicepresidenciaDeletedSuccess);
        setMessageType(MessageBarType.success);
        await fetchVicepresidencias();
      } catch (err) {
        setError(strings.ErrorDeletingVicepresidencia);
        setMessageType(MessageBarType.error);
        console.error(err);
      } finally {
        setIsLoading(false);
      }
    }
  };*/

  const onSaveVicepresidencia = async () => {
    if (!validateForm() || !currentVicepresidencia) { 
        setTitleError(strings.FormErrorsWarning );
      return;
    }

    setIsLoading(true); 
    setMessage(undefined);

    try {
      if (currentVicepresidencia.Id) {
        // Update existing vicepresidencia
        await spService.updateVicepresidencia(currentVicepresidencia);
        //setMessage(strings.VicepresidenciaUpdatedSuccess ); 
        setMessage({ type: MessageBarType.success, text: strings.VicepresidenciaUpdatedSuccess });
      } else {
        // Create new vicepresidencia
        await spService.createVicepresidencia(currentVicepresidencia);
        //setMessage(strings.VicepresidenciaAddedSuccess ); 
        setMessage({ type: MessageBarType.success, text: strings.VicepresidenciaAddedSuccess });
      }
      setIsModalOpen(false);
      setCurrentVicepresidencia(undefined);
      await fetchVicepresidencias(); // Refresh the list
    } catch (err) {
      //setError(currentVicepresidencia.Id ? strings.ErrorUpdatingVicepresidencia : strings.ErrorAddingVicepresidencia ); 
      setMessage({ type: MessageBarType.error, text: currentVicepresidencia.Id ? strings.ErrorUpdatingVicepresidencia : strings.ErrorAddingVicepresidencia });
      console.error(err);
    } finally {
      setIsLoading(false);
    }
  };

  const onCancel = (): void => {
    setIsModalOpen(false);
    setCurrentVicepresidencia(undefined); 
    setMessage(undefined);
    setTitleError(undefined);
  };

  const columns: IColumn[] = [
    /*{
      key: 'idColumn',
      name: 'ID',
      fieldName: 'Id',
      minWidth: 40,
      maxWidth: 60, // Keep ID small
      isResizable: true,
    },*/
    {
      key: 'titleColumn',
      name: strings.VicepresidenciaNameLabel,
      fieldName: 'Title',
      minWidth: 200, // Main column
      isResizable: true,
    },
    {
      key: 'activoColumn',
      name: strings.VicepresidenciaActiveLabel,
      fieldName: 'activo',
      minWidth: 80,
      isResizable: true,
      onRender: (item: IVicepresidenciaItem) => (item.activo ? strings.YesLabel : strings.NoLabel),
    },
    {
      key: 'actionsColumn',
      name: strings.AccionesColumn,
      minWidth: 100,
      isResizable: true,
      onRender: (item: IVicepresidenciaItem) => (
        <Stack horizontal tokens={{ childrenGap: 5 }} wrap>
          <TooltipHost content={strings.EditVicepresidenciaButton}>
            <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => onEditVicepresidencia(item)} />
          </TooltipHost>
          <TooltipHost content={strings.DeleteVicepresidenciaButton}>
            <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => {
              if (item.Id) {
                void handleDelete(item)// onDeleteVicepresidencia(item.Id, item.Title);
              } else {
                //setError(strings.CannotDeleteWithoutId);
                setMessage({ type: MessageBarType.error, text: strings.CannotDeleteWithoutId });
              }
            }} />
          </TooltipHost>
        </Stack>
      ),
    },
  ];

  return (
    <div style={{ padding: 20 }}>
      <h2>{strings.MantenedorVicepresidenciasView}</h2>

      
      {message && (
        <MessageBar 
          messageBarType={message.type} 
          isMultiline={false} 
          onDismiss={() => setMessage(undefined)} 
          dismissButtonAriaLabel="Cerrar">
          {message.text}
        </MessageBar>
      )}

      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 10 }}>
        <PrimaryButton text={strings.AddVicepresidenciaButton} onClick={onAddVicepresidencia} iconProps={{ iconName: 'Add' }} />
      </Stack>

      {isLoading ? (
        <Spinner size={SpinnerSize.large} label={strings.LoadingVicepresidencias} />
      ) : (
        vicepresidencias.length > 0 ? (
          <DetailsList
            items={vicepresidencias}
            columns={columns}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
            isHeaderVisible={true}
          />
        ) : (
          <MessageBar>{strings.NoVicepresidenciasFound}</MessageBar>
        )
      )}

      <Dialog
        hidden={!isModalOpen}
        onDismiss={onCancel}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: currentVicepresidencia?.Id ? strings.EditVicepresidenciaButton : strings.AddVicepresidenciaButton,
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
            label={strings.VicepresidenciaNameLabel}
            required
            value={currentVicepresidencia?.Title || ''}
            onChange={(e, newValue) =>
              setCurrentVicepresidencia((prev: IVicepresidenciaItem | undefined) => ({ ...prev, Title: newValue || '' } as IVicepresidenciaItem))
            }
            errorMessage={strings.RequiredField}
          />
          <Toggle
            label={strings.VicepresidenciaActiveLabel}
            onText={strings.YesLabel}
            offText={strings.NoLabel}
            checked={currentVicepresidencia?.activo || false}
            onChange={(e, checked) =>
              setCurrentVicepresidencia((prev: IVicepresidenciaItem | undefined) => ({ ...prev, activo: checked || false } as IVicepresidenciaItem))
            }
          />
        </Stack>

        <DialogFooter>
          <PrimaryButton onClick={onSaveVicepresidencia} text={strings.SaveButton} disabled={isLoading} />
          <DefaultButton onClick={onCancel} text={strings.CancelButton} disabled={isLoading} />
        </DialogFooter>
      </Dialog>

      <Dialog
          hidden={!showDeleteConfirm}
          onDismiss={() => setShowDeleteConfirm(false)}
          dialogContentProps={{
              type: DialogType.normal,
              title: strings.ConfirmDeleteVicepresidencia.replace('{0}', vicepresidenciaToDelete?.Title || ''),
          }}
          modalProps={{
              isBlocking: true,
              styles: { main: { maxWidth: 450 } }
          }}
      >
          <DialogFooter>
              <PrimaryButton onClick={confirmDelete} text={strings.DeleteVicepresidenciaButton} />
              <DefaultButton onClick={onCancel} text={strings.CancelButton} />
          </DialogFooter>
      </Dialog>


    </div>
  );
};

export default MantenedorVicepresidencias;
