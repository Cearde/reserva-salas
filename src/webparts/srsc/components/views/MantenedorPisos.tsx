import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { IViewProps } from './IViewProps';
import * as strings from 'SrscWebPartStrings';
import {
  PrimaryButton,
  DefaultButton,
  Dialog,
  DialogType,
  DialogFooter,
  TextField,
  Dropdown,
  IDropdownOption,
  Toggle,
  Stack,
  MessageBar,
  MessageBarType,
  Image,
  ImageFit,
  ActionButton,
  Label,
  Spinner,
  SpinnerSize,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  Checkbox // Added Checkbox
} from '@fluentui/react';
import { useSPFxContext } from '../../contexts/SPFxContext';
import { SPService  } from '../../services/sp';
import {IPisoItem } from '../models/entities';
//import { Validate } from '@microsoft/sp-core-library';

const MantenedorPisos: React.FC<IViewProps> = () => {
  const spfxContext = useSPFxContext();
  const spService = React.useMemo(() => new SPService(spfxContext), [spfxContext]);
  const [message, setMessage] = React.useState<{ type: MessageBarType, text: string } | undefined>(undefined);
  const [pisos, setPisos] = useState<IPisoItem[]>([]);
  const [plantas, setPlantas] = useState<IDropdownOption[]>([]);
  const [isModalOpen, setIsModalOpen] = useState<boolean>(false);
  const [currentPiso, setCurrentPiso] = useState<IPisoItem | undefined>(undefined);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [titleError, setTitleError] = useState<string | undefined>(undefined);
  const [file, setFile] = useState<File | undefined>(undefined);
  const [filePreviewUrl, setFilePreviewUrl] = useState<string | undefined>(undefined);
  const [showDeleteConfirm, setShowDeleteConfirm] = React.useState<boolean>(false);
  const [planificacionHoras, setPlanificacionHoras] = useState<IDropdownOption[]>([]); // New state for planificacion horas
  const [pisoToDelete, setPisoToDelete] = React.useState<IPisoItem | undefined>(undefined);
  const [selectedHoras, setSelectedHoras] = useState<number[]>([]); // New state for selected hours
  const [selectAllHoras, setSelectAllHoras] = useState<boolean>(false); // State for "Marcar todos" checkbox
  const [formErrors, setFormErrors] = React.useState<{ Title?: string, Planta?: string, Imagen?: string, Horarios?: string }>({});

  const fetchPisosAndPlantas = useCallback(async () => {
    setIsLoading(true);
    //setError(undefined);
    try {
      const fetchedPisos = await spService.getPisos(true); // Fetch all pisos, including inactive
      setPisos(fetchedPisos);

      const fetchedPlantas = await spService.fetchActiveListItems('LM_PLANTAS');
      setPlantas(fetchedPlantas);
    } catch (err) {
      //setError(`Error al cargar datos: ${err.message}`);
      setMessage({ type: MessageBarType.error, text: `Error al cargar datos: ${err.message}` });
      console.error(err);
    } finally {
      setIsLoading(false);
    }
  }, [spService]);

  const fetchPlanificacionHoras = useCallback(async () => {
    try {
      const fetchedHoras = await spService.getPlanificacionHoras();
      setPlanificacionHoras(fetchedHoras);
    } catch (err) {
      //setError(`Error al cargar horarios de planificación: ${err.message}`);
      setMessage({ type: MessageBarType.error, text: `Error al cargar horarios de planificación: ${err.message}` });
      console.error(err);
    }
  }, [spService]);

  useEffect(() => {
    void fetchPisosAndPlantas();
    void fetchPlanificacionHoras();
  }, [fetchPisosAndPlantas, fetchPlanificacionHoras]);

  const onAddPiso = () => {
    setCurrentPiso({
      Title: '',
      activo: true,
      PlantaId: undefined, // Initialize with undefined
      IMAGEN: undefined,
      idImgPiso: undefined, // Initialize with 0 or any default value
      Horarios: [], // Initialize Horarios for new piso
    } as IPisoItem);
    setFile(undefined);
    setFilePreviewUrl(undefined);
    setSelectedHoras([]); // Clear selected hours for new piso
    setIsModalOpen(true);
  };

  const onEditPiso = async (piso: IPisoItem) => {
    setCurrentPiso({ ...piso });
    setFile(undefined);
    setFilePreviewUrl(piso.IMAGEN); // Set preview to existing image

    // Fetch existing horarios for this piso
    if (piso.Id === undefined || piso.PlantaId === undefined) { // If Id or PlantaId is undefined, we cannot fetch horarios
      setSelectedHoras([]);
      setIsModalOpen(true);
      return; // Exit early
    }

    setIsLoading(true);
    try {
      const existingHorarioEntry = await spService.getHorariosForPiso(piso.Id, piso.PlantaId);

      setSelectedHoras(existingHorarioEntry ? existingHorarioEntry.HORAS.map(h => h.ID) : []);
      console.log(`[MantenedorPisos.tsx] onEditPiso - fetched existing horarios:`, selectedHoras);
    } catch (err) {
      //setError(`Error al cargar horarios del piso: ${err.message}`);
      setMessage({ type: MessageBarType.error, text: `Error al cargar horarios del piso: ${err.message}` });
      console.error(err);
      setSelectedHoras([]); // Fallback to empty if error
    } finally {
      setIsLoading(false);
    }
    setIsModalOpen(true);
  };

  /*const onDeletePiso = async (id: number): Promise<void> => {
    if (window.confirm('¿Está seguro que desea eliminar este piso?')) {
      setIsLoading(true);
      setError(undefined);
      try {
        await spService.deletePiso(id);
        await fetchPisosAndPlantas();
      } catch (err) {
        setError(`Error al eliminar piso: ${err.message}`);
        console.error(err);
      } finally {
        setIsLoading(false);
      }
    }
  };*/

   const onCancel = (): void => {
    setIsModalOpen(false);
    setShowDeleteConfirm(false);
    setCurrentPiso(undefined);
   // setError(undefined);
    setMessage(undefined);
  };

  const confirmDelete = async () => {
        if (pisoToDelete?.Id) {
            try {
                await spService.deletePiso(pisoToDelete.Id);
                setMessage({ type: MessageBarType.success, text: strings.PisoDeletedSuccess });
                fetchPisosAndPlantas();
            } catch (err) {
                //setError(strings.ErrorDeletingPiso + " " + err.message);
                setMessage({ type: MessageBarType.error, text: strings.ErrorDeletingPiso + " " + err.message });
                console.error("Error eliminando piso:", err);
            } finally {
                setShowDeleteConfirm(false);
                setPisoToDelete(undefined);
            }
        } else {
            setMessage({ type: MessageBarType.error, text: strings.CannotDeletePisoWithoutId });
            setShowDeleteConfirm(false);
            setPisoToDelete(undefined);
        }
    };

  const onSavePiso = async () => {
    
    const errors: { Title?: string, Planta?: string, Imagen?: string, Horarios?: string } = {};
    if (!currentPiso )
    {return;}
        
    if (!currentPiso.Title) {
      errors.Title = strings.RequiredField;
      //return;
    }
    if (currentPiso.PlantaId === undefined || currentPiso.PlantaId === 0) {
      errors.Planta = strings.RequiredField;
    }
    if (selectedHoras.length === 0) {
      errors.Horarios = strings.RequiredField;
    }

    if (filePreviewUrl === undefined || filePreviewUrl === "" || currentPiso?.IMAGEN === undefined) {
      errors.Imagen = strings.RequiredField;
    }
    setFormErrors(errors);
    if (Object.keys(errors).length > 0) {
      setTitleError(strings.FormErrorsWarning);
      return;
    }

    setIsLoading(true);
    //setError(undefined);
    try {
      let imageUrl = currentPiso.IMAGEN;
      let idImgPiso = currentPiso.idImgPiso;

      if (file) {
        // Upload new image
        const folderPath = 'LM_IMAGENESPISOS'; // SharePoint library name
        const uploadedFileRelativeUrl = await spService.uploadFile(file.name, await file.arrayBuffer(), folderPath);
        imageUrl = uploadedFileRelativeUrl.Url;
        idImgPiso = uploadedFileRelativeUrl.Id;
      }

      const pisoToSave: IPisoItem = {
        ...currentPiso,
        IMAGEN: imageUrl,
        idImgPiso: idImgPiso,
      };

      let savedPiso: IPisoItem;
      if (pisoToSave.Id) {
        // Update existing piso
        savedPiso = await spService.updatePiso(pisoToSave);
      } else {
        // Create new piso
        savedPiso = await spService.createPiso(pisoToSave);
      }

      // Create or update LM_Horario entry for the saved piso
      if (savedPiso.Id && savedPiso.PlantaId) {
        await spService.createOrUpdateHorarioEntry(savedPiso.Id, savedPiso.PlantaId, selectedHoras);

      } else {
        //setError('Error: No se pudo guardar los horarios. ID de Piso o Planta no disponible.');
        setMessage({ type: MessageBarType.error, text: 'Error: No se pudo guardar los horarios. ID de Piso o Planta no disponible.' });
        setIsLoading(false);
        return;
      }
        setMessage({ type: MessageBarType.success, text: strings.PisoSavedSuccessfully });
       
      setIsModalOpen(false);
      await fetchPisosAndPlantas();
    } catch (err) {
      setMessage({ type: MessageBarType.error, text: `Error al guardar piso: ${err.message}` });
      console.error(err);
    } finally {
      setIsLoading(false);
    }
  };

  const onFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const selectedFile = event.target.files ? event.target.files[0] : undefined;
    setFile(selectedFile);
    if (selectedFile) {
      setFilePreviewUrl(URL.createObjectURL(selectedFile));
    } else {
      setFilePreviewUrl(undefined);
    }
  };
  const handleDelete = (item: IPisoItem) => {
          setPisoToDelete(item);
          setShowDeleteConfirm(true);
      };

  const columns: IColumn[] = [
    {
      key: 'idColumn',
      name: 'ID',
      fieldName: 'Id',
      minWidth: 40,
      maxWidth: 60, // Keep a small max-width for ID
      isResizable: true,
    },
    {
      key: 'titleColumn',
      name: 'Título',
      fieldName: 'Title',
      minWidth: 150, // Give more space to title
      isResizable: true,
    },
    {
      key: 'plantaColumn',
      name: 'Planta',
      fieldName: 'PlantaTitle',
      minWidth: 100,
      isResizable: true,
    },
    {
      key: 'activoColumn',
      name: 'Activo',
      fieldName: 'activo',
      minWidth: 50,
      maxWidth: 70, // Keep this small
      isResizable: true,
      onRender: (item: IPisoItem) => (item.activo ? 'Sí' : 'No'),
    },
    {
      key: 'imagenColumn',
      name: 'Imagen',
      fieldName: 'IMAGEN',
      minWidth: 80,
      isResizable: true,
      onRender: (item: IPisoItem) =>
        item.IMAGEN ? (
          <Image src={item.IMAGEN} alt="Imagen del Piso" width={50} height={50} imageFit={ImageFit.contain} />
        ) : (
          <span>Sin imagen</span>
        ),
    },
    {
      key: 'actionsColumn',
      name: 'Acciones',
      minWidth: 120, // Ensure enough space for both buttons
      isResizable: true,
      onRender: (item: IPisoItem) => (
        <Stack horizontal tokens={{ childrenGap: 5 }} wrap>
          <ActionButton iconProps={{ iconName: 'Edit' }} onClick={() => onEditPiso(item)}>            
          </ActionButton>
          <ActionButton iconProps={{ iconName: 'Delete' }} onClick={() => {
            if (item.Id) {
              void handleDelete(item) // onDeletePiso(item.Id as number);
            } else {
              //setError('No se puede eliminar un piso sin ID.');
              setMessage({ type: MessageBarType.error, text: 'No se puede eliminar un piso sin ID.' });
            }
          }}>            
          </ActionButton>
        </Stack>
      ),
    },
  ];

  return (
    <div style={{ padding: 20 }}>
      <h2>Mantenedor de Pisos</h2>

      {message && (
          <MessageBar
              messageBarType={message.type}
              isMultiline={false}
              onDismiss={() => setMessage(undefined)}
              dismissButtonAriaLabel="Cerrar"
          >
              {message.text}
          </MessageBar>
      )}
 

      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 10 }}>
        <PrimaryButton text="Agregar Piso" onClick={onAddPiso} />
      </Stack>

      {isLoading ? (
        <Spinner size={SpinnerSize.large} label="Cargando pisos..." />
      ) : (
        <DetailsList
          items={pisos}
          columns={columns}
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
        />
      )}

      <Dialog
        hidden={!isModalOpen}
        onDismiss={() => setIsModalOpen(false)}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: currentPiso?.Id ? 'Editar Piso' : 'Agregar Nuevo Piso',
        }}
        modalProps={{
          isBlocking: isLoading,
        }}
      >
        {titleError && (
            <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                {titleError}
            </MessageBar>
        )   }
        <Stack tokens={{ childrenGap: 15 }}>
          <TextField
            label="Título"
            required
            value={currentPiso?.Title || ''}
            onChange={(e, newValue) =>
              setCurrentPiso((prev: IPisoItem | undefined) => ({ ...prev, Title: newValue || '' } as IPisoItem))
            }
            errorMessage={formErrors.Title}
          />
          <Dropdown
            label="Planta"
            required
            options={plantas}
            selectedKey={currentPiso?.PlantaId || null}
            onChange={(e, option) =>
              setCurrentPiso((prev: IPisoItem | undefined) => ({ ...prev, PlantaId: option?.key as number } as IPisoItem))
            }
            placeholder="Seleccione una planta"
            errorMessage={formErrors.Planta}
          />
          <Toggle
            label="Activo"
            onText="Sí"
            offText="No"
            checked={currentPiso?.activo || false}
            onChange={(e, checked) =>
              setCurrentPiso((prev: IPisoItem | undefined) => ({ ...prev, activo: checked || false } as IPisoItem))
            }
          />
          <Stack>
            <Label>Imagen</Label>
            {formErrors.Imagen && (
              <div style={{ color: '#a80000', fontSize: '12px', fontWeight: 'normal', fontFamily: 'Segoe UI, "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif' }}>
                {formErrors.Imagen}
              </div>
            )}
            <input type="file" accept="image/*" onChange={onFileChange} />
            {filePreviewUrl && (
              <Image
                src={filePreviewUrl}
                alt="Previsualización de la imagen"
                width={100}
                height={100}
                imageFit={ImageFit.contain}
                style={{ marginTop: 10 }}
              />
            )}
            {!filePreviewUrl && currentPiso?.IMAGEN && (
              <Image
                src={currentPiso.IMAGEN}
                alt="Imagen actual del Piso"
                width={100}
                height={100}
                imageFit={ImageFit.contain}
                style={{ marginTop: 10 }}
              />
            )}
          </Stack>

          <Stack tokens={{ childrenGap: 5 }}>
            <Label required>Horarios</Label>
            {formErrors.Horarios && (
              <div style={{ color: '#a80000', fontSize: '12px', fontWeight: 'normal', fontFamily: 'Segoe UI, "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif' }}>
                {formErrors.Horarios}
              </div>
            )}
            <Checkbox
              label="Marcar todos"
              checked={selectAllHoras}
              indeterminate={selectedHoras.length > 0 && selectedHoras.length < planificacionHoras.length}
              onChange={(ev, checked) => {
                if (checked) {
                  setSelectedHoras(planificacionHoras.map(hora => Number(hora.key)));
                } else {
                  setSelectedHoras([]);
                }
                setSelectAllHoras(checked || false);
              }}
            />
            <Stack horizontal wrap tokens={{ childrenGap: 10 }}>
              {planificacionHoras.map(hora => (
                <Checkbox
                  key={hora.key}
                  label={hora.text}
                  checked={selectedHoras.includes(Number(hora.key))}
                  onChange={(ev: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked?: boolean) => {
                    const horaId = Number(hora.key);
                    setSelectedHoras(prev => {
                      const newSelected = isChecked ? [...prev, horaId] : prev.filter(id => id !== horaId);
                      setSelectAllHoras(newSelected.length === planificacionHoras.length && planificacionHoras.length > 0);
                      return newSelected;
                    });
                  }}
                />
              ))}
            </Stack>
          </Stack>
        </Stack>

        <DialogFooter>
          <PrimaryButton onClick={onSavePiso} text="Guardar" disabled={isLoading} />
          <DefaultButton onClick={() => setIsModalOpen(false)} text="Cancelar" disabled={isLoading} />
        </DialogFooter>
      </Dialog>

      <Dialog
                hidden={!showDeleteConfirm}
                onDismiss={() => setShowDeleteConfirm(false)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: strings.ConfirmDeletePiso.replace('{0}', pisoToDelete?.Title || ''),
                    //subText: strings.ConfirmDeletePiso.replace('{0}', pisoToDelete?.Title || '')
                }}
                modalProps={{
                    isBlocking: true,
                    styles: { main: { maxWidth: 450 } }
                }}
            >
                <DialogFooter>
                    <PrimaryButton onClick={confirmDelete} text={strings.DeletePisoButton} />
                    <DefaultButton onClick={onCancel} text={strings.CancelButton} />
                </DialogFooter>
            </Dialog>
    </div>
  );
};

export default MantenedorPisos;
