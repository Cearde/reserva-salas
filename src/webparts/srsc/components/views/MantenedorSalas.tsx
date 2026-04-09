import * as React from 'react';
import { IViewProps } from './IViewProps';
import { SPService  } from '../../services/sp';
import {ISPSalaItem } from '../models/entities';
import { useSPFxContext } from '../../contexts/SPFxContext';
import * as strings from 'SrscWebPartStrings';
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  PrimaryButton,
  DefaultButton,
  TextField,
  Dropdown,
  IDropdownOption,
  //Toggle,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Dialog,
  DialogType,
  DialogFooter,
  Stack,
  IconButton,
  TooltipHost
} from '@fluentui/react';
import styles from '../Srsc.module.scss';
//import { Validate } from '@microsoft/sp-core-library';
//import { useSalasByPiso } from '../../hooks/useReservaSalaData';

const MantenedorSalas: React.FC<IViewProps> = () => {
  const context = useSPFxContext();
  const spService = React.useMemo(() => new SPService(context), [context]);
  //const { data: salas, loading: salasLoading, error: salasError } = useSalasByPiso(spService, selectedPiso);

  const [salas, setSalas] = React.useState<ISPSalaItem[]>([]);
  const [pisos, setPisos] = React.useState<IDropdownOption[]>([]);
  const [plantas, setPlantas] = React.useState<IDropdownOption[]>([]); // New state for plantas
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [message, setMessage] =  React.useState<{ type: MessageBarType, text: string } | undefined>(undefined);//React.useState<string | undefined>(undefined);
 // const [messageType, setMessageType] = React.useState<MessageBarType>(MessageBarType.info);

  const [isModalOpen, setIsModalOpen] = React.useState<boolean>(false);
  const [currentSala, setCurrentSala] = React.useState<ISPSalaItem | undefined>(undefined);
  const [filterPisoId, setFilterPisoId]: [number | undefined, React.Dispatch<React.SetStateAction<number | undefined>>] = React.useState<number | undefined>(undefined);


  const [showDeleteConfirm, setShowDeleteConfirm] = React.useState<boolean>(false);
  const [salaToDelete, setSalaToDelete] = React.useState<ISPSalaItem | undefined>(undefined);
  // States for modal's image selection
  const [selectedPlantaInModal, setSelectedPlantaInModal] = React.useState<number | undefined>(undefined);
  const [selectedPisoInModal, setSelectedPisoInModal] = React.useState<number | undefined>(undefined);
  const [floorImageUrl, setFloorImageUrl] = React.useState<string | undefined>(undefined);
  const [imageDimensions, setImageDimensions] = React.useState<{ width: number; height: number } | undefined>(undefined);
  const [isImageLoading, setIsImageLoading] = React.useState<boolean>(false); // New state for image loading - Added comment to force re-evaluation

  // Form validation state
  const [salaNameError, setSalaNameError] = React.useState<string | undefined>(undefined);
  const [salaPisoError, setSalaPisoError] = React.useState<string | undefined>(undefined);
  const [salaCoordinatesError, setSalaCoordinatesError] = React.useState<string | undefined>(undefined);
  const [salaCapacityError, setSalaCapacityError] = React.useState<string | undefined>(undefined);
  const [titleError, setTitleError] = React.useState<string | undefined>(undefined);


  ///VER QR
  const [isQRModalOpen, setIsQRModalOpen] = React.useState<boolean>(false);
  const [qrImageUrl, setQrImageUrl] = React.useState<string | undefined>(undefined);
  const [selectedSalaName, setSelectedSalaName] = React.useState<string>("");


  const fetchSalas = React.useCallback(async () => {
    setLoading(true);
    setError(undefined);
    try {
      const fetchedSalas = await spService.getSalas(filterPisoId, true); // Fetch all (active/inactive) for mantenedor
      setSalas(fetchedSalas);
    } catch (err) {
      console.error("Error fetching salas:", err);
      setError(strings.ErrorFetchingSalas);
    } finally {
      setLoading(false);
    }
  }, [spService, filterPisoId]);

  const fetchPisos = React.useCallback(async () => {
    try {
      const [fetchedPisos, fetchedPlantas] = await Promise.all([
        spService.getPisosWithPlantaId(),
        spService.fetchActiveListItems('LM_PLANTAS')
      ]);
      setPisos([{ key: 'all', text: strings.AllPlaceholder }, ...fetchedPisos]);
      setPlantas(fetchedPlantas);
    } catch (err) {
      console.error("Error fetching pisos:", err);
      setError(strings.ErrorFetchingPisos);
    }
  }, [spService]);

  React.useEffect(() => {
    void fetchPisos();
  }, [fetchPisos]);

  React.useEffect(() => {
    void fetchSalas();
  }, [fetchSalas]);

  // Effect to update floor image when selectedPisoInModal changes
  React.useEffect(() => {
    if (selectedPisoInModal) {
      const pisoOption = pisos.find(p => Number(p.key) === selectedPisoInModal);
      if (pisoOption && pisoOption.data && pisoOption.data.IMAGEN) {
        setFloorImageUrl(pisoOption.data.IMAGEN);
      } else {
        setFloorImageUrl(undefined);
      }
    } else {
      setFloorImageUrl(undefined);
    }
  }, [selectedPisoInModal, pisos]);

  const validateForm = (): boolean => {
    let isValid = true;

    if (!currentSala?.Title || currentSala.Title.trim() === '') {
      setSalaNameError(strings.RequiredField);
      isValid = false;
    } else {
      setSalaNameError(undefined);
    }



    if (!currentSala?.PISOId && !selectedPisoInModal) {
      setSalaPisoError(strings.RequiredField);
      isValid = false;
    } else {
      setSalaPisoError(undefined);
    }

    if (!currentSala?.COORDENADA || !/^\d+,\d+$/.test(currentSala.COORDENADA)) {
      setSalaCoordinatesError(strings.InvalidCoordinates);
      isValid = false;
    } else {
      setSalaCoordinatesError(undefined);
    }

    if (currentSala?.CAPACIDAD === undefined || isNaN(currentSala.CAPACIDAD) || currentSala.CAPACIDAD <= 0) {
      setSalaCapacityError(strings.InvalidNumber);
      isValid = false;
    } else {
      setSalaCapacityError(undefined);
    }


    return isValid;
  };

  const handleImageClick = (event: React.MouseEvent<HTMLImageElement>): void => {
    const img = event.currentTarget;
    const rect = img.getBoundingClientRect();
    const offsetX = event.clientX - rect.left;
    const offsetY = event.clientY - rect.top;

    const naturalWidth = img.naturalWidth;
    const naturalHeight = img.naturalHeight;
    const clientWidth = img.clientWidth;
    const clientHeight = img.clientHeight;

    // Calculate scaled coordinates
    const scaledX = Math.round((offsetX / clientWidth) * naturalWidth);
    const scaledY = Math.round((offsetY / clientHeight) * naturalHeight);

    setCurrentSala(prev => ({ ...prev, COORDENADA: `${scaledX},${scaledY}` } as ISPSalaItem));
  };

  const handleAddClick = (): void => {
    setCurrentSala({
      Title: '',
      COORDENADA: '', // Start with empty coordinates
      PISOId: undefined,
      CAPACIDAD: 0,
      activo: true,
    } as ISPSalaItem);
    setSelectedPlantaInModal(undefined);
    setSelectedPisoInModal(undefined);
    setFloorImageUrl(undefined);
    setImageDimensions(undefined);
    setIsModalOpen(true);
    setError(undefined);
    setMessage(undefined);
    setSalaNameError(undefined);
    setSalaPisoError(undefined);
    setSalaCoordinatesError(undefined);
    setSalaCapacityError(undefined);
  };

  const handleEditClick = (sala: ISPSalaItem): void => {
    setCurrentSala({ ...sala }); // Create a copy to edit
    setSelectedPlantaInModal(pisos.find(p => Number(p.key) === sala.PISOId)?.data?.plantaId);
    setSelectedPisoInModal(sala.PISOId);
    // floorImageUrl will be set by the useEffect when selectedPisoInModal changes
    setImageDimensions(undefined);
    setIsModalOpen(true);
    setError(undefined);
    setMessage(undefined);
    setSalaNameError(undefined);
    setSalaPisoError(undefined);
    setSalaCoordinatesError(undefined);
    setSalaCapacityError(undefined);
  };

  const handleVerQRClick = async (sala: ISPSalaItem): Promise<void> => {
    // Construye la URL. Ajusta 'QRSalas' al nombre real de tu biblioteca
    // y la extensión (.png, .jpg) según corresponda.
    //const siteUrl = context.pageContext.web.absoluteUrl;
    
    const imageUrl = await spService.getQR(sala.Id, sala.PISOId);
    
    //`${siteUrl}/LO_QRPUESTOS/${sala.planta}/${sala.Title}/QR_Sala_${sala.Title}.png`; 

    const fullUrl = `${window.location.origin}${imageUrl}`;

    setSelectedSalaName(sala.Title);
    setQrImageUrl(fullUrl);
    setIsQRModalOpen(true);

  };

  const descargarQR = async () => {
    if (!qrImageUrl) return;

    try {
      // Para evitar problemas de CORS al descargar, lo ideal es convertir a Blob
      const response = await fetch(qrImageUrl);
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      
      const link = document.createElement('a');
      link.href = url;
      link.download = `QR_${selectedSalaName}.png`; // Nombre del archivo
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error("Error al descargar la imagen:", error);
      // Fallback simple si el fetch falla por CORS
      window.open(qrImageUrl, '_blank');
    }
  };

  const imprimirSoloQR = () => {
    const ventanaImpresion = window.open('', '_blank');
    if (ventanaImpresion) {
      ventanaImpresion.document.write(`
        <html>
          <head>
            <title>Imprimir QR - ${selectedSalaName}</title>
            <style>
              body { display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
              img { max-width: 100%; height: auto; }
            </style>
          </head>
          <body onload="window.print();window.close()">
            <img src="${qrImageUrl}" />
          </body>
        </html>
      `);
      ventanaImpresion.document.close();
    }
  };
  /*
  const handleDeleteClick = async (salaId: number, salaName: string): Promise<void> => {
    if (window.confirm(strings.ConfirmDeleteSala.replace('{0}', salaName))) {
      setLoading(true);
      try {
        await spService.deleteSala(salaId);
        setMessage(strings.SalaDeletedSuccess);
        setMessageType(MessageBarType.success);
        void fetchSalas();
      } catch (err) {
        console.error("Error deleting sala:", err);
        setError(strings.ErrorDeletingSala);
        setMessageType(MessageBarType.error);
      } finally {
        setLoading(false);
      }
    }
  };*/

  const handleDelete = (item: ISPSalaItem) => {
          setSalaToDelete(item);
          setShowDeleteConfirm(true);
      };
  
  const confirmDelete = async () => {
      if (salaToDelete?.Id) {
          try {
              await spService.deleteSala(salaToDelete.Id);
              setMessage({ type: MessageBarType.success, text: strings.SalaDeletedSuccess });
              fetchSalas();
          } catch (err) {
              const msg = err instanceof Error ? err.message : String(err);
              setError(strings.ErrorDeletingSala + " " + msg);
              console.error("Error deleting sala:", err);
          } finally {
              setShowDeleteConfirm(false);
              setSalaToDelete(undefined);
          }
      } else {
          setMessage({ type: MessageBarType.error, text: strings.CannotDeleteSalaWithoutId });
          setShowDeleteConfirm(false);
          setSalaToDelete(undefined);
      }
  };

  const handleSave = async (): Promise<void> => {
    if (!validateForm() || !currentSala) {
      //setMessageType(MessageBarType.error);
      setTitleError(strings.FormErrorsWarning);
      //setMessage({ type: MessageBarType.error, text: "Por favor, corrija los errores del formulario." });
      return;
    }

    const salaToSave: ISPSalaItem = currentSala; // Ensure currentSala is treated as ISPSalaItem

    setLoading(true);
    setError(undefined);
    setMessage(undefined);

    try {
      if (salaToSave.Id) {
        // Update existing sala
        await spService.updateSala(salaToSave);
        setMessage({ type: MessageBarType.success, text: strings.SalaUpdatedSuccess });
        //setMessageType(MessageBarType.success);
      } else {
        // Create new sala
        await spService.createSala(salaToSave);
        setMessage({ type: MessageBarType.success, text: strings.SalaAddedSuccess });
        //setMessageType(MessageBarType.success);
      }
      setIsModalOpen(false);
      setCurrentSala(undefined);
      void fetchSalas(); // Refresh the list
    } catch (err) {
      console.error("Error saving sala:", err);
      const msg = err instanceof Error ? err.message : String(err);
      setError((currentSala?.Id ? strings.ErrorUpdatingSala : strings.ErrorAddingSala) + " " + msg);
      //setMessageType(MessageBarType.error);
    } finally {
      setLoading(false);
    }
  };

  const handleCancel = (): void => {
    setIsModalOpen(false);
    setShowDeleteConfirm(false);
    setCurrentSala(undefined);
    setError(undefined);
    setMessage(undefined);
    setSalaNameError(undefined);
    setSalaPisoError(undefined);
    setSalaCoordinatesError(undefined);
    setSalaCapacityError(undefined);
  };

  const validarCoordenadas = (salasExistentes: ISPSalaItem[]): boolean => {
    if (!currentSala?.COORDENADA) return false;

    const [currentX, currentY] = currentSala.COORDENADA.split(',').map(num => parseInt(num, 10));

    // Definimos el margen mínimo para evitar superposición. 
    // Si tu marcador mide 80x30, un margen de 60-80px en X y 25-30px en Y es ideal.
    const MARGEN_X = 85; 
    const MARGEN_Y = 35;

    // Verificamos si alguna sala existente está "demasiado cerca"
    const haySuperposicion = salasExistentes.some(sala => {
      // Si estamos editando una sala existente, no queremos compararla con ella misma
      if (sala.Id === currentSala.Id) return false;

      const [existenteX, existenteY] = sala.COORDENADA.split(',').map(num => parseInt(num, 10));

      const diferenciaX = Math.abs(currentX - existenteX);
      const diferenciaY = Math.abs(currentY - existenteY);

    // Si la distancia en ambos ejes es menor al margen, están chocando
      return diferenciaX < MARGEN_X && diferenciaY < MARGEN_Y;
  });

  return !haySuperposicion; // Retorna true si el lugar está libre
};




  const columns: IColumn[] = [
    { key: 'name', name: strings.SalaNameLabel, fieldName: 'Title', minWidth: 250, maxWidth: 260, isResizable: true }, 
    { key: 'division', name: strings.DivisionLabel, fieldName: 'PISOId', minWidth: 150, maxWidth: 160, isResizable: true, isMultiline:true,
      onRender: (item: ISPSalaItem) => {
        if (item.PISOId === undefined) return 'N/A'; // Handle undefined PISOId
        const piso = pisos.find(p => p.key === item.PISOId);
        const divisionId = piso ? piso.data?.plantaId : item.PISOId;
        return plantas.find(planta => planta.key === divisionId)?.text || divisionId;

        //return piso ? piso.data?.plantaId : item.PISOId;
      }
    },
      //setSelectedPlantaInModal(pisos.find(p => Number(p.key) === sala.PISOId)?.data?.plantaId);
    { key: 'piso', name: strings.SalaPisoLabel, fieldName: 'PISOId', minWidth: 150, maxWidth: 160, isResizable: true, isMultiline:true,
      onRender: (item: ISPSalaItem) => {
        if (item.PISOId === undefined) return 'N/A'; // Handle undefined PISOId
        const piso = pisos.find(p => p.key === item.PISOId);
        return piso ? piso.text : item.PISOId;
      }
    },
    { key: 'capacity', name: strings.SalaCupoLabel, fieldName: 'CAPACIDAD', minWidth: 50, isResizable: true },
   // { key: 'coordinates', name: strings.SalaCoordinatesLabel, fieldName: 'COORDENADA', minWidth: 100, isResizable: true },
    { key: 'active', name: strings.SalaActiveLabel, fieldName: 'activo', minWidth: 80, isResizable: true,
      onRender: (item: ISPSalaItem) => (item.activo ? 'Sí' : 'No')
    },
    {
      key: 'actions', name: strings.AccionesColumn, minWidth: 150, isResizable: true,
      onRender: (item: ISPSalaItem) => (
        <Stack horizontal tokens={{ childrenGap: 5 }} wrap>
          <TooltipHost content={strings.VerQR}>
            <IconButton iconProps={{ iconName: 'QRCode' }} onClick={() => handleVerQRClick(item)} />
          </TooltipHost>
          <TooltipHost content={strings.EditSalaButton}>
            <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => handleEditClick(item)} />
          </TooltipHost>
          <TooltipHost content={strings.DeleteSalaButton}>
            <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => 
               handleDelete (item)}  //handleDeleteClick(item.Id!, item.Title)} 
               />
          </TooltipHost>
        </Stack>
      )
    }
  ];

  const handleFilterPisoChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option && option.key !== 'all') {
      const keyAsNumber = Number(option.key);
      if (!isNaN(keyAsNumber)) {
        (setFilterPisoId as any)(keyAsNumber); // Cast to any
      } else {
        (setFilterPisoId as any)(undefined); // Cast to any
      }
    } else {
      (setFilterPisoId as any)(undefined); // Cast to any
    }
  };

  return (
    <div className={styles.mantenedorSalas}>
      <h2>{strings.MantenedorSalasTitle}</h2>

      {loading && <Spinner size={SpinnerSize.large} label={strings.LoadingSalas} />}
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {message && !error && <MessageBar messageBarType={message.type} onDismiss={() => setMessage(undefined)}>{message.text}</MessageBar>}

      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 15 }}>
        <PrimaryButton iconProps={{ iconName: 'Add' }} onClick={handleAddClick}>
          {strings.AddSalaButton}
        </PrimaryButton>
        <Dropdown
          placeholder={strings.SelectPisoFilterPlaceholder}
          options={pisos}
          selectedKey={filterPisoId  || null}
          onChange={handleFilterPisoChange}
          style={{ width: 200 }}
        />
      </Stack>

      {!loading && !error && salas.length === 0 && (
        <MessageBar messageBarType={MessageBarType.info}>{strings.NoSalasFound}</MessageBar>
      )}

      {!loading && !error && salas.length > 0 && (
        <DetailsList
          items={salas}
          columns={columns}
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
        />
      )}

      <Dialog
        hidden={!isModalOpen}
        onDismiss={handleCancel}
        className={styles.customSalaModal}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: currentSala?.Id ? strings.EditSalaButton : strings.AddSalaButton,
        }}
        modalProps={{
          isBlocking: true,
        }}
      >
        {titleError && (
            <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                {titleError}
            </MessageBar>
        ) }
        <Stack tokens={{ childrenGap: 15 }}>
          <div style={{ maxWidth: '300px' }}>
            {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
            <TextField
              label={strings.SalaNameLabel}
              value={currentSala?.Title || ''}
              onChange={(e, newValue) => {
                const sanitizedValue = (newValue || '').replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ\s\-]/g, '');
                //setCurrentSala(prev => ({ ...prev, Title: newValue || '' } as ISPSalaItem))}
                setCurrentSala(prev => ({ 
                  ...prev, 
                  Title: sanitizedValue 
                } as ISPSalaItem));
              }}
              required
              errorMessage={salaNameError}
            />
          </div>
          <div style={{ maxWidth: '300px' }}>
            <Dropdown
              label={strings.PlantaLabel}
              placeholder={strings.SelectPlantaPlaceholder}
              options={plantas}
              selectedKey={selectedPlantaInModal || null}
              onChange={(e, option) => {
                setSelectedPlantaInModal(option ? Number(option.key) : undefined);
                setSelectedPisoInModal(undefined); // Reset piso when planta changes
                setCurrentSala(prev => ({ ...prev, PISOId: undefined } as ISPSalaItem)); // Reset PISOId in currentSala
              }}
              required
            />
          </div>
          <div style={{ maxWidth: '300px' }}>
            <Dropdown
              label={strings.SalaPisoLabel}
              placeholder={strings.SelectPisoPlaceholder}
              options={pisos.filter(p => p.key !== 'all' && (selectedPlantaInModal ? p.data?.plantaId === selectedPlantaInModal : true))}
              selectedKey={selectedPisoInModal || null}
              onChange={(e, option) => {
                setSelectedPisoInModal(option ? Number(option.key) : undefined);
                setCurrentSala(prev => ({ ...prev, PISOId: option ? Number(option.key) : undefined } as ISPSalaItem));
              }}
              required
              errorMessage={salaPisoError}
              disabled={!selectedPlantaInModal}
            />
          </div>
          <div style={{ maxWidth: '300px' }}>
            <TextField
              label={strings.SalaCupoLabel}
              type="number"
              value={currentSala?.CAPACIDAD?.toString() || '0'}
              onChange={(e, newValue) => {
                const sanitizedValue = (newValue || '').replace(/\D/g, '');
                
                //setCurrentSala(prev => ({ ...prev, CAPACIDAD: Number(newValue) || 0 } as ISPSalaItem))}
                setCurrentSala(prev => ({ 
                  ...prev, 
                  CAPACIDAD: sanitizedValue === '' ? 0 : parseInt(sanitizedValue, 10) 
                } as ISPSalaItem));
              }}
              required
              errorMessage={salaCapacityError}
            />
          </div>

          {isImageLoading && <Spinner size={SpinnerSize.large} label="Cargando imagen del plano..." />}
          {!isImageLoading && selectedPisoInModal && !floorImageUrl && (
            <MessageBar messageBarType={MessageBarType.warning}>
              No se encontró imagen de plano para el piso seleccionado.
            </MessageBar>
          )}

          {floorImageUrl && !isImageLoading && (
            <>
              <div className={styles.planContainer}>
                <img
                  src={floorImageUrl}
                  alt={strings.PlanoPisoTitle}
                  className={styles.planImage}
                  onClick={handleImageClick}
                  onLoad={(e) => {
                    const img = e.currentTarget;
                    setImageDimensions({ width: img.naturalWidth, height: img.naturalHeight });
                    setIsImageLoading(false); // Image loaded
                  }}
                  onError={() => {
                    setFloorImageUrl(undefined); // Clear image on error
                    setIsImageLoading(false); // Image failed to load
                    setMessage({ type: MessageBarType.error, text: "Error al cargar la imagen del plano." }); // Provide feedback
                    //setMessageType(MessageBarType.error);
                  }}
                />
 
                {salas.map(sala => (
                  <PrimaryButton
                    key={sala.Id}
                    text={sala.Title}
                    disabled={true}
                    title={sala.Title}
                    style={{
                      position: 'absolute',
                      left: imageDimensions ? `${(parseInt(sala.COORDENADA.split(',')[0], 10) / imageDimensions.width * 100)}%` : '0',
                      top: imageDimensions ? `${(parseInt(sala.COORDENADA.split(',')[1], 10) / imageDimensions.height * 100)}%` : '0',
                      transform: 'translate(-50%, -50%)',
                      visibility: imageDimensions ? 'visible' : 'hidden',
                      backgroundColor: 'rgba(80, 80, 80, 0.5)',
                    }}
                  />
                ))}
                {currentSala?.COORDENADA && validarCoordenadas(salas) && (
                  //alert("hpolas"),
                  <div
                    style={{
                      position: 'absolute',
                      left: `${(parseInt(currentSala.COORDENADA.split(',')[0], 10) / (imageDimensions?.width || 1)) * 100}%`,
                      top: `${(parseInt(currentSala.COORDENADA.split(',')[1], 10) / (imageDimensions?.height || 1)) * 100}%`,
                      transform: 'translate(-50%, -50%)',
                      width: '80px',
                      height: '30px',
                      //borderRadius: '50%',
                      backgroundColor: 'red',
                      border: '2px solid white',
                      boxShadow: '0 0 5px rgba(0,0,0,0.5)',
                      pointerEvents: 'none', // Prevent this div from capturing clicks
                    }}
                  />
                )}
              </div>
              {salaCoordinatesError && <div style={{ color: '#a80000', fontSize: '12px', fontWeight: 'normal', fontFamily: 'Segoe UI, "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif' }}>{salaCoordinatesError}</div>}
            </>
          )}
        </Stack>
        <DialogFooter>
          <PrimaryButton onClick={handleSave} text={strings.SaveButton} disabled={loading} />
          <DefaultButton onClick={handleCancel} text={strings.CancelButton} disabled={loading} />
        </DialogFooter>
      </Dialog>
      <Dialog
            hidden={!showDeleteConfirm}
            onDismiss={() => setShowDeleteConfirm(false)}
            dialogContentProps={{
                type: DialogType.normal,
                title: strings.ConfirmDeleteSala.replace('{0}', salaToDelete?.Title || ''),
                //subText: strings.ConfirmDeleteDivision.replace('{0}', divisionToDelete?.Title || '')
            }}
            modalProps={{
                isBlocking: true,
                styles: { main: { maxWidth: 450 } }
            }}
        >
            <DialogFooter>
                <PrimaryButton onClick={confirmDelete} text={strings.DeleteSalaButton} />
                <DefaultButton onClick={handleCancel} text={strings.CancelButton} />
            </DialogFooter>
        </Dialog>

        {/* Modal para visualizar el QR */}
        <Dialog
          hidden={!isQRModalOpen}
          onDismiss={() => setIsQRModalOpen(false)}
          dialogContentProps={{
            type: DialogType.normal,
            title: `Código QR - ${selectedSalaName}`
          }}
          modalProps={{ isBlocking: false }}
        >
          <Stack horizontalAlign="center" tokens={{ childrenGap: 20 }} style={{ marginTop: '10px' }}>
            {qrImageUrl ? (
              <img 
                src={qrImageUrl} 
                alt={`QR ${selectedSalaName}`} 
                style={{ width: '250px', height: '250px', border: '1px solid #eee' }}
                onError={(e) => {
                  // Si la imagen no existe, podrías mostrar un placeholder o error
                  (e.target as HTMLImageElement).src = 'https://via.placeholder.com/250?text=QR+No+Encontrado';
                }}
              />
            ) : (
              <Spinner size={SpinnerSize.large} />
            )}
          </Stack>
          <DialogFooter>
            
            {/* Botón de Descarga */} 
            <PrimaryButton
              onClick={descargarQR} 
              iconProps={{ iconName: 'Download' }} 
              text="Descargar" 
            />
            {/* Botón de Impresión Limpia */}
            <DefaultButton 
              onClick={imprimirSoloQR} 
              iconProps={{ iconName: 'Print' }} 
              text="Imprimir" 
            />

            <DefaultButton 
              onClick={() => setIsQRModalOpen(false)} 
              text={strings.CloseButton || "Cerrar"} 
            />
          </DialogFooter>
        </Dialog>
    </div>

    
  );
};

export default MantenedorSalas;