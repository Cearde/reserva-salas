import * as React from 'react';
//import { useState, useEffect, useContext } from 'react';
import { useState } from 'react';
import { Dropdown, IDropdownOption, PrimaryButton, Stack, Image, MessageBar, MessageBarType, Spinner, SpinnerSize, Text } from '@fluentui/react';
import * as qrcode from 'qrcode';
import { useSPFxContext } from '../../contexts/SPFxContext';
import { SPService } from '../../services/sp';
//import { IDivisionItem, IPisoItem, ISala } from '../../models/entities';
import * as strings from 'SrscWebPartStrings';
import styles from '../Srsc.module.scss';
//import { FieldAttachmentsRenderer } from '@pnp/spfx-controls-react';

const GenerarQR: React.FC = () => {
  //const spfxContext = useContext(SPFxContext);
  const context = useSPFxContext();
  //const spService = React.useMemo(() => spfxContext ? new SPService(spfxContext) : undefined, [spfxContext]);
  const spService = React.useMemo(() => new SPService(context), [context]);
  //const [plantas, setPlantas] = useState<IDropdownOption[]>([]);
  const [selectedPlantaId, setSelectedPlantaId] = useState<string | number | undefined>(undefined);
  const [selectedPlantaName, setSelectedPlantaName] = useState<string | undefined>(undefined);
  
  //const [pisos, setPisos] = useState<IDropdownOption[]>([]);
  const [selectedPisoId, setSelectedPisoId] = useState<string | number | undefined>(undefined);
  const [selectedPisoName, setSelectedPisoName] = useState<string | undefined>(undefined);
  
  const [salas,setSalas] = useState<IDropdownOption[]>([]);
  const [selectedSalaId, setSelectedSalaId] = useState<string | number | undefined>(undefined);
 // const [selectedSalaName, setSelectedSalaName] = useState<string | undefined>(undefined);
  
  const [plantas, setPlantas] = React.useState<IDropdownOption[]>([]);
  const [pisos, setPisos] = React.useState<IDropdownOption[]>([]);
  //const [usuarios, setUsuarios] = React.useState<IDropdownOption[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [message, setMessage] = React.useState<string | undefined>(undefined);
  const [messageType, setMessageType] = React.useState<MessageBarType>(MessageBarType.info);

  //const [qrCodeIndividualUrl, setQrCodeIndividualUrl] = useState<string | undefined>(undefined);
  //const [qrCodesMassive, setQrCodesMassive] = useState<{ salaName: string; url: string; }[]>([]);
  const [qrCodes, setQrCodes] = useState<{ 
                                          plantaId: string | number | undefined; 
                                          pisoId: string | number | undefined;
                                          salaId: string | number | undefined;
                                          salaName: string; 
                                          fileBase64: string }[]>([]);
  //const [loading, setLoading] = useState<boolean>(false);
  //const [error, setMessage] = useState<string | undefined>(undefined);
  const [showIndividualSalaSelector, setShowIndividualSalaSelector] = useState<boolean>(false);
  const generatedQRs: { plantaId: string | number | undefined;
                        pisoId: string | number | undefined;
                        salaId: string | number | undefined;
                        salaName: string;
                        fileBase64: string }[] = [];
  //const baseUrl = "https://codelcochile.sharepoint.com/sites/pacm/SitePages/RegistrarPuesto.aspx";
/*
  useEffect(() => {
    if (!spService) {
      setMessage("SPService no está inicializado.");
      return;
    }

    setLoading(true);
    spService.getDivisiones(true) // Fetch all plantas, including inactive ones if needed for QR generation
      .then(data => {
        const options: IDropdownOption[] = data.map(item => ({
          key: item.Id,
          text: item.Title,
          data: item
        }));
        setPlantas(options);
        setLoading(false);
      })
      .catch(err => {
        setMessage(`Error al cargar las Divisiones: ${err.message}`);
        setLoading(false);
      });
  }, [spService]);

  useEffect(() => {
    if (!spService || selectedPlantaId === undefined) {
      setPisos([]);
      setSelectedPisoId(undefined);
      setSalas([]);
      setSelectedSalaId(undefined);
      setQrCodeIndividualUrl(undefined);
      setQrCodesMassive([]);
      return;
    }

    setLoading(true);
    spService.getPisos(true) // Fetch all pisos, including inactive ones
      .then(data => {
        const filteredPisos = data.filter(piso => piso.PlantaId === selectedPlantaId);
        const options: IDropdownOption[] = filteredPisos.map(item => ({
          key: item.Id,
          text: item.Title,
          data: item
        }));
        setPisos(options);
        setLoading(false);
      })
      .catch(err => {
        setMessage(`Error al cargar los Pisos: ${err.message}`);
        setLoading(false);
      });
  }, [spService, selectedPlantaId]);

  useEffect(() => {
    if (!spService || selectedPisoId === undefined || !showIndividualSalaSelector) {
      setSalas([]);
      setSelectedSalaId(undefined);
      setQrCodeIndividualUrl(undefined);
      return;
    }

    setLoading(true);
    spService.getSalas(selectedPisoId as number, true) // Fetch all salas for the selected piso
      .then(data => {
        const options: IDropdownOption[] = data.map(item => ({
          key: item.Id,
          text: item.Title,
          data: item
        }));
        setSalas(options);
        setLoading(false);
      })
      .catch(err => {
        setMessage(`Error al cargar las Salas: ${err.message}`);
        setLoading(false);
      });
  }, [spService, selectedPisoId, showIndividualSalaSelector]);
*/

React.useEffect(() => {
    setLoading(true);
    Promise.all([
      spService.fetchActiveListItems('LM_PLANTAS'),
      spService.getPisosWithPlantaId(),
     // spService.getHotspotsForFloor(),
    ]).then(([plantasData, pisosData]) => {
      setPlantas(plantasData);
      setPisos(pisosData);
    }).catch(error => {
      console.error("Error fetching dropdown options for SalasDisponibles", error);
      setMessage(strings.ErrorLoadingFilterOptions);
      setMessageType(MessageBarType.error);
    }).finally(() => {
      setLoading(false);
    });
  }, [spService]);

  const onPlantaChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption | undefined): void => {
    setSelectedPlantaId(item?.key as number);
    setSelectedPlantaName(item?.text as string);
    setSelectedPisoId(undefined);
    setSelectedSalaId(undefined);
    setQrCodes([]);
    //setQrCodesMassive([]);
    setMessage(undefined);
  };

  const onPisoChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption | undefined): void => {
    setSelectedPisoId(item?.key as number);
    setSelectedPisoName(item?.text as string);
    console.log("Piso seleccionado ID:", item?.key);
    setSelectedSalaId(undefined);
    setQrCodes([]);
    //setQrCodesMassive([]);
    setMessage(undefined);
  };

  React.useEffect(() => {
  // Solo disparamos la carga si hay un ID válido
  if (selectedPisoId) {
    setLoading(true);
    
    spService.fetchActiveListItems('LM_PUESTOSPISO', ['PISOId'], `PISOId eq ${selectedPisoId}`)
      .then((salasData) => {
        // Suponiendo que tienes un setSalas o similar para las opciones del dropdown
        setSalas(salasData); 
      })
      .catch(error => {
        setMessage(strings.ErrorLoadingFilterOptions);
        setMessageType(MessageBarType.error);
      })
      .finally(() => {
        setLoading(false);
      });
  } else {
    // Si el ID es undefined (porque cambiaron la planta), limpiamos las opciones
    setSalas([]);
  }
}, [selectedPisoId, spService]); // Se ejecuta cuando cambia el piso o el servicio


React.useEffect(() => {
  // Solo disparamos la carga si hay un ID válido
  if (selectedPlantaId) {
    setLoading(true);
    
    spService.fetchActiveListItems('LM_PISOS', ['PLANTAId'], `PLANTAId eq ${selectedPlantaId}`)
      .then((divisiondata) => { 
        setPisos(divisiondata); 
      })
      .catch(error => {
        setMessage(strings.ErrorLoadingFilterOptions);
        setMessageType(MessageBarType.error);
      })
      .finally(() => {
        setLoading(false);
      });
  } else {
    // Si el ID es undefined (porque cambiaron la planta), limpiamos las opciones
    setPisos([]);
  }
}, [selectedPlantaId, spService]);


  const onSalaChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption | undefined): void => {
    setSelectedSalaId(item?.key as number);
   // setSelectedSalaName(item?.text as string);
    setQrCodes([]);
    setMessage(undefined);
  };

  const generateQrCode = async (pisoId: string, salaId: string): Promise<string> => {
    //const url = `${baseUrl}?piso=${encodeURIComponent(pisoId)}&sala=${encodeURIComponent(salaId)}`;
    const url = JSON.stringify({
                'piso':pisoId,
                'Sala':salaId
                });
    return qrcode.toDataURL(url);
  };

  const handleGenerateQrIndividual = async (): Promise<void> => {
    setMessage(undefined);
    //setQrCodeIndividualUrl(undefined);
    setQrCodes([]);
    

    if (!selectedPlantaId || !selectedPisoId || !selectedSalaId) {
      setMessage(strings.GenerarQRMissingSelection);
      return;
    }
    setMessage(undefined);
    
    setLoading(true);
    try {
      const selectedPiso = pisos.find(p => p.key === selectedPisoId)?.text;
      const selectedSala = salas.find(s => s.key === selectedSalaId)?.text;

      if (!selectedPiso || !selectedSala) {
        setMessage(strings.GenerarQRInvalidSelection);
        setLoading(false);
        return;
      }

      const fileBase64 = await generateQrCode(selectedPisoId.toString(), selectedSalaId.toString());
      //setQrCodeIndividualUrl(qrUrl);
      generatedQRs.push({ plantaId: selectedPlantaId,
                          pisoId: selectedPisoId,                          
                          salaId: selectedSalaId, 
                          salaName: selectedSala,
                          fileBase64: fileBase64
                        }); 
      setQrCodes(generatedQRs);
    } catch (err) {
      setMessage(`Error al generar QR individual: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const handleGenerateQrMassive = async (): Promise<void> => {
    setMessage(undefined);
    //setQrCodeIndividualUrl(undefined);
    setQrCodes([]);

    if (!selectedPlantaId || !selectedPisoId) {
      setMessage(strings.GenerarQRMissingSelection);
      return;
    }

    setLoading(true);
    try {
      const selectedPiso = pisos.find(p => p.key === selectedPisoId)?.text;
      if (!selectedPiso) {
        setMessage(strings.GenerarQRInvalidSelection);
        setLoading(false);
        return;
      }

      const allSalasForPiso = await spService?.getSalas(selectedPisoId as number, true);
      if (!allSalasForPiso || allSalasForPiso.length === 0) {
        setMessage(strings.GenerarQRNoRoomsFound);
        setLoading(false);
        return;
      }

      for (const sala of allSalasForPiso) {
        const fileBase64 = await generateQrCode(selectedPiso, sala.Title);
        generatedQRs.push({ plantaId: selectedPlantaId.toString(),
                            salaId: sala.Id!.toString(), 
                            salaName: sala.Title,
                            pisoId: selectedPisoId.toString(),
                            fileBase64: fileBase64 });
      }
      setQrCodes(generatedQRs);

    } catch (err) {
      setMessage(strings.GenerarQRNoRoomsFound);
      setLoading(false);
    } finally {
      setLoading(false);
    }
  };


  const resetForm = () => {
    
    setSelectedPlantaId(undefined); // O el valor inicial que tengas
    //setPlantas([]);
    setSelectedPisoId(undefined);
    setQrCodes([]); // Limpia la lista de QRs generados
    };

  const saveGenerateQrMassive = async (): Promise<void> => {

    console.log("Guardar QR masivo");
    //console.log(qrCodesMassive);
    console.log(qrCodes);
    //const codigosQR 
    try { 
    
      setLoading(true);

      for (const qr of qrCodes) {
        const base64Clean = qr.fileBase64.split(',').pop() || ""; 
        const binaryString = window.atob(base64Clean);
        const len = binaryString.length;
        
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
          bytes[i] = binaryString.charCodeAt(i);
        } 

        const folterPath = `LO_QRPUESTOS/${selectedPlantaName}`;
        let  folderDivision = await spService.ensureFolderPath(folterPath);

        if(folderDivision !== undefined){
            folderDivision += `/${selectedPisoName}`
            const folderPiso = await spService.ensureFolderPath(folderDivision);
            if(folderPiso !== undefined){
                const fileId =spService.uploadFile(`QR_Sala_${qr.salaName}.png`,bytes.buffer, folderPiso);
                if(fileId !== undefined){
                    const metadata = {
                                        //"__metadata": { "type": "SP.Data.LO_QRPUESTOSItem" },
                                        "PLANTAId": selectedPlantaId,
                                        "PISOId": selectedPisoId,
                                        "SALAId": qr.salaId,
                                      }
                      spService.uploadFileWithMetadata('LO_QRPUESTOS',(await fileId).Id,metadata)


                }
            }
        }
      }
      resetForm();
      setMessage(strings.guardarQRSuccessMessage);
      setMessageType(MessageBarType.success);
    } catch (error) {
        setMessage("Error subiendo archivo Base64:");
        setMessageType(MessageBarType.error);
    }finally {
      setLoading(false);
    }  
  };

  return (
    <div className={styles.generarQRContainer}>
      <Text variant="xxLarge" className={styles.headerText}>{strings.GenerarQRTitle}</Text>

      {loading && <Spinner size={SpinnerSize.large} label={strings.LoadingConfig} />}
      {message && <MessageBar messageBarType={messageType} isMultiline={true}>{message}</MessageBar>}

      <Stack horizontal wrap tokens={{ childrenGap: 15 }}>
        <div className={`${styles.formGroup} ${styles.formControl}`}>
          <Dropdown
            label={strings.PlantaLabel}
            options={plantas}
            selectedKey={selectedPlantaId}
            onChange={onPlantaChange}
            placeholder={strings.SelectPlantaPlaceholder}
            disabled={loading}
          />
        </div>
        <div className={`${styles.formGroup} ${styles.formControl}`}>
          <Dropdown
            label={strings.PisoLabel}
            options={pisos}
            selectedKey={selectedPisoId}
            //options={pisos.filter(p => p.plantaId === selectedPlantaId)}
            onChange={onPisoChange}
            placeholder={strings.SelectPisoPlaceholder}
            disabled={loading || selectedPlantaId === undefined}
          />
        </div>
      </Stack>
      <Stack horizontal tokens={{ childrenGap: 10 }} className={styles.buttonGroup}>
        <div className={`${styles.formGroup} ${styles.formControl}`}>
          <PrimaryButton
            text={strings.GenerarQRIndividualButton}
            onClick={() => {
              setShowIndividualSalaSelector(true);
              //setQrCodeIndividualUrl(undefined);
              setQrCodes([]);
              setSelectedSalaId(undefined);
              setMessage(undefined);
            }}
            disabled={loading || selectedPisoId === undefined}
          /> 
          <PrimaryButton  
            style={{ marginLeft: '10px' }}
            text={strings.GenerarQRMassiveButton}
            onClick={() => {
              setShowIndividualSalaSelector(false);
              //setQrCodeIndividualUrl(undefined);
              setQrCodes([]);
              setSelectedSalaId(undefined);
              setMessage(undefined);
              handleGenerateQrMassive();
            }}
            disabled={loading || selectedPisoId === undefined}
          />
        </div>
      </Stack>

        {showIndividualSalaSelector && selectedPisoId !== undefined && (
          <Stack horizontal tokens={{ childrenGap: 10 }} className={styles.buttonGroup}>
            <div className={`${styles.formGroup} ${styles.formControl}`}>
              <Dropdown
                label={strings.SalaLabel}
                options={salas}
                selectedKey={selectedSalaId}
                onChange={onSalaChange}
                placeholder={strings.SelectSalaPlaceholder}
                disabled={loading || selectedPisoId === undefined}
              />
            </div>
          </Stack>
        )}

        {showIndividualSalaSelector && selectedSalaId !== undefined && (
          <PrimaryButton
            text={strings.GenerarQRButton}
            onClick={handleGenerateQrIndividual}
            disabled={loading || selectedSalaId === undefined}
          />
        )}
      

      <Stack tokens={{ childrenGap: 20 }} className={styles.qrDisplayArea}>
        {qrCodes.length === 1 && (
          <Stack horizontalAlign="center" verticalAlign="center" className={styles.qrCodeItem}>
            <Text variant="large">{strings.GenerarQRIndividualResult}</Text>
            <Image src={qrCodes[0].fileBase64} alt="QR Code" width={200} height={200} />
          </Stack>
        )}

        {qrCodes.length > 1 && (
          <Stack tokens={{ childrenGap: 10 }}>
            <Text variant="large">{strings.GenerarQRMassiveResult}</Text>
            <div className={styles.massiveQrGrid}>
              {qrCodes.map((qr, index) => (
                <Stack key={index} horizontalAlign="center" verticalAlign="center" className={styles.qrCodeItem}>
                  <Text variant="medium" className={styles.qrLabel}>{qr.salaName}</Text>
                  <Image src={qr.fileBase64} alt={`QR Code for ${qr.salaName}`} width={150} height={150} />
                </Stack>
              ))}
            </div>
          </Stack>
        )}
      </Stack>

      <div>

        {(qrCodes.length > 0) && (
          <PrimaryButton
            text={strings.GuardarQRButton}
            onClick={saveGenerateQrMassive}
            disabled={qrCodes.length === 0}
          />
        )}
      </div>
    </div>
  );
};

export default GenerarQR;
