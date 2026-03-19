import * as React from 'react';
import { IReservaSalaScheduleProps, Iperson } from '../models/entities';
import { DefaultButton, PrimaryButton, Spinner, SpinnerSize, Checkbox, Stack } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { useSPFxContext } from '../../contexts/SPFxContext';
//import { WebPartContext } from '@microsoft/sp-webpart-base';
//import { IPersonaProps } from '@fluentui/react/lib/Persona';
import * as strings from 'SrscWebPartStrings';

const overlayStyle: React.CSSProperties = {
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
};

const modalStyle: React.CSSProperties = {
  backgroundColor: '#fff',
  padding: '20px',
  borderRadius: '5px',
  minWidth: '800px',
  maxWidth: '90%',
  boxShadow: '0 5px 15px rgba(0, 0, 0, 0.3)',
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  maxHeight: '90vh',       
  overflowY: 'auto',       
};

const headerInfoContainerStyle: React.CSSProperties = {
  display: 'flex',
  justifyContent: 'space-around',
  alignItems: 'center',
  width: '100%',
  padding: '15px',
  backgroundColor: '#f7f7f7',
  borderRadius: '4px',
  marginBottom: '20px',
  border: '1px solid #e9e9e9'
};

const infoCardStyle: React.CSSProperties = {
  textAlign: 'center',
  padding: '0 15px'
};

const infoCardLabelStyle: React.CSSProperties = {
  fontSize: '0.9em',
  color: '#666',
  marginBottom: '4px',
  display: 'block'
};

const infoCardValueStyle: React.CSSProperties = {
  margin: 0,
  fontSize: '1.1em',
  fontWeight: 600,
  color: '#333'
};

const tableStyle: React.CSSProperties = {
  width: '100%',
  borderCollapse: 'collapse',
  marginTop: '15px',
  marginBottom: '15px',
};

const thStyle: React.CSSProperties = {
  border: '1px solid #ccc',
  padding: '8px',
  textAlign: 'left',
  backgroundColor: '#f2f2f2',
};

const tdStyle: React.CSSProperties = {
  border: '1px solid #ccc',
  padding: '8px',
  textAlign: 'left',
};

const ReservaSalaSchedule: React.FC<IReservaSalaScheduleProps> = ({
  isOpen,
  roomName,
  selectedPisoName,
  selectedDate,
  usuarioId,
  horarios,
  onClose,
  onSelectSlot,
  isLoading = false,
  CAPACIDAD 
}) => {
  const context = useSPFxContext();
  console.log(`[ReservaSalaSchedule.tsx] Render. isOpen: ${isOpen}, isLoading: ${isLoading}, horarios.length: ${horarios.length}`);
  console.log(`[ReservaSalaSchedule.tsx] Props:`, { roomName, selectedPisoName, selectedDate, usuarioId });
  console.log(`[ReservaSalaSchedule.tsx] capacidad 29: `, CAPACIDAD );
  
  const [selectedHourIds, setSelectedHourIds] = React.useState<string[]>([]);
  const [selectedPeople, setSelectedPeople] = React.useState<Iperson[]>([]);// React.useState<IPersonaProps[]>([]);
 // const [listaFinal, setListaFinal] = React.useState<any[]>([]);
  const [pickerKey, setPickerKey] = React.useState<number>(0);

  React.useEffect(() => {
    if (!isOpen) {
      setSelectedHourIds([]); // Clear selections when modal closes
      setSelectedPeople([]); // Clear people picker
    }
  }, [isOpen]);

  if (!isOpen) {
    console.log('[ReservaSalaSchedule.tsx] Modal is not open, returning null.');
    return null;
  }

  const _getPeoplePickerItems = (items: any[]) => {// (items: IPersonaProps[]): void => {
    //setSelectedPeople(items);
    if (items.length > 0) {
    // 1. Agregamos el nuevo usuario a nuestra lista acumulada
    // Evitamos duplicados comparando por loginName
        const nuevoUsuario = items[0];
        setSelectedPeople(prev => {
          if (prev.some(u => u.loginName === nuevoUsuario.loginName)) return prev;
          return [...prev, nuevoUsuario];
        });

        // 2. ¡EL TRUCO! Cambiamos la key para limpiar el buscador
        setPickerKey(prev => prev + 1);
  }
  }

  const _eliminarUsuario = (loginName: string) => {
  setSelectedPeople(prev => prev.filter(u => u.loginName !== loginName));
};

  const formattedDate = selectedDate ? selectedDate.toLocaleDateString() : strings.FechaNotSelected;

  const handleCheckboxChange = (horarioId: string, isChecked: boolean | undefined): void => {
    console.log(`[handleCheckboxChange] horarioId: ${horarioId}, isChecked: ${isChecked}`);
    setSelectedHourIds(prev => {
      const newSelected = isChecked ? [...prev, horarioId] : prev.filter(id => id !== horarioId);
      console.log(`[handleCheckboxChange] newSelectedHourIds:`, newSelected);
      return newSelected;
    });
  };

  const handleConfirmReservation = (): void => {
    // Map selectedHourIds back to IHorarioSala objects
    const selectedHorarios = horarios.filter(horario => selectedHourIds.includes(horario.ID));
    console.log('[ReservaSalaSchedule.tsx] Confirming reservation with selected horarios:', selectedHorarios);
    onSelectSlot(selectedHorarios, selectedPeople);
  };

  return (
    <div style={overlayStyle}>
      <div style={modalStyle}>
        <h3>{strings.ReservationDetailsTitle}</h3>
        <div style={headerInfoContainerStyle}>
          <div style={infoCardStyle}>
            <span style={infoCardLabelStyle}>{strings.PisoColumn}</span>
            <p style={infoCardValueStyle}>{selectedPisoName}</p>
          </div>
          <div style={infoCardStyle}>
            <span style={infoCardLabelStyle}>{strings.SalaColumn}</span>
            <p style={infoCardValueStyle}>{roomName}</p>
          </div>
          <div style={infoCardStyle}>
            <span style={infoCardLabelStyle}>{strings.FechaReservaColumn}</span>
            <p style={infoCardValueStyle}>{formattedDate}</p>
          </div>
        </div>

        {isLoading ? (
          <Spinner size={SpinnerSize.large} label={strings.LoadingHorarios} style={{ margin: '20px 0' }} />
        ) : (
          <>
            {horarios.length === 0 ? (
              <p>{strings.NoHorariosAvailable}</p>
            ) : (
              <table style={tableStyle}>
                <thead>
                  <tr>
                    <th style={thStyle}>{strings.HorasDisponiblesTable}</th>
                    <th style={thStyle}>{strings.DisponibilidadTable}</th>
                    <th style={thStyle}>{strings.ReservadoPorTable}</th>
                  </tr>
                </thead>
                <tbody>
                  {horarios.map((horario) => {
                    const isChecked = selectedHourIds.includes(horario.ID);
                    console.log(`[Checkbox Render] Horario ID: ${horario.ID}, HORA: ${horario.HORA}, isChecked: ${isChecked}, selectedHourIds:`, selectedHourIds);
                    return (
                      <tr key={horario.ID} style={{ color: horario.statusColor }}>
                        <td style={tdStyle}>
                          <Checkbox
                            label={horario.HORA}
                            disabled={!horario.Disponibilidad}
                            checked={isChecked}
                            onChange={(ev, isChecked) => handleCheckboxChange(horario.ID, isChecked)}
                            title={horario.HORA}
                          />
                        </td>
                        <td style={tdStyle}>
                          <strong>{horario.statusText}</strong>
                        </td>
                        <td style={tdStyle}>{horario.reservadoPorText}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}

            <div style={{width: '100%', marginTop: '10px', marginBottom: '20px'}}>
              <p style={infoCardLabelStyle}>
               Esta sala de reunión tiene cupo disponible para {CAPACIDAD} persona(s) adicionales al solicitante.
              </p>
              
              <PeoplePicker
                key={pickerKey}
                context={context as any}
                disabled= {selectedPeople.length >= CAPACIDAD} // Deshabilitar si se alcanzó la capacidad
                /*context={{
                  pageContext: {
                    web: {
                      absoluteUrl: 'https://codelcochile.sharepoint.com/sites/srsc', //context.pageContext.web.absoluteUrl,
                      serverRelativeUrl: 'https://codelcochile.sharepoint.com/sites/srsc' //context.pageContext.web.serverRelativeUrl
                    }
                  },
                  spHttpClient: context.spHttpClient
                } as any}*/
                titleText={strings.AsistentesPeoplePicker}
                webAbsoluteUrl={context.pageContext.web.absoluteUrl} 
                //showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
                onChange={_getPeoplePickerItems}
                defaultSelectedUsers={[]}
              />

              {/* Lista de usuarios debajo */}
                <div style={{ marginTop: '20px' }}>
                  <h4>Usuarios seleccionados:</h4>
                  <ul style={{ listStyle: 'none', padding: 0 }}>
                    {selectedPeople.map(user => (
                      <li key={user.loginName} style={{ 
                        display: 'flex', 
                        justifyContent: 'space-between', 
                        padding: '8px', 
                        borderBottom: '1px solid #eee',
                        alignItems: 'center' 
                      }}>
                        <span>{user.text} <small>({user.secondaryText})</small></span>
                        <button 
                          onClick={() => _eliminarUsuario(user.loginName)}
                          style={{ color: 'red', cursor: 'pointer', border: 'none', background: 'none' }}
                        >
                          X
                        </button>
                      </li>
                    ))}
                  </ul>
                </div>

            </div>

            <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: '20px' }}>
              <PrimaryButton
                onClick={handleConfirmReservation}
                disabled={selectedHourIds.length === 0}
              >
                {strings.ConfirmReservationButton} ({selectedHourIds.length})
              </PrimaryButton>
              <DefaultButton onClick={onClose}>{strings.CloseButton}</DefaultButton>
            </Stack>
          </>
        )}
      </div>
    </div>
  );
};

export default ReservaSalaSchedule;
