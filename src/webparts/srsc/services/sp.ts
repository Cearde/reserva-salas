import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import * as strings from 'SrscWebPartStrings';
import { IDropdownOption } from "@fluentui/react"; 
 
import { Web } from  "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

import { ISPItem, ISPListItem, ISPPisoItem,IPisoItem,IReservationSPItem,
         IHotspot,IHorarioItem,IExistingReservation,IAvailableRoom,IReservationReportItem,IVicepresidenciaItem,
         IPlanificacionHoraItem,ISPSalaItem,IDivisionItem,
         IGerenciaItem,
         IUsuarioItem,
         ISPUsuarioItem,
         IUploadResult,
        IHorarioSala,ISala, 
        ISalaItem} from '../components/models/entities';


export class SPService {
  private context: WebPartContext; 

  constructor(context: WebPartContext) {
    this.context = context;
    
  }

  /**
   * Fetches items from a SharePoint list where the 'activo' field is true.
   * @param listName The name of the SharePoint list.
   * @returns A promise that resolves to an array of IDropdownOption.
   */
  public async fetchActiveListItems(listName: string, additionalSelects: string[] = [], filters?: string): Promise<IDropdownOption[]> {
    const selectFields = ['Id', 'Title', ...additionalSelects].join(',');
    const filterConditions = filters? `&$filter=${filters}` : ``;
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=${selectFields}${filterConditions}`;
    
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);

      if (response.ok) {
        const data = await response.json();
        if (data.value) {
          // Use the specific ISPListItem interface instead of 'any'
          return data.value.map((item: ISPListItem) => ({ 
            key: item.Id, 
            text: item.Title,
            data: item // Store the whole item for additional properties
          }));
        }
        return [];
      } else {
        const errorData = await response.json();
        console.error(`Error fetching list ${listName}:`, errorData);
        throw new Error(`Could not fetch ${listName}`);
      }
    } catch (error) {
      console.error(`Exception while fetching list ${listName}:`, error);
      throw error; // Re-throw the error to be caught by the caller
    }
  }

  /**
   * Gets all the necessary dropdown options for the reservation form.
   * @returns A promise that resolves to an object containing arrays of options for plantas, pisos, and usuarios.
   */
  public async getFormDropdownOptions(): Promise<{ plantas: IDropdownOption[], pisos: IDropdownOption[], usuarios: IDropdownOption[] }> {
    const [usuariosData, pisosData,plantasData] = await Promise.all([
      // Fetch only active plantas
      this.fetchActiveListItems('LM_USUARIOS'),
      this.fetchActiveListItems('LM_PISOS', ['IMAGEN','PLANTAId']), // Also fetch the IMAGEN field for floors
     this.fetchActiveListItems('LM_PLANTAS',[]) // Filter plantas by the user's division, 
    ]);

    return {
      plantas: plantasData,
      pisos: pisosData,
      usuarios: usuariosData
    };
  }

  public async getPisosWithPlantaId(): Promise<IDropdownOption[]> {
    const pisosApiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('LM_PISOS')/items?$select=Id,Title,activo,IMAGEN,PLANTA/Id,PLANTA/Title&$expand=PLANTA&$filter=activo eq 1`; // Added PLANTA/Title
    try {
      const response = await this.context.spHttpClient.get(pisosApiUrl, SPHttpClient.configurations.v1);
      if (response.ok) {
        const pisosJson = await response.json();
        if (pisosJson.value) {
          return pisosJson.value.map((item: ISPPisoItem) => { // ISPPisoItem now includes PLANTA.Title
            let imageUrl = item.IMAGEN;
            if (imageUrl && imageUrl.startsWith('/')) {
              imageUrl = `${this.context.pageContext.web.absoluteUrl}${imageUrl}`;
            }
            const data = {
              ...item,
              plantaId: item.PLANTA ? item.PLANTA.Id : null,
              plantaTitle: item.PLANTA ? item.PLANTA.Title : 'N/A', // Added plantaTitle
              IMAGEN: imageUrl // Store the resolved image URL
            };
            return {
              key: item.Id,
              text: item.Title,
              data: data
            };
          });
        }
      } else {
        const errorData = await response.json();
        console.error('Error fetching pisos list with plantaId:', errorData);
        throw new Error('Could not fetch pisos with plantaId');
      }
      return [];
    } catch (error) {
      console.error('Exception while getting pisos with plantaId:', error);
      throw error;
    }
  }

  /**
   * Fetches the clickable hotspots for a given floor.
   * @param floorId The ID of the floor.
   * @returns A promise that resolves to an array of hotspot items.
   */
  public async getHotspotsForFloor(floorId: number): Promise<IHotspot[]> {
    const listName = 'LM_PUESTOSPISO';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,COORDENADA,PISOId,CAPACIDAD&$filter=TIPO eq 'SalaReunion' and PISOId eq ${floorId}`;
    
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);

      if (response.ok) {
        const data = await response.json();
        return data.value || [];
      } else {
        const errorData = await response.json();
        console.error(`Error fetching list ${listName}:`, errorData);
        throw new Error(`Could not fetch hotspots for floor ${floorId}`);
      }
    } catch (error) {
      console.error(`Exception while fetching hotspots for floor ${floorId}:`, error);
      throw error;
    }
  }

  

  /**
   * Fetches room data for a given floor and maps it to the ISala interface.
   * @param pisoId The ID of the floor.
   * @returns A promise that resolves to an array of ISala objects.
   */
  public async getSalasByPisoPlanta(pisoId: number, fecha: Date, plantaId: number): Promise<ISala[]> {
    try {
      const hotspots: IHotspot[] = await this.getHotspotsForFloor(pisoId);
      
      const promesasDeSalas = hotspots.map(async (hotspot: IHotspot) => {
          const coordenadaString = hotspot.COORDENADA || "0,0";
          const [x, y] = coordenadaString.split(',').map(coord => parseInt(coord.trim(), 10));


          const sala: ISala = {
            ID: hotspot.Id,
            Nombre: hotspot.Title,
            PisoID: hotspot.PISOId,
            PosicionX: isNaN(x) ? 0 : x,
            PosicionY: isNaN(y) ? 0 : y,
            CAPACIDAD: hotspot.CAPACIDAD,
            Disponibilidad: 'full'
          };


          const horasDeSala = await this.getScheduleForSala(sala, pisoId, plantaId, fecha); // Assuming plantaId is 1 for now
          let disponibilidad = 'full'
          const horasTotales = horasDeSala.length;
          //const horasConResarve = horasDeSala.filter(res => res.Disponibilidad === false).length;
          const horasLibres = horasDeSala.filter(res => res.Disponibilidad === true).length;
          
          if(horasLibres === 0){
            disponibilidad = "full";
          } else if(horasLibres === horasTotales){
            disponibilidad = "empty";
          } else {
            disponibilidad = "partial";
          }
          sala.Disponibilidad = disponibilidad;

          return sala;
        });


      return Promise.all(promesasDeSalas);
    } catch (error) {
      console.error(`Exception while getting salas for piso ${pisoId}:`, error);
      throw error; // Re-throw to be handled by the component
    }
  }

  public async getDisponibilidadSalas(pisoId: number, fecha: Date, plantaId: number): Promise<ISala[]> {
    try {
      const hotspots: IHotspot[] = await this.getHotspotsForFloor(pisoId);
      
      const promesasDeSalas = hotspots.map(async (hotspot: IHotspot) => {
          const coordenadaString = hotspot.COORDENADA || "0,0";
          const [x, y] = coordenadaString.split(',').map(coord => parseInt(coord.trim(), 10));

        const sala: ISala = {
          ID: hotspot.Id,
          Nombre: hotspot.Title,
          PisoID: hotspot.PISOId,
          PosicionX: isNaN(x) ? 0 : x,
          PosicionY: isNaN(y) ? 0 : y,
          CAPACIDAD: hotspot.CAPACIDAD,
          Disponibilidad: "full"
        }

          const horasDeSala = await this.getScheduleForSala(sala, pisoId, plantaId, fecha); // Assuming plantaId is 1 for now
          let disponibilidad = 'full'
          const horasTotales = horasDeSala.length;
          //const horasConResarve = horasDeSala.filter(res => res.Disponibilidad === false).length;
          const horasLibres = horasDeSala.filter(res => res.Disponibilidad === true).length;
          
          if(horasLibres === 0){
            disponibilidad = "full";
          } else if(horasLibres === horasTotales){
            disponibilidad = "empty";
          } else {
            disponibilidad = "partial";
          }
          sala.Disponibilidad = disponibilidad;
          
          return sala;
        });

      return Promise.all(promesasDeSalas);
    } 
    catch (error) {
      console.error(`Exception while getting salas for piso ${pisoId}:`, error);
      throw error; // Re-throw to be handled by the component
    }
  }



  /**
   * Fetches available hours from 'LM_Horario' list, filtered by PISO, PLANTA, and Day of Week.
   * @param pisoId The ID of the selected floor (PISO).
   * @param plantaId The ID of the selected building (PLANTA).
   * @param dayOfWeek The day of the week (e.g., "Lunes", "Martes").
   * @returns A promise that resolves to an array of IHorarioItem.
   */
  public async getHorarios(pisoId: number, plantaId: number, dayOfWeek: string): Promise<IHorarioItem[]> {
    const listName = 'LM_Horario';
    // Assuming the lookup field for Planta is named 'PlantaId'
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=*,PISO/Title,PLANTA/Title,HORAS/Title,HORAS/ID,HORAS&$expand=PISO,PLANTA,HORAS&$filter=PISOId eq ${pisoId} and PLANTAId eq ${plantaId}`;// and Title eq '${dayOfWeek}'`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      
      if (response.ok) {
        const data = await response.json();
        return data.value || [];
      } else {
        const errorData = await response.json();
        console.error(`Error fetching horarios for piso ${pisoId}, planta ${plantaId}, day ${dayOfWeek}:`, errorData);
        throw new Error(`Could not fetch horarios`);
      }
    } catch (error) {
      console.error(`Exception while fetching horarios:`, error);
      throw error;
    }
  }


public async ensureUserInGroup(groupName: string, userEmail: string): Promise<any> {
  try {
    // 1. Obtenemos el grupo por su nombre
    const members = await `${this.context.pageContext.web.absoluteUrl}/_api/web/sitegroups/getByName('${groupName}')/users?$filter=Email eq '${userEmail}'`;
    
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(members, SPHttpClient.configurations.v1);

    if (response.ok) {
      const data = await response.json();
      return data; // Si el array tiene elementos, el usuario ya está en el grupo

    }
    

  } catch (error) {
    console.error("Error en ensureUserInGroup:", error);
    throw error; 
  }
}


public async removeUserFromGroup(groupName: string, userId: number): Promise<boolean> {
  try {
    //const encodedLogin = encodeURIComponent(userId);
    const removeUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/sitegroups/getbyname('${groupName}')/users/removebyid(${userId})`;
    
    const spOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'X-HTTP-Method': 'DELETE' // Algunas versiones de SP prefieren el override de método
      }
    };

    const addRes = await this.context.spHttpClient.post(removeUrl, SPHttpClient.configurations.v1, spOptions);

    if (!addRes.ok) {
      const error = await addRes.json();
      //throw new Error(error.error.message.value);
      console.error(error.error.message.value);
      return false; // Si no se pudo agregar, devolvemos false para no bloquear el acceso
    }

    
    return true; // Si no se pudo verificar, devolvemos false para no bloquear el acceso


  } catch (error) {
    console.error("Error en addUserInGroup:", error);
    return false; // En caso de error, devolvemos false para no bloquear el acceso
  }
}

public async addUserInGroup(groupName: string, userEmail: string): Promise<any> {
  try {
    const addUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/sitegroups/getbyname('${groupName}')/users`;

    const addOptions = {
        headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
        },
        body: JSON.stringify({ LoginName: userEmail })
    };

    const addRes = await this.context.spHttpClient.post(addUrl, SPHttpClient.configurations.v1, addOptions);

    if (!addRes.ok) {
      const error = await addRes.json();
      console.error(error.error.message.value); 
      throw new Error(error.error.message.value);
    }
    return await addRes.json(); // Devolvemos la respuesta completa para que el caller pueda verificar el resultado
  } catch (error) {
    console.error("Error en addUserInGroup:", error);
    throw error; // En caso de error, lanzamos el error para que el caller pueda manejarlo
  }
}

  /**
   * Checks if the current user is a member of the specified SharePoint group.
   * @param groupName The name of the SharePoint group.
   * @returns A promise that resolves to true if the user is in the group, false otherwise.
   */
  public async isCurrentUserInGroup(groupName: string): Promise<boolean> {
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser?$expand=groups`;
    
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      
      if (response.ok) {
        const user = await response.json();
        const groups: { Title: string }[] = user.Groups;
        return groups.some(group => group.Title === groupName);
      } else {
        console.error(`Failed to get user groups: ${response.statusText}`);
        return false;
      }
    } catch (error) {
      console.error(`Exception while checking user group membership for group ${groupName}:`, error);
      return false;
    }
  }

   public async deleteReservasParaSala(id: number): Promise<void> {

    const listName = 'LO_PUESTORESERVADO';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'odata-version': '3.0',
        'IF-MATCH': '*', // Required for delete
        'X-HTTP-Method': 'DELETE' // Required for delete
      }
    };

    try {

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (!response.ok) {
        const errorData = await response.json();
        console.error(`Error deleting reserva:`, errorData);
        throw new Error(`Could not delete reserva. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while deleting reserva:`, error);
      throw error;
    }
   };

  public async getReservasParaSala(salaId: number, fecha: Date): Promise<IExistingReservation[]> {
    const listName = 'LO_PUESTORESERVADO';

    const fechaInicio = new Date(fecha);
    fechaInicio.setHours(0, 0, 0, 0);

    const fechaFin = new Date(fecha);
    fechaFin.setHours(23, 59, 59, 999);

    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?` +
      `$select=HORARESERVADA,USUARIO/Title,USUARIO/Id&$expand=USUARIO` +
      `&$filter=PUESTOId eq ${salaId} and FECHAINICIORESERVA ge '${fechaInicio.toISOString()}' and FECHAINICIORESERVA le '${fechaFin.toISOString()}'`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);

      if (response.ok) {
        const data = await response.json();
        return data.value || [];
      } else {
        const errorData = await response.json();
        console.error(`Error fetching reservations for sala ${salaId}:`, errorData);
        throw new Error(`Could not fetch reservations`);
      }
    } catch (error) {
      console.error(`Exception while fetching reservations:`, error);
      throw error;
    }
  }

  /**
   * Creates a new reservation item by constructing the item and calling addReserva.
   * This method contains the business logic for creating a reservation.
   * @param params The raw data needed to create the reservation.
   */
  public async createReservation(params: {
    pisoId: number;
    salaId: number;
    usuarioId: number;
    selectedDate: Date;
    selectedHorarios: IHorarioSala[];
    attendees: { id?: string }[];
  }): Promise<IReservationSPItem> {

    const { pisoId, salaId, usuarioId, selectedDate, selectedHorarios, attendees } = params;

    // 1. Sort hours and determine start and end times
    const sortedHoras = selectedHorarios.map(h => h.HORA).sort((a, b) => a.localeCompare(b));
    const primerHorarioStr = sortedHoras[0]; // e.g., "09:00 - 09:30"
    const ultimoHorarioStr = sortedHoras[sortedHoras.length - 1]; // e.g., "10:00 - 10:30"

    const [startHours, startMinutes] = primerHorarioStr.split(' ')[0].split(':').map(Number);
    const [endHours, endMinutes] = ultimoHorarioStr.split(' ')[0].split(':').map(Number);

    const fechaComienzoReserva = new Date(selectedDate);
    fechaComienzoReserva.setHours(startHours, startMinutes, 0, 0);

    const fechaTerminoReserva = new Date(selectedDate);
    fechaTerminoReserva.setHours(endHours, endMinutes, 0, 0);

    // 2. Calculate check-in deadline (20 mins after start)
    const fechaLimiteCheck = new Date(fechaComienzoReserva);
    fechaLimiteCheck.setMinutes(fechaLimiteCheck.getMinutes() + 20);

    // 3. Process attendees
    const attendeeIds = attendees
    ? attendees
      //  .filter(person => person.id)
        .map(person =>  person.id as string)
      : [];

    const asistentesIds =attendeeIds;// { results: attendeeIds };

    // 4. Construct the SharePoint item
    const newItem: IReservationSPItem = {
      PISOId: pisoId,
      PUESTOId: salaId,
      USUARIOId: usuarioId,
      FECHAINICIORESERVA: fechaComienzoReserva.toISOString(),
      FECHATERMINORESERVA: fechaTerminoReserva.toISOString(),
      ESTADO: "RESERVADO",
      HORARESERVADA: sortedHoras.join(','),
      FECHACOMIENZORESERVA: fechaComienzoReserva.toISOString(),
      FECHALIMITECHECK: fechaLimiteCheck.toISOString(),
      USUARIORESERVAId: this.context.pageContext.legacyPageContext.userId,
      puestoid: salaId//,
      //ASISTENTESId: asistentesIds
    };

    //const sp = spfi(this.context.pageContext.web.absoluteUrl).using(SPFx(this.context));

    if (asistentesIds.length > 0) {
      const webBase = Web(this.context.pageContext.web.absoluteUrl);
      //const userRes = await sp.web.ensureUser(user.loginName);
      //const idNumerico = userRes.data.Id;
      const idsPromesas = asistentesIds.map(async (usuario) => {
          // Aseguramos al usuario en el sitio para obtener su ID numérico
          const result = await webBase.ensureUser(usuario);
          return result.data.Id;
      });

      const idsNumericos = await Promise.all(idsPromesas);
      newItem.ASISTENTESId = idsNumericos;
    }

    // 5. Call the private method to add the item to the list
    return this.addReserva(newItem);
  }

  public async blockRoom(reservationData: IReservationSPItem): Promise<IReservationSPItem> {
    return this.addReserva(reservationData);
  }

  /**
   * Creates a new reservation item in the 'LO_PUESTORESERVADO' list.
   * @param reservationData The data for the new reservation.
   */
  private async addReserva(reservationData: IReservationSPItem): Promise<IReservationSPItem> {
    const listName = 'LO_PUESTORESERVADO';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

    const spHttpClientOptions = {
        headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
        },
        body: JSON.stringify(reservationData)
    };

    try {
        const response: SPHttpClientResponse = await this.context.spHttpClient.post(
            apiUrl,
            SPHttpClient.configurations.v1,
            spHttpClientOptions
        );

        if (response.ok) {
            const addedItem = await response.json();
            return addedItem;
        } else {
            const errorData = await response.json();
            console.error(`Error creating item in ${listName}:`, errorData);
            throw new Error(`Could not create reservation. Status: ${response.statusText}`);
        }
    } catch (error) {
        console.error(`Exception while creating item in ${listName}:`, error);
        throw error;
    }
  }

  public async getAvailableRoomsByPiso(pisoId: number, pisoName: string, plantaId: number, date: Date): Promise<IAvailableRoom[]> {
    const dayOfWeek = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'][date.getDay()];
    const now = new Date();
    const isToday = date.getFullYear() === now.getFullYear() &&
                    date.getMonth() === now.getMonth() &&
                    date.getDate() === now.getDate();

    try {
      const [salas, horarioItems] = await Promise.all([
        this.getSalasByPisoPlanta(pisoId, date,plantaId),
        this.getHorarios(pisoId, plantaId, dayOfWeek)
      ]);

      if (!horarioItems || horarioItems.length === 0 || !horarioItems[0].HORAS) {
        return []; // No schedule defined for this day/floor/plant
      }

      const allPossibleHoras = horarioItems[0].HORAS.sort((a, b) => a.Title.localeCompare(b.Title));

      const availableRoomsPromises = salas.map(async (sala) => {
        const reservations = await this.getReservasParaSala(sala.ID, date);
        const reservedHours = new Set<string>();
        reservations.forEach(res => {
          if (res.HORARESERVADA) {
            res.HORARESERVADA.split(',').forEach(hora => reservedHours.add(hora));
          }
        });

        const availableHours = allPossibleHoras.filter(hora => {
          if (reservedHours.has(hora.Title)) {
            return false;
          }
          if (isToday) {
            const [startHours, startMinutes] = hora.Title.split(' ')[0].split(':').map(Number);
            const slotTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), startHours, startMinutes, 0, 0);
            return slotTime >= now;
          }
          return true;
        });

        if (availableHours.length > 0) {
          return {
            pisoId: pisoId,
            pisoName: pisoName,
            salaId: sala.ID,
            salaName: sala.Nombre,
            CAPACIDAD: sala.CAPACIDAD || 0,
            availableHours: availableHours.map(h => h.Title).join(', '),
          };
        }
        return null;
      });

      const results = await Promise.all(availableRoomsPromises);
      return results.filter((room): room is IAvailableRoom => room !== null);

    } catch (error) {
      console.error(`Error getting available rooms for piso ${pisoId}:`, error);
      return []; // Return empty array on error
    }
  }

  public async getScheduleForSala(sala: ISala, pisoId: number, plantaId: number, date: Date): Promise<IHorarioSala[]> {
    console.log(`[sp.ts] getScheduleForSala called with:`, { salaId: sala.ID, pisoId, plantaId, date });

    const getDayOfWeekString = (d: Date): string => {
      return ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'][d.getDay()];
    };

    const dayOfWeek = getDayOfWeekString(date);
    console.log(`[sp.ts] Calculated day of week: ${dayOfWeek}`);
    const now = new Date();
    const isToday = date.getFullYear() === now.getFullYear() &&
                    date.getMonth() === now.getMonth() &&
                    date.getDate() === now.getDate();

    try {
      const [horarioItems, reservations] = await Promise.all([
        this.getHorarios(pisoId, plantaId, dayOfWeek),
        this.getReservasParaSala(sala.ID, date)
      ]);
      console.log(`[sp.ts] Fetched horarioItems:`, horarioItems);
      console.log(`[sp.ts] Fetched reservations:`, reservations);

      const reservationMap = new Map<string, string>();
      reservations.forEach(reserva => {
        if (reserva.HORARESERVADA && reserva.USUARIO) {
          const horas = reserva.HORARESERVADA.split(',');
          horas.forEach(hora => {
            reservationMap.set(hora, reserva.USUARIO.Title);
          });
        }
      });
     // console.log(`[sp.ts] Created reservationMap:`, reservationMap);
      //console.log(`[sp.ts] Created horarioItems[0].HORAS:`, horarioItems[0].HORAS);
      let mappedHorarios: IHorarioSala[] = [];

      if (horarioItems.length > 0 && horarioItems[0].HORAS && horarioItems[0].HORAS.length > 0) {
        mappedHorarios = horarioItems[0].HORAS.map((horaLookup: { ID: number; Title: string; }) => {
          const isReserved = reservationMap.has(horaLookup.Title);
          
          let isPast = false;
          if (isToday) {
            const [startHours, startMinutes] = horaLookup.Title.split(' ')[0].split(':').map(Number);
            const slotTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), startHours, startMinutes, 0, 0);
            if (slotTime < now) {
              isPast = true;
            }
          }

          if (isReserved) {
            return {
              ID: `${sala.ID}-${horaLookup.ID}`,
              IDSala: sala.ID,
              HORA: horaLookup.Title,
              Disponibilidad: false,
              ReservadoPara: reservationMap.get(horaLookup.Title),
              statusText: strings.ReservedStatus,
              reservadoPorText: reservationMap.get(horaLookup.Title) || '',
              statusColor: 'red'
            };
          } else if (isPast) {
            return {
              ID: `${sala.ID}-${horaLookup.ID}`,
              IDSala: sala.ID,
              HORA: horaLookup.Title,
              Disponibilidad: false,
              ReservadoPara: strings.NoReservationStatus,
              statusText: strings.NotAvailableStatus,
              reservadoPorText: strings.NoReservationStatus,
              statusColor: 'black'
            };
          } else { // Available
            return {
              ID: `${sala.ID}-${horaLookup.ID}`,
              IDSala: sala.ID,
              HORA: horaLookup.Title,
              Disponibilidad: true,
              ReservadoPara: undefined,
              statusText: strings.AvailableStatus,
              reservadoPorText: strings.NAStatus,
              statusColor: 'green'
            };
          }
        });

        // Sort the horarios by time before setting the state
        mappedHorarios.sort((a, b) => {
          const timeA = a.HORA.split(' - ')[0];
          const timeB = b.HORA.split(' - ')[0];
          return timeA.localeCompare(timeB);
        });
      }
      console.log(`[sp.ts] Final mappedHorarios:`, mappedHorarios);
      return mappedHorarios;
    } catch (error) {
      console.error('[sp.ts] CRITICAL ERROR in getScheduleForSala:', error);
      throw error; // Re-throw error to be caught by the calling component
    }
  }

  public async getReportData(filters: { startDate?: Date, endDate?: Date, pisoId?: number, usuarioId?: number }): Promise<IReservationReportItem[]> { // Return raw SP items
    const listName = 'LO_PUESTORESERVADO';
    const filterParts: string[] = [];
    
    if (filters.startDate) {
        filterParts.push(`FECHAINICIORESERVA ge '${filters.startDate.toISOString()}'`);
    }
    if (filters.endDate) {
        const adjustedEndDate = new Date(filters.endDate);
        adjustedEndDate.setHours(23, 59, 59, 999); // Include the whole end day
        filterParts.push(`FECHAINICIORESERVA le '${adjustedEndDate.toISOString()}'`);
    }
    if (filters.pisoId) {
        filterParts.push(`PISOId eq ${filters.pisoId}`);
    }

    if (filters.usuarioId) {
        filterParts.push(`USUARIOId eq ${filters.usuarioId}`);
    }

    const filterQuery = filterParts.join(' and ');

    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?` +
      `$select=Id,FECHAINICIORESERVA,FECHATERMINORESERVA,PISO/Title,PISO/IMAGEN,PISO/ID,ESTADO,HORARESERVADA,FECHASALIDA,FECHAENTRADA,PUESTO/Title,PUESTO/COORDENADA,PUESTO/ID,USUARIO/Title,USUARIO/ID` +
      `&$expand=PISO,PUESTO,USUARIO` +
      (filterQuery ? `&$filter=${filterQuery}` : '') +
      `&$orderby=FECHAINICIORESERVA desc&$top=500`; // Add a $top to prevent massive data returns

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);

      if (response.ok) {
        const data = await response.json();
        if (data.value) {
            return data.value.map((item: IReservationReportItem): IReservationReportItem => ({ // Map to IReservationReportItem
                Id: item.Id,
                FECHAINICIORESERVA: item.FECHAINICIORESERVA,
                FECHATERMINORESERVA: item.FECHATERMINORESERVA,
                PISO: item.PISO,
                ESTADO: item.ESTADO,
                HORARESERVADA: item.HORARESERVADA,
                FECHASALIDA: item.FECHASALIDA,
                FECHAENTRADA: item.FECHAENTRADA,
                PUESTO: item.PUESTO,
                USUARIO: item.USUARIO
            }));
        }
        return [];
      } else {
        const errorData = await response.json();
        console.error(`Error fetching report data:`, errorData);
        throw new Error(`Could not fetch report data`);
      }
    } catch (error) {
      console.error(`Exception while fetching report data:`, error);
      throw error;
    }
  }
  public async getSalasByPiso(pisoId: number): Promise<ISalaItem[]> {

    const listName = 'LM_PUESTOSPISO';
    let filter = `TIPO eq 'SalaReunion' and activo eq 1`;
    if (pisoId) {
      filter += ` and PISOId eq ${pisoId}`;
    }

    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,COORDENADA,PISO/Id,PISO/Title,CAPACIDAD,activo,PLANTA/Title,PLANTA/Id&$filter=${filter}&$expand=PLANTA,PISO`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if (response.ok) {
        const data = await response.json();
        if (data.value) {
          return data.value.map((item: ISPSalaItem) => ({
            Id: item.Id,
            Title: item.Title,
            activo: item.activo,
            PISOId: item.PISO?.Id,
            PISOTitle: item.PISO?.Title,
            PLANTAId: item.PLANTA?.Id,
            PLANTATitle: item.PLANTA?.Title,
            CAPACIDAD: item.CAPACIDAD,
            COORDENADA: item.COORDENADA
          })); 
        } else {
          const errorData = await response.json();
          console.error(`Error fetching salas:`, errorData);
          throw new Error(`Could not fetch salas`);
          // Return empty array on error
        }
      }
      return []; 
    } catch (error) {
      console.error(`Exception while fetching salas:`, error);
      throw error;
    }
  }

  public async getSalas( filterPlantaId?: number, includeInactive: boolean = false): Promise<ISalaItem[]> {
    const listName = 'LM_PUESTOSPISO';
    let filter = `TIPO eq 'SalaReunion'`;
    if (filterPlantaId) {
      filter += ` and PLANTAId eq ${filterPlantaId}`;
    }
    if (!includeInactive) {
      filter += ` and activo eq 1`;
    }

    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,COORDENADA,PISO/Id,PISO/Title,CAPACIDAD,activo,PLANTA/Title,PLANTA/Id&$filter=${filter}&$expand=PLANTA,PISO`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if (response.ok) {
        const data = await response.json();
        if (data.value) {
          return data.value.map((item: ISPSalaItem) => ({
            Id: item.Id,
            Title: item.Title,
            activo: item.activo,
            PISOId: item.PISO?.Id,
            PISOTitle: item.PISO?.Title,
            PLANTAId: item.PLANTA?.Id,
            PLANTATitle: item.PLANTA?.Title,
            CAPACIDAD: item.CAPACIDAD,
            COORDENADA: item.COORDENADA
          })); 
        } else {
          const errorData = await response.json();
          console.error(`Error fetching salas:`, errorData);
          throw new Error(`Could not fetch salas`);
          // Return empty array on error
        }
      }
      return []; 
    } catch (error) {
      console.error(`Exception while fetching salas:`, error);
      throw error;
    }
  }

  public async createSala(sala: ISalaItem): Promise<ISalaItem> {
    const listName = 'LM_PUESTOSPISO';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify({
        Title: sala.Title,
        COORDENADA: sala.COORDENADA,
        PLANTAId: sala.PLANTAId,
        PISOId: sala.PISOId,
        CAPACIDAD: sala.CAPACIDAD,
        activo: sala.activo,
       // IMAGEN: sala.IMAGEN,
        TIPO: 'SalaReunion' // Ensure it's always created as a SalaReunion
      })
    };

    try {

      const entidadEnUso = await this.getEntityByFilter("LM_PUESTOSPISO", "PISOId eq " + sala.PISOId + ' and Title eq \'' + sala.Title + '\'');

      if (entidadEnUso) {
        throw new Error(`No se puede agregar la Sala porque ya existe.`);
      }


      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.ok) {
        const addedItem = await response.json();
        return addedItem as ISalaItem;
      } else {
        const errorData = await response.json();
        console.error(`Error creating sala:`, errorData);
        throw new Error(`Could not create sala. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while creating sala:`, error);
      throw error;
    }
  }

  public async updateSala(sala: ISalaItem): Promise<ISalaItem> {
    const listName = 'LM_PUESTOSPISO';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${sala.Id})`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*', // Required for update
        'X-HTTP-Method': 'MERGE' // Required for update
      },
      body: JSON.stringify({
        Title: sala.Title,
        COORDENADA: sala.COORDENADA,
        PLANTAId: sala.PLANTAId, // Assuming you want to allow changing the associated PLANTA
        PISOId: sala.PISOId,
        CAPACIDAD: sala.CAPACIDAD,
        activo: sala.activo,
       // IMAGEN: sala.IMAGEN,
        TIPO: 'SalaReunion' // Ensure it remains a SalaReunion
      })
    };

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.ok) {
        // Update operations typically return 204 No Content, so no JSON to parse
        return sala; // Return the updated sala item
      } else {
        const errorData = await response.json();
        console.error(`Error updating sala:`, errorData);
        throw new Error(`Could not update sala. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while updating sala:`, error);
      throw error;
    }
  }

  public async deleteSala(id: number): Promise<void> {

    const listName = 'LM_PUESTOSPISO';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'odata-version': '3.0',
        'IF-MATCH': '*', // Required for delete
        'X-HTTP-Method': 'DELETE' // Required for delete
      }
    };

    try {

      const entidadEnUso = await this.getEntityByFilter("LO_PUESTORESERVADO", "PUESTOId eq " + id);

      if (entidadEnUso) {
        throw new Error(`tiene reservas asociadas.`);
      }

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (!response.ok) {
        const errorData = await response.json();
        console.error(`Error deleting sala:`, errorData);
        throw new Error(`Could not delete sala. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while deleting sala:`, error);
      throw error;
    }
  }

  /**
   * Fetches all items from the LM_PISOS list, expanding the PLANTA lookup field.
   * @param includeInactive Optional: Whether to include inactive pisos. Defaults to false.
   * @returns A promise that resolves to an array of IPisoItem.
   */
  public async getPisos(includeInactive: boolean = false, filterPlantaId?: number): Promise<IPisoItem[]> {
    const listName = 'LM_PISOS';
    let filter = '&$filter=activo eq 1';
    
    if (filterPlantaId) {
      filter += ` and PLANTAId eq ${filterPlantaId}`;
    }
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,activo,IMAGEN,PLANTA/Id,PLANTA/Title&$expand=PLANTA${filter}`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if (response.ok) {
        const data = await response.json();
        if (data.value) {
          return data.value.map((item: ISPPisoItem) => ({
            Id: item.Id,
            Title: item.Title,
            activo: item.activo,
            IMAGEN: item.IMAGEN,
            PlantaId: item.PLANTA ? item.PLANTA.Id : null,
            PlantaTitle: item.PLANTA ? item.PLANTA.Title : 'N/A',
          }));
        }
        return [];
      } else {
        const errorData = await response.json();
        console.error(`Error fetching pisos:`, errorData);
        throw new Error(`Could not fetch pisos. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while fetching pisos:`, error);
      throw error;
    }
  }

  /**
   * Creates a new item in the LM_PISOS list.
   * @param piso The IPisoItem object to create.
   * @returns A promise that resolves to the created IPisoItem.
   */
  public async createPiso(piso: IPisoItem): Promise<IPisoItem> {
    const listName = 'LM_PISOS';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify({
        Title: piso.Title,
        activo: piso.activo,
        IMAGEN: piso.IMAGEN,
        PLANTAId: piso.PlantaId, // SharePoint expects 'PLANTAId' for lookup field
        idImgPiso: piso.idImgPiso // Campo adicional para la imagen del piso
      })
    };

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.ok) {
        const addedItem = await response.json();
        return {
          ...piso,
          Id: addedItem.Id,
          PlantaTitle: addedItem.PLANTA ? addedItem.PLANTA.Title : piso.PlantaTitle // Update PlantaTitle if available
        };
      } else {
        const errorData = await response.json();
        console.error(`Error creating piso:`, errorData);
        throw new Error(`Could not create piso. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while creating piso:`, error);
      throw error;
    }
  }

  /**
   * Updates an existing item in the LM_PISOS list.
   * @param piso The IPisoItem object to update.
   * @returns A promise that resolves to the updated IPisoItem.
   */
  public async updatePiso(piso: IPisoItem): Promise<IPisoItem> {
    const listName = 'LM_PISOS';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${piso.Id})`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify({
        Title: piso.Title,
        activo: piso.activo,
        IMAGEN: piso.IMAGEN,
        PLANTAId: piso.PlantaId,
        idImgPiso: piso.idImgPiso // Campo adicional para la imagen del piso
      })
    };

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.ok) {
        return piso; // Return the updated item
      } else {
        const errorData = await response.json();
        console.error(`Error updating piso:`, errorData);
        throw new Error(`Could not update piso. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while updating piso:`, error);
      throw error;
    }
  }

  /**
   * Deletes an item from the LM_PISOS list.
   * @param id The ID of the piso to delete.
   * @returns A promise that resolves when the item is deleted.
   */
  public async deletePiso(id: number): Promise<void> {

    const listName = 'LM_PISOS';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'OData-Version': '3.0',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
      }
    };

    try {

      const entidadEnUso = await this.getEntityByFilter("LM_PUESTOSPISO", "PISOId eq " + id);

      if (entidadEnUso) {
        throw new Error(`tiene salas asociadas.`);
      }


      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (!response.ok) {
        const errorData = await response.json();
        console.error(`Error deleting piso:`, errorData);
        throw new Error(`Could not delete piso. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while deleting piso:`, error);
      throw error;
    }
  }

  /**
   * Fetches all items from the LM_PlanificacionHoras list.
   * @returns A promise that resolves to an array of IDropdownOption.
   */
  public async getPlanificacionHoras(): Promise<IDropdownOption[]> {
    const listName = 'LM_PlanificacionHoras';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title&$orderby=Title asc`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);

      if (response.ok) {
        const data = await response.json();
        if (data.value) {
          return data.value.map((item: ISPItem) => ({
            key: item.Id,
            text: item.Title,
          }));
        }
        return [];
      } else {
        const errorData = await response.json();
        console.error(`Error fetching ${listName}:`, errorData);
        throw new Error(`Could not fetch ${listName}`);
      }
    } catch (error) {
      console.error(`Exception while fetching ${listName}:`, error);
      throw error;
    }
  }

  /**
   * Creates a new item in the LM_Horario list.
   * @param pisoId The ID of the associated Piso.
   * @param plantaId The ID of the associated Planta.
   * @param horaTitle The title of the hour (e.g., "09:00 - 10:00").
   * @returns A promise that resolves to the created item.
   */
  public async createHorario(pisoId: number, plantaId: number, horaTitle: string): Promise<any> {
    const listName = 'LM_Horario';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify({
        Title: horaTitle,
        PISOId: pisoId,
        PLANTAId: plantaId,
      })
    };

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.ok) {
        return await response.json();
      } else {
        const errorData = await response.json();
        console.error(`Error creating horario:`, errorData);
        throw new Error(`Could not create horario. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while creating horario:`, error);
      throw error;
    }
  }

  /**
   * Fetches all LM_Horario items associated with a specific Piso.
   * @param pisoId The ID of the Piso.
   * @returns A promise that resolves to an array of LM_Horario items.
   */
  public async getHorariosForPiso(pisoId?: number, plantaId?: number): Promise<IHorarioItem | undefined> {
    if (pisoId === undefined || plantaId === undefined) {
      return undefined; // If no pisoId or plantaId, return undefined
    }

    const listName = 'LM_Horario';
    // Assuming a default day of the week for general hours, e.g., "Lunes"
    //const defaultDayOfWeek = "Lunes"; 
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,PISOId,PLANTAId,HORAS/ID,HORAS/Title&$expand=HORAS&$filter=PISOId eq ${pisoId} and PLANTAId eq ${plantaId}`;// and Title eq '${defaultDayOfWeek}'`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);

      if (response.ok) {
        const data = await response.json();
        if (data.value && data.value.length > 0) {
          console.log("Raw LM_Horario entry from SharePoint:", data.value[0]); // Log the raw data
          return data.value[0] as IHorarioItem;
        }
        return undefined;
      } else {
        const errorData = await response.json();
        console.error(`Error fetching horarios for piso ${pisoId}, planta ${plantaId}:`, errorData);
        throw new Error(`Could not fetch horarios for piso. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while fetching horarios for piso ${pisoId}, planta ${plantaId}:`, error);
      throw error;
    }
  }

  /**
   * Creates or updates an LM_Horario item for a given Piso and Planta.
   * @param pisoId The ID of the associated Piso.
   * @param plantaId The ID of the associated Planta.
   * @param selectedHoraIds An array of IDs from LM_PlanificacionHoras.
   * @returns A promise that resolves to the created or updated item.
   *  const dayOfWeek = [ 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes'];
   */
  /*
  public async createOrUpdateHorarioEntry(pisoId: number, plantaId: number, selectedHoraIds: number[]): Promise<IHorarioItem> {
    const listName = 'LM_Horario';
    const diasSemana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"];
  
    const defaultDayOfWeek = "Martes"; // Assuming a default day for general hours

    const existingHorario = await this.getHorariosForPiso(pisoId, plantaId);

    const itemPayload = {
      Title: defaultDayOfWeek,
      PISOId: pisoId,
      PLANTAId: plantaId,
      HORASId: selectedHoraIds//{ results: selectedHoraIds } // Correct format for Multi-lookup field
    };

    if (existingHorario) {
      // Update existing item
      const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${existingHorario.Id})`;
      const spHttpClientOptions = {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': '',
          'IF-MATCH': '*', // Required for update
          'X-HTTP-Method': 'MERGE' // Required for update
        },
        body: JSON.stringify(itemPayload)
      };

      try {
        const response: SPHttpClientResponse = await this.context.spHttpClient.post(
          apiUrl,
          SPHttpClient.configurations.v1,
          spHttpClientOptions
        );

        if (response.ok) {
          return { ...existingHorario, HORAS: selectedHoraIds.map(id => ({ ID: id, Title: '' })) }; // Return updated item
        } else {
          const errorData = await response.json();
          console.error(`Error updating horario entry:`, errorData);
          throw new Error(`Could not update horario entry. Status: ${response.statusText}`);
        }
      } catch (error) {
        console.error(`Exception while updating horario entry:`, error);
        throw error;
      }
    } else {
      // Create new item
      const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
      const spHttpClientOptions = {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: JSON.stringify(itemPayload)
      };

      try {
        const response: SPHttpClientResponse = await this.context.spHttpClient.post(
          apiUrl,
          SPHttpClient.configurations.v1,
          spHttpClientOptions
        );

        if (response.ok) {
          const addedItem = await response.json();
          return { ...addedItem, HORAS: selectedHoraIds.map(id => ({ ID: id, Title: '' })) }; // Return created item
        } else {
          const errorData = await response.json();
          console.error(`Error creating horario entry:`, errorData);
          throw new Error(`Could not create horario entry. Status: ${response.statusText}`);
        }
      } catch (error) {
        console.error(`Exception while creating horario entry:`, error);
        throw error;
      }
    }
  }

*/

public async createOrUpdateHorarioEntry(pisoId: number, plantaId: number, selectedHoraIds: number[]): Promise<void> {

  const listName = 'LM_Horario';

  
  const headers: any = {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
        };

  

  try {
      
      const existingHorario = await this.getHorariosForPiso(pisoId, plantaId);      
      const apiUrl = existingHorario 
        ? `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${existingHorario.Id})`
        : `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
        
      if (existingHorario) {
          headers['IF-MATCH'] = '*';
          headers['X-HTTP-Method'] = 'MERGE';
        
      } 
      this.saveSingleDay(pisoId, plantaId,  selectedHoraIds, apiUrl, headers, existingHorario!);


       
  } catch (error) {
    console.error(`Exception while updating horario entry:`, error);
    throw error;
  }
}

private async saveSingleDay(pisoId: number, 
                            plantaId: number, 
                            selectedHoraIds: number[], 
                            apiUrl: string, 
                            headers: any, 
                            existingHorario: IHorarioItem): Promise<IHorarioItem> {

  const itemPayload = {
        Title: '',
        PISOId: pisoId,
        PLANTAId: plantaId,
        HORASId: selectedHoraIds //{ results: selectedHoraIds } // Correct format for Multi-lookup field
      }; 
  
    const spHttpClientOptions = {
        headers: headers,
        body: JSON.stringify(itemPayload)
      };

      const response: SPHttpClientResponse = await  this.context.spHttpClient.post(
          apiUrl,
          SPHttpClient.configurations.v1,
          spHttpClientOptions
        );
        
       if (response.ok) {
          return { ...existingHorario, HORAS: selectedHoraIds.map(id => ({ ID: id, Title: '' })) }; // Return updated item
        } else {
          const errorData = response.json();
          console.error(`Error updating horario entry:`, errorData);
          throw new Error(`Could not update horario entry. Status: ${response.statusText}`);
        }
}

public async uploadFileWithMetadata(
  listName: string,
  docId: number, 
  metadata: any // Objeto con los nombres de columna y valores
): Promise<void>{

  //const itemEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${folderPath}')/ListItemAllFields`;
  const itemEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${docId})`
  // Paso 3: Actualizar los metadatoslistName
        const updateResponse: SPHttpClientResponse = await this.context.spHttpClient.post(
          itemEndpoint,
          SPHttpClient.configurations.v1,
          {
            body: JSON.stringify(metadata),
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=nometadata',
              'odata-version': '',
              'IF-MATCH': '*',
              'X-HTTP-Method': 'MERGE'
            }
          }
        );

        if (!updateResponse.ok) {
          const error = await updateResponse.json();
          console.error("Error actualizando metadatos:", error);
          throw new Error("El archivo se subió pero los metadatos fallaron.");
        }
      }



  /**
   * Fetches the request digest for SharePoint API calls.
   * @returns A promise that resolves to the X-RequestDigest value.
   */
  private async getRequestDigest(): Promise<string> {
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/contextinfo`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    };

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.ok) {
        const data = await response.json();
        return data.FormDigestValue;
      } else {
        const errorData = await response.json();
        console.error(`Error fetching request digest:`, errorData);
        throw new Error(`Could not fetch request digest. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while fetching request digest:`, error);
      throw error;
    }
  }

  /**
   * Uploads a file to a SharePoint document library.
   * @param fileName The name of the file to upload.
   * @param fileContent The content of the file as an ArrayBuffer.
   * @param folderPath The server-relative URL of the folder to upload to (e.g., 'Shared Documents/Images').
   * @returns A promise that resolves to the server-relative URL of the uploaded file.
   */
  public async uploadFile(fileName: string, fileContent: ArrayBuffer, folderPath: string): Promise<IUploadResult> {
    const serverRelativeUrl = `${this.context.pageContext.web.serverRelativeUrl}/${folderPath}`;
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRelativeUrl}')/Files/add(overwrite=true,url='${fileName}')?$expand=ListItemAllFields`;

    const requestDigest = await this.getRequestDigest(); // Get the request digest

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': '',
        'X-RequestDigest': requestDigest // Add the request digest header
      },
      body: fileContent
    };

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.ok) {
        const data = await response.json();
        // Construct the absolute URL using the site collection's absolute URL
        //const absoluteUrl = `${this.context.pageContext.site.absoluteUrl.split('/sites')[0]}${data.ServerRelativeUrl}`;
        return {
                  Id: data.ListItemAllFields.Id, // SharePoint asocia un ID único a cada archivo subido
                  Url: `${this.context.pageContext.site.absoluteUrl.split('/sites')[0]}${data.ServerRelativeUrl}`   // La ruta del archivo
                };;
      } else {
        const errorData = await response.json();
        console.error(`Error uploading file:`, errorData);
        throw new Error(`Could not upload file. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while uploading file:`, error);
      throw error;
    }
  }




  /**
 * Asegura que toda la estructura de carpetas exista.
 * @param libraryName Nombre de la biblioteca (ej: 'LO_QRPUESTOS')
 * @param folderPath Ruta de subcarpetas (ej: 'Casa Matriz/piso 2')
 * @returns La URL relativa final de la carpeta
 */
public async ensureFolderPath(folderPath: string): Promise<string> {
  // Ejemplo sin PnPjs (usando SPHttpClient nativo)
const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/folders`;

const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': '',
        //'X-RequestDigest': requestDigest // Add the request digest header
        },
        body: JSON.stringify({
          //'__metadata': { 'type': 'SP.Folder' },
          'ServerRelativeUrl': `${this.context.pageContext.web.serverRelativeUrl}/${folderPath}`
        })
      };

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (!response.ok) {
        const errorJson = await response.json();
        // Código de error: "La carpeta ya existe". Si es este, simplemente ignoramos y seguimos.
        if (errorJson.error && errorJson.error.code.indexOf("-2147024713") !== -1) {
          console.log(`La carpeta ${folderPath} ya existe, continuando...`);
          //continue; 
        } else {
          throw new Error(`Error creando ${folderPath}: ${response.statusText}`);
        }
      }

  return folderPath; // Retornamos la ruta completa verificada
}

public async getQR(salaID:number = 1, pisoId:number = 1): Promise<string> {
    //const listName = 'lista';

    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('LO_QRPUESTOS')/items?$filter=PISOId eq ${pisoId} and SALAId eq ${salaID}&$select=File/ServerRelativeUrl&$expand=File`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if (response.ok) {
        const data = await response.json();
        return data.value[0].File.ServerRelativeUrl;
      } else {
        const errorData = await response.json();
        console.error('Error obteniendo el QR', errorData);
        throw new Error(`Error obteniendo el QR. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Error obteniendo el QR:`, error);
      throw error;
    }
  }

  // CRUD operations for LM_VICEPRESIDENCIAS
  public async getVicepresidencias(includeInactive: boolean = false): Promise<IVicepresidenciaItem[]> {
    const listName = 'LM_VICEPRESIDENCIAS';
    let filter = '';
    if (!includeInactive) {
      filter = `&$filter=activo eq 1`;
    }
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,activo${filter}`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if (response.ok) {
        const data = await response.json();
        return data.value as IVicepresidenciaItem[];
      } else {
        const errorData = await response.json();
        console.error(`Error fetching vicepresidencias:`, errorData);
        throw new Error(`Could not fetch vicepresidencias. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while fetching vicepresidencias:`, error);
      throw error;
    }
  }

  public async createVicepresidencia(vicepresidencia: IVicepresidenciaItem): Promise<IVicepresidenciaItem> {
    const listName = 'LM_VICEPRESIDENCIAS';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify({
        Title: vicepresidencia.Title,
        activo: vicepresidencia.activo,
      })
    };

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.ok) {
        const addedItem = await response.json();
        return addedItem as IVicepresidenciaItem;
      } else {
        const errorData = await response.json();
        console.error(`Error creating vicepresidencia:`, errorData);
        throw new Error(`Could not create vicepresidencia. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while creating vicepresidencia:`, error);
      throw error;
    }
  }

  public async updateVicepresidencia(vicepresidencia: IVicepresidenciaItem): Promise<IVicepresidenciaItem> {
    const listName = 'LM_VICEPRESIDENCIAS';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${vicepresidencia.Id})`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify({
        Title: vicepresidencia.Title,
        activo: vicepresidencia.activo,
      })
    };

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.ok) {
        return vicepresidencia;
      } else {
        const errorData = await response.json();
        console.error(`Error updating vicepresidencia:`, errorData);
        throw new Error(`Could not update vicepresidencia. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while updating vicepresidencia:`, error);
      throw error;
    }
  }

  public async deleteVicepresidencia(id: number): Promise<void> {
  
    const listName = 'LM_VICEPRESIDENCIAS';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'OData-Version': '3.0',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
      }
    };

    try {

        
      const entidadEnUso = await this.getEntityByFilter("LM_GERENCIAS", "VICEPRESIDENCIAId eq " + id);

      if (entidadEnUso) {
        throw new Error(`tiene gerencias asociadas.`);
      }


      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (!response.ok) {
        const errorData = await response.json();
        console.error(`Error deleting vicepresidencia:`, errorData);
        throw new Error(`Could not delete vicepresidencia. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while deleting vicepresidencia:`, error);
      throw error;
    }
  }

  // CRUD operations for LM_GERENCIAS
  public async getGerencias(includeInactive: boolean = false): Promise<IGerenciaItem[]> {
    const listName = 'LM_GERENCIAS';
    let filter = '';
    if (!includeInactive) {
      filter = `&$filter=activo eq 1`;
    }
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,activo,VICEPRESIDENCIA/Id,VICEPRESIDENCIA/Title&$expand=VICEPRESIDENCIA${filter}`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if (response.ok) {
        const data = await response.json();
        if (data.value) {
          
          return data.value.map((item: any) => ({
         // return data.value.map((item: ISPGerenciaItem): IGerenciaItem => ({
            Id: item.Id,
            Title: item.Title,
            activo: item.activo,
            VicepresidenciaId: item.VICEPRESIDENCIA ? item.VICEPRESIDENCIA.Id : null,
            VicepresidenciaTitle: item.VICEPRESIDENCIA ? item.VICEPRESIDENCIA.Title : 'N/A',
          }));
        }
        return [];
      } else {
        const errorData = await response.json();
        console.error(`Error fetching gerencias:`, errorData);
        throw new Error(`Could not fetch gerencias. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while fetching gerencias:`, error);
      throw error;
    }
  }

  public async createGerencia(gerencia: IGerenciaItem): Promise<IGerenciaItem> {
    const listName = 'LM_GERENCIAS';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify({
        Title: gerencia.Title,
        activo: gerencia.activo,
        VICEPRESIDENCIAId: gerencia.VicepresidenciaId,
      })
    };

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.ok) {
        const addedItem = await response.json();
        return {
          ...gerencia,
          Id: addedItem.Id,
          Title: addedItem.VICEPRESIDENCIA ? addedItem.VICEPRESIDENCIA.Title : gerencia.VicepresidenciaTitle
        };
      } else {
        const errorData = await response.json();
        console.error(`Error creating gerencia:`, errorData);
        throw new Error(`Could not create gerencia. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while creating gerencia:`, error);
      throw error;
    }
  }

  public async updateGerencia(gerencia: IGerenciaItem): Promise<IGerenciaItem> {
    const listName = 'LM_GERENCIAS';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${gerencia.Id})`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify({
        Title: gerencia.Title,
        activo: gerencia.activo,
        VICEPRESIDENCIAId: gerencia.VicepresidenciaId,
      })
    };

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.ok) {
        return gerencia;
      } else {
        const errorData = await response.json();
        console.error(`Error updating gerencia:`, errorData);
        throw new Error(`Could not update gerencia. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while updating gerencia:`, error);
      throw error;
    }
  }

  public async deleteGerencia(id: number): Promise<void> {

    const listName = 'LM_GERENCIAS';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'OData-Version': '3.0',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
      }
    };

    try {
      const entidadEnUso = await this.getEntityByFilter("LM_USUARIOS", "GERENCIAId eq " + id);

      if (entidadEnUso) {
        throw new Error(`No se puede eliminar la Gerencia porque tiene usuarios asociados.`);
      }

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (!response.ok) {
        const errorData = await response.json();
        console.error(`Error eliminando la Gerencia:`, errorData);
        throw new Error(`no se pudo eliminar la gerencia. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception mientras se eliminaaba la Gerencia:`, error);
      throw error;
    }
  }

  // CRUD operations for LM_USUARIOS
  public async getUsuarios(includeInactive: boolean = false, plantaId?: number): Promise<IUsuarioItem[]> {
    const listName = 'LM_USUARIOS';
    let filter = '';
    /*if (!includeInactive) {
      filter = '`&$filter=activo eq 1`';
    }*/
    if (plantaId !== undefined) {
      filter += `&$filter= divisionId eq ${plantaId}`;
    }
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,activo,esAdmin,USUARIO/Id,USUARIO/Name,USUARIO/Title,USUARIO/EMail,division/Id,division/Title,GERENCIA/Id,GERENCIA/Title&$expand=USUARIO,division,GERENCIA${filter}`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if (response.ok) {
        const data = await response.json();
        if (data.value) {
          return data.value.map((item: ISPUsuarioItem): IUsuarioItem => ({
            Id: item.Id,
            usuarioId: item.USUARIO ? item.USUARIO.Id : 0,
            LoginName: item.USUARIO ? item.USUARIO.Name : '',
            text: item.USUARIO ? item.USUARIO.Title : 'N/A',
            email: item.USUARIO ? item.USUARIO.EMail : 'N/A',
            //VicepresidenciaId: item.VICEPRESIDENCIA ? item.VICEPRESIDENCIA.Id : 0,
            //VicepresidenciaTitle: item.VICEPRESIDENCIA ? item.VICEPRESIDENCIA.Title : 'N/A',
            divisionId: item.division ? item.division.Id : 0,
            divisionTitle: item.division ? item.division.Title : 'N/A',
            GerenciaId: item.GERENCIA ? item.GERENCIA.Id : 0,
            GerenciaTitle: item.GERENCIA ? item.GERENCIA.Title : 'N/A',
            activo: item.activo,
            esAdmin: item.esAdmin,
          }));
        }
        return [];
      } else {
        const errorData = await response.json();
        console.error(`Error fetching usuarios:`, errorData);
        throw new Error(`Could not fetch usuarios. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while fetching usuarios:`, error);
      throw error;
    }
  }

  public async createUsuario(usuario: IUsuarioItem): Promise<IUsuarioItem> {
    const listName = 'LM_USUARIOS';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
    
    const webBase = Web(this.context.pageContext.web.absoluteUrl);
      //const userRes = await sp.web.ensureUser(user.loginName);
      //const idNumerico = userRes.data.Id;
    const idUsuario = await webBase.ensureUser(usuario.LoginName);

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify({
        USUARIOId: idUsuario.data.Id,//usuario.id,
        //VICEPRESIDENCIAId: usuario.VicepresidenciaId,
        divisionId: usuario.divisionId,
        GERENCIAId: usuario.GerenciaId,
        activo: usuario.activo,
        Title: idUsuario.data.Title, 
        esAdmin: usuario.esAdmin,
      })
    };

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.ok) {
        const addedItem = await response.json();
        return {
          ...usuario,
          Id: addedItem.Id,
        };
      } else {
        const errorData = await response.json();

        //const errorMessage = errorData["odata.error"]?.message?.value || "Error desconocido";
        const errorCode = errorData["odata.error"]?.code || "Sin código";
        if(errorCode.includes("DuplicateValues")) {
          
          throw new Error(`usuario duplicado`);
        }
        else{
          throw new Error(`Error creando el usuario. Status: ${response.statusText}`);
        }
      }
    } catch ( error) {
      console.error(`Error mientras se creaaba el usuario:`,error);
      const msg = error instanceof Error ? error.message : String(error);
      throw new Error(`Error mientras se creaba el usuario: ${msg}`);
    }
  }

  public async updateUsuario(usuario: IUsuarioItem): Promise<IUsuarioItem> {
    const listName = 'LM_USUARIOS';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${usuario.Id})`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify({
        USUARIOId: usuario.usuarioId,
        //VICEPRESIDENCIAId: usuario.VicepresidenciaId,
        GERENCIAId: usuario.GerenciaId,
        activo: usuario.activo,
        esAdmin: usuario.esAdmin,
        divisionId: usuario.divisionId,
      })
    };

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.ok) {
        return usuario;
      } else {
        const errorData = await response.json();
        console.error(`Error updating usuario:`, errorData);
        throw new Error(`Could not update usuario. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while updating usuario:`, error);
      throw error;
    }
  }

  public async deleteUsuario(id: number): Promise<void> {
    const listName = 'LM_USUARIOS';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
      }
    };

    try {

      const entidadEnUso = await this.getEntityByFilter("LO_PUESTORESERVADO", "USUARIOId eq " + id);

      if (entidadEnUso) {
        throw new Error(`No se puede eliminar el Usuario porque tiene reservas asociadas.`);
      }


      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (!response.ok) {
        const errorData = await response.json();
        console.error(`Error soft deleting usuario:`, errorData);
        throw new Error(`Could not soft delete usuario. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while soft deleting usuario:`, error);
      throw error;
    }
  }

  // CRUD operations for LM_PLANTAS (referred to as Divisiones)
  public async getDivisiones(includeInactive: boolean = false): Promise<IDivisionItem[]> {
    const listName = 'LM_PLANTAS';
    let filter = '';
    if (!includeInactive) {
      filter = `&$filter=activo eq 1`;
    }
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,activo${filter}`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if (response.ok) {
        const data = await response.json();
        return data.value as IDivisionItem[];
      } else {
        const errorData = await response.json();
        console.error(`Error fetching divisiones:`, errorData);
        throw new Error(`Could not fetch divisiones. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while fetching divisiones:`, error);
      throw error;
    }
  }

  public async createDivision(division: IDivisionItem): Promise<IDivisionItem> {
    const listName = 'LM_PLANTAS';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify({
        Title: division.Title,
        activo: division.activo,
      })
    };

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.ok) {
        const addedItem = await response.json();
        return addedItem as IDivisionItem;
      } else {
        const errorData = await response.json();
        console.error(`Error creating division:`, errorData);
        throw new Error(`Could not create division. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while creating division:`, error);
      throw error;
    }
  }

  public async updateDivision(division: IDivisionItem): Promise<IDivisionItem> {
    const listName = 'LM_PLANTAS';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${division.Id})`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify({
        Title: division.Title,
        activo: division.activo,
      })
    };

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.ok) {
        return division;
      } else {
        const errorData = await response.json();
        console.error(`Error updating division:`, errorData);
        throw new Error(`Could not update division. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while updating division:`, error);
      throw error;
    }
  }

  public async deleteDivision(id: number): Promise<void> {

    


    const listName = 'LM_PLANTAS';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'OData-Version': '3.0',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
      }
    };

    try {

      const entidadEnUso = await this.getEntityByFilter("LM_PISOS", "PLANTAId eq " + id);

      if (entidadEnUso) {
        throw new Error(`No se puede eliminar la división porque tiene pisos asociados.`);
      }

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (!response.ok) {
        const errorData = await response.json();
        console.error(`Error deleting division:`, errorData);
        throw new Error(`Error al eliminar la división: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Excepcion mientras se eliminaba la división:`, error);
      throw error;
    }
  }


  public async getEntityByFilter(listName:string = '', filtro:string = ''): Promise<boolean> {
    //const listName = 'lista';
    let filter = filtro? '&$filter=' + filtro : '';
    
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title${filter}`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if (response.ok) {
        const data = await response.json();
        return data.value.length > 0;
      } else {
        const errorData = await response.json();
        console.error(`Error obteniendo datos de ${listName}:`, errorData);
        throw new Error(`no se pudo obtener datos de la lista ${listName}. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Excecione mientras se consultaban la lista ${listName}:`, error);
      throw error;
    }
  }

public async getItemsByDivision<T>(
  listName: string, 
  divisionId?: number, 
  selectFields: string[] = ['*'], 
  expandFields: string[] = []
): Promise<T[]> {
  try {
    // Construimos la query base
    let endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
    
    // Agregamos el filtro por División
    // Asegúrate de que todas tus listas tengan la columna 'DivisionId' o cambia el nombre aquí
    const filter = divisionId? `&$filter=PLANTAId eq ${divisionId}` : '';
    const select = `?$select=${selectFields.join(',')}`;
    const expand = expandFields.length > 0 ? `&$expand=${expandFields.join(',')}` : '';

    //const url = `${endpoint}${filter}${select}${expand}`;
    const url = `${endpoint}${select}${filter}${expand}`;

    const response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (response.ok) {
      const data = await response.json();
      return data.value as T[]; // Aquí es donde ocurre la "magia" del genérico
    } else {
      throw new Error(`Error obteniendo items de la lista ${listName}`);
    }
  } catch (error) {
    console.error("Error en getItemsByDivision:", error);
    return [];
  }
}

  // CRUD operations for LM_PlanificacionHoras
  public async getPlanificacionHorasItems(): Promise<IPlanificacionHoraItem[]> {
    const listName = 'LM_PlanificacionHoras';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if (response.ok) {
        const data = await response.json();
        return data.value as IPlanificacionHoraItem[];
      } else {
        const errorData = await response.json();
        console.error(`Error fetching planificacion horas:`, errorData);
        throw new Error(`Could not fetch planificacion horas. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while fetching planificacion horas:`, error);
      throw error;
    }
  }

  public async createPlanificacionHora(hora: IPlanificacionHoraItem): Promise<IPlanificacionHoraItem> {
    const listName = 'LM_PlanificacionHoras';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify({
        Title: hora.Title,
      })
    };

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.ok) {
        const addedItem = await response.json();
        return addedItem as IPlanificacionHoraItem;
      } else {
        const errorData = await response.json();
        console.error(`Error creating planificacion hora:`, errorData);
        throw new Error(`Could not create planificacion hora. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while creating planificacion hora:`, error);
      throw error;
    }
  }

  public async updatePlanificacionHora(hora: IPlanificacionHoraItem): Promise<IPlanificacionHoraItem> {
    const listName = 'LM_PlanificacionHoras';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${hora.Id})`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify({
        Title: hora.Title,
      })
    };

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.ok) {
        return hora;
      } else {
        const errorData = await response.json();
        console.error(`Error updating planificacion hora:`, errorData);
        throw new Error(`Could not update planificacion hora. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while updating planificacion hora:`, error);
      throw error;
    }
  }

  public async deletePlanificacionHora(id: number): Promise<void> {

    const listName = 'LM_PlanificacionHoras';
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`;

    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'OData-Version': '3.0',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
      }
    };

    try {

      const entidadEnUso = await this.getEntityByFilter("LM_Horario", "HORASId eq " + id);

      if (entidadEnUso) {
        throw new Error(`No se puede eliminar el Horario porque esta asociado a Pisos.`);
      }

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        apiUrl,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (!response.ok) {
        const errorData = await response.json();
        console.error(`Error deleting planificacion hora:`, errorData);
        throw new Error(`Could not delete planificacion hora. Status: ${response.statusText}`);
      }
    } catch (error) {
      console.error(`Exception while deleting planificacion hora:`, error);
      throw error;
    }
  }

/*public async uploadBase64File(
    base64String: string, 
    fileName: string, 
    libraryRelativeUrl: string
  ): Promise<any> {
    try {
      // 1. Limpiar el prefijo del base64 (ej: "data:image/png;base64,") si existe
      const base64Data = base64String.split(',').pop() || "";
      
      // 2. Convertir Base64 a un ArrayBuffer
      const byteCharacters = atob(base64Data);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);

      // 3. Subir a SharePoint
      // Usamos .addChunked para archivos que podrían ser pesados, o .add para normales
      const fileAdded = await sp.web
        .getFolderByServerRelativeUrl(libraryRelativeUrl)
        .files.add(fileName, byteArray, true); // 'true' permite sobrescribir

      return fileAdded.data;
    } catch (error) {
      console.error("Error subiendo archivo Base64:", error);
      throw error;
    }
  } */
}

