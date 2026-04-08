import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { IViewProps } from './IViewProps';
import {
  PrimaryButton,
  DefaultButton,
  Dialog,
  DialogType,
  DialogFooter,
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
  Dropdown
 // IDropdownOption
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { useSPFxContext } from '../../contexts/SPFxContext';
import { SPService } from '../../services/sp';
import { IUsuarioItem, IGerenciaItem, IDivisionItem } from '../models/entities';
import * as strings from 'SrscWebPartStrings';

const MantenedorUsuarios: React.FC<IViewProps> = () => {
  const context = useSPFxContext();
  const spService = React.useMemo(() => new SPService(context), [context]);

  const [usuarios, setUsuarios] = useState<IUsuarioItem[]>([]);
  //const [vicepresidencias, setVicepresidencias] = useState<IDropdownOption[]>([]);
  const [divisiones, setdivisiones] = useState<IDivisionItem[]>([]);
  const [allGerencias, setAllGerencias] = useState<IGerenciaItem[]>([]);
  //const [filteredGerencias, setFilteredGerencias] = useState<IDropdownOption[]>([]);
  
  const [isModalOpen, setIsModalOpen] = useState<boolean>(false);
  const [currentUser, setCurrentUser] = useState<IUsuarioItem | undefined>(undefined);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [errorTitle, setErrorTitle] = useState<string | undefined>(undefined);
  //const [message, setMessage] = useState<string | undefined>(undefined);
 // const [messageType, setMessageType] = useState<MessageBarType>(MessageBarType.info);
  const [showDeleteConfirm, setShowDeleteConfirm] = React.useState<boolean>(false);
  const [usuarioToDelete, setUsuarioToDelete] = React.useState<IUsuarioItem | undefined>(undefined);
  const [message, setMessage] = React.useState<{ type: MessageBarType, text: string } | undefined>(undefined);
  // Form validation state
  const [userError, setUserError] = useState<string | undefined>(undefined);
  //const [vicepresidenciaError, setVicepresidenciaError] = useState<string | undefined>(undefined);
  const [divisionError, setDivisionError] = useState<string | undefined>(undefined);
  const [gerenciaError, setGerenciaError] = useState<string | undefined>(undefined);

  const fetchUsuario = useCallback(async () => {
    setIsLoading(true);
    setErrorTitle(undefined);
    try {
      const [fetchedUsuarios, /*fetchedVicepresidencias,*/fetchdivisiones, fetchedGerencias] = await Promise.all([
        spService.getUsuarios(true), // Fetch all users, including inactive
        //spService.getVicepresidencias(false), // Fetch only active vicepresidencias for dropdown
        spService.getDivisiones(false), // Fetch only active divisiones for dropdown
        spService.getGerencias(false) // Fetch only active gerencias for dropdown
      ]);
      setUsuarios(fetchedUsuarios);
      //setVicepresidencias(fetchedVicepresidencias.map(vp => ({ key: vp.Id!, text: vp.Title })));
      setdivisiones(fetchdivisiones);//.map(vp => ({ key: vp.Id!, text: vp.Title })));
      setAllGerencias(fetchedGerencias);
    } catch (err) {
      //setError(strings.ErrorFetchingData);
      setMessage({ type: MessageBarType.error, text: strings.ErrorFetchingData });
      console.error(err);
    } finally {
      setIsLoading(false);
    }
  }, [spService]);

  useEffect(() => {
    void fetchUsuario();
  }, [fetchUsuario]);

  // Effect for cascading dropdown
  /*useEffect(() => {
    if (currentUser?.VicepresidenciaId && allGerencias.length > 0) {
      const filtered = allGerencias
        .filter(g => g.VicepresidenciaId === currentUser.VicepresidenciaId)
        .map(g => ({ key: g.Id!, text: g.Title }));
      setFilteredGerencias(filtered);
    } else {
      setFilteredGerencias([]);
    }
  }, [currentUser?.VicepresidenciaId, allGerencias]);
*/

  const validateForm = (): boolean => {
    let isValid = true;
    if (!currentUser?.LoginName) {
      setUserError(strings.RequiredField);
      isValid = false;
    } else {
      setUserError(undefined);
    }

    if (!currentUser?.divisionId){//VicepresidenciaId) {
      //setVicepresidenciaError(strings.RequiredField);
      setDivisionError(strings.RequiredField);

      isValid = false;
    } else {
      //setVicepresidenciaError(undefined);
      setDivisionError(undefined);
    }

    if (!currentUser?.GerenciaId) {
        setGerenciaError(strings.RequiredField);
        isValid = false;
    } else {
        setGerenciaError(undefined);
    }
    return isValid;
  };

  const onAddUser = () => {
    setCurrentUser({
      id: "",
      text: '',
      email: '',
      //secondaryText: '',
      usuarioId: 0,
      LoginName: '',
     // VicepresidenciaId: 0,
      divisionId: 0,
      GerenciaId: 0,
      esAdmin: false,
      activo: true,
    } as IUsuarioItem);
    setIsModalOpen(true);
    setErrorTitle(undefined);
    setMessage(undefined);
    setUserError(undefined);
    //setVicepresidenciaError(undefined);
    setDivisionError(undefined);
    setGerenciaError(undefined);
  };

  const onEditUser = (item: IUsuarioItem) => {
    setCurrentUser({ ...item }); // Create a copy to edit
    setIsModalOpen(true);
    setErrorTitle(undefined);
    setMessage(undefined);
    setUserError(undefined);
    //setVicepresidenciaError(undefined);
    setDivisionError(undefined);
    setGerenciaError(undefined);
  };
/*
  const onDeleteUser = async (id: number, title: string): Promise<void> => {
    if (window.confirm(strings.ConfirmDeleteUser.replace('{0}', title))) {
      setIsLoading(true);
      setError(undefined);
      setMessage(undefined);
      try {
        await spService.softDeleteUsuario(id);
        setMessage(strings.UserDeletedSuccess);
        setMessageType(MessageBarType.success);
        await fetchData();
      } catch (err) {
        setError(strings.ErrorDeletingUser);
        setMessageType(MessageBarType.error);
        console.error(err);
      } finally {
        setIsLoading(false);
      }
    }
  };*/

   const handleDelete = (item: IUsuarioItem) => {
      setUsuarioToDelete(item);
      setShowDeleteConfirm(true);
  };

  const confirmDelete = async () => {
      if (usuarioToDelete?.Id) {
          try {
              await spService.deleteUsuario(usuarioToDelete.Id);
              const buscaEnGrupoAdmin = await spService.ensureUserInGroup('SRSC_ADMINISTRADOR',usuarioToDelete.email);
              const buscaEnColaboradores = await spService.ensureUserInGroup('SRSC-Colaboradores',usuarioToDelete.email);
              if(buscaEnGrupoAdmin.value.length > 0){
                await spService.removeUserFromGroup('SRSC_ADMINISTRADOR', usuarioToDelete.usuarioId);
              }

              if(buscaEnColaboradores.value.length > 0){
                await spService.removeUserFromGroup('SRSC-Colaboradores', usuarioToDelete.usuarioId);
              }

              setMessage({ type: MessageBarType.success, text: strings.UsuarioDeletedSuccess });

              fetchUsuario();
          } catch (err) {
              //setError(strings.ErrorDeletingUsuario + " " + err.message);
              const msg = err instanceof Error ? err.message : String(err);
              setMessage({ type: MessageBarType.error, text: strings.ErrorDeletingUsuario + ": " + msg });
              console.error("Error eliminando usuario:", err);
          } finally {
              setShowDeleteConfirm(false);
              setUsuarioToDelete(undefined);
          }
      } else {
          setMessage({ type: MessageBarType.error, text: strings.CannotDeleteDivisionWithoutId });
          setShowDeleteConfirm(false);
          setUsuarioToDelete(undefined);
      }
  };

  const onSaveUser = async () => {
    if (!validateForm() || !currentUser) {
      //setMessageType(MessageBarType.error);
      setMessage({ type: MessageBarType.error, text: strings.FormErrorsWarning });
      return;
    }

    setIsLoading(true);
    setErrorTitle(undefined);
    //setMessage(undefined);

    try {
      if (currentUser.Id) {
        await spService.updateUsuario(currentUser);
        adminGrupoUsuario(currentUser);
        setMessage({ type: MessageBarType.success, text: strings.UserUpdatedSuccess });
        //setMessageType(MessageBarType.success);
      } else {
        await spService.createUsuario(currentUser);
        adminGrupoUsuario(currentUser);
        setMessage({ type: MessageBarType.success, text: strings.UserAddedSuccess });
        //setMessageType(MessageBarType.success);
      }
      setIsModalOpen(false);
      setCurrentUser(undefined);

      await fetchUsuario(); // Refresh the list
    } catch (err) {
      //setError(currentUser.Id ? strings.ErrorUpdatingUser : strings.ErrorAddingUser);
      const msg = err instanceof Error ? err.message : String(err);
      if(msg.toLowerCase().includes('duplicado')){
        setErrorTitle("El usuario ya se encuentra registrado.");
      } else {
        
        setMessage({ type: MessageBarType.error, text: currentUser.Id ? strings.ErrorUpdatingUser : strings.ErrorAddingUser + " " + msg });
      }
      console.error(err);
    } finally {
      setIsLoading(false);
    }
  };

  const adminGrupoUsuario = async (currentUser: IUsuarioItem) => {

    const buscaEnGrupoAdmin = await spService.ensureUserInGroup('SRSC_ADMINISTRADOR',currentUser.email);
    const bscaEnColaboradores = await spService.ensureUserInGroup('SRSC-Colaboradores',currentUser.email);

    const estaEnGrupoAdmin = buscaEnGrupoAdmin.value.length > 0;
    const estaEnColaboradores = bscaEnColaboradores.value.length > 0;

    if(currentUser.esAdmin && !estaEnGrupoAdmin){
      await spService.addUserInGroup('SRSC_ADMINISTRADOR', currentUser.LoginName);

    }

    if(!currentUser.esAdmin && estaEnGrupoAdmin){
      const userId = buscaEnGrupoAdmin.value[0].Id;
      await spService.removeUserFromGroup('SRSC_ADMINISTRADOR', userId);
      await spService.addUserInGroup('SRSC-Colaboradores', currentUser.LoginName);

    }

    if(!currentUser.esAdmin &&  !estaEnColaboradores){
      await spService.addUserInGroup('SRSC-Colaboradores', currentUser.LoginName);
    }
  };



  const seleccionarUsuario = (items: any[]) => {
    if (items && items.length > 0) {
        setCurrentUser((prev: IUsuarioItem | undefined) => ({ ...prev, LoginName: items[0].id ? items[0].id : 'N/A', 
                                                                      text: items[0].text || '', 
                                                                      email: items[0].secondaryText || '' } as IUsuarioItem));
        setUserError(undefined);
    } else {
        setCurrentUser((prev: IUsuarioItem | undefined) => ({ ...prev, LoginName: 'N/A' } as IUsuarioItem));
    }
  };


  const onCancel = (): void => {
    setIsModalOpen(false);
    setShowDeleteConfirm(false);
    setCurrentUser(undefined);
    setErrorTitle(undefined);
    setMessage(undefined);
    setUserError(undefined);
    //setVicepresidenciaError(undefined);
    setDivisionError(undefined);
    setGerenciaError(undefined);
  };

  const columns: IColumn[] = [
   // { key: 'idColumn', name: 'ID', fieldName: 'Id', minWidth: 20,  isResizable: true },
    { key: 'userColumn', name: strings.UserLabel, fieldName: 'text', minWidth: 150, isResizable: true },
    { key: 'emailColumn', name: strings.EmailLabel, fieldName: 'email', minWidth: 150, isResizable: true, isMultiline: true },
    { key: 'division', name: strings.PisoPlantaLabel, fieldName: 'divisionTitle', minWidth: 80, isResizable: true, isMultiline: true },
   // { key: 'vpColumn', name: strings.GerenciaVicepresidenciaLabel, fieldName: 'VicepresidenciaTitle', minWidth: 150, isResizable: true, isMultiline: true },
    { key: 'gColumn', name: strings.GerenciaLabel, fieldName: 'divisionTitle', minWidth: 150, isResizable: true },
    { key: 'esAdminColumn', name: strings.UsuarioROLLabel, fieldName: 'esAdmin', minWidth: 80, isResizable: true, onRender: (item: IUsuarioItem) => (item.esAdmin ? "Admin" : "Colaborador") },
    { key: 'activoColumn', name: strings.UsuarioActivoLabel, fieldName: 'activo', minWidth: 50, isResizable: true, onRender: (item: IUsuarioItem) => (item.activo ? strings.YesLabel : strings.NoLabel) },
    {
      key: 'actionsColumn',
      name: strings.AccionesColumn,
      minWidth: 100,
      isResizable: true,
      onRender: (item: IUsuarioItem) => (
        <Stack horizontal tokens={{ childrenGap: 5 }} wrap>
          <TooltipHost content={strings.EditUserButton}>
            <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => onEditUser(item)} />
          </TooltipHost>
          <TooltipHost content={strings.DeleteUserButton}>
            <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => {
              if (item.Id) {
                void handleDelete(item)// onDeleteUser(item.Id, item.text);
              }
            }} />
          </TooltipHost>
        </Stack>
      ),
    },
  ];

  return (
    <div style={{ padding: 20 }}>
      <h2>{strings.MantenedorUsuariosView}</h2>

      
      {message  && <MessageBar messageBarType={message.type} onDismiss={() => setMessage(undefined)}>{message.text}</MessageBar>}

      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 10 }}>
        <PrimaryButton text={strings.AddUserButton} onClick={onAddUser} iconProps={{ iconName: 'Add' }} />
      </Stack>

      {isLoading ? (
        <Spinner size={SpinnerSize.large} label={strings.LoadingUsers} />
      ) : (
        <DetailsList
          items={usuarios}
          columns={columns}
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
        />
      )}

      <Dialog
        hidden={!isModalOpen}
        onDismiss={onCancel}
        dialogContentProps={{ type: DialogType.largeHeader, title: currentUser?.Id ? strings.EditUserButton : strings.AddUserButton }}
        modalProps={{ isBlocking: isLoading }}
      >
        {errorTitle && <MessageBar 
          messageBarType={MessageBarType.error}
          onDismiss={() => setMessage(undefined)}>
            {errorTitle}
          </MessageBar>}

        <Stack tokens={{ childrenGap: 15 }}>
          <PeoplePicker
            context={context as any}
            titleText={strings.UserLabel}
            personSelectionLimit={1}
            webAbsoluteUrl={context.pageContext.web.absoluteUrl}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
            required
            onChange={seleccionarUsuario}
            defaultSelectedUsers={currentUser?.email ? [currentUser.email] : []}
            errorMessage={userError}
          />
          {/* Dropdowns de Vicepresidencia y Gerencia con validación 
          <Dropdown
            label={strings.GerenciaVicepresidenciaLabel}
            required
            options={vicepresidencias}
            selectedKey={currentUser?.VicepresidenciaId || null}
            onChange={(e, option) => {
                setCurrentUser((prev: IUsuarioItem | undefined) => ({ ...prev, VicepresidenciaId: option?.key as number, GerenciaId: 0 } as IUsuarioItem));
                setVicepresidenciaError(undefined);
            }}
            placeholder={strings.SelectVicepresidenciaPlaceholder}
            errorMessage={vicepresidenciaError}
          />*/}
          <Dropdown
            label={strings.PisoPlantaLabel}
            required
            options={divisiones.map(d => ({ key: d.Id!, text: d.Title }))}
            selectedKey={currentUser?.divisionId || null}
            onChange={(e, option) => {
                setCurrentUser((prev: IUsuarioItem | undefined) => ({ ...prev, divisionId: option?.key as number } as IUsuarioItem));
                //setVicepresidenciaError(undefined);
                setDivisionError(undefined);
            }}
            placeholder={strings.SelectPlantaPlaceholder}
            //errorMessage={vicepresidenciaError}
            errorMessage={divisionError}
          />
          <Dropdown
            label={strings.GerenciaLabel}
            required
            options={allGerencias.map(g => ({ key: g.Id!, text: g.Title }))}
            selectedKey={currentUser?.GerenciaId || null}
            onChange={(e, option) => {
                setCurrentUser((prev: IUsuarioItem | undefined) => ({ ...prev, GerenciaId: option?.key as number } as IUsuarioItem));
                setGerenciaError(undefined);
            }}
            placeholder={strings.SelectGerenciaPlaceholder}
            errorMessage={gerenciaError}
           //disabled={!currentUser?.VicepresidenciaId || filteredGerencias.length === 0}
          />
          
          <Toggle
            label="Usuario Administrador"
            onText={strings.YesLabel}
            offText={strings.NoLabel}
            checked={currentUser?.esAdmin || false}
            onChange={(e, checked) =>
              setCurrentUser((prev: IUsuarioItem | undefined) => ({ ...prev, esAdmin: checked || false } as IUsuarioItem))
            }
          />
          <Toggle
            label="Activo"
            onText={strings.YesLabel}
            offText={strings.NoLabel}
            checked={currentUser?.activo || false}
            onChange={(e, checked) =>
              setCurrentUser((prev: IUsuarioItem | undefined) => ({ ...prev, activo: checked || false } as IUsuarioItem))
            }
          />
        </Stack>

        <DialogFooter>
          <PrimaryButton onClick={onSaveUser} text={strings.SaveButton} disabled={isLoading} />
          <DefaultButton onClick={onCancel} text={strings.CancelButton} disabled={isLoading} />
        </DialogFooter>
      </Dialog>

      <Dialog
          hidden={!showDeleteConfirm}
          onDismiss={() => setShowDeleteConfirm(false)}
          dialogContentProps={{
              type: DialogType.normal,
              title: strings.ConfirmDeleteUser.replace('{0}', usuarioToDelete?.text || ''),
          }}
          modalProps={{
              isBlocking: true,
              styles: { main: { maxWidth: 450 } }
          }}
      >
          <DialogFooter>
              <PrimaryButton onClick={confirmDelete} text={strings.UsuarioDeleteButton} />
              <DefaultButton onClick={onCancel} text={strings.CancelButton} />
          </DialogFooter>
      </Dialog>
    </div>
  );
};

export default MantenedorUsuarios;