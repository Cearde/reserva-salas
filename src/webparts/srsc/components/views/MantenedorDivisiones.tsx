
import * as React from 'react';
import { useState } from 'react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
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
import SPFxContext from '../../contexts/SPFxContext';
import { SPService } from '../../services/sp';
import * as strings from 'SrscWebPartStrings';
import { IDivisionItem } from '../models/entities';

const MantenedorDivisiones: React.FC = () => {
    const spContext = React.useContext(SPFxContext);
    const spService = React.useMemo(() => new SPService(spContext as WebPartContext), [spContext]);

    const [divisiones, setDivisiones] = React.useState<IDivisionItem[]>([]);
    const [loading, setLoading] = React.useState<boolean>(true);
   // const [error, setError] = React.useState<string | undefined>(undefined);
    const [isModalOpen, setIsModalOpen] = React.useState<boolean>(false);
    const [currentDivision, setCurrentDivision] = React.useState<IDivisionItem | undefined>(undefined);
    const [formErrors, setFormErrors] = React.useState<{ Title?: string }>({});
    const [showDeleteConfirm, setShowDeleteConfirm] = React.useState<boolean>(false);
    const [divisionToDelete, setDivisionToDelete] = React.useState<IDivisionItem | undefined>(undefined);
    const [message, setMessage] = React.useState<{ type: MessageBarType, text: string } | undefined>(undefined);
    const [titleError, setTitleError] = useState<string | undefined>(undefined);
    
    const fetchDivisiones = React.useCallback(async () => {
        setLoading(true);
        setTitleError(undefined);
        try {
            const items = await spService.getDivisiones(true); // Fetch all, including inactive
            setDivisiones(items);
        } catch (err) {
            const msg = err instanceof Error ? err.message : String(err);
            setMessage({ type: MessageBarType.error, text: strings.ErrorFetchingDivisiones + " " + msg });
            //setTitleError(strings.ErrorFetchingDivisiones);
            console.error("Error fetching divisiones:", err);
        } finally {
            setLoading(false);
        }
    }, [spService]);

    React.useEffect(() => {
        fetchDivisiones();
    }, [fetchDivisiones]);

    const validateForm = (): boolean => {
        const errors: { Title?: string } = {};
        if (!currentDivision?.Title) {
            errors.Title = strings.RequiredField;
        }
        setFormErrors(errors);
        return Object.keys(errors).length === 0;
    };

    const handleAdd = () => {
        setCurrentDivision({ Id: undefined, Title: '', activo: true });
        setFormErrors({});
        setIsModalOpen(true);
    };

    const handleEdit = (item: IDivisionItem) => {
        setCurrentDivision({ ...item });
        setFormErrors({});
        setIsModalOpen(true);
    };

    const handleDelete = (item: IDivisionItem) => {
        setDivisionToDelete(item);
        setShowDeleteConfirm(true);
    };

    const confirmDelete = async () => {
        if (divisionToDelete?.Id) {
            try {
                await spService.deleteDivision(divisionToDelete.Id);
                setMessage({ type: MessageBarType.success, text: strings.DivisionDeletedSuccess });
                fetchDivisiones();
            } catch (err) {
                //setTitleError(strings.ErrorDeletingDivision + " " + err.message);
                const msg = err instanceof Error ? err.message : String(err);
                setMessage({ type: MessageBarType.error, text: strings.ErrorDeletingDivision + " " + msg });
                console.error("Error eliminando division:", err);
            } finally {
                setShowDeleteConfirm(false);
                setDivisionToDelete(undefined);
            }
        } else {
            setMessage({ type: MessageBarType.error, text: strings.CannotDeleteDivisionWithoutId });
            setShowDeleteConfirm(false);
            setDivisionToDelete(undefined);
        }
    };

    const onCancel = (): void => {
    setIsModalOpen(false);
    setShowDeleteConfirm(false);
    setCurrentDivision(undefined);
    setTitleError(undefined);
    setMessage(undefined);
    setTitleError(undefined);
  };
    const handleSave = async () => {
        if (!validateForm()) {
            setTitleError(strings.FormErrorsWarning );
            return;
        }

        if (!currentDivision) return;

        try {
            if (currentDivision.Id) {
                await spService.updateDivision(currentDivision);
                setMessage({ type: MessageBarType.success, text: strings.DivisionUpdatedSuccess });
            } else {
                await spService.createDivision(currentDivision);
                setMessage({ type: MessageBarType.success, text: strings.DivisionAddedSuccess });
            }
            setIsModalOpen(false);
            setCurrentDivision(undefined);
            fetchDivisiones();
        } catch (err) {
            //setTitleError(currentDivision.Id ? strings.ErrorUpdatingDivision : strings.ErrorAddingDivision);
            setMessage({ type: MessageBarType.error, text: currentDivision.Id ? strings.ErrorUpdatingDivision : strings.ErrorAddingDivision });
            console.error("Error saving division:", err);
        }
    };

    const columns: IColumn[] = [
       /* {
            key: 'column1',
            name: 'ID',
            fieldName: 'Id',
            minWidth: 40,
            maxWidth: 60, // Keep ID small
            isResizable: true
        },*/
        {
            key: 'column2',
            name: strings.DivisionNameLabel,
            fieldName: 'Title',
            minWidth: 200, // Main column, give it more space
            isResizable: true
        },
        {
            key: 'column3',
            name: strings.DivisionActiveLabel,
            fieldName: 'activo',
            minWidth: 80,
            isResizable: true,
            onRender: (item: IDivisionItem) => (item.activo ? strings.YesLabel : strings.NoLabel),
        },
        {
            key: 'actionsColumn',
                  name: strings.AccionesColumn,
                  minWidth: 100,
                  isResizable: true,
                  onRender: (item: IDivisionItem) => (
                    <Stack horizontal tokens={{ childrenGap: 5 }} wrap>
                      <TooltipHost content={strings.EditDivisionButton}>
                        <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => handleEdit(item)} />
                      </TooltipHost>
                      <TooltipHost content={strings.DeleteDivisionButton}>
                        <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => {
                          if (item.Id) {
                            void handleDelete(item);
                          } else {
                            setTitleError(strings.CannotDeleteWithoutId);
                          }
                        }} />
                      </TooltipHost>
                    </Stack>
            )
        }
    ];

    return (
        <div style={{ padding: 20 }}>
            <h2>{strings.MantenedorDivisionesView}</h2>

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

            

            <PrimaryButton onClick={handleAdd} style={{ marginBottom: 20 }}>
                {strings.AddDivisionButton}
            </PrimaryButton>

            {loading ? (
                <Spinner size={SpinnerSize.large} label={strings.LoadingDivisiones} />
            ) : (
                divisiones.length > 0 ? (
                    <DetailsList
                        items={divisiones}
                        columns={columns}
                        setKey="Id"
                        layoutMode={DetailsListLayoutMode.justified}
                        selectionMode={SelectionMode.none}
                        isHeaderVisible={true}
                    />
                ) : (
                    <MessageBar messageBarType={MessageBarType.info}>
                        {strings.NoDivisionesFound}
                    </MessageBar>
                )
            )}

            <Dialog
                hidden={!isModalOpen}
                onDismiss={onCancel}
                dialogContentProps={{
                type: DialogType.largeHeader,
                title: currentDivision?.Id ? strings.EditDivisionButton : strings.AddDivisionButton,
        }}
            >
                {titleError && (
                <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                    {titleError}
                </MessageBar>
            )   }
                <div style={{ padding: 20 }}>
                    <h3>{currentDivision?.Id ? strings.EditDivisionButton : strings.AddDivisionButton}</h3>
                    <TextField
                        label={strings.DivisionNameLabel}
                        value={currentDivision?.Title || ''}
                        onChange={(e, newValue) => setCurrentDivision(prev => prev ? { ...prev, Title: newValue || '' } : undefined)}
                        required
                        errorMessage={formErrors.Title}
                    />
                    <Toggle
                        label={strings.DivisionActiveLabel}
                        onText={strings.YesLabel}
                        offText={strings.NoLabel}
                        checked={currentDivision?.activo}
                        onChange={(e, checked) => setCurrentDivision(prev => prev ? { ...prev, activo: checked || false } : undefined)}
                        style={{ marginTop: 10 }}
                    />
                    <div style={{ marginTop: 20 }}>
                        <PrimaryButton onClick={handleSave}>{strings.SaveButton}</PrimaryButton>
                        <DefaultButton onClick={() => setIsModalOpen(false)} style={{ marginLeft: 8 }}>{strings.CancelButton}</DefaultButton>
                    </div>
                </div>
            </Dialog>

            <Dialog
                hidden={!showDeleteConfirm}
                onDismiss={() => setShowDeleteConfirm(false)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: strings.ConfirmDeleteDivision.replace('{0}', divisionToDelete?.Title || ''),
                    //subText: strings.ConfirmDeleteDivision.replace('{0}', divisionToDelete?.Title || '')
                }}
                modalProps={{
                    isBlocking: true,
                    styles: { main: { maxWidth: 450 } }
                }}
            >
                <DialogFooter>
                    <PrimaryButton onClick={confirmDelete} text={strings.DeleteDivisionButton} />
                    <DefaultButton onClick={onCancel} text={strings.CancelButton} />
                </DialogFooter>
            </Dialog>
        </div>
    );
};

export default MantenedorDivisiones;