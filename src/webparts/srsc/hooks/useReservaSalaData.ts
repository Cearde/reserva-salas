import * as React from 'react';
import { SPService } from '../services/sp';
import { IDropdownOption } from '@fluentui/react';
import { ISala } from '../components/models/entities';
import * as strings from 'SrscWebPartStrings';
import { MessageBarType } from '@fluentui/react';

interface DropdownOptions {
  plantas: IDropdownOption[];
  pisos: IDropdownOption[];
  usuarios: IDropdownOption[];
}

interface HookState<T> {
  data: T;
  loading: boolean;
  error?: { message: string; type: MessageBarType };
}

// Custom hook for fetching form dropdown options
export function useFormDropdownOptions(spService: SPService): HookState<DropdownOptions> {
  const [state, setState] = React.useState<HookState<DropdownOptions>>({
    data: { plantas: [], pisos: [], usuarios: [] },
    loading: true,
  });

  React.useEffect(() => {
    let isMounted = true;
    setState(prev => ({ ...prev, loading: true, error: undefined }));

    spService.getFormDropdownOptions()
      .then(options => {
        if (isMounted) {
          setState(prev => ({ ...prev, data: options, loading: false }));
        }
      })
      .catch(error => {
        console.error("Error fetching dropdown options", error);
        if (isMounted) {
          setState(prev => ({
            ...prev,
            loading: false,
            error: { message: strings.ErrorLoadingFilterOptions, type: MessageBarType.error },
          }));
        }
      });

    return () => {
      isMounted = false;
    };
  }, [spService]);

  return state;
}

// Custom hook for fetching salas by piso
export function useSalasByPiso(spService: SPService, selectedPiso?: number, fecha?: Date, plantaId?: number): HookState<ISala[]> {
  const [state, setState] = React.useState<HookState<ISala[]>>({
    data: [],
    loading: false,
  });

  React.useEffect(() => {
    let isMounted = true;
    if (selectedPiso && fecha && plantaId) {
      setState(prev => ({ ...prev, loading: true, error: undefined }));
      spService.getSalasByPiso(selectedPiso, fecha, plantaId)
        .then(salasData => {
          if (isMounted) {
            setState(prev => ({ ...prev, data: salasData, loading: false }));
          }
        })
        .catch(error => {
          console.error(`Error fetching salas for piso ${selectedPiso}:`, error);
          if (isMounted) {
            setState(prev => ({
              ...prev,
              loading: false,
              error: { message: `${strings.ErrorLoadingSalasForPiso} ${selectedPiso}.`, type: MessageBarType.error },
            }));
          }
        });
    } else {
      setState(prev => ({ ...prev, data: [], loading: false, error: undefined }));
    }

    return () => {
      isMounted = false;
    };
  }, [ selectedPiso,fecha,plantaId]);//[spService, selectedPiso,fecha]);

  return state;
}
