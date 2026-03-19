import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

// Define the context with a default value of undefined
const SPFxContext = React.createContext<WebPartContext | undefined>(undefined);

/**
 * Custom hook to consume the SPFxContext.
 * It ensures the context is not undefined when used.
 */
export const useSPFxContext = (): WebPartContext => {
  const context = React.useContext(SPFxContext);
  if (context === undefined) {
    throw new Error('useSPFxContext must be used within a SPFxContextProvider');
  }
  return context;
};

export default SPFxContext;
