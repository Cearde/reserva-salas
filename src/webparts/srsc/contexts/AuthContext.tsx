import * as React from 'react';
import { createContext, useContext, useState, useEffect, useMemo } from 'react';
import { SPService } from '../services/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
//import { IUsuarioItem } from '../components/models/entities';

// The shape of the context
interface IAuthContext {
    isAdmin: boolean;
    isAuthLoading: boolean;
    isUserValido:boolean
}

// Create the context
const AuthContext = createContext<IAuthContext | undefined>(undefined);

/**
 * The provider component that wraps the app and provides the auth status.
 */
export const AuthProvider: React.FC<{ children: React.ReactNode, context: WebPartContext }> = ({ children, context }) => {
    const [isAdmin, setIsAdmin] = useState(false);
    const [isAuthLoading, setIsAuthLoading] = useState(true);
    const [isUserValido, setIsUserValido] = useState(false);

    const spService = useMemo(() => new SPService(context), [context]);

    useEffect((): void => {
        const checkAdminStatus = async (): Promise<void> => {
            try {
                // Check if the user is in the specific admin group
                const adminStatus = await spService.isCurrentUserInGroup('SRSC_ADMINISTRADOR');
                const usuarios= await spService.getUsuarios(true);
                const usuarioValido = usuarios.filter(usuario => usuario.secondaryText.toLowerCase() === context.pageContext.user.loginName.toLowerCase() && usuario.activo).length > 0;
                setIsAdmin(adminStatus);
                setIsUserValido(usuarioValido);
            } catch (error: unknown) {
                console.error("Failed to check admin status", error);
                setIsAdmin(false); // Default to not admin on error
            } finally {
                setIsAuthLoading(false);
            }
        };

        checkAdminStatus().catch(err => {
            console.error("An unexpected error occurred in checkAdminStatus:", err);
        });
    }, [spService]);

    const value = useMemo(() => ({ isAdmin, isAuthLoading, isUserValido }), [isAdmin, isAuthLoading, isUserValido]);

    return (
        <AuthContext.Provider value={value}>
            {children}
        </AuthContext.Provider>
    );
};

/**
 * Custom hook to easily consume the AuthContext.
 */
export const useAuth = (): IAuthContext => {
    const context = useContext(AuthContext);
    if (context === undefined) {
        throw new Error('useAuth must be used within an AuthProvider');
    }
    return context;
};
