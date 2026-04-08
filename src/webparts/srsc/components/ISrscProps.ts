import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISrscProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  loginName: string;
  usuarioIDLista: number; // Agregamos el userId como opcional, se establecerá después de obtenerlo
  usuarioIDDivision: number;
}
