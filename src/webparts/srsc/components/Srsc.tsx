import * as React from 'react';
import styles from '../components/Srsc.module.scss';
import { ISrscProps } from './ISrscProps';
import SPFxContext from '../contexts/SPFxContext';
import {useAuth} from '../contexts/AuthContext';
import * as strings from 'SrscWebPartStrings';

// Import the new view components
import ReservaSala from './views/ReservaSala';
import SalasDisponibles from './views/SalasDisponibles';
import BloqueoSala from './views/BloqueoSala';
import GenerarQR from './views/GenerarQR';
import Reportes from './views/Reportes';
import MisReservas from './views/MisReservar';
import MantenedorHorarios from './views/MantenedorHorarios';
import MantenedorUsuarios from './views/MantenedorUsuarios';
import MantenedorSalas from './views/MantenedorSalas';
import MantenedorPisos from './views/MantenedorPisos'; // New import
import MantenedorVicepresidencias from './views/MantenedorVicepresidencias'; // New import
import MantenedorGerencia from './views/MantenedorGerencia'; // New import
import MantenedorDivisiones from './views/MantenedorDivisiones';
import { APP_VERSION } from '../utils/utils';
import { Icon } from '@fluentui/react';
//import { DialogFooter } from '@fluentui/react';

// Define a type for menu items that can have sub-items
interface MenuItem {
  key: string;
  text: string;
  view?: string; // The view to render if it's a direct link
  icon?: string; // Optional icon class for the menu item
  url?: string;
  subItems?: MenuItem[]; // Sub-items for a dropdown menu
  adminOnly?: boolean; // Flag to indicate if the item is for admins only
}

const menuItems: MenuItem[] = [
  { key: 'reservaSala', text: strings.ReservaSalaView, view: strings.ReservaSalaView,  icon :'fa-solid fa-calendar-check', adminOnly: false },
  { key: 'salasDisponibles', text: strings.SalasDisponiblesView, view: strings.SalasDisponiblesView, icon: 'fa-solid fa-chair', adminOnly: false },//fa-building
  { key: 'bloqueoSala', text: strings.BloqueoSalaView, view: strings.BloqueoSalaView , icon: 'fa-solid fa-user-lock', adminOnly: true }, //fa-user-lock
  { key: 'misReservas', text: strings.MisReservasView, view: strings.MisReservasView, icon: 'fa-solid fa-calendar-alt', adminOnly: false }, //fa-calendar-alt
  { key: 'generarQR', text: strings.GenerarQRView, view: strings.GenerarQRView, icon: 'fa-solid fa-qrcode', adminOnly: true},
  { key: 'reportes', text: strings.ReportesView, view: strings.ReportesView, icon: 'fa-solid fa-chart-bar' , adminOnly: true},
  {
    key: 'mantenedores',
    text: strings.MantenedorView, icon: 'fa-solid fa-tools',
    adminOnly: true,
    subItems: [
      { key: 'mantenedorDivisiones', text: strings.MantenedorDivisionesView, view: strings.MantenedorDivisionesView },
      { key: 'mantenedorPisos', text: strings.MantenedorPisosView, view: strings.MantenedorPisosView }, // New item
      { key: 'mantenedorSalas', text: strings.MantenedorSalasView, view: strings.MantenedorSalasView },
      { key: 'mantenedorHorarios', text: strings.MantenedorHorariosView, view: strings.MantenedorHorariosView },
      { key: 'mantenedorVicepresidencias', text: strings.MantenedorVicepresidenciasView, view: strings.MantenedorVicepresidenciasView }, // New item
      { key: 'mantenedorGerencia', text: strings.MantenedorGerenciaView, view: strings.MantenedorGerenciaView }, // New item
      { key: 'mantenedorUsuarios', text: strings.MantenedorUsuariosView, view: strings.MantenedorUsuariosView },
      { key:'contenidos', text: 'Contenido del Sitio', view: '', url: '/_layouts/15/viewlsts.aspx?view=14'}, // Example of a non-view item
      { key: 'version', text: 'V. ' + APP_VERSION, view: '' },
    ],
  },
];

// The View type will now be derived from all possible 'view' properties in menuItems
type View =
  typeof strings.ReservaSalaView |
  typeof strings.SalasDisponiblesView |
  typeof strings.BloqueoSalaView |
  typeof strings.GenerarQRView |
  typeof strings.ReportesView |
  typeof strings.MisReservasView |
  typeof strings.MantenedorPisosView | // New view type
  typeof strings.MantenedorHorariosView |
  typeof strings.MantenedorUsuariosView |
  typeof strings.MantenedorSalasView |
  typeof strings.MantenedorDivisionesView |
  typeof strings.MantenedorVicepresidenciasView; // New view type

const Srsc: React.FC<ISrscProps> = (props) => {
  const [currentView, setCurrentView] = React.useState<View>(menuItems[0].view as View); // Initialize with the first item's view
  const [isMantenedoresOpen, setIsMantenedoresOpen] = React.useState<boolean>(false);

  const renderView = (): JSX.Element => {
    switch (currentView) {
      case strings.ReservaSalaView:
        return <ReservaSala {... props}/>;
      case strings.SalasDisponiblesView:
        return <SalasDisponibles {... props}/>;
      case strings.BloqueoSalaView:
        return <BloqueoSala {... props}/>;
      case strings.GenerarQRView:
        return <GenerarQR />;
      case strings.ReportesView:
        return <Reportes {... props}/>;
      case strings.MisReservasView:
        return <MisReservas {... props}/>;
      case strings.MantenedorPisosView: // New case
        return <MantenedorPisos {... props}{... props}/>;
      case strings.MantenedorHorariosView:
        return <MantenedorHorarios {... props}/>;
      case strings.MantenedorUsuariosView:
        return <MantenedorUsuarios {... props}/>;
      case strings.MantenedorSalasView:
        return <MantenedorSalas {... props}/>;
      case strings.MantenedorVicepresidenciasView: // New case
        return <MantenedorVicepresidencias {... props}/>;
      case strings.MantenedorGerenciaView: // New case
        return <MantenedorGerencia {... props}/>;
      case strings.MantenedorDivisionesView:
        return <MantenedorDivisiones />;
      default:
        // Fallback to a default view if currentView is somehow not in menuItems
        // This case should ideally not be reached with proper type handling
        return <ReservaSala {... props}/>;
    }
  };

  const { isAdmin, isUserValido, isAuthLoading } = useAuth();

  const onItemClick = (item: View): void => {
  // Si el usuario hace clic en 'version', simplemente retornamos y no hacemos nada
  if (item && item === 'version') {
    //ev?.preventDefault(); // Evita cualquier comportamiento por defecto
    return; 
  }

  // Si no es la versión, cambiamos la vista normalmente
  if (item && item) {
    setCurrentView(item);
  }
  };

  if (isAuthLoading) {
    return (
      <div className={styles.srsc}>
        <div className={styles.titleSection}>
          {/* Puedes poner un Spinner de Fluent UI o un mensaje discreto */}
          <h1 className={styles.mainTitle}>Cargando aplicación...</h1>
        </div>
      </div>
    );
  }

  if (!isUserValido) {
    return (
      <div className={styles.srsc}>
        <div className={styles.userSection}>
          <span className={styles.userName}>👤 {props.userDisplayName}</span>
        </div>
        <div className={styles.titleSection}>
          <h1 className={styles.mainTitle}>Acceso restringido</h1>
          <p className={styles.mainTitle} style={{ fontSize: '18px' }}>
            El usuario no está correctamente configurado en el sistema.
          </p>
        </div>
      </div>
    );
  }

  return (
    <SPFxContext.Provider value={props.context}>
    {isUserValido && (
      <div className={styles.srsc}>
        <div className={styles.srsc}>
            <div className={styles.userSection}>
            <span className={styles.userName}>
              👤 {props.userDisplayName}
            </span>
          </div>

          {/* Sección Central: Título */}
          <div className={styles.titleSection}>
            <img src='https://codelcochile.sharepoint.com/sites/srsc/SiteAssets/logoCodelco.png' alt="Logo" className={styles.logo} />
            <h1 className={styles.mainTitle}>Reserva Salas Corporativa</h1>
          </div>
          
          {/* Sección Derecha: Espaciador para mantener el equilibrio del centro */}
          <div className={styles.spacer}></div>
        </div>
        

        <header className={styles.appHeader}>
          <nav className={styles.appNav}> 
              {menuItems.map(item => {
                if (item.subItems && isAdmin) {
                  return (
                    <div key={item.key} className={styles.dropdown}>
                      <button
                        className={`${styles.navButton} ${isMantenedoresOpen ? styles.active : ''}`}
                        onClick={() => setIsMantenedoresOpen(!isMantenedoresOpen)}
                      >
                        {item.icon && (
                          item.icon.indexOf('fa-') > -1 ? (
                            <i className={item.icon} style={{ marginRight: '8px' }}></i>
                          ) : (
                            <Icon iconName={item.icon} style={{ marginRight: '8px' }} />
                          )
                        )}
                        {item.text}
                      </button>
                      {isMantenedoresOpen &&  isAdmin &&(
                        <div className={styles.dropdownContent}>
                          {item.subItems.map(subItem => (
                            <button
                              key={subItem.key}
                              className={`${styles.navButton} ${currentView === subItem.view ? styles.active : ''}`}
                              onClick={() => {
                                if(subItem.url) {
                                  window.open(props.context.pageContext.site.absoluteUrl + subItem.url, '_blank');
                                  console.log('Abriendo URL:', subItem.url);
                                  return;
                                }else{
                                  onItemClick(subItem.view as View);// setCurrentView(subItem.view as View);
                                  setIsMantenedoresOpen(false); // Close dropdown after selection
                                }
                              }}
                            >
                              {subItem.text}
                            </button>
                          ))}
                        </div>
                      )}
                    </div>
                  );
                } else {
                  if ( (isAdmin) || (!item.adminOnly && !isAdmin) ) {
                    return (
                      <button
                        key={item.key}
                        className={`${styles.navButton} ${currentView === item.view ? styles.active : ''}`}
                        onClick={() => onItemClick(item.view as View)}//setCurrentView(item.view as View)}
                      >
                        
                          {item.icon && (
                            item.icon.indexOf('fa-') > -1 ? (
                              <i className={item.icon} style={{ marginRight: '8px' }}></i>
                            ) : (
                              <Icon iconName={item.icon} style={{ marginRight: '8px' }} />
                            )
                          )}

                        {item.text}
                      </button>
                    );
                  }
                }
              })} 
          </nav>
        </header>
        <main className={styles.appMain}>
          <div className={styles.viewContent}>
            {renderView()}
          </div>
        </main>
      </div>
      ) }
      {!isUserValido && (
        <div className={styles.srsc}> 
          <div className={styles.srsc}>
              <div className={styles.userSection}>
              <span className={styles.userName}>
                👤 {props.userDisplayName}
              </span>
            </div>

            {/* Sección Central: Título */}
            <div className={styles.titleSection}>
              <h1 className={styles.mainTitle}>Reserva Salas Corporativa</h1>
            </div>
            
            {/* Sección Derecha: Espaciador para mantener el equilibrio del centro */}
            <div className={styles.spacer}></div>
          </div>
          <div className={styles.titleSection}>
              <h1 className={styles.mainTitle}>El usuario no esta correctamente configurado</h1>
          </div>
        </div>
      )}
    </SPFxContext.Provider>
  
  );
};

export default Srsc;