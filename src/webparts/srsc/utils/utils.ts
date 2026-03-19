export const getRestrictedDates = (rangeDays: number = 365): Date[] => {
  const restricted: Date[] = [];
  const start = new Date();
  // Empezamos un poco antes de hoy por si el usuario navega hacia atrás
  start.setDate(start.getDate() - 30); 

  for (let i = 0; i < rangeDays; i++) {
    const date = new Date(start);
    date.setDate(start.getDate() + i);
    
    // 0 = Domingo, 6 = Sábado
    if (date.getDay() === 0 || date.getDay() === 6) {
      restricted.push(new Date(date));
    }
  }
  return restricted;
};

export const getStatusColor = (disponibilidad: string): string => {
  switch (disponibilidad.toLowerCase()) {
    case 'full':
      return '#E74C3C';//'red';
    case 'empty':
      return '#27AE60';//'green';
    case 'partial':
      return '#F1C40F';'yellow';
    default:
      return 'grey'; // Color de respaldo por si el dato viene mal
  }
};

export const calcularPermisos = (isAdmin: boolean) => {
  return isAdmin ? 'Acceso Total' : 'Acceso Limitado';
};

export const onFormatDate = (date?: Date): string => {
    if (!date) return '';
    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0'); // +1 because January is 0
    const year = date.getFullYear();
    return `${day}-${month}-${year}`;
  };
export const APP_VERSION = "1.2.0.1"; // Aquí pones la versión de tu package-solution

