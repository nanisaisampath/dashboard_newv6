import { read, utils } from 'xlsx';

/**
 * Parses an Excel file and returns the data as an array of objects
 * @param file - The Excel file to parse
 * @returns Promise that resolves to an array of objects representing the Excel data
 */
export const parseExcelFile = async (file: File): Promise<any[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = read(data, { type: 'binary' });

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        const jsonData = utils.sheet_to_json(worksheet, { defval: '', raw: false });

        console.log('Parsed Excel Sample:', jsonData.slice(0, 1));
        resolve(jsonData as any[]);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = (error) => reject(error);

    reader.readAsBinaryString(file);
  });
};

/**
 * Calculates metrics from ticket data
 * @param data - Array of ticket data
 * @returns Object containing calculated metrics
 */
export const calculateMetrics = (data: any[]) => {
  if (!data || data.length === 0) {
    return {
      totalTickets: 0,
      openTickets: 0,
      resolvedTickets: 0
    };
  }

  const totalTickets = data.length;

  const openStatuses = ['open', 'in review', 'hold', 'in progress'];
  const resolvedStatuses = ['closed'];

  const openTickets = data.filter(item => {
    const status = (item.Status || item.status || '').toLowerCase();
    return openStatuses.includes(status);
  }).length;

  const resolvedTickets = data.filter(item => {
    const status = (item.Status || item.status || '').toLowerCase();
    return resolvedStatuses.includes(status);
  }).length;

  return {
    totalTickets,
    openTickets,
    resolvedTickets
  };
};

/**
 * Prepare data for time series chart (tickets per day)
 * @param data - Array of ticket data
 * @returns Prepared data for chart
 */
export const prepareTimeSeriesData = (data: any[]) => {
  if (!data || data.length === 0) return [];

  const ticketsByDate: Record<string, number> = {};

  data.forEach(ticket => {
    const rawDate = ticket.Date || ticket.date || ticket['Created Date'] || ticket['Created On'];
    if (rawDate) {
      const dateKey = new Date(rawDate).toISOString().split('T')[0];
      ticketsByDate[dateKey] = (ticketsByDate[dateKey] || 0) + 1;
    }
  });

  return Object.entries(ticketsByDate)
    .map(([date, tickets]) => ({ date, tickets }))
    .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
};

/**
 * Prepare data for bar/pie charts based on a given category
 * @param data - Array of ticket data
 * @param category - The category key (e.g., 'technology', 'client', etc.)
 * @returns Aggregated chart data
 */
export const prepareCategoryData = (data: any[], category: string) => {
  if (!data || data.length === 0 || !category) return [];

  const categoryMap: Record<string, number> = {};

  data.forEach(ticket => {
    let categoryValue;

    switch (category) {
      case 'technology':
        categoryValue = ticket['Technology/Platform'] || ticket.Technology || ticket.technology;
        break;
      case 'client':
        categoryValue = ticket.Client || ticket.client;
        break;
      case 'ticketType':
        categoryValue = ticket['Ticket Type'] || ticket.TicketType || ticket.ticketType;
        break;
      case 'assignedTo':
        categoryValue = ticket['Assigned To'] || ticket['Assigned to'] || ticket.AssignedTo || ticket.assignedTo;
        break;
      case 'status':
        categoryValue = ticket.Status || ticket.status;
        break;
      default:
        categoryValue = ticket[category] || ticket[category.toLowerCase()] || ticket[capitalize(category)];
        break;
    }

    const value = categoryValue ? String(categoryValue).trim() : 'Unknown';
    categoryMap[value] = (categoryMap[value] || 0) + 1;
  });

  return Object.entries(categoryMap).map(([name, value]) => ({ name, value }));
};

/**
 * Capitalizes first letter of a string
 */
const capitalize = (text: string) => text.charAt(0).toUpperCase() + text.slice(1);
