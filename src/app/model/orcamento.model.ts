import { type Company } from '../data/empresas';

export type QuoteItem = {
  id: number;
  product: string;
  quantity: number;
  unit: string;
  unitPrice: number;

  code?: string;
  reference?: string;
  discount?: number;
};

export type CustomerData = {
  client?: string;
  fantasyName?: string;
  address?: string;
  district?: string;
  city?: string;
  state?: string;
  postalCode?: string;
  document?: string;
  stateRegistration?: string;
  region?: string;
  freight?: string;
  phone?: string;
  whatsapp?: string;
  priceTable?: string;
  contact?: string;
  email?: string;
  salesperson?: string;
  externalSalesperson?: string;
};

export type ExcelImportResult = {
  items: QuoteItem[];
  customerData?: CustomerData;
  companyId?: string;
  deliveryTime?: string;
  paymentMethod?: string;
};

export type Quote = {
  label: string;
  company: Company;
  deliveryTime: string;
  paymentMethod: string;
  customerData?: CustomerData;
  items: QuoteItem[];
  adjustment: number;
  source: string;
  date?: Date;
};
