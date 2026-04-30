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

export type ExcelImportResult = {
  items: QuoteItem[];
  deliveryTime?: string;
  paymentMethod?: string;
};

export type Quote = {
  label: string;
  company: Company;
  deliveryTime: string;
  paymentMethod: string;
  items: QuoteItem[];
  adjustment: number;
  source: string;
};