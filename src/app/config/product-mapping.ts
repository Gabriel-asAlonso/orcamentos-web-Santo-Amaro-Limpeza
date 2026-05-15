export const PRODUCT_MAPPING: Record<string, string> = {
  'Lavanda Pro SA': 'Limpador Perfumado Concentrado',
  'Algas Pro SA': 'Limpador Perfumado Concentrado',
  'SA-24': 'Desencrustante Concentrado',
  'Pine Limp SA': 'Detergente Desengraxante Concentrado',
  'Peroxy Pro SA': 'Limpador de Uso Geral Concentrado (Peróxido de Hidrogênio)'
};

export function mapProductName(productName: string): string {
  return PRODUCT_MAPPING[productName] || productName;
}
