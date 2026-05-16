export const PRODUCT_MAPPING: Record<string, string> = {
  'Lavanda Pro SA': 'Limpador Perfumado Concentrado',
  'Algas Pro SA': 'Limpador Perfumado Concentrado',
  'SA-24': 'Desencrustante Concentrado',
  'Pine Limp SA': 'Detergente Desengraxante Concentrado',
  'Peroxy Pro SA': 'Limpador de Uso Geral Concentrado (Peróxido de Hidrogênio)'
};

export function mapProductName(productName: string): string {
  const trimmed = productName.trim();

  // Match exato (comportamento original, sem litragem no nome)
  if (PRODUCT_MAPPING[trimmed]) {
    return PRODUCT_MAPPING[trimmed];
  }

  // Separar sufixo de litragem no final do nome: "5L", "2 L", "1,5L", "20L"
  const volumeMatch = trimmed.match(/\s+(\d+(?:[.,]\d+)?\s*L)\s*$/i);

  if (volumeMatch?.index !== undefined) {
    const volume = volumeMatch[1].replace(/\s+/g, '').toUpperCase();
    const baseName = trimmed.slice(0, volumeMatch.index).trim();
    const mappedBase = PRODUCT_MAPPING[baseName];

    if (mappedBase) {
      return `${mappedBase} ${volume}`;
    }
  }

  return productName;
}
