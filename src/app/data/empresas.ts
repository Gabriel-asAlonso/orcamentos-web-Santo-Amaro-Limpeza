export type Company = {
  id: string;
  name: string;
  address: string;
  identity: string;
  taxId: string;
  logoPath: string;
  contact?: string;
  phone?: string;
  email?: string;
  postalCode?: string;
  cnpj?: string;
};

export const COMPANIES: Company[] = [
  {
    id: 'splash',
    name: 'SPLASH COMERCIO E SERVIÇO DE PROD.DE LIMPEZA LTDA.ME',
    address: 'Rua Onofre Silveira, 117 - Vila Fachini - CEP 04334-100',
    identity: 'Contato: Ricardo / Andrea - Fone: 3384-5096',
    taxId: '',
    logoPath: '/logoSantoAmaroLimpeza.jpg',
    contact: 'Ricardo / Andrea',
    phone: '3384-5096',
    email: 'splash@splashconsultoria.com.br',
    postalCode: '04334-100',
    cnpj: '12.584.220/0001-76'
  },
  {
    id: 'aurora',
    name: 'Aurora Eventos Corporativos',
    address: 'Av. das Nacoes, 420 - Sao Paulo/SP',
    identity: 'Verde institucional, selo executivo e rodape fiscal',
    taxId: '12.584.220/0001-76',
    logoPath: '/logoExemplo.jpeg',
    contact: 'Ricardo / Andrea',
    phone: '8888-8888',
    email: 'splash@splashconsultoria.com.br',
    postalCode: '04334-100',
    cnpj: '12.888.888/0001-76'

  },
  {
    id: 'norte',
    name: 'Norte Food Service',
    address: 'Rua Itapua, 88 - Curitiba/PR',
    identity: 'Marca ambar, cabecalho comercial e assinatura digital',
    taxId: '08.991.310/0001-44',
    logoPath: '/logoExemplo2.jpg'
  },
  {
    id: 'delta',
    name: 'Delta Facilities',
    address: 'Alameda Central, 1370 - Belo Horizonte/MG',
    identity: 'Azul tecnico, tabela compacta e dados operacionais',
    taxId: '31.773.005/0001-20',
    logoPath: 'assets/logos/delta.png'
  }
];
