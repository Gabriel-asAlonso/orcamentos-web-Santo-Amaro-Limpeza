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
  stateRegistration?: string;
  website?: string;
};

export const COMPANIES: Company[] = [
  {
    id: 'splash',
    name: 'SPLASH COMERCIO E SERVIÇO DE PRODDE LIMPEZA LTDA. ME',
    address: 'Rua Onofre Silveira, 117 - Vila Fachini - CEP 04334-100',
    identity: 'Contato: Ricardo / Andrea - Fone: 3384-5096',
    taxId: '',
    logoPath: '/logoSplash.png',
    contact: 'Ricardo / Andrea',
    phone: '3384-5096',
    email: 'splash@splashconsultoria.com.br',
    postalCode: '04334-100',
    cnpj: '13.054.383/0001-22',
    stateRegistration: '',
    website: ''
  },
  {
    id: 'peg-limp',
    name: 'PEGLIMP - DISTR. DE PRODUTOS DE LIMPEZA',
    address: 'Rua São Leonardo, 35 - Freguesia do Ó - São Paulo - SP',
    identity: 'Contato: Antonio Roberto',
    taxId: '61.621.272/0001-01',
    logoPath: '/logoPegLimp.png',
    contact: 'Antonio Roberto',
    phone: '3991-0045 / 93092-8003',
    email: 'peglimp@terra.com.br',
    postalCode: '61.621.272/0001-01',
    cnpj: '61.621.272/0001-01',
    stateRegistration: '',
    website: ''
  },
  {
    id: 'grapa-clean',
    name: 'GRAPACLEAN INSUMOS CORPORATIVOS EIRELI',
    address: 'Av. Manuel Alves Soares, 199 - Jd. Colonial - São Paulo SP - CEP 04821-270',
    identity: 'Telefone: (11) 5663-5800 At.: Rafael',
    taxId: '',
    logoPath: '/logoGrapaClean.png',
    contact: 'Rafael',
    phone: '(11) 5663-5800',
    email: '',
    postalCode: '04821-270',
    cnpj: '33.807.887/0001-83',
    stateRegistration: '126.188.105.117',
    website: 'www.grapaclean.com.br'
  }
];