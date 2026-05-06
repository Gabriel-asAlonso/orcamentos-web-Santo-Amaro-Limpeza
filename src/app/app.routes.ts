import { Routes } from '@angular/router';
import { Login } from './login/login';
import { orcamentos } from './pages/orcamentos/orcamentos';

export const routes: Routes = [
  {
    path: '',
    pathMatch: 'full',
    redirectTo: 'login'
  },
  {
    path: 'login',
    component: Login
  },
  {
    path: 'orcamentos',
    component: orcamentos
  },
  {
    path: '**',
    redirectTo: 'login'
  }
];
