import { Routes } from '@angular/router';

import { authGuard } from './core/guards/auth-guard';
import { MainLayoutComponent } from './layout/main-layout/main-layout';


export const routes: Routes = [
  {
    path: '',
    component: MainLayoutComponent,
    canActivate: [authGuard], // Protegeremos todas las rutas internas
    children: [
      {
        path: '',
        redirectTo: 'dashboard',
        pathMatch: 'full'
      },
      {
        path: 'dashboard',
        loadComponent: () => import('./features/dashboard/dashboard').then(m => m.Dashboard)
      },
      {
        path: 'programacion',
        loadComponent: () => import('./features/programacion-rutas/programacion-rutas').then(m => m.ProgramacionRutasComponent)
      },
      {
        path: 'ventas-retornos',
        loadComponent: () => import('./features/ventas-retornos/ventas-retornos').then(m => m.VentasRetornos)
      },
      {
        path: 'control-combustible',
        loadComponent: () => import('./features/combustible-peaje/combustible-peaje').then(m => m.CombustiblePeaje)
      },
      {
        path: 'productividad-nocturna',
        loadComponent: () => import('./features/productividad-nocturna/productividad-nocturna').then(m => m.ProductividadNocturna)
      },
      {
        path: 'clientes',
        loadComponent: () => import('./features/clientes/clientes').then(m => m.Clientes)
      },
      {
        path: 'usuarios',
        loadComponent: () => import('./features/usuarios/usuarios').then(m => m.Usuarios)
      },
      {
        path: 'manual-usuario',
        loadComponent: () => import('./features/manual-usuario/manual-usuario').then(m => m.ManualUsuario)
      }
    ]
  },
  {
    path: 'operador',
    loadComponent: () => import('./features/operador/operador').then(m => m.OperadorComponent)
  },
  {
    path: 'login',
    loadComponent: () => import('./features/login/login').then(m => m.Login)
  },
  {
    path: '**',
    redirectTo: 'login'
  }
];