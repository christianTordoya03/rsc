import { Routes } from '@angular/router';
import { authGuard } from './core/guards/auth-guard';
import { MainLayoutComponent } from './layout/main-layout/main-layout';

export const routes: Routes = [
  // 1. ZONA ADMINISTRATIVA (Con Menú Lateral)
  {
    path: '',
    component: MainLayoutComponent,
    canActivate: [authGuard], 
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
      },
      // AQUÍ ESTÁ EL MÓDULO (Integrado al sistema y con su nuevo nombre)
      {
        path: 'operacion-nocturna',
        loadComponent: () => import('./features/productividad-nocturna/operacion-nocturna').then(m => m.OperacionNocturnaComponent)
      }
    ]
  },

  // 2. ZONAS OPERATIVAS (A Pantalla Completa, Sin Menú Lateral)
  {
    path: 'operador',
    loadComponent: () => import('./features/operador/operador').then(m => m.OperadorComponent)
  },

  // 3. LOGIN Y RUTAS DESCONOCIDAS
  {
    path: 'login',
    loadComponent: () => import('./features/login/login').then(m => m.Login)
  },
  {
    path: '**',
    redirectTo: 'login'
  }
];