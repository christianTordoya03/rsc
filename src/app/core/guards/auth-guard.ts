import { CanActivateFn } from '@angular/router';

export const authGuard: CanActivateFn = async (route, state) => {
  // PLAN B: Candado abierto temporalmente para crear el primer usuario
  return true; 
};