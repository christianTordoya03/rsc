import { Component, inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ToastService } from '../../../core/services/toast';

@Component({
  selector: 'app-toast-container',
  standalone: true,
  imports: [CommonModule],
  template: `
    <!-- Contenedor centrado dinámico: 95% de ancho en móvil, se ajusta en PC -->
    <div class="fixed top-4 left-1/2 -translate-x-1/2 z-[999999] flex flex-col gap-3 pointer-events-none w-[95%] sm:w-auto items-center">
      
      <div *ngFor="let toast of toastService.toasts()"
           class="pointer-events-auto flex items-center gap-3 w-full sm:w-96 p-4 rounded-xl shadow-2xl border transform transition-all duration-300 animate-slide-down"
           [ngClass]="{
             'bg-emerald-50 border-emerald-200 text-emerald-800': toast.type === 'success',
             'bg-red-50 border-red-200 text-red-800': toast.type === 'error',
             'bg-amber-50 border-amber-200 text-amber-800': toast.type === 'warning',
             'bg-blue-50 border-blue-200 text-blue-800': toast.type === 'info'
           }">
        
        <!-- Icono Éxito -->
        <svg *ngIf="toast.type === 'success'" class="w-6 h-6 text-emerald-500 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
        <!-- Icono Error -->
        <svg *ngIf="toast.type === 'error'" class="w-6 h-6 text-red-500 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
        <!-- Icono Alerta -->
        <svg *ngIf="toast.type === 'warning'" class="w-6 h-6 text-amber-500 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"></path></svg>
        <!-- Icono Info (Agregado por precaución) -->
        <svg *ngIf="toast.type === 'info'" class="w-6 h-6 text-blue-500 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
        
        <p class="text-[13px] font-bold flex-1 leading-snug">{{ toast.message }}</p>

        <button (click)="toastService.removeToast(toast.id)" class="text-slate-400 hover:text-slate-600 focus:outline-none p-1 bg-white/50 rounded-full transition-colors">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
        </button>
      </div>
    </div>
  `,
  styles: [`
    /* Cambiamos la animación a slideDown porque ahora sale desde arriba */
    .animate-slide-down { animation: slideDown 0.4s cubic-bezier(0.16, 1, 0.3, 1) forwards; }
    @keyframes slideDown {
      from { transform: translateY(-100%); opacity: 0; }
      to { transform: translateY(0); opacity: 1; }
    }
  `]
})
export class ToastContainerComponent {
  toastService = inject(ToastService);
}