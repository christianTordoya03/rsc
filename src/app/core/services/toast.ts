import { Injectable, signal } from '@angular/core';

export interface Toast {
  id: number;
  message: string;
  type: 'success' | 'error' | 'warning' | 'info';
}

@Injectable({
  providedIn: 'root'
})
export class ToastService {
  toasts = signal<Toast[]>([]);
  private counter = 0;

  showSuccess(message: string) { this.addToast(message, 'success'); }
  showError(message: string) { this.addToast(message, 'error'); }
  showWarning(message: string) { this.addToast(message, 'warning'); }
  showInfo(message: string) { this.addToast(message, 'info'); }

  private addToast(message: string, type: Toast['type']) {
    const id = this.counter++;
    this.toasts.update(current => [...current, { id, message, type }]);
    // La notificación desaparece sola en 3 segundos
    setTimeout(() => this.removeToast(id), 3000);
  }

  removeToast(id: number) {
    this.toasts.update(current => current.filter(t => t.id !== id));
  }
}