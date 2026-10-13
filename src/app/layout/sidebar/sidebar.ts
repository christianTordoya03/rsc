import { Component, EventEmitter, Input, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RouterLink, RouterLinkActive } from '@angular/router'; // Importaciones específicas

@Component({
  selector: 'app-sidebar',
  standalone: true,
  imports: [CommonModule, RouterLink, RouterLinkActive],
  templateUrl: './sidebar.html'
})
export class SidebarComponent {
  @Input() isOpen = false;
  @Output() closeSidebar = new EventEmitter<void>();

  // Estructura de las 6 secciones solicitadas
  menuItems = [
    { path: '/dashboard', label: 'Panel Principal' },
    { path: '/ventas-retornos', label: 'Ventas y Retornos' },
    { path: '/control-combustible', label: 'Control Combustible' },
    { path: '/productividad-nocturna', label: 'Venta Nocturna' },
    { path: '/clientes', label: 'Clientes' },
    { path: '/usuarios', label: 'Usuarios' },
    { path: '/programacion', label: 'Programación de Rutas' },
    { path: '/manual-usuario', label: 'Manual de Usuario' }
  ];

  onClose() {
    this.closeSidebar.emit();
  }
}