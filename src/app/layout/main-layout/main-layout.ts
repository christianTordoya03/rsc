import { Component, OnInit, inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RouterOutlet } from '@angular/router';
// Importa tus componentes Sidebar y Topbar
import { SidebarComponent } from '../sidebar/sidebar'; 
import { TopbarComponent } from '../topbar/topbar';
import { ToastContainerComponent } from '../../shared/components/toast/toast.component';
import { SupabaseService } from '../../core/services/supabase';

@Component({
  selector: 'app-main-layout',
  standalone: true,
  imports: [CommonModule, RouterOutlet, SidebarComponent, TopbarComponent, ToastContainerComponent],
  templateUrl: './main-layout.html'
})
export class MainLayoutComponent implements OnInit {
  isSidebarOpen = false;
  rolUsuario = ''; // <-- Nueva variable para guardar el rol
  
  private supabase = inject(SupabaseService);

  async ngOnInit() {
    // Al cargar el layout, preguntamos quién está conectado
    const { data: { session } } = await this.supabase.auth.getSession();
    if (session) {
      const { data } = await this.supabase.from('perfiles')
        .select('rol')
        .eq('id', session.user.id)
        .single();
        
      if (data) {
        this.rolUsuario = data.rol;
      }
    }
  }

  toggleSidebar() {
    this.isSidebarOpen = !this.isSidebarOpen;
  }
}