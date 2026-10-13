import { Component, EventEmitter, OnInit, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Router } from '@angular/router';
import { SupabaseService } from '../../core/services/supabase';

@Component({
  selector: 'app-topbar',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './topbar.html'
})
export class TopbarComponent implements OnInit {
  @Output() openSidebar = new EventEmitter<void>();
  
  // Variables dinámicas para la interfaz
  userName = 'Cargando...';
  userRole = '';
  userInitials = '';

  constructor(
    private supabase: SupabaseService,
    private router: Router
  ) {}

  async ngOnInit() {
    await this.cargarPerfilUsuario();
  }

  async cargarPerfilUsuario() {
    try {
      // 1. Obtenemos la sesión actual
      const { data: { session } } = await this.supabase.auth.getSession();
      
      if (session) {
        // 2. Buscamos el nombre y rol en nuestra tabla perfiles
        const { data: perfil } = await this.supabase.perfiles
          .select('nombre_completo, rol')
          .eq('id', session.user.id)
          .single();

        if (perfil) {
          this.userName = perfil.nombre_completo;
          this.userRole = perfil.rol;
          this.userInitials = this.obtenerIniciales(perfil.nombre_completo);
        }
      }
    } catch (error) {
      console.error('Error cargando perfil:', error);
    }
  }

  // Función para extraer iniciales (Ej: Juan Perez -> JP)
  obtenerIniciales(nombre: string): string {
    const partes = nombre.trim().split(' ');
    if (partes.length === 0) return 'U';
    if (partes.length === 1) return partes[0].charAt(0).toUpperCase();
    return (partes[0].charAt(0) + partes[1].charAt(0)).toUpperCase();
  }

  async cerrarSesion() {
    await this.supabase.auth.signOut();
    this.router.navigate(['/login']);
  }
}