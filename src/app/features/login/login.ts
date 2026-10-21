import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormBuilder, FormGroup, ReactiveFormsModule, Validators } from '@angular/forms';
import { Router } from '@angular/router';
import { SupabaseService } from '../../core/services/supabase';
import { ToastService } from '../../core/services/toast';

@Component({
  selector: 'app-login',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule],
  templateUrl: './login.html'
})
export class Login {
  loginForm: FormGroup;
  isLoading = false;

  constructor(
    private fb: FormBuilder,
    private supabase: SupabaseService,
    private router: Router,
    private toast: ToastService
  ) {
    // Un solo campo que acepta hasta 12 caracteres (DNI o CE)
    this.loginForm = this.fb.group({
      documento: ['', [Validators.required, Validators.maxLength(12)]]
    });
  }

  async onSubmit() {
    if (this.loginForm.invalid) {
      this.loginForm.markAllAsTouched();
      return;
    }

    this.isLoading = true;
    const documento = this.loginForm.value.documento.trim();
    
    // Usamos el documento tanto para el email como para la contraseña
    const pseudoEmail = `${documento.toLowerCase()}@rscdiesel.com`;

    try {
      const { data, error } = await this.supabase.auth.signInWithPassword({
        email: pseudoEmail,
        password: documento
      });

      if (error) throw error;

      if (data.session) {
        // SOLO pedimos rol y nombre_completo. ¡Nada de estado!
        const { data: perfil, error: perfilError } = await this.supabase.from('perfiles')
          .select('rol, nombre_completo')
          .eq('id', data.session.user.id)
          .single();

        if (perfilError) throw perfilError;

        const rol = perfil?.rol;
        this.toast.showSuccess(`Bienvenido, ${perfil?.nombre_completo?.split(' ')[0] || 'Usuario'}`);

        // REDIRECCIONES (Corregidas para coincidir con app.routes.ts)
        if (rol === 'NOCTURNO' || documento === '47677704') {
          this.router.navigate(['/operacion-nocturna']); // <- Nuevo nombre
        } else if (rol === 'ADMIN') {
          this.router.navigate(['/ventas-retornos']);
        } else if (rol === 'LOGISTICA') {
          this.router.navigate(['/programacion']); 
        } else if (rol === 'OPERACIONES') {
          this.router.navigate(['/operador']); 
        } else {
          // Si el rol en Supabase no es ninguno de los de arriba, no hará nada, pero avisará:
          this.toast.showWarning(`Rol no reconocido: ${rol}`);
          console.warn(`El rol '${rol}' no coincide con las opciones permitidas para redirección.`);
        }
      }
    } catch (error: any) {
      console.error('Error en Login:', error);
      this.toast.showError('Acceso denegado. Documento incorrecto.');
    } finally {
      this.isLoading = false;
    }
  }
}