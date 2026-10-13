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
    this.loginForm = this.fb.group({
      dni: ['', [Validators.required, Validators.pattern(/^\d{8}$/)]]
    });
  }

  async onSubmit() {
    if (this.loginForm.invalid) {
      this.loginForm.markAllAsTouched();
      return;
    }

    this.isLoading = true;
    const dni = this.loginForm.value.dni;
    const pseudoEmail = `${dni}@rscdiesel.com`;

    try {
      // 1. Iniciar sesión
      const { data, error } = await this.supabase.auth.signInWithPassword({
        email: pseudoEmail,
        password: dni
      });

      if (error) throw error;

      if (data.session) {
        // 2. Leer el rol del usuario en la tabla perfiles
        const { data: perfil } = await this.supabase.perfiles
          .select('rol')
          .eq('id', data.session.user.id)
          .single();

        this.toast.showSuccess('Bienvenido al sistema');

        // 3. Redirección inteligente basada en el rol
        if (perfil && perfil.rol === 'OPERACIONES') {
          this.router.navigate(['/operador']); // Pantalla móvil exclusiva
        } else {
          this.router.navigate(['/dashboard']); // Pantalla administrativa
        }
      }
    } catch (error: any) {
      this.toast.showError('Acceso denegado. DNI no registrado o incorrecto.');
    } finally {
      this.isLoading = false;
    }
  }
}