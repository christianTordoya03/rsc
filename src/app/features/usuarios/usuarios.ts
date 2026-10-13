import { Component, OnInit, ElementRef, ViewChild } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormBuilder, FormGroup, ReactiveFormsModule, Validators, FormsModule } from '@angular/forms';
import * as ExcelJS from 'exceljs';
import { SupabaseService } from '../../core/services/supabase';
import { ToastService } from '../../core/services/toast';
import { ConfirmModalComponent } from '../../shared/components/confirm-modal/confirm-modal';

@Component({
  selector: 'app-usuarios',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule, FormsModule, ConfirmModalComponent],
  templateUrl: './usuarios.html'
})
export class Usuarios implements OnInit {
  @ViewChild('fileInput') fileInput!: ElementRef;
  
  showForm = false;
  isEditing = false;
  isLoading = false;
  currentEditId: string | null = null;
  showDeleteConfirm = false;
  recordToDelete: string | null = null;
  
  registros: any[] = [];
  usuarioForm: FormGroup;
  userRole = '';

  // --- VARIABLES DE FILTRO Y PAGINACIÓN ---
  searchTerm = '';
  roleFilter = '';
  currentPage = 1;
  itemsPerPage = 15;

  constructor(
    private fb: FormBuilder,
    private supabase: SupabaseService,
    private toast: ToastService
  ) {
    this.usuarioForm = this.fb.group({
      dni: ['', [Validators.required, Validators.pattern(/^[a-zA-Z0-9]{8,12}$/)]],
      nombre_completo: ['', Validators.required],
      rol: ['OPERACIONES', Validators.required]
    });
  }

  async ngOnInit() {
    const { data: { session } } = await this.supabase.auth.getSession();
    if (session) {
      const { data } = await this.supabase.perfiles.select('rol').eq('id', session.user.id).single();
      this.userRole = data?.rol || '';
    }
    await this.cargarUsuarios();
  }

  async cargarUsuarios() {
    this.isLoading = true;
    try {
      const { data, error } = await this.supabase.perfiles.select('*');
      if (error) throw error;
      
      // ORDENAMIENTO ESTRICTO: 1° ADMIN, 2° LOGISTICA, 3° OPERACIONES
      const rolesOrder: { [key: string]: number } = {
        'ADMIN': 1,
        'LOGISTICA': 2,
        'OPERACIONES': 3
      };

      this.registros = (data || []).sort((a, b) => {
        const rankA = rolesOrder[a.rol] || 4;
        const rankB = rolesOrder[b.rol] || 4;
        
        // Si tienen distinto rol, los ordenamos por el número asignado (1, 2, 3)
        if (rankA !== rankB) {
          return rankA - rankB;
        }
        // Si tienen el MISMO rol, los ordenamos alfabéticamente por nombre
        return a.nombre_completo.localeCompare(b.nombre_completo);
      });

    } catch (error: any) {
      this.toast.showError('Error al cargar usuarios');
    } finally {
      this.isLoading = false;
    }
  }

  // --- LÓGICA DE FILTROS Y PAGINACIÓN ---
  get filteredRegistros() {
    return this.registros.filter(r => {
      const matchSearch = !this.searchTerm || 
        r.nombre_completo.toLowerCase().includes(this.searchTerm.toLowerCase()) || 
        r.dni.toLowerCase().includes(this.searchTerm.toLowerCase());
      const matchRole = !this.roleFilter || r.rol === this.roleFilter;
      return matchSearch && matchRole;
    });
  }

  get totalPages() {
    return Math.ceil(this.filteredRegistros.length / this.itemsPerPage) || 1;
  }

  get paginatedRegistros() {
    const start = (this.currentPage - 1) * this.itemsPerPage;
    return this.filteredRegistros.slice(start, start + this.itemsPerPage);
  }

  onFilterChange() {
    this.currentPage = 1; 
  }

  nextPage() {
    if (this.currentPage < this.totalPages) this.currentPage++;
  }

  prevPage() {
    if (this.currentPage > 1) this.currentPage--;
  }
  // ---------------------------------------

  async importarExcel(event: any) {
    const file = event.target.files[0];
    if (!file) return;

    this.isLoading = true;
    this.toast.showInfo('Creando cuentas de usuario. Esto puede tardar...');

    try {
      const workbook = new ExcelJS.Workbook();
      const arrayBuffer = await file.arrayBuffer();
      await workbook.xlsx.load(arrayBuffer);
      const worksheet = workbook.worksheets[0];

      const nuevosUsuarios: any[] = [];
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {
          nuevosUsuarios.push({
            dni: row.getCell(1).text?.replace(/[^a-zA-Z0-9]/g, '').trim(),
            nombre_completo: row.getCell(2).text?.trim(),
            rol: row.getCell(3).text?.trim().toUpperCase() || 'OPERACIONES'
          });
        }
      });

      const dataValida = nuevosUsuarios.filter(u => u.dni && u.nombre_completo);
      let exitosos = 0;

      for (const u of dataValida) {
        const pseudoEmail = `${u.dni}@rscdiesel.com`;
        const { data: authData, error: authErr } = await this.supabase.auth.signUp({
          email: pseudoEmail, password: u.dni
        });

        if (!authErr && authData.user) {
          const { data: perfilData } = await this.supabase.perfiles.insert({
            id: authData.user.id, dni: u.dni, nombre_completo: u.nombre_completo, rol: u.rol
          }).select().single();
          
          if (perfilData) {
            this.registros.unshift(perfilData);
            exitosos++;
          }
        }
      }
      
      await this.cargarUsuarios(); 
      this.toast.showSuccess(`${exitosos} usuarios creados exitosamente`);
    } catch (error: any) {
      this.toast.showError('Error al procesar Excel');
    } finally {
      this.isLoading = false;
      this.fileInput.nativeElement.value = '';
    }
  }

  async onSubmit() {
    if (this.usuarioForm.invalid) return this.toast.showError('Revise el DNI y los nombres.');
    this.isLoading = true;
    const v = this.usuarioForm.value;
    try {
      if (!this.isEditing) {
        const pseudoEmail = `${v.dni}@rscdiesel.com`;
        const { data: authData, error: authErr } = await this.supabase.auth.signUp({ email: pseudoEmail, password: v.dni });
        if (authErr) throw authErr;
        
        if (authData.user) {
          await this.supabase.perfiles.insert({
            id: authData.user.id, dni: v.dni, nombre_completo: v.nombre_completo, rol: v.rol
          });
          this.toast.showSuccess('Usuario creado correctamente');
        }
      } else {
        await this.supabase.perfiles.update({ nombre_completo: v.nombre_completo, rol: v.rol }).eq('id', this.currentEditId);
        this.toast.showInfo('Perfil actualizado');
      }
      
      await this.cargarUsuarios(); 
      this.closeForm();
    } catch (error: any) {
      this.toast.showError('Error: ' + error.message);
    } finally {
      this.isLoading = false;
    }
  }

  async executeDelete() {
    if (!this.recordToDelete) return;
    this.isLoading = true;
    try {
      const { error } = await this.supabase.perfiles.delete().eq('id', this.recordToDelete);
      if (error) throw error;
      
      await this.cargarUsuarios(); 
      this.toast.showSuccess('Acceso de usuario revocado');
      this.cancelDelete();
    } catch (error: any) {
      this.toast.showError('Error: ' + error.message);
    } finally {
      this.isLoading = false;
    }
  }

  openForm(registro?: any) {
    if (registro) {
      this.isEditing = true; this.currentEditId = registro.id;
      this.usuarioForm.patchValue(registro);
      this.usuarioForm.get('dni')?.disable();
    } else {
      this.isEditing = false; this.currentEditId = null;
      this.usuarioForm.enable();
      this.usuarioForm.reset({ rol: 'OPERACIONES' });
    }
    this.showForm = true;
  }
  closeForm() { this.showForm = false; }
  confirmDelete(id: string) { this.recordToDelete = id; this.showDeleteConfirm = true; }
  cancelDelete() { this.showDeleteConfirm = false; this.recordToDelete = null; }
}