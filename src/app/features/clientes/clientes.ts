import { Component, OnInit, ElementRef, ViewChild } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormBuilder, FormGroup, ReactiveFormsModule, Validators } from '@angular/forms';
import * as ExcelJS from 'exceljs';
import { SupabaseService } from '../../core/services/supabase';
import { ToastService } from '../../core/services/toast';
import { ConfirmModalComponent } from '../../shared/components/confirm-modal/confirm-modal';

@Component({
  selector: 'app-clientes',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule, ConfirmModalComponent],
  templateUrl: './clientes.html'
})
export class Clientes implements OnInit {
  @ViewChild('fileInput') fileInput!: ElementRef;
  
  showForm = false;
  isEditing = false;
  isLoading = false;
  currentEditId: string | null = null;
  showDeleteConfirm = false;
  recordToDelete: string | null = null;
  
  registros: any[] = [];
  clienteForm: FormGroup;
  userRole = ''; // <-- Nueva variable para la jerarquía

  constructor(
    private fb: FormBuilder,
    private supabase: SupabaseService,
    private toast: ToastService
  ) {
    this.clienteForm = this.fb.group({
      ruc: ['', [Validators.required, Validators.pattern(/^\d{11}$/)]],
      razon_social: ['', Validators.required],
      direccion: ['', Validators.required],
      distrito: ['', Validators.required],
      telefono: [''], contacto: [''], correo: ['', Validators.email], estado: [true]
    });
  }

  async ngOnInit() {
    // 1. Obtenemos quién está conectado para darle permisos
    const { data: { session } } = await this.supabase.auth.getSession();
    if (session) {
      const { data } = await this.supabase.perfiles.select('rol').eq('id', session.user.id).single();
      this.userRole = data?.rol || '';
    }
    
    await this.cargarClientes();
  }

  async cargarClientes() {
    this.isLoading = true;
    try {
      const { data, error } = await this.supabase.clientes.select('*').order('creado_en', { ascending: false });
      if (error) throw error;
      this.registros = data || [];
    } catch (error: any) {
      this.toast.showError('Error al cargar clientes');
    } finally {
      this.isLoading = false;
    }
  }

  // --- NUEVO: FUNCIÓN DE CARGA MASIVA DE EXCEL ---
  async importarExcel(event: any) {
    const file = event.target.files[0];
    if (!file) return;

    this.isLoading = true;
    this.toast.showInfo('Leyendo archivo Excel...');

    try {
      const workbook = new ExcelJS.Workbook();
      const arrayBuffer = await file.arrayBuffer();
      await workbook.xlsx.load(arrayBuffer);
      const worksheet = workbook.worksheets[0]; // Lee la primera hoja

      const nuevosClientes: any[] = [];
      
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) { // Saltamos la fila 1 (Cabeceras)
          nuevosClientes.push({
            ruc: row.getCell(1).text?.trim(),
            razon_social: row.getCell(2).text?.trim(),
            direccion: row.getCell(3).text?.trim() || 'Sin dirección',
            distrito: row.getCell(4).text?.trim() || 'No especificado',
            telefono: row.getCell(5).text?.trim() || '',
            contacto: row.getCell(6).text?.trim() || '',
            correo: row.getCell(7).text?.trim() || '',
            estado: true
          });
        }
      });

      // Filtramos filas vacías
      const dataValida = nuevosClientes.filter(c => c.ruc && c.razon_social);

      if (dataValida.length > 0) {
        // Inserción Masiva a Supabase
        const { data, error } = await this.supabase.clientes.insert(dataValida).select();
        if (error) throw error;
        
        this.registros = [...(data || []), ...this.registros];
        this.toast.showSuccess(`${dataValida.length} clientes importados correctamente`);
      } else {
        this.toast.showWarning('El Excel está vacío o tiene formato incorrecto');
      }
    } catch (error: any) {
      this.toast.showError('Error al procesar Excel: ' + error.message);
    } finally {
      this.isLoading = false;
      this.fileInput.nativeElement.value = ''; // Resetea el input
    }
  }

  async onSubmit() {
    if (this.clienteForm.invalid) return this.toast.showError('Complete los campos');
    this.isLoading = true;
    const v = this.clienteForm.value;
    try {
      if (!this.isEditing) {
        const { data, error } = await this.supabase.clientes.insert(v).select().single();
        if (error) throw error;
        this.registros.unshift(data);
        this.toast.showSuccess('Cliente registrado');
      } else {
        const { data, error } = await this.supabase.clientes.update(v).eq('id', this.currentEditId).select().single();
        if (error) throw error;
        const idx = this.registros.findIndex(r => r.id === this.currentEditId);
        if (idx !== -1) this.registros[idx] = data;
        this.toast.showInfo('Cliente actualizado');
      }
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
      const { error } = await this.supabase.clientes.delete().eq('id', this.recordToDelete);
      if (error) throw error;
      this.registros = this.registros.filter(r => r.id !== this.recordToDelete);
      this.toast.showSuccess('Cliente eliminado permanentemente');
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
      this.clienteForm.patchValue(registro);
    } else {
      this.isEditing = false; this.currentEditId = null;
      this.clienteForm.reset({ estado: true });
    }
    this.showForm = true;
  }
  closeForm() { this.showForm = false; }
  confirmDelete(id: string) { this.recordToDelete = id; this.showDeleteConfirm = true; }
  cancelDelete() { this.showDeleteConfirm = false; this.recordToDelete = null; }
}