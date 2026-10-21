import { Component, OnInit, ElementRef, ViewChild } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormBuilder, FormGroup, ReactiveFormsModule, Validators, FormsModule } from '@angular/forms';
import * as ExcelJS from 'exceljs';
import { SupabaseService } from '../../core/services/supabase';
import { ToastService } from '../../core/services/toast';
import { ConfirmModalComponent } from '../../shared/components/confirm-modal/confirm-modal';

@Component({
  selector: 'app-clientes',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule, FormsModule, ConfirmModalComponent],
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

  // Variables para la eliminación masiva
  showDeleteAllConfirm = false;
  
  registros: any[] = [];
  clienteForm: FormGroup;
  userRole = ''; 

  terminoBusqueda = '';
  paginaActual = 1;
  itemsPorPagina = 50;

  constructor(
    private fb: FormBuilder,
    private supabase: SupabaseService,
    private toast: ToastService
  ) {
    this.clienteForm = this.fb.group({
      ruc: [''],
      razon_social: ['', Validators.required],
      codigo_cliente: ['', Validators.required],
      direccion: [''],
      distrito: [''],
      zona: [''],
      tipo_cliente: [''],
      telefono: [''], 
      contacto: [''], 
      correo: ['', Validators.email], 
      estado: [true]
    });
  }

  async ngOnInit() {
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
      let todosLosClientes: any[] = [];
      let hayMasDatos = true;
      let desde = 0;
      const limite = 1000; // Máximo permitido por Supabase por viaje

      // Hacemos viajes iterativos hasta traer a los 1800+ clientes
      while (hayMasDatos) {
        const { data, error } = await this.supabase.from('clientes')
          .select('*')
          .order('razon_social', { ascending: true })
          .range(desde, desde + limite - 1); // Rango dinámico: 0-999, luego 1000-1999, etc.

        if (error) throw error;

        if (data && data.length > 0) {
          todosLosClientes = [...todosLosClientes, ...data];
          desde += limite; // Preparamos el inicio del siguiente viaje
          
          // Si nos trajo menos de 1000, significa que es la última página y terminamos
          if (data.length < limite) {
            hayMasDatos = false;
          }
        } else {
          hayMasDatos = false;
        }
      }

      this.registros = todosLosClientes;
    } catch (error: any) {
      this.toast.showError('Error al cargar la base de datos completa: ' + error.message);
    } finally {
      this.isLoading = false;
    }
  }

  get clientesFiltrados() {
    if (!this.terminoBusqueda) return this.registros;
    
    const term = this.terminoBusqueda.toLowerCase();
    return this.registros.filter(c => 
      (c.razon_social && c.razon_social.toLowerCase().includes(term)) ||
      (c.ruc && c.ruc.includes(term)) ||
      (c.codigo_cliente && c.codigo_cliente.toLowerCase().includes(term))
    );
  }

  onSearchChange() {
    this.paginaActual = 1; 
  }

  get totalPaginas() {
    return Math.ceil(this.clientesFiltrados.length / this.itemsPorPagina) || 1;
  }

  cambiarPagina(delta: number) {
    const nueva = this.paginaActual + delta;
    if (nueva >= 1 && nueva <= this.totalPaginas) {
      this.paginaActual = nueva;
    }
  }

  async importarExcel(event: any) {
    const file = event.target.files[0];
    if (!file) return;

    this.isLoading = true;
    this.toast.showInfo('Leyendo base de datos...');

    try {
      const workbook = new ExcelJS.Workbook();
      const arrayBuffer = await file.arrayBuffer();
      await workbook.xlsx.load(arrayBuffer);
      const worksheet = workbook.worksheets[0]; 

      const clientesParaSubir: any[] = [];
      
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) { 
          const destinat = row.getCell(1).text?.trim(); 
          const cliente = row.getCell(2).text?.trim();  

          if (destinat && cliente) {
            clientesParaSubir.push({
              codigo_cliente: destinat, 
              razon_social: cliente,
              ruc: row.getCell(3).text?.trim() || null,
              distrito: row.getCell(4).text?.trim() || null,
              zona: row.getCell(5).text?.trim() || null,
              direccion: row.getCell(6).text?.trim() || 'Sin dirección',
              tipo_cliente: row.getCell(8).text?.trim() || null,
              telefono: row.getCell(9).text?.trim() || '',
              contacto: '',
              correo: '',
              estado: true
            });
          }
        }
      });

      if (clientesParaSubir.length > 0) {
        // FILTRO ANTI-DUPLICADOS DENTRO DEL MISMO EXCEL
        const clientesUnicosMap = new Map();
        for (const c of clientesParaSubir) {
          clientesUnicosMap.set(c.codigo_cliente, c);
        }
        const dataLimpia = Array.from(clientesUnicosMap.values());

        this.toast.showInfo(`Sincronizando ${dataLimpia.length} clientes únicos...`);

        const chunkSize = 500;
        for (let i = 0; i < dataLimpia.length; i += chunkSize) {
          const chunk = dataLimpia.slice(i, i + chunkSize);
          const { error } = await this.supabase.from('clientes')
            .upsert(chunk, { onConflict: 'codigo_cliente' });
          if (error) throw error;
        }

        this.toast.showSuccess(`Base de datos actualizada con éxito`);
        await this.cargarClientes(); 
      } else {
        this.toast.showWarning('El Excel está vacío o le faltan códigos de Destinatario');
      }
    } catch (error: any) {
      this.toast.showError('Error crítico al procesar Excel: ' + error.message);
    } finally {
      this.isLoading = false;
      this.fileInput.nativeElement.value = ''; 
    }
  }

  async onSubmit() {
    if (this.clienteForm.invalid) return this.toast.showError('Complete la Razón Social y el Código');
    this.isLoading = true;
    const v = this.clienteForm.value;
    try {
      if (!this.isEditing) {
        const { error } = await this.supabase.from('clientes').upsert([v], { onConflict: 'codigo_cliente' });
        if (error) throw error;
        this.toast.showSuccess('Cliente guardado con éxito');
      } else {
        const { error } = await this.supabase.from('clientes').update(v).eq('id', this.currentEditId);
        if (error) throw error;
        this.toast.showInfo('Cliente actualizado');
      }
      
      await this.cargarClientes(); 
      this.closeForm();
    } catch (error: any) {
      this.toast.showError('Error: ' + error.message);
    } finally {
      this.isLoading = false;
    }
  }

  // ELIMINAR UNO SOLO
  async executeDelete() {
    if (!this.recordToDelete) return;
    this.isLoading = true;
    try {
      // Magia aquí: Cambiamos .delete() por .update({ estado: false })
      const { error } = await this.supabase.clientes.update({ estado: false }).eq('id', this.recordToDelete);
      
      if (error) throw error;
      
      // Lo ocultamos visualmente de la tabla para que parezca eliminado
      this.registros = this.registros.filter(r => r.id !== this.recordToDelete);
      this.toast.showSuccess('Cliente eliminado correctamente');
      
      this.cancelDelete();
      
      if (this.clientesFiltrados.length === 0 && this.paginaActual > 1) this.paginaActual--;
    } catch (error: any) {
      this.toast.showError('Error: ' + error.message);
    } finally {
      this.isLoading = false;
    }
  }

  // NUEVO: ELIMINAR TODOS MASIVAMENTE
  async executeDeleteAll() {
    this.isLoading = true;
    try {
      // Borra todos los registros que tengan un ID (Es decir, todos)
      const { error } = await this.supabase.clientes.delete().not('id', 'is', null);
      if (error) throw error;
      
      this.registros = [];
      this.paginaActual = 1;
      this.toast.showSuccess('Base de datos formateada correctamente.');
      this.cancelDeleteAll();
    } catch (error: any) {
      this.toast.showError('Error al vaciar DB: ' + error.message);
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

  confirmDeleteAll() { this.showDeleteAllConfirm = true; }
  cancelDeleteAll() { this.showDeleteAllConfirm = false; }
}