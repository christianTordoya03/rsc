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
  
  modoVista: 'DISTRIBUIDORES' | 'INDUSTRIAS' = 'DISTRIBUIDORES';

  showForm = false;
  isEditing = false;
  isLoading = false;
  currentEditId: string | null = null;
  showDeleteConfirm = false;
  recordToDelete: string | null = null;
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

  cambiarModoVista(modo: 'DISTRIBUIDORES' | 'INDUSTRIAS') {
    this.modoVista = modo;
    this.paginaActual = 1;
  }

  async cargarClientes() {
    this.isLoading = true;
    try {
      let todosLosClientes: any[] = [];
      let desde = 0;
      const paso = 500; 

      while (true) {
        const { data, error } = await this.supabase.from('clientes')
          .select('*')
          .order('id', { ascending: true }) 
          .range(desde, desde + paso - 1); 

        if (error) throw error;
        if (!data || data.length === 0) break;

        todosLosClientes.push(...data);
        if (data.length < paso) break;
        desde += paso; 
      }
      
      todosLosClientes.sort((a, b) => (a.razon_social || '').localeCompare(b.razon_social || ''));
      this.registros = todosLosClientes;
    } catch (error: any) {
      this.toast.showError('Error al cargar la base de datos: ' + error.message);
    } finally {
      this.isLoading = false;
    }
  }

  // ==== FILTRO BLINDADO POR PESTAÑA ====
  get clientesFiltrados() {
    let base = this.registros.filter(c => {
      if (this.modoVista === 'INDUSTRIAS') return c.tipo_cliente === 'B2B/INDUSTRIAS';
      return c.tipo_cliente !== 'B2B/INDUSTRIAS'; // Distribuidores
    });

    if (!this.terminoBusqueda) return base;
    
    const term = this.terminoBusqueda.toLowerCase();
    return base.filter(c => 
      (c.razon_social && c.razon_social.toLowerCase().includes(term)) ||
      (c.ruc && c.ruc.includes(term)) ||
      (c.codigo_cliente && c.codigo_cliente.toLowerCase().includes(term))
    );
  }

  onSearchChange() { this.paginaActual = 1; }

  get totalPaginas() { return Math.ceil(this.clientesFiltrados.length / this.itemsPorPagina) || 1; }

  cambiarPagina(delta: number) {
    const nueva = this.paginaActual + delta;
    if (nueva >= 1 && nueva <= this.totalPaginas) {
      this.paginaActual = nueva;
    }
  }

  getCellValueText(cell: any): string {
    if (!cell || cell.value === null || cell.value === undefined) return '';
    if (typeof cell.value === 'string') return cell.value.trim().toUpperCase();
    if (typeof cell.value === 'number') return String(cell.value).trim().toUpperCase();
    if (typeof cell.value === 'object') {
      if (cell.value.result !== undefined) return String(cell.value.result).trim().toUpperCase();
      if (cell.value.richText) return cell.value.richText.map((rt: any) => rt.text).join('').trim().toUpperCase();
      if (cell.value.text) return String(cell.value.text).trim().toUpperCase();
    }
    return String(cell.value).trim().toUpperCase();
  }

  async importarExcel(event: any) {
    const file = event.target.files[0];
    if (!file) return;

    this.isLoading = true;
    this.toast.showInfo(`Sincronizando Excel en ${this.modoVista}...`);

    try {
      const workbook = new ExcelJS.Workbook();
      const arrayBuffer = await file.arrayBuffer();
      await workbook.xlsx.load(arrayBuffer);
      const worksheet = workbook.worksheets[0]; 

      let headerRow = -1;
      let cCod = -1, cNom = -1, cRuc = -1, cDir = -1, cDist = -1, cZona = -1, cTel = -1;

      worksheet.eachRow((row: any, rowNum: number) => {
        if (rowNum <= 5 && headerRow === -1) {
          row.eachCell((cell: any, colNum: number) => {
            const val = this.getCellValueText(cell);
            
            if (val === 'DESTINAT.' || val === 'CODIGO' || val === 'CÓDIGO') { cCod = colNum; headerRow = rowNum; }
            if (val === 'CLIENTE' || val.includes('RAZON SOCIAL')) { cNom = colNum; headerRow = rowNum; }
            if (val === 'RUC') cRuc = colNum;
            if (val === 'DISTRITO DM' || val === 'DISTRITO') cDist = colNum;
            if (val === 'DIRECCION DM' || (val.includes('DIRECCION') && !val.includes('DM'))) cDir = colNum;
            if (val === 'ZONA') cZona = colNum;
            if (val.includes('TEL')) cTel = colNum;
          });
        }
      });

      if (headerRow === -1 || cNom === -1 || cCod === -1) {
        throw new Error('No se detectó el formato correcto. Faltan columnas de Código (o DESTINAT.) y Razón Social (o CLIENTE).');
      }

      const clientesParaSubir: any[] = [];
      const clientesUnicosMap = new Map();
      
      worksheet.eachRow((row: any, rowNum: number) => {
        if (rowNum > headerRow) { 
          const codigoBruto = this.getCellValueText(row.getCell(cCod)); 
          const nombre = this.getCellValueText(row.getCell(cNom));  

          if (codigoBruto && nombre && nombre !== '0') {
            // BLINDAJE: Aislamos los códigos B2B con un prefijo
            const codigo = this.modoVista === 'INDUSTRIAS' 
                           ? 'B2B-' + codigoBruto.replace('B2B-', '') 
                           : codigoBruto.replace('B2B-', '');

            const clienteObj = {
              codigo_cliente: codigo, 
              razon_social: nombre,
              ruc: cRuc > -1 ? this.getCellValueText(row.getCell(cRuc)) : null,
              distrito: cDist > -1 ? this.getCellValueText(row.getCell(cDist)) : null,
              zona: cZona > -1 ? this.getCellValueText(row.getCell(cZona)) : null,
              direccion: cDir > -1 ? this.getCellValueText(row.getCell(cDir)) : 'Sin dirección',
              tipo_cliente: this.modoVista === 'INDUSTRIAS' ? 'B2B/INDUSTRIAS' : 'DISTRIBUIDORES',
              telefono: cTel > -1 ? this.getCellValueText(row.getCell(cTel)) : '',
              contacto: '',
              correo: '',
              estado: true
            };
            clientesUnicosMap.set(codigo, clienteObj);
          }
        }
      });

      const dataLimpia = Array.from(clientesUnicosMap.values());

      if (dataLimpia.length > 0) {
        this.toast.showInfo(`Sincronizando ${dataLimpia.length} clientes en ${this.modoVista}...`);

        const chunkSize = 500;
        for (let i = 0; i < dataLimpia.length; i += chunkSize) {
          const chunk = dataLimpia.slice(i, i + chunkSize);
          const { error } = await this.supabase.from('clientes').upsert(chunk, { onConflict: 'codigo_cliente' });
          if (error) throw error;
        }

        this.toast.showSuccess(`Base de datos actualizada correctamente en ${this.modoVista}`);
        await this.cargarClientes(); 
      } else {
        this.toast.showWarning('El Excel está vacío o le faltan códigos.');
      }
    } catch (error: any) {
      this.toast.showError('Error al procesar Excel: ' + error.message);
    } finally {
      this.isLoading = false;
      this.fileInput.nativeElement.value = ''; 
    }
  }

  async onSubmit() {
    if (this.clienteForm.invalid) return this.toast.showError('Complete la Razón Social y el Código');
    this.isLoading = true;
    const v = this.clienteForm.value;
    
    // BLINDAJE EN FORMULARIO MANUAL
    let codFinal = v.codigo_cliente?.trim().toUpperCase() || null;
    if (codFinal && this.modoVista === 'INDUSTRIAS') {
       codFinal = 'B2B-' + codFinal.replace('B2B-', '');
    }

    const payload = {
      codigo_cliente: codFinal,
      ruc: v.ruc?.trim() || null,
      razon_social: v.razon_social.trim().toUpperCase(),
      direccion: v.direccion?.trim().toUpperCase() || null,
      distrito: v.distrito?.trim().toUpperCase() || null,
      zona: v.zona?.trim().toUpperCase() || null,
      tipo_cliente: this.modoVista === 'INDUSTRIAS' ? 'B2B/INDUSTRIAS' : 'DISTRIBUIDORES',
      telefono: v.telefono?.trim() || null,
      contacto: v.contacto?.trim().toUpperCase() || null,
      correo: v.correo?.trim() || null,
      estado: v.estado
    };

    try {
      if (!this.isEditing) {
        const { error } = await this.supabase.from('clientes').upsert([payload], { onConflict: 'codigo_cliente' });
        if (error) throw error;
        this.toast.showSuccess('Cliente guardado con éxito');
      } else {
        const { error } = await this.supabase.from('clientes').update(payload).eq('id', this.currentEditId);
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

  async executeDelete() {
    if (!this.recordToDelete) return;
    this.isLoading = true;
    try {
      const { error } = await this.supabase.clientes.update({ estado: false }).eq('id', this.recordToDelete);
      if (error) throw error;
      
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

  // ==== BORRADO MASIVO BLINDADO POR PESTAÑA ====
  async executeDeleteAll() {
    this.isLoading = true;
    try {
      const clientesABorrar = this.registros.filter(c => {
        if (this.modoVista === 'INDUSTRIAS') return c.tipo_cliente === 'B2B/INDUSTRIAS';
        return c.tipo_cliente !== 'B2B/INDUSTRIAS';
      });

      const ids = clientesABorrar.map(c => c.id);

      if (ids.length === 0) {
        this.toast.showWarning(`No hay clientes de ${this.modoVista} para eliminar.`);
        this.cancelDeleteAll();
        this.isLoading = false;
        return;
      }

      const chunkSize = 300;
      for (let i = 0; i < ids.length; i += chunkSize) {
        const chunk = ids.slice(i, i + chunkSize);
        const { error } = await this.supabase.clientes.delete().in('id', chunk);
        if (error) throw error;
      }
      
      this.registros = this.registros.filter(c => !ids.includes(c.id));
      this.paginaActual = 1;
      
      this.toast.showSuccess(`Directorio de ${this.modoVista} vaciado correctamente.`);
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
      
      // Ocultamos el prefijo visualmente para no confundir al usuario
      let codVisual = registro.codigo_cliente;
      if (this.modoVista === 'INDUSTRIAS' && codVisual?.startsWith('B2B-')) {
         codVisual = codVisual.replace('B2B-', '');
      }

      this.clienteForm.patchValue({
        ...registro,
        codigo_cliente: codVisual
      });
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