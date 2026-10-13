import { Component, OnInit, ElementRef, ViewChild } from '@angular/core';
import { CommonModule, DatePipe } from '@angular/common';
import { FormBuilder, FormGroup, ReactiveFormsModule, Validators, FormsModule } from '@angular/forms';
import * as ExcelJS from 'exceljs';
import { SupabaseService } from '../../core/services/supabase';
import { ToastService } from '../../core/services/toast';

@Component({
  selector: 'app-programacion-rutas',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule, FormsModule],
  providers: [DatePipe], // Añadido para formatear la hora de edición
  templateUrl: './programacion-rutas.html'
})
export class ProgramacionRutasComponent implements OnInit {
  @ViewChild('fileInput') fileInput!: ElementRef;

  showForm = false;
  isLoading = false;
  isEditing = false;
  currentEditId: string | null = null;
  
  rutas: any[] = [];
  clientesActivos: any[] = [];
  
  placas = ['BBU-840', 'ACK-826', 'BHH-935', 'BBT-751', 'F4D-855', 'AYJ-784', 'BAN-786', 'B0L-756', 'F4B-804', 'BEL-713'];

  rutaForm: FormGroup;
  filtroCliente = '';
  filtroPlaca = '';
  
  userName = ''; // Para guardar quién hace la edición

  constructor(
    private fb: FormBuilder,
    private supabase: SupabaseService,
    private toast: ToastService,
    private datePipe: DatePipe
  ) {
    const hoy = new Date().toISOString().split('T')[0];
    
    this.rutaForm = this.fb.group({
      fecha_despacho: [hoy, Validators.required],
      placa: ['', Validators.required],
      cliente_id: ['', Validators.required],
      p05: [0, Validators.min(0)], p10: [0, Validators.min(0)], pm15: [0, Validators.min(0)], p15: [0, Validators.min(0)], p45: [0, Validators.min(0)]
    });
  }

  async ngOnInit() {
    // Obtenemos el nombre del usuario conectado
    const { data: { session } } = await this.supabase.auth.getSession();
    if (session) {
      const { data } = await this.supabase.perfiles.select('nombre_completo').eq('id', session.user.id).single();
      if (data) this.userName = data.nombre_completo;
    }

    await this.cargarDatosBase();
    await this.cargarRutasDelDia();
  }

  async cargarDatosBase() {
    try {
      const { data: clientes } = await this.supabase.clientes.select('id, razon_social, ruc').eq('estado', true).order('razon_social');
      this.clientesActivos = clientes || [];
    } catch (error) {
      this.toast.showError('Error al cargar clientes');
    }
  }

  async cargarRutasDelDia() {
    this.isLoading = true;
    try {
      const { data, error } = await this.supabase.despachos
        .select(`*, clientes ( razon_social ), perfiles!despachos_operador_id_fkey ( nombre_completo )`)
        .order('creado_en', { ascending: false });

      if (error) throw error;
      this.rutas = data || [];
    } catch (error: any) {
      this.toast.showError('Error al cargar rutas: ' + error.message);
    } finally {
      this.isLoading = false;
    }
  }

  cumpleFiltros(r: any): boolean {
    const matchPlaca = !this.filtroPlaca || r.placa.toLowerCase().includes(this.filtroPlaca.toLowerCase());
    const nombreCliente = r.clientes?.razon_social || '';
    const matchCliente = !this.filtroCliente || nombreCliente.toLowerCase().includes(this.filtroCliente.toLowerCase());
    return matchPlaca && matchCliente;
  }

  get rutasPendientes() { return this.rutas.filter(r => r.estado === 'PENDIENTE' && this.cumpleFiltros(r)); }
  get rutasHistorial() { return this.rutas.filter(r => r.estado !== 'PENDIENTE' && this.cumpleFiltros(r)); }

  async importarExcel(event: any) {
    // ... (Mismo código de Excel que ya tenías)
    const file = event.target.files[0];
    if (!file) return;

    this.isLoading = true;
    this.toast.showInfo('Procesando programación masiva...');

    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await file.arrayBuffer());
      const worksheet = workbook.worksheets[0];

      const mapaClientes = new Map(this.clientesActivos.map(c => [c.ruc, c.id]));
      const nuevasRutas: any[] = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) { 
          let fechaStr = '';
          const fechaCelda = row.getCell(1).value;
          if (fechaCelda instanceof Date) {
            fechaStr = fechaCelda.toISOString().split('T')[0];
          } else {
            fechaStr = row.getCell(1).text?.trim(); 
          }

          const placa = row.getCell(2).text?.trim().toUpperCase();
          const rucCliente = row.getCell(3).text?.trim();
          
          const p05 = Number(row.getCell(4).value) || 0;
          const p10 = Number(row.getCell(5).value) || 0;
          const pm15 = Number(row.getCell(6).value) || 0;
          const p15 = Number(row.getCell(7).value) || 0;
          const p45 = Number(row.getCell(8).value) || 0;

          const clienteId = mapaClientes.get(rucCliente);

          if (fechaStr && placa && clienteId) {
            const total = p05 + p10 + pm15 + p15 + p45;
            if (total > 0) {
              nuevasRutas.push({
                fecha_despacho: fechaStr, placa, cliente_id: clienteId,
                p05, p10, pm15, p15, p45, estado: 'PENDIENTE', venta_real: 0, merma: 0,
                programador_nombre: this.userName // <--- Añade esta línea
              });
            }
          }
        }
      });

      if (nuevasRutas.length > 0) {
        const { data, error } = await this.supabase.despachos.insert(nuevasRutas).select(`*, clientes ( razon_social ), perfiles!despachos_operador_id_fkey ( nombre_completo )`);
        if (error) throw error;
        this.rutas = [...(data || []), ...this.rutas];
        this.toast.showSuccess(`${nuevasRutas.length} rutas asignadas`);
      } else {
        this.toast.showWarning('No se generaron rutas.');
      }
    } catch (error: any) {
      this.toast.showError('Error al importar Excel: ' + error.message);
    } finally {
      this.isLoading = false;
      this.fileInput.nativeElement.value = '';
    }
  }

  async onSubmit() {
    if (this.rutaForm.invalid) return this.toast.showError('Complete los campos');
    this.isLoading = true;
    const v = this.rutaForm.value;

    try {
      const total = v.p05 + v.p10 + v.pm15 + v.p15 + v.p45;
      if (total === 0) throw new Error('Programe al menos 1 cilindro');

      if (!this.isEditing) {
        // CREAR NUEVO
        const payload = {
          fecha_despacho: v.fecha_despacho, placa: v.placa, cliente_id: v.cliente_id,
          p05: v.p05, p10: v.p10, pm15: v.pm15, p15: v.p15, p45: v.p45,
          estado: 'PENDIENTE', venta_real: 0, merma: 0,
          programador_nombre: this.userName
        };

        const { data, error } = await this.supabase.despachos.insert(payload).select(`*, clientes ( razon_social ), perfiles!despachos_operador_id_fkey ( nombre_completo )`).single();
        if (error) throw error;
        this.rutas.unshift(data);
        this.toast.showSuccess('Carga asignada');
      } else {
        // ACTUALIZAR EXISTENTE
        const payloadUpdate = {
          fecha_despacho: v.fecha_despacho, placa: v.placa, cliente_id: v.cliente_id,
          p05: v.p05, p10: v.p10, pm15: v.pm15, p15: v.p15, p45: v.p45,
          editado_en: new Date().toISOString(),
          editor_nombre: this.userName
        };

        const { data, error } = await this.supabase.despachos.update(payloadUpdate).eq('id', this.currentEditId).select(`*, clientes ( razon_social ), perfiles!despachos_operador_id_fkey ( nombre_completo )`).single();
        if (error) throw error;
        
        const idx = this.rutas.findIndex(r => r.id === this.currentEditId);
        if (idx !== -1) this.rutas[idx] = data;
        this.toast.showInfo('Ruta actualizada');
      }
      this.closeForm();
    } catch (error: any) {
      this.toast.showError(error.message);
    } finally {
      this.isLoading = false;
    }
  }

  async anularRuta(id: string) {
    if (!confirm('¿Está seguro de anular esta ruta?')) return;
    try {
      const { error } = await this.supabase.despachos.update({ estado: 'ANULADO' }).eq('id', id);
      if (error) throw error;
      const idx = this.rutas.findIndex(r => r.id === id);
      if (idx !== -1) this.rutas[idx].estado = 'ANULADO';
      this.toast.showInfo('Ruta anulada');
    } catch (error: any) {
      this.toast.showError('Error al anular: ' + error.message);
    }
  }

  openForm(registro?: any) {
    if (registro) {
      this.isEditing = true;
      this.currentEditId = registro.id;
      this.rutaForm.patchValue({
        fecha_despacho: registro.fecha_despacho, placa: registro.placa, cliente_id: registro.cliente_id,
        p05: registro.p05, p10: registro.p10, pm15: registro.pm15, p15: registro.p15, p45: registro.p45
      });
    } else {
      this.isEditing = false;
      this.currentEditId = null;
      this.rutaForm.reset({ p05: 0, p10: 0, pm15: 0, p15: 0, p45: 0, fecha_despacho: new Date().toISOString().split('T')[0] });
    }
    this.showForm = true;
  }
  
  closeForm() { this.showForm = false; }
}