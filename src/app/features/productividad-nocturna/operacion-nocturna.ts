import { Component, OnInit, inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormBuilder, FormGroup, FormArray, ReactiveFormsModule, Validators } from '@angular/forms';
import { ToastService } from '../../core/services/toast';
import { SupabaseService } from '../../core/services/supabase';
import { ConfirmModalComponent } from '../../shared/components/confirm-modal/confirm-modal';
import * as ExcelJS from 'exceljs';

@Component({
  selector: 'app-operacion-nocturna',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule, ConfirmModalComponent],
  templateUrl: './operacion-nocturna.html'
})
export class OperacionNocturnaComponent implements OnInit {
  private fb = inject(FormBuilder);
  private toast = inject(ToastService);
  private supabase = inject(SupabaseService);

  showForm = false;
  isLoading = false;
  registros: any[] = [];
  nombreResponsable = 'Cargando...'; 
  showDeleteConfirm = false;
  recordToDelete: string | null = null;
  reporteForm!: FormGroup;

  empleadosDB: string[] = [];

  unidades = [
    { placa: 'AYJ-784', capacidad: 3750 }, { placa: 'F4D-855', capacidad: 1750 },
    { placa: 'F4B-804', capacidad: 1750 }, { placa: 'ACK-826', capacidad: 4875 },
    { placa: 'B0L-756', capacidad: 4375 }, { placa: 'B7M-872', capacidad: 9310 },
    { placa: 'BAN-786', capacidad: 2500 }, { placa: 'BBT-751', capacidad: 3500 },
    { placa: 'BBU-840', capacidad: 3500 }, { placa: 'F6Y-818', capacidad: 1600 },
    { placa: 'BEL-713', capacidad: 8250 }, { placa: 'BHH-935', capacidad: 2500 }
  ];

  estadosCilindros = ['Llenos', 'Vacíos'];

  async ngOnInit() {
    this.inicializarFormulario();
    await this.obtenerUsuario();
    await this.cargarEmpleados();
    await this.cargarHistorial();
  }

  async cargarEmpleados() {
    // Traemos a todos los usuarios registrados (puedes filtrar por rol si deseas)
    const { data, error } = await this.supabase.from('perfiles').select('nombre_completo');
    if (data && !error) {
      // Guardamos solo los nombres válidos
      this.empleadosDB = data.map(p => p.nombre_completo).filter(nombre => nombre);
    }
  }

  async obtenerUsuario() {
    const { data: { session } } = await this.supabase.auth.getSession();
    if (session) {
      const { data } = await this.supabase.from('perfiles').select('nombre_completo').eq('id', session.user.id).single();
      if (data) {
        this.nombreResponsable = data.nombre_completo;
        this.reporteForm.patchValue({ responsable: this.nombreResponsable });
      }
    }
  }

  inicializarFormulario() {
    this.reporteForm = this.fb.group({
      fecha: [new Date().toISOString().split('T')[0], Validators.required],
      responsable: [{ value: '', disabled: true }],
      cargas: this.fb.array([]),
      descargas: this.fb.array([]),
      hora_inicio: [''], hora_fin: [''],
      tiempo_total: [{ value: '', disabled: true }],
      dificultad_carga: [''], dificultad_operacion: [''],
      colaboradores: this.fb.array([])
    });

    this.reporteForm.get('hora_inicio')?.valueChanges.subscribe(() => this.calcularTiempoEnCarga());
    this.reporteForm.get('hora_fin')?.valueChanges.subscribe(() => this.calcularTiempoEnCarga());
  }

  get cargas() { return this.reporteForm.get('cargas') as FormArray; }
  get descargas() { return this.reporteForm.get('descargas') as FormArray; }
  get colaboradores() { return this.reporteForm.get('colaboradores') as FormArray; }

  agregarCarga() {
    const fila = this.fb.group({
      placa: ['', Validators.required],
      capacidad: [{ value: 0, disabled: true }],
      p05: [0, Validators.min(0)], p10: [0, Validators.min(0)], p15: [0, Validators.min(0)], p45: [0, Validators.min(0)],
      total_kg: [{ value: 0, disabled: true }],
      factor_carga: [{ value: '0%', disabled: true }],
      constancia: [''], obs: ['']
    });

    fila.valueChanges.subscribe(() => {
      const f = fila.getRawValue();
      const unidad = this.unidades.find(u => u.placa === f.placa);
      const cap = unidad ? unidad.capacidad : 0;
      const tot = (Number(f.p05) * 5) + (Number(f.p10) * 10) + (Number(f.p15) * 15) + (Number(f.p45) * 45);
      const factor = cap > 0 ? (tot / cap) * 100 : 0;
      fila.patchValue({ capacidad: cap, total_kg: tot, factor_carga: factor.toFixed(2) + '%' }, { emitEvent: false });
    });
    this.cargas.push(fila);
  }
  eliminarCarga(i: number) { this.cargas.removeAt(i); }

  agregarDescarga() {
    this.descargas.push(this.fb.group({
      placa: ['', Validators.required], estado: ['Vacíos', Validators.required],
      p05: [0], p10: [0], p15: [0], p45: [0], comentarios: ['']
    }));
  }
  eliminarDescarga(i: number) { this.descargas.removeAt(i); }

  // AHORA EL COLABORADOR TIENE NOMBRE Y OBSERVACIONES CLARAS
  agregarColaborador() {
    this.colaboradores.push(this.fb.group({ nombre: ['', Validators.required], observaciones: [''] }));
  }
  eliminarColaborador(i: number) { this.colaboradores.removeAt(i); }

  calcularTiempoEnCarga() {
    const i = this.reporteForm.get('hora_inicio')?.value;
    const f = this.reporteForm.get('hora_fin')?.value;
    if (i && f) {
      const [hi, mi] = i.split(':').map(Number); const [hf, mf] = f.split(':').map(Number);
      let diff = (hf * 60 + mf) - (hi * 60 + mi);
      if (diff < 0) diff += 24 * 60;
      this.reporteForm.patchValue({ tiempo_total: `${Math.floor(diff / 60)}h ${diff % 60}m` }, { emitEvent: false });
    }
  }

  openForm() {
    this.inicializarFormulario();
    this.obtenerUsuario();
    this.agregarCarga(); this.agregarDescarga();
    this.showForm = true;
  }
  closeForm() { this.showForm = false; }

  async cargarHistorial() {
    this.isLoading = true;
    try {
      const { data, error } = await this.supabase.from('productividad_nocturna').select('*').order('fecha', { ascending: false });
      if (error && error.code !== '42P01') throw error;
      this.registros = data || [];
    } catch (e) { this.toast.showError('Error al cargar.'); } finally { this.isLoading = false; }
  }

  async onSubmit() {
    if (this.reporteForm.invalid) {
      this.reporteForm.markAllAsTouched();
      return this.toast.showError('Complete los campos obligatorios.');
    }
    this.isLoading = true;
    try {
      const payload = { fecha: this.reporteForm.getRawValue().fecha, responsable: this.nombreResponsable, datos_formulario: this.reporteForm.getRawValue() };
      await this.supabase.from('productividad_nocturna').insert(payload);
      this.toast.showSuccess('Operación registrada.');
      this.closeForm(); await this.cargarHistorial();
    } catch (e:any) { this.toast.showError('Error: ' + e.message); } finally { this.isLoading = false; }
  }

  confirmDelete(id: string) { this.recordToDelete = id; this.showDeleteConfirm = true; }
  cancelDelete() { this.showDeleteConfirm = false; this.recordToDelete = null; }
  async executeDelete() {
    if (!this.recordToDelete) return;
    this.isLoading = true;
    await this.supabase.from('productividad_nocturna').delete().eq('id', this.recordToDelete);
    this.cancelDelete(); await this.cargarHistorial();
    this.isLoading = false;
  }

  async exportarExcel() {
    // CAMBIA 'this.registros' POR EL ARRAY QUE USES EN ESTE MÓDULO
    if (!this.registros || this.registros.length === 0) {
      this.toast.showWarning('No hay datos para exportar.');
      return;
    }

    this.toast.showInfo('Generando reporte de Operación Nocturna...');
    
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Operación Nocturna');

      // Ajusta las columnas según lo que necesites mostrar en este reporte
      worksheet.columns = [
        { header: 'FECHA', key: 'fecha', width: 15 },
        { header: 'PLACA', key: 'placa', width: 15 },
        { header: 'CLIENTE', key: 'cliente', width: 45 },
        { header: 'ESTADO', key: 'estado', width: 15 },
        { header: '05 KG', key: 'p05', width: 10 },
        { header: '10 KG', key: 'p10', width: 10 },
        { header: '15 KG', key: 'p15', width: 10 },
        { header: '45 KG', key: 'p45', width: 10 },
        { header: 'OBSERVACIÓN', key: 'observacion', width: 25 }
      ];

      worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E1B4B' } };

      this.registros.forEach((r: any) => {
        worksheet.addRow({
          fecha: r.fecha_despacho,
          placa: r.placa,
          cliente: r.clientes?.razon_social || r.cliente_nombre || 'Sin cliente',
          estado: r.estado,
          p05: r.p05 || 0,
          p10: r.p10 || 0,
          p15: r.p15 || 0,
          p45: r.p45 || 0,
          observacion: r.observacion || ''
        });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Operacion_Nocturna_${new Date().toISOString().split('T')[0]}.xlsx`;
      a.click();
      window.URL.revokeObjectURL(url);
      
      this.toast.showSuccess('Excel exportado correctamente');
    } catch (error) {
      this.toast.showError('Error al exportar el archivo');
    }
  }
}