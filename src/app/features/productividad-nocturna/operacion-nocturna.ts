import { Component, OnInit, inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormBuilder, FormGroup, FormArray, ReactiveFormsModule, Validators, FormsModule } from '@angular/forms';
import { ToastService } from '../../core/services/toast';
import { SupabaseService } from '../../core/services/supabase';
import { ConfirmModalComponent } from '../../shared/components/confirm-modal/confirm-modal';
import * as ExcelJS from 'exceljs';

@Component({
  selector: 'app-operacion-nocturna',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule, FormsModule, ConfirmModalComponent],
  templateUrl: './operacion-nocturna.html'
})
export class OperacionNocturnaComponent implements OnInit {
  private fb = inject(FormBuilder);
  private toast = inject(ToastService);
  private supabase = inject(SupabaseService);

  showForm = false;
  showAsistenciaModal = false;
  isLoading = false;
  registros: any[] = [];
  nombreResponsable = 'Cargando...'; 
  showDeleteConfirm = false;
  recordToDelete: string | null = null;
  reporteForm!: FormGroup;

  filtroMes = new Date().toISOString().slice(0, 7); 
  filtroEmpleado = '';
  filtroEstado = 'Todos';
  asistenciaConsolidada: any[] = [];

  // Variables para el modal de Detalles
  showDetalleModal = false;
  registroDetalleActual: any = null;

  empleadosDB: string[] = [
    'YAHUANA CORDOVA YAMPIER DAVID / DNI: 62727297',
    'RÍOS CUMARI ELTON / DNI: 76559442',
    'LOPEZ CASTILLO ELVES ABEL / DNI: 74211559',
    'PINO MENENDEZ MARIANO ANDERSON'
  ];

  estadosAsistencia = ['Asistió', 'Falta', 'Tardanza', 'Permiso'];

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
    await this.cargarHistorial();
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

  agregarColaborador() {
    this.colaboradores.push(this.fb.group({ 
      nombre: ['', Validators.required], 
      estado: ['Asistió', Validators.required], 
      observaciones: [''] 
    }));
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
    this.agregarCarga(); 
    this.showForm = true;
  }
  closeForm() { this.showForm = false; }

  async cargarHistorial() {
    this.isLoading = true;
    try {
      const { data, error } = await this.supabase.from('productividad_nocturna').select('*').order('fecha', { ascending: false });
      if (error && error.code !== '42P01') throw error;
      this.registros = data || [];
    } catch (e) { this.toast.showError('Error al cargar historial.'); } finally { this.isLoading = false; }
  }

  async onSubmit() {
    if (this.cargas.length === 0) {
      return this.toast.showError('Debe registrar al menos una Salida de Unidad (Carga).');
    }

    if (this.reporteForm.invalid) {
      this.reporteForm.markAllAsTouched();
      return this.toast.showError('Complete los campos obligatorios en rojo.');
    }
    
    this.isLoading = true;
    try {
      const rawForm = this.reporteForm.getRawValue();

      // Consolidamos todas las placas para cumplir con el NOT NULL de Supabase
      const placasCarga = rawForm.cargas.map((c: any) => c.placa).filter(Boolean);
      const placasDescarga = rawForm.descargas.map((d: any) => d.placa).filter(Boolean);
      const placasUnicas = Array.from(new Set([...placasCarga, ...placasDescarga])).join(' / ') || 'SIN PLACA';

      const payload = { 
        fecha: rawForm.fecha, 
        responsable: this.nombreResponsable, 
        placa: placasUnicas, 
        // ❌ Se eliminó la propiedad "estado" que causaba el conflicto
        datos_formulario: rawForm 
      };

      const { error } = await this.supabase.from('productividad_nocturna').insert(payload);
      if (error) throw error;

      this.toast.showSuccess('Operación registrada exitosamente.');
      this.closeForm(); 
      await this.cargarHistorial();
    } catch (e:any) { 
      this.toast.showError('Error de Base de Datos: ' + e.message); 
    } finally { 
      this.isLoading = false; 
    }
  }

  confirmDelete(id: string) { this.recordToDelete = id; this.showDeleteConfirm = true; }
  cancelDelete() { this.showDeleteConfirm = false; this.recordToDelete = null; }
  async executeDelete() {
    if (!this.recordToDelete) return;
    this.isLoading = true;
    const { error } = await this.supabase.from('productividad_nocturna').delete().eq('id', this.recordToDelete);
    if (error) this.toast.showError('Error al eliminar: ' + error.message);
    else this.toast.showSuccess('Reporte eliminado.');
    
    this.cancelDelete(); 
    await this.cargarHistorial();
    this.isLoading = false;
  }

  // ===============================================
  // MÓDULO DE CONTROL DE ASISTENCIA Y FALTAS
  // ===============================================
  abrirAsistencia() {
    this.showAsistenciaModal = true;
    this.generarReporteAsistencia();
  }

  cerrarAsistencia() {
    this.showAsistenciaModal = false;
  }

  generarReporteAsistencia() {
    this.asistenciaConsolidada = [];
    
    this.registros.forEach((reg: any) => {
      if (this.filtroMes && !reg.fecha.startsWith(this.filtroMes)) return;

      const personal = reg.datos_formulario?.colaboradores || [];
      
      personal.forEach((p: any) => {
        const term = this.filtroEmpleado.toLowerCase().trim();
        if (term && !p.nombre.toLowerCase().includes(term)) return;
        
        if (this.filtroEstado !== 'Todos' && p.estado !== this.filtroEstado) return;

        this.asistenciaConsolidada.push({
          fecha: reg.fecha,
          responsable: reg.responsable,
          nombre: p.nombre,
          estado: p.estado || 'Asistió',
          observaciones: p.observaciones || ''
        });
      });
    });

    this.asistenciaConsolidada.sort((a, b) => b.fecha.localeCompare(a.fecha));
  }

  async exportarAsistencia() {
    if (this.asistenciaConsolidada.length === 0) {
      this.toast.showWarning('No hay datos en la tabla para exportar.');
      return;
    }

    this.toast.showInfo('Generando Excel de Asistencia...');
    
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Control de Personal');

      worksheet.columns = [
        { header: 'FECHA TURNO', key: 'fecha', width: 15 },
        { header: 'EMPLEADO', key: 'nombre', width: 45 },
        { header: 'ESTADO', key: 'estado', width: 15 },
        { header: 'OBSERVACIONES / DETALLE', key: 'observaciones', width: 40 },
        { header: 'REPORTADO POR', key: 'responsable', width: 30 }
      ];

      worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E1B4B' } };

      this.asistenciaConsolidada.forEach(r => {
        const row = worksheet.addRow({
          fecha: r.fecha,
          nombre: r.nombre,
          estado: r.estado,
          observaciones: r.observaciones,
          responsable: r.responsable
        });

        const estadoCell = row.getCell('estado');
        estadoCell.font = { bold: true };
        if (r.estado === 'Falta') estadoCell.font.color = { argb: 'FFDC2626' }; 
        else if (r.estado === 'Tardanza') estadoCell.font.color = { argb: 'FFD97706' }; 
        else if (r.estado === 'Permiso') estadoCell.font.color = { argb: 'FF2563EB' }; 
        else estadoCell.font.color = { argb: 'FF16A34A' }; 
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Asistencia_Nocturna_${this.filtroMes}.xlsx`;
      a.click();
      window.URL.revokeObjectURL(url);
      
      this.toast.showSuccess(`Excel exportado con ${this.asistenciaConsolidada.length} registros.`);
    } catch (error) {
      this.toast.showError('Error al exportar el archivo');
    }
  }

  // ===============================================
  // EXPORTADOR DEL REPORTE OPERATIVO (CORREGIDO PARA JSON MULTIPLE)
  // ===============================================
  async exportarExcel() {
    if (!this.registros || this.registros.length === 0) {
      this.toast.showWarning('No hay datos operativos para exportar.');
      return;
    }

    this.toast.showInfo('Generando reporte logístico...');
    try {
      const workbook = new ExcelJS.Workbook();
      
      // Pestaña 1: CARGAS
      const wsCargas = workbook.addWorksheet('Salidas (Cargas)');
      wsCargas.columns = [
        { header: 'FECHA TURNO', key: 'fecha', width: 15 },
        { header: 'RESPONSABLE', key: 'responsable', width: 25 },
        { header: 'PLACA', key: 'placa', width: 15 },
        { header: 'CAPACIDAD', key: 'capacidad', width: 15 },
        { header: '05 KG', key: 'p05', width: 10 },
        { header: '10 KG', key: 'p10', width: 10 },
        { header: '15 KG', key: 'p15', width: 10 },
        { header: '45 KG', key: 'p45', width: 10 },
        { header: 'TOTAL KG', key: 'total_kg', width: 15 },
        { header: '% FACTOR', key: 'factor', width: 15 },
        { header: 'CONSTANCIA', key: 'constancia', width: 15 },
        { header: 'OBSERVACIÓN', key: 'obs', width: 25 }
      ];
      wsCargas.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      wsCargas.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E1B4B' } };

      // Pestaña 2: DESCARGAS
      const wsDescargas = workbook.addWorksheet('Retornos (Descargas)');
      wsDescargas.columns = [
        { header: 'FECHA TURNO', key: 'fecha', width: 15 },
        { header: 'RESPONSABLE', key: 'responsable', width: 25 },
        { header: 'PLACA', key: 'placa', width: 15 },
        { header: 'ESTADO', key: 'estado', width: 15 },
        { header: '05 KG', key: 'p05', width: 10 },
        { header: '10 KG', key: 'p10', width: 10 },
        { header: '15 KG', key: 'p15', width: 10 },
        { header: '45 KG', key: 'p45', width: 10 },
        { header: 'COMENTARIOS', key: 'comentarios', width: 25 }
      ];
      wsDescargas.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      wsDescargas.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E1B4B' } };

      this.registros.forEach((r: any) => {
        const f = r.datos_formulario;
        if (f?.cargas) {
          f.cargas.forEach((c: any) => {
            wsCargas.addRow({
              fecha: r.fecha, responsable: r.responsable,
              placa: c.placa, capacidad: c.capacidad,
              p05: c.p05, p10: c.p10, p15: c.p15, p45: c.p45,
              total_kg: c.total_kg, factor: c.factor_carga,
              constancia: c.constancia, obs: c.obs
            });
          });
        }
        if (f?.descargas) {
          f.descargas.forEach((d: any) => {
            wsDescargas.addRow({
              fecha: r.fecha, responsable: r.responsable,
              placa: d.placa, estado: d.estado,
              p05: d.p05, p10: d.p10, p15: d.p15, p45: d.p45,
              comentarios: d.comentarios
            });
          });
        }
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Operacion_Nocturna_Logistica_${new Date().toISOString().split('T')[0]}.xlsx`;
      a.click();
      window.URL.revokeObjectURL(url);
      
      this.toast.showSuccess('Excel exportado correctamente');
    } catch (error) {
      this.toast.showError('Error al exportar el archivo');
    }
  }

    // Funciones para abrir y cerrar
  verDetalle(registro: any) {
    this.registroDetalleActual = registro;
    this.showDetalleModal = true;
  }

  cerrarDetalle() {
    this.showDetalleModal = false;
    this.registroDetalleActual = null;
  }
}