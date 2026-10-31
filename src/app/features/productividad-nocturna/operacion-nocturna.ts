import { Component, OnInit, inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormBuilder, FormGroup, FormArray, ReactiveFormsModule, Validators, FormsModule } from '@angular/forms';
import { ToastService } from '../../core/services/toast';
import { SupabaseService } from '../../core/services/supabase';
import { ConfirmModalComponent } from '../../shared/components/confirm-modal/confirm-modal';
import * as ExcelJS from 'exceljs';
import html2canvas from 'html2canvas';

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
  showDetalleModal = false;
  isLoading = false;
  registros: any[] = [];
  registrosCompletos: any[] = []; 
  nombreResponsable = 'Cargando...'; 
  
  isEditing = false;
  currentEditId: string | null = null;
  showDeleteConfirm = false;
  recordToDelete: string | null = null;
  registroDetalleActual: any = null;
  
  // Función para obtener la fecha de Lima (GMT-5) sin desfases
  private obtenerFechaLocal() {
    const fecha = new Date();
    fecha.setMinutes(fecha.getMinutes() - fecha.getTimezoneOffset());
    return fecha.toISOString().split('T')[0];
  }

  filtroMainInicio = new Date(new Date().getFullYear(), new Date().getMonth(), 1).toISOString().split('T')[0];
  filtroMainFin = this.obtenerFechaLocal();

  // Variables del Dashboard Principal
  dashTotalTurnos = 0;
  dashTotalSalidas = 0;
  dashTotalRetornos = 0;
  dashTotalKg = 0;
  dashDescargasVaciosQty = 0; // NUEVO: Unidades Vacías
  dashDescargasLlenosKg = 0;  // NUEVO: Kilos Llenos
  dashDataPlacas: any[] = []; 

  reporteForm!: FormGroup;
  
  filtroAsistenciaInicio = new Date(new Date().getFullYear(), new Date().getMonth(), 1).toISOString().split('T')[0];
  filtroAsistenciaFin = this.obtenerFechaLocal();
  
  filtroEmpleado = '';
  filtroEstado = 'Todos';
  asistenciaConsolidada: any[] = [];

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
    { placa: 'BEL-713', capacidad: 8250 }, { placa: 'BHH-935', capacidad: 2500 },

    { placa: 'BSY-931', capacidad: 2500 }, { placa: 'AUH-927', capacidad: 2500 },
    { placa: 'F6Z-867', capacidad: 1600 }, { placa: 'BHI-940', capacidad: 2500 },
    { placa: 'BTB-817', capacidad: 2500 }, { placa: 'AUH-916', capacidad: 2500 },
    { placa: 'F4D-707', capacidad: 1750 }
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
      fecha: [this.obtenerFechaLocal(), Validators.required],
      responsable: [{ value: this.nombreResponsable, disabled: true }],
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

  private aplicarSuscripcionCarga(fila: FormGroup) {
    fila.valueChanges.subscribe(() => {
      const f = fila.getRawValue();
      const unidad = this.unidades.find(u => u.placa === f.placa);
      const cap = unidad?.capacidad || 0;
      const tot = (Number(f.p05) * 5) + (Number(f.p10) * 10) + (Number(f.p15) * 15) + (Number(f.p45) * 45);
      const factor = cap > 0 ? (tot / cap) * 100 : 0;
      fila.patchValue({ capacidad: cap, total_kg: tot, factor_carga: factor.toFixed(2) + '%' }, { emitEvent: false });
    });
  }

  agregarCarga(data?: any) {
    const fila = this.fb.group({
      placa: [data?.placa || '', Validators.required],
      capacidad: [{ value: data?.capacidad || 0, disabled: true }],
      p05: [data?.p05 || 0, Validators.min(0)], p10: [data?.p10 || 0, Validators.min(0)], 
      p15: [data?.p15 || 0, Validators.min(0)], p45: [data?.p45 || 0, Validators.min(0)],
      total_kg: [{ value: data?.total_kg || 0, disabled: true }],
      factor_carga: [{ value: data?.factor_carga || '0%', disabled: true }],
      constancia: [data?.constancia || ''], obs: [data?.obs || '']
    });
    this.aplicarSuscripcionCarga(fila);
    this.cargas.push(fila);
  }
  eliminarCarga(i: number) { this.cargas.removeAt(i); }

  agregarDescarga(data?: any) {
    this.descargas.push(this.fb.group({
      placa: [data?.placa || '', Validators.required], estado: [data?.estado || 'Vacíos', Validators.required],
      p05: [data?.p05 || 0], p10: [data?.p10 || 0], p15: [data?.p15 || 0], p45: [data?.p45 || 0], comentarios: [data?.comentarios || '']
    }));
  }
  eliminarDescarga(i: number) { this.descargas.removeAt(i); }

  agregarColaborador(data?: any) {
    this.colaboradores.push(this.fb.group({ 
      nombre: [data?.nombre || '', Validators.required], 
      estado: [data?.estado || 'Asistió', Validators.required], 
      observaciones: [data?.observaciones || ''] 
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
    this.isEditing = false;
    this.currentEditId = null;
    this.inicializarFormulario();
    this.reporteForm.patchValue({ responsable: this.nombreResponsable });
    this.agregarCarga(); 
    this.showForm = true;
  }

  editarOperacion(reg: any) {
    this.isEditing = true;
    this.currentEditId = reg.id;
    this.inicializarFormulario();
    
    this.reporteForm.patchValue({
      fecha: reg.fecha,
      responsable: reg.responsable,
      hora_inicio: reg.datos_formulario?.hora_inicio || '',
      hora_fin: reg.datos_formulario?.hora_fin || '',
      tiempo_total: reg.datos_formulario?.tiempo_total || '',
      dificultad_carga: reg.datos_formulario?.dificultad_carga || ''
    });

    const cargasData = reg.datos_formulario?.cargas || [];
    cargasData.forEach((c: any) => this.agregarCarga(c));
    if (this.cargas.length === 0) this.agregarCarga();

    const descargasData = reg.datos_formulario?.descargas || [];
    descargasData.forEach((d: any) => this.agregarDescarga(d));

    const colabsData = reg.datos_formulario?.colaboradores || [];
    colabsData.forEach((col: any) => this.agregarColaborador(col));

    this.showForm = true;
  }

  closeForm() { this.showForm = false; }

  verDetalle(reg: any) { this.registroDetalleActual = reg; this.showDetalleModal = true; }
  cerrarDetalle() { this.showDetalleModal = false; this.registroDetalleActual = null; }

  async cargarHistorial() {
    this.isLoading = true;
    try {
      const { data, error } = await this.supabase.from('productividad_nocturna').select('*').order('fecha', { ascending: false });
      if (error && error.code !== '42P01') throw error;
      
      this.registrosCompletos = data || [];
      this.aplicarFiltroMain(); 
    } catch (e) { 
      this.toast.showError('Error al cargar historial.'); 
    } finally { 
      this.isLoading = false; 
    }
  }

  aplicarFiltroMain() {
    this.registros = this.registrosCompletos.filter(reg => {
      return reg.fecha >= this.filtroMainInicio && reg.fecha <= this.filtroMainFin;
    });
    this.actualizarDashboard();
  }

  actualizarDashboard() {
    this.dashTotalTurnos = this.registros.length;
    this.dashTotalSalidas = 0;
    this.dashTotalRetornos = 0;
    this.dashTotalKg = 0;
    
    // Reset de nuevas variables
    this.dashDescargasVaciosQty = 0;
    this.dashDescargasLlenosKg = 0;

    const placasMap = new Map<string, { kg: number, cap: number }>();

    this.registros.forEach(r => {
      // CARGAS
      if (r.datos_formulario?.cargas) {
        this.dashTotalSalidas += r.datos_formulario.cargas.length;
        r.datos_formulario.cargas.forEach((c: any) => {
          const k = Number(c.total_kg || 0);
          const cp = Number(c.capacidad || 0);
          this.dashTotalKg += k;
          
          if (c.placa) {
            const current = placasMap.get(c.placa) || { kg: 0, cap: 0 };
            placasMap.set(c.placa, { kg: current.kg + k, cap: current.cap + cp });
          }
        });
      }
      
      // DESCARGAS (Lógica separada por Estado)
      if (r.datos_formulario?.descargas) {
        this.dashTotalRetornos += r.datos_formulario.descargas.length;
        r.datos_formulario.descargas.forEach((d: any) => {
          const q05 = Number(d.p05 || 0);
          const q10 = Number(d.p10 || 0);
          const q15 = Number(d.p15 || 0);
          const q45 = Number(d.p45 || 0);
          
          if (d.estado === 'Vacíos') {
             // Solo suma las unidades
             this.dashDescargasVaciosQty += (q05 + q10 + q15 + q45);
          } else if (d.estado === 'Llenos') {
             // Suma los kilos
             this.dashDescargasLlenosKg += (q05 * 5) + (q10 * 10) + (q15 * 15) + (q45 * 45);
          }
        });
      }
    });

    const sortedPlacas = Array.from(placasMap.entries())
      .map(([placa, data]) => {
        const factorPorcentaje = data.cap > 0 ? ((data.kg / data.cap) * 100).toFixed(1) : '0.0';
        return { placa, kg: data.kg, factor: factorPorcentaje + '%' };
      })
      .sort((a, b) => b.kg - a.kg)
      .slice(0, 10); 

    const maxKg = Math.max(...sortedPlacas.map(d => d.kg), 1);

    this.dashDataPlacas = sortedPlacas.map(d => ({
      nombreCorto: d.placa,
      kg: d.kg,
      factor: d.factor,
      porcentaje: (d.kg / maxKg) * 100 
    }));
  }

  async onSubmit() {
    if (this.cargas.length === 0) return this.toast.showError('Debe registrar al menos una Salida de Unidad (Carga).');
    if (this.reporteForm.invalid) {
      this.reporteForm.markAllAsTouched();
      return this.toast.showError('Complete los campos obligatorios en rojo.');
    }
    
    this.isLoading = true;
    try {
      const rawForm = this.reporteForm.getRawValue();
      const placasCarga = rawForm.cargas.map((c: any) => c.placa).filter(Boolean);
      const placasDescarga = rawForm.descargas.map((d: any) => d.placa).filter(Boolean);
      const placasUnicas = Array.from(new Set([...placasCarga, ...placasDescarga])).join(' / ') || 'SIN PLACA';

      const payload = { 
        fecha: rawForm.fecha, 
        responsable: rawForm.responsable || this.nombreResponsable, 
        placa: placasUnicas, 
        datos_formulario: rawForm 
      };

      if (this.isEditing && this.currentEditId) {
        const { error } = await this.supabase.from('productividad_nocturna').update(payload).eq('id', this.currentEditId);
        if (error) throw error;
        this.toast.showSuccess('Operación actualizada exitosamente.');
      } else {
        const { error } = await this.supabase.from('productividad_nocturna').insert(payload);
        if (error) throw error;
        this.toast.showSuccess('Operación registrada exitosamente.');
      }

      this.closeForm(); 
      await this.cargarHistorial();
    } catch (e:any) { this.toast.showError('Error de Base de Datos: ' + e.message); } finally { this.isLoading = false; }
  }

  confirmDelete(id: string) { this.recordToDelete = id; this.showDeleteConfirm = true; }
  cancelDelete() { this.showDeleteConfirm = false; this.recordToDelete = null; }
  async executeDelete() {
    if (!this.recordToDelete) return;
    this.isLoading = true;
    const { error } = await this.supabase.from('productividad_nocturna').delete().eq('id', this.recordToDelete);
    if (error) this.toast.showError('Error al eliminar: ' + error.message);
    else this.toast.showSuccess('Reporte eliminado.');
    
    this.cancelDelete(); await this.cargarHistorial();
    this.isLoading = false;
  }

  abrirAsistencia() { this.showAsistenciaModal = true; this.generarReporteAsistencia(); }
  cerrarAsistencia() { this.showAsistenciaModal = false; }
  
  generarReporteAsistencia() {
    this.asistenciaConsolidada = [];
    this.registrosCompletos.forEach((reg: any) => {
      if (reg.fecha < this.filtroAsistenciaInicio || reg.fecha > this.filtroAsistenciaFin) return;
      
      const personal = reg.datos_formulario?.colaboradores || [];
      personal.forEach((p: any) => {
        const term = this.filtroEmpleado.toLowerCase().trim();
        if (term && !p.nombre.toLowerCase().includes(term)) return;
        if (this.filtroEstado !== 'Todos' && p.estado !== this.filtroEstado) return;

        this.asistenciaConsolidada.push({
          fecha: reg.fecha, responsable: reg.responsable, nombre: p.nombre,
          estado: p.estado || 'Asistió', observaciones: p.observaciones || ''
        });
      });
    });
    this.asistenciaConsolidada.sort((a, b) => b.fecha.localeCompare(a.fecha));
  }

  async exportarAsistencia() {
    if (this.asistenciaConsolidada.length === 0) return this.toast.showWarning('No hay datos en el rango seleccionado.');
    this.toast.showInfo('Generando Excel de Asistencia...');
    try {
      const workbook = new ExcelJS.Workbook();
      const ws = workbook.addWorksheet('Control de Personal');
      ws.columns = [
        { header: 'FECHA', key: 'fecha', width: 15 }, { header: 'EMPLEADO', key: 'nombre', width: 45 },
        { header: 'ESTADO', key: 'estado', width: 15 }, { header: 'OBSERVACIONES', key: 'observaciones', width: 40 }
      ];
      ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E1B4B' } };

      this.asistenciaConsolidada.forEach(r => {
        const row = ws.addRow(r);
        const cell = row.getCell('estado');
        cell.font = { bold: true };
        if (r.estado === 'Falta') cell.font.color = { argb: 'FFDC2626' }; 
        else if (r.estado === 'Tardanza') cell.font.color = { argb: 'FFD97706' }; 
        else if (r.estado === 'Permiso') cell.font.color = { argb: 'FF2563EB' }; 
        else cell.font.color = { argb: 'FF16A34A' }; 
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; 
      a.download = `Asistencia_${this.filtroAsistenciaInicio}_al_${this.filtroAsistenciaFin}.xlsx`; 
      a.click();
    } catch (e) { this.toast.showError('Error al exportar'); }
  }

  async exportarExcel() {
    if (!this.registros || this.registros.length === 0) return this.toast.showWarning('No hay datos en el rango seleccionado.');
    this.toast.showInfo('Capturando gráfica y generando Excel Gerencial...');
    
    let graficaBase64 = '';
    const chartElement = document.getElementById('grafica-nocturna');
    if (chartElement) {
      try {
        const canvas = await html2canvas(chartElement, { scale: 2, backgroundColor: '#ffffff' });
        graficaBase64 = canvas.toDataURL('image/png');
      } catch (e) {
        console.error('Error al capturar la gráfica', e);
      }
    }

    try {
      const workbook = new ExcelJS.Workbook();
      
      const wsDash = workbook.addWorksheet('Dashboard Resumen', { views: [{ showGridLines: false }] });
      
      wsDash.getColumn('A').width = 3;  
      wsDash.getColumn('B').width = 37; 
      wsDash.getColumn('C').width = 15; 
      wsDash.getColumn('D').width = 4;  
      wsDash.getColumn('E').width = 15; 
      wsDash.getColumn('F').width = 18; 
      wsDash.getColumn('G').width = 50; 

      wsDash.mergeCells('B2:G4');
      const titleCell = wsDash.getCell('B2');
      titleCell.value = `📊 REPORTE GERENCIAL - OPERACIÓN NOCTURNA (${this.filtroMainInicio} al ${this.filtroMainFin})`;
      titleCell.font = { size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
      titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } }; 
      titleCell.alignment = { vertical: 'middle', horizontal: 'center' };

      wsDash.mergeCells('B6:C6');
      const kpiTitle = wsDash.getCell('B6');
      kpiTitle.value = 'MÉTRICAS CLAVE DEL RANGO';
      kpiTitle.font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
      kpiTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } }; 
      kpiTitle.alignment = { vertical: 'middle', horizontal: 'center' };

      const kpis = [
        { label: 'Total Turnos Registrados', val: this.dashTotalTurnos },
        { label: 'Total Camiones Cargados (Salidas)', val: this.dashTotalSalidas },
        { label: 'Total Unidades con Cilindros Descargados', val: this.dashTotalRetornos },
        { label: 'Volumen Total Cargado (Kg Venta)', val: this.dashTotalKg },
        { label: 'Cilindros Vacíos Descargados (Unidades)', val: this.dashDescargasVaciosQty },
        { label: 'Cilindros Llenos Descargados (Kg)', val: this.dashDescargasLlenosKg }
      ];

      let rowLine = 8;
      kpis.forEach((kpi, idx) => {
        const lblCell = wsDash.getCell(`B${rowLine}`);
        const valCell = wsDash.getCell(`C${rowLine}`);

        lblCell.value = kpi.label;
        lblCell.font = { bold: true, color: { argb: 'FF475569' } };
        lblCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        lblCell.border = { left: {style: 'thin', color: {argb: 'FFCBD5E1'}}, bottom: {style: 'thin', color: {argb: 'FFCBD5E1'}}, top: {style: 'thin', color: {argb: 'FFCBD5E1'}} };

        valCell.value = kpi.val;
        valCell.font = { bold: true, size: 12, color: { argb: 'FF0F766E' } };
        valCell.alignment = { horizontal: 'right', vertical: 'middle' };
        valCell.border = { right: {style: 'thin', color: {argb: 'FFCBD5E1'}}, bottom: {style: 'thin', color: {argb: 'FFCBD5E1'}}, top: {style: 'thin', color: {argb: 'FFCBD5E1'}} };
        
        if (idx === 3 || idx === 5) valCell.numFmt = '#,##0 "Kg"';
        if (idx === 4) valCell.numFmt = '#,##0 "Und"';
        rowLine++;
      });

      if (graficaBase64) {
        const imageId = workbook.addImage({ base64: graficaBase64, extension: 'png' });
        wsDash.addImage(imageId, { tl: { col: 4, row: 5 }, ext: { width: 550, height: 350 } });
      }

      const wsCargas = workbook.addWorksheet('Detalle Cargas');
      wsCargas.columns = [
        { header: 'FECHA TURNO', key: 'fecha', width: 15 }, { header: 'RESPONSABLE', key: 'responsable', width: 25 },
        { header: 'PLACA', key: 'placa', width: 15 }, { header: 'CAPACIDAD', key: 'capacidad', width: 15 },
        { header: '05 KG', key: 'p05', width: 10 }, { header: '10 KG', key: 'p10', width: 10 },
        { header: '15 KG', key: 'p15', width: 10 }, { header: '45 KG', key: 'p45', width: 10 },
        { header: 'TOTAL KG', key: 'total_kg', width: 15 }, { header: '% FACTOR', key: 'factor', width: 15 },
        { header: 'OBSERVACIÓN', key: 'obs', width: 30 }
      ];
      wsCargas.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      wsCargas.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };

      this.registros.forEach((r: any) => {
        if (r.datos_formulario?.cargas) {
          r.datos_formulario.cargas.forEach((c: any) => {
            wsCargas.addRow({
              fecha: r.fecha, responsable: r.responsable, placa: c.placa, capacidad: c.capacidad,
              p05: c.p05, p10: c.p10, p15: c.p15, p45: c.p45, total_kg: c.total_kg, factor: c.factor_carga, obs: c.obs
            });
          });
        }
      });

      const wsDescargas = workbook.addWorksheet('Control Cilindros Descargados');
      wsDescargas.columns = [
        { header: 'FECHA TURNO', key: 'fecha', width: 15 }, { header: 'PLACA', key: 'placa', width: 15 },
        { header: 'ESTADO CILINDROS', key: 'estado', width: 20 },
        { header: '05 KG', key: 'p05', width: 10 }, { header: '10 KG', key: 'p10', width: 10 },
        { header: '15 KG', key: 'p15', width: 10 }, { header: '45 KG', key: 'p45', width: 10 },
        { header: 'COMENTARIOS', key: 'comentarios', width: 35 }
      ];
      wsDescargas.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      wsDescargas.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD97706' } }; 

      this.registros.forEach((r: any) => {
        if (r.datos_formulario?.descargas) {
          r.datos_formulario.descargas.forEach((d: any) => {
            wsDescargas.addRow({
              fecha: r.fecha, placa: d.placa, estado: d.estado,
              p05: d.p05, p10: d.p10, p15: d.p15, p45: d.p45, comentarios: d.comentarios
            });
          });
        }
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url;
      a.download = `Reporte_Nocturno_${this.filtroMainInicio}_al_${this.filtroMainFin}.xlsx`;
      a.click();
      window.URL.revokeObjectURL(url);
      
      this.toast.showSuccess('Reporte Gerencial descargado con éxito.');
    } catch (error) {
      this.toast.showError('Error al exportar Excel. Intente nuevamente.');
    }
  }
}