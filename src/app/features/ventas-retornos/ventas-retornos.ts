import { Component, OnInit, inject, AfterViewInit, ViewChild, ElementRef, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormBuilder, FormGroup, ReactiveFormsModule, Validators, FormsModule } from '@angular/forms';
import { Chart, registerables } from 'chart.js';
import saveAs from 'file-saver';
import * as ExcelJS from 'exceljs';
import html2canvas from 'html2canvas';
import { ToastService } from '../../core/services/toast';
import { SupabaseService } from '../../core/services/supabase';
import { ConfirmModalComponent } from '../../shared/components/confirm-modal/confirm-modal';

Chart.register(...registerables);

@Component({
  selector: 'app-ventas-retornos',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule, FormsModule, ConfirmModalComponent],
  templateUrl: './ventas-retornos.html'
})
export class VentasRetornos implements OnInit, AfterViewInit, OnDestroy {
  @ViewChild('comparativaChart') comparativaChartRef!: ElementRef;
  @ViewChild('contenedorGraficoVR') contenedorGraficoVRRef!: ElementRef;
  chartInstance: Chart | null = null;
  private autoRefreshSub: any;

  private fb = inject(FormBuilder);
  private toast = inject(ToastService);
  private supabase = inject(SupabaseService); 
  
  showForm = false;
  isEditing = false;
  isLoading = false; 
  currentEditId: string | null = null;
  
  showDeleteConfirm = false;
  recordToDelete: string | null = null;
  
  kpiTotalProgramado = 0;
  kpiTotalVenta = 0;
  kpiTotalRetorno = 0;

  fechaInicio = '';
  fechaFin = '';
  terminoBusqueda = '';

  paginaActual = 1;
  itemsPorPagina = 15;

  ventasForm: FormGroup = this.fb.group({
    fecha: ['', Validators.required],
    placa: ['', Validators.required],
    conductor: ['', Validators.required],
    distribuidor: ['', Validators.required],
    distrito: ['', Validators.required],
    condicion: ['P (Programado)', Validators.required],
    prog_05: [0], prog_10: [0], prog_15: [0], prog_45: [0],
    venta_05: [0], venta_10: [0], venta_15: [0], venta_45: [0],
    ret_05: [0], ret_10: [0], ret_15: [0], ret_45: [0],
    observaciones: ['']
  });

  registrosBase: any[] = []; 
  registros: any[] = []; 

  async ngOnInit() {
    const hoy = new Date();
    this.fechaInicio = new Date(hoy.getFullYear(), hoy.getMonth(), 1).toISOString().split('T')[0];
    this.fechaFin = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0).toISOString().split('T')[0];

    await this.cargarDespachos(); 
    this.iniciarSincronizacionSilenciosa();

    this.ventasForm.valueChanges.subscribe(val => {
      if (!this.isEditing && this.showForm) {
        localStorage.setItem('ventas_draft', JSON.stringify(val));
      }
    });
  }

  ngAfterViewInit() {
    setTimeout(() => {
      this.actualizarGrafico();
    }, 100);
  }

  ngOnDestroy() {
    if (this.autoRefreshSub) {
      clearInterval(this.autoRefreshSub);
    }
  }

  iniciarSincronizacionSilenciosa() {
    this.autoRefreshSub = setInterval(async () => {
      if (!this.isLoading && !this.showForm) {
        try {
          await this.cargarDespachos(true); 
        } catch (err) {}
      }
    }, 10000); 
  }

  trackById(index: number, item: any): string {
    return item.id;
  }

  async cargarDespachos(isSilent = false) {
    if (!isSilent) this.isLoading = true;
    try {
      const { data, error } = await this.supabase.from('despachos')
        .select(`*, clientes ( razon_social, distrito ), perfiles!despachos_operador_id_fkey ( nombre_completo )`)
        .neq('estado', 'ANULADO') 
        .order('fecha_despacho', { ascending: false });

      if (error) throw error;

      this.registrosBase = (data || []).map(d => ({
        id: d.id,
        fecha: d.fecha_despacho,
        placa: d.placa,
        conductor: d.perfiles?.nombre_completo || 'Sin operador',
        cliente: d.clientes?.razon_social || d.cliente_id || 'Desconocido',
        distrito: d.clientes?.distrito || d.distrito || '-',
        condicion: 'P (Programado)',
        estado: d.estado || 'PENDIENTE',
        p05: d.p05 || 0, p10: d.p10 || 0, p15: d.p15 || 0, p45: d.p45 || 0,
        v05: d.v05 || 0, v10: d.v10 || 0, v15: d.v15 || 0, v45: d.v45 || 0,
        m05: d.m05 || 0, m10: d.m10 || 0, m15: d.m15 || 0, m45: d.m45 || 0, 
        obs: d.observaciones || ''
      }));

      this.aplicarFiltros();
    } catch (error: any) {
      if (!isSilent) this.toast.showError('Error al cargar la base de datos');
    } finally {
      if (!isSilent) this.isLoading = false;
    }
  }

  onSearchChange() {
    this.aplicarFiltros();
  }

  aplicarFiltros() {
    let filtrados = this.registrosBase.filter(r => {
      if (this.fechaInicio && r.fecha < this.fechaInicio) return false;
      if (this.fechaFin && r.fecha > this.fechaFin) return false;
      return true;
    });

    if (this.terminoBusqueda) {
      const term = this.terminoBusqueda.toLowerCase();
      filtrados = filtrados.filter(r => 
        (r.placa && r.placa.toLowerCase().includes(term)) ||
        (r.cliente && r.cliente.toLowerCase().includes(term)) ||
        (r.conductor && r.conductor.toLowerCase().includes(term)) ||
        (r.distrito && r.distrito.toLowerCase().includes(term))
      );
    }

    this.registros = filtrados;
    this.paginaActual = 1; 
    this.calcularKPIs();
    this.actualizarGrafico();
  }

  get totalPaginas() {
    return Math.ceil(this.registros.length / this.itemsPorPagina) || 1;
  }

  cambiarPagina(delta: number) {
    const nueva = this.paginaActual + delta;
    if (nueva >= 1 && nueva <= this.totalPaginas) {
      this.paginaActual = nueva;
    }
  }

  calcularKPIs() {
    this.kpiTotalProgramado = this.registros.reduce((acc, r) => acc + r.p05 + r.p10 + r.p15 + r.p45, 0);
    this.kpiTotalVenta = this.registros.reduce((acc, r) => acc + r.v05 + r.v10 + r.v15 + r.v45, 0);
    this.kpiTotalRetorno = this.registros.reduce((acc, r) => acc + r.m05 + r.m10 + r.m15 + r.m45, 0);
  }

  actualizarGrafico() {
    if (!this.comparativaChartRef) return;

    const sumVentas = [
      this.registros.reduce((acc, r) => acc + r.v05, 0),
      this.registros.reduce((acc, r) => acc + r.v10, 0),
      this.registros.reduce((acc, r) => acc + r.v15, 0),
      this.registros.reduce((acc, r) => acc + r.v45, 0)
    ];

    const sumRetornos = [
      this.registros.reduce((acc, r) => acc + r.m05, 0),
      this.registros.reduce((acc, r) => acc + r.m10, 0),
      this.registros.reduce((acc, r) => acc + r.m15, 0),
      this.registros.reduce((acc, r) => acc + r.m45, 0)
    ];

    if (this.chartInstance) {
      this.chartInstance.data.datasets[0].data = sumVentas;
      this.chartInstance.data.datasets[1].data = sumRetornos;
      this.chartInstance.update();
    } else {
      this.chartInstance = new Chart(this.comparativaChartRef.nativeElement, {
        type: 'bar',
        data: {
          labels: ['05 KG', '10 KG', '15 KG', '45 KG'],
          datasets: [
            { label: 'Ventas Reales', data: sumVentas, backgroundColor: '#4F46E5', borderRadius: 4 },
            { label: 'Retornos / Mermas', data: sumRetornos, backgroundColor: '#EF4444', borderRadius: 4 }
          ]
        },
        options: { responsive: true, maintainAspectRatio: false }
      });
    }
  }

  checkDraft() {
    const draft = localStorage.getItem('ventas_draft');
    if (draft) this.ventasForm.patchValue(JSON.parse(draft), { emitEvent: false });
  }

  async getBase64ImageFromUrl(imageUrl: string): Promise<string> {
    const response = await fetch(imageUrl);
    const blob = await response.blob();
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result as string);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  openForm(registro?: any) {
    if (registro) {
      this.isEditing = true;
      this.currentEditId = registro.id;
      this.ventasForm.patchValue({
        fecha: registro.fecha, placa: registro.placa, conductor: registro.conductor, distribuidor: registro.cliente, distrito: registro.distrito, condicion: registro.condicion,
        prog_05: registro.p05, prog_10: registro.p10, prog_15: registro.p15, prog_45: registro.p45,
        venta_05: registro.v05, venta_10: registro.v10, venta_15: registro.v15, venta_45: registro.v45,
        ret_05: registro.m05, ret_10: registro.m10, ret_15: registro.m15, ret_45: registro.m45,
        observaciones: registro.obs
      }, { emitEvent: false });
    } else {
      this.isEditing = false;
      this.currentEditId = null;
      this.ventasForm.reset({ condicion: 'P (Programado)' }, { emitEvent: false });
      this.checkDraft();
    }
    this.showForm = true;
  }

  closeForm() { 
    this.showForm = false; 
  }

  confirmDelete(id: string) {
    this.recordToDelete = id;
    this.showDeleteConfirm = true;
  }

  async executeDelete() {
    if (!this.recordToDelete) return;
    this.isLoading = true;
    try {
      const { error } = await this.supabase.from('despachos').delete().eq('id', this.recordToDelete);
      if (error) throw error;

      await this.cargarDespachos(); 
      this.toast.showSuccess('Registro eliminado de la base de datos'); 
      this.cancelDelete();
    } catch (error: any) {
      this.toast.showError('Error al eliminar: ' + error.message);
    } finally {
      this.isLoading = false;
    }
  }

  cancelDelete() {
    this.showDeleteConfirm = false;
    this.recordToDelete = null;
  }

  async onSubmit() {
    if (this.ventasForm.valid) {
      this.isLoading = true;
      const v = this.ventasForm.value;

      const totalVendido = v.venta_05 + v.venta_10 + v.venta_15 + v.venta_45;
      const totalMerma = v.ret_05 + v.ret_10 + v.ret_15 + v.ret_45;

      const payload = {
        fecha_despacho: v.fecha,
        placa: v.placa,
        distrito: v.distrito,
        p05: v.prog_05, p10: v.prog_10, p15: v.prog_15, p45: v.prog_45,
        v05: v.venta_05, v10: v.venta_10, v15: v.venta_15, v45: v.venta_45,
        m05: v.ret_05, m10: v.ret_10, m15: v.ret_15, m45: v.ret_45,
        venta_real: totalVendido, 
        merma: totalMerma, 
        observaciones: v.observaciones,
        estado: 'COMPLETADO' 
      };

      try {
        if (!this.isEditing) {
          const { error } = await this.supabase.from('despachos').insert(payload);
          if (error) throw error;
          localStorage.removeItem('ventas_draft');
          this.toast.showSuccess('Ruta registrada con éxito en la nube');
        } else {
          const { error } = await this.supabase.from('despachos').update(payload).eq('id', this.currentEditId);
          if (error) throw error;
          this.toast.showInfo('Registro y Dashboard actualizados');
        }
        
        await this.cargarDespachos();
        this.closeForm();
      } catch (error: any) {
        this.toast.showError('Error al guardar: ' + error.message);
      } finally {
        this.isLoading = false;
      }
    } else { 
      this.ventasForm.markAllAsTouched(); 
      this.toast.showError('Complete los campos obligatorios antes de guardar');
    }
  }

  async exportarExcel() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Lista Maestra');

    const logoBase64 = await this.getBase64ImageFromUrl('/rsc.png');
    const logoId = workbook.addImage({ base64: logoBase64, extension: 'png' });

    worksheet.addImage(logoId, { tl: { col: 0.5, row: 0.5 }, ext: { width: 140, height: 50 } });

    worksheet.mergeCells('D1:T2');
    const titleCell = worksheet.getCell('D1');
    titleCell.value = `LISTA MAESTRA DE GLP ENVASADO (${this.fechaInicio} al ${this.fechaFin})`;
    titleCell.font = { bold: true, size: 16 };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

    const fechaActual = new Date().toLocaleString('es-PE');
    worksheet.getCell('A4').value = 'FECHA DE EMISIÓN:';
    worksheet.getCell('A4').font = { bold: true };
    worksheet.getCell('C4').value = fechaActual;

    worksheet.getCell('A5').value = 'RESPONSABLE:';
    worksheet.getCell('A5').font = { bold: true };
    worksheet.getCell('C5').value = 'Administrador de Almacén';

    const headerFill: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6F2FF' } };
    const headerFont: Partial<ExcelJS.Font> = { name: 'Arial', bold: true, size: 10 };
    const borderStyle: Partial<ExcelJS.Borders> = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

    worksheet.mergeCells('A7:G7'); worksheet.getCell('A7').value = 'DATOS GENERALES';
    worksheet.mergeCells('H7:K7'); worksheet.getCell('H7').value = 'PROGRAMADO'; worksheet.getCell('H7').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
    worksheet.mergeCells('L7:O7'); worksheet.getCell('L7').value = 'VENTA REAL'; worksheet.getCell('L7').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E7FF' } };
    worksheet.mergeCells('P7:S7'); worksheet.getCell('P7').value = 'RETORNO / MERMA'; worksheet.getCell('P7').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE4E6' } };
    worksheet.getCell('T7').value = 'OBSERVACIONES';

    worksheet.getRow(7).eachCell(cell => { cell.font = headerFont; cell.alignment = { horizontal: 'center', vertical: 'middle' }; cell.border = borderStyle; });

    worksheet.getRow(8).values = ['FECHA', 'DISTRIBUIDOR / CLIENTE', 'DISTRITO', 'CONDUCTOR', 'PROMOTOR', 'PLACA', 'CONDICION', '05KG', '10KG', '15KG', '45KG', '05KG', '10KG', '15KG', '45KG', '05KG', '10KG', '15KG', '45KG', 'DETALLE'];
    worksheet.getRow(8).eachCell(cell => { cell.fill = headerFill; cell.font = headerFont; cell.alignment = { horizontal: 'center', vertical: 'middle' }; cell.border = borderStyle; });

    worksheet.columns = [{ width: 12 }, { width: 30 }, { width: 20 }, { width: 20 }, { width: 15 }, { width: 12 }, { width: 12 }, { width: 8 }, { width: 8 }, { width: 8 }, { width: 8 }, { width: 8 }, { width: 8 }, { width: 8 }, { width: 8 }, { width: 8 }, { width: 8 }, { width: 8 }, { width: 8 }, { width: 35 }];

    this.registros.forEach(reg => {
      const row = worksheet.addRow([reg.fecha, reg.cliente, reg.distrito, reg.conductor, '-', reg.placa, reg.condicion, reg.p05, reg.p10, reg.p15, reg.p45, reg.v05, reg.v10, reg.v15, reg.v45, reg.m05, reg.m10, reg.m15, reg.m45, reg.obs]);
      row.eachCell(cell => { cell.border = borderStyle; cell.alignment = { vertical: 'middle' }; });
    });

    if (this.contenedorGraficoVRRef) {
      try {
        const canvas = await html2canvas(this.contenedorGraficoVRRef.nativeElement, { scale: 2, backgroundColor: '#ffffff' });
        const graficaBase64 = canvas.toDataURL('image/png');
        const chartImgId = workbook.addImage({ base64: graficaBase64, extension: 'png' });
        worksheet.addImage(chartImgId, { tl: { col: 1, row: worksheet.rowCount + 2 }, ext: { width: 600, height: 350 } });
      } catch (e) {
        console.error('Error al capturar la gráfica', e);
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `GO-F-007_Ventas_Retornos_${new Date().getTime()}.xlsx`);
  }
}