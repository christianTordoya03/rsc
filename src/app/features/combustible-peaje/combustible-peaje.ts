import { Component, OnInit, inject, AfterViewInit, ViewChild, ElementRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormBuilder, FormGroup, ReactiveFormsModule, Validators } from '@angular/forms';
import { Chart, registerables } from 'chart.js';
import saveAs from 'file-saver';
import * as ExcelJS from 'exceljs';
import { ToastService } from '../../core/services/toast';
import { SupabaseService } from '../../core/services/supabase';
import { ConfirmModalComponent } from '../../shared/components/confirm-modal/confirm-modal';

Chart.register(...registerables);

@Component({
  selector: 'app-combustible-peaje',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule, ConfirmModalComponent],
  templateUrl: './combustible-peaje.html'
})
export class CombustiblePeaje implements OnInit, AfterViewInit {
  @ViewChild('rendimientoChart') rendimientoChartRef!: ElementRef;
  chartInstance: Chart | null = null;

  private fb = inject(FormBuilder);
  private toast = inject(ToastService);
  private supabase = inject(SupabaseService); 
  
  showForm = false;
  isEditing = false;
  isLoading = false;
  currentEditId: string | null = null;
  showDeleteConfirm = false;
  recordToDelete: string | null = null;
  
  userId = '';
  registros: any[] = []; 

  // KPIs
  kpiOptimas = 0; kpiAlertas = 0; kpiGastoTotal = 0;

  combustibleForm: FormGroup = this.fb.group({
    fecha: ['', Validators.required],
    placa: ['', Validators.required],
    tipoRecorrido: ['Primera vuelta', Validators.required],
    estacion: [''], // Opcional, el chofer no lo llena
    conductor: [''], // Opcional, el chofer no lo llena
    odoInicial: [0, [Validators.required, Validators.min(0)]],
    odoFinal: [0, [Validators.required, Validators.min(0)]],
    galones: [0, [Validators.required, Validators.min(0.1)]],
    precioGalon: [0, [Validators.required, Validators.min(0.1)]],
    gasto_peaje: [0, [Validators.min(0)]],
    observaciones: ['']
  });

  async ngOnInit() {
    const { data: { session } } = await this.supabase.auth.getSession();
    if (session) this.userId = session.user.id;
    
    await this.cargarTickets(); 

    this.combustibleForm.valueChanges.subscribe(val => {
      if (!this.isEditing && this.showForm) {
        localStorage.setItem('combustible_draft', JSON.stringify(val));
      }
    });
  }

  ngAfterViewInit() { setTimeout(() => this.actualizarGrafico(), 100); }

  async cargarTickets() {
    this.isLoading = true;
    try {
      const { data, error } = await this.supabase.from('control_combustible')
        .select('*')
        .order('fecha', { ascending: false });

      if (error) throw error;
      this.registros = data || [];
      this.calcularKPIs();
      this.actualizarGrafico();
    } catch (error: any) {
      this.toast.showError('Error al cargar tickets de combustible');
    } finally {
      this.isLoading = false;
    }
  }

  // Se actualizan las variables para que coincidan con la base de datos real
  get kmRecorrido() { return Math.max(0, this.combustibleForm.value.odoFinal - this.combustibleForm.value.odoInicial); }
  get rendimiento() { return this.combustibleForm.value.galones > 0 ? (this.kmRecorrido / this.combustibleForm.value.galones) : 0; }
  get costoTotal() { return (this.combustibleForm.value.galones * this.combustibleForm.value.precioGalon) + (this.combustibleForm.value.gasto_peaje || 0); }

  calcularKPIs() {
    this.kpiOptimas = 0; this.kpiAlertas = 0; this.kpiGastoTotal = 0;
    this.registros.forEach(r => {
      const km = (r.km_final || 0) - (r.km_inicial || 0);
      const rend = (r.galones && r.galones > 0) ? km / r.galones : 0;
      if (rend >= 5) this.kpiOptimas++; else this.kpiAlertas++;
      this.kpiGastoTotal += (r.costo_total || 0) + (r.gasto_peaje || 0);
    });
  }

  actualizarGrafico() {
    if (!this.rendimientoChartRef) return;
    if (this.chartInstance) this.chartInstance.destroy();

    const ordenados = [...this.registros].sort((a, b) => new Date(a.fecha).getTime() - new Date(b.fecha).getTime());
    const labels = ordenados.map(r => `${r.fecha} (${r.placa})`);
    const dataRendimiento = ordenados.map(r => ((r.km_final || 0) - (r.km_inicial || 0)) / (r.galones || 1));

    this.chartInstance = new Chart(this.rendimientoChartRef.nativeElement, {
      type: 'line',
      data: {
        labels: labels,
        datasets: [{
          label: 'Rendimiento (Km/gal)', data: dataRendimiento, borderColor: '#10B981',
          backgroundColor: 'rgba(16, 185, 129, 0.1)', borderWidth: 2, pointBackgroundColor: '#10B981', fill: true, tension: 0.4
        }]
      },
      options: { responsive: true, maintainAspectRatio: false }
    });
  }

  async onSubmit() {
    if (this.combustibleForm.invalid) return this.combustibleForm.markAllAsTouched();
    
    this.isLoading = true;
    const v = this.combustibleForm.value;

    try {
      const payloadCombustible = {
        fecha: v.fecha, 
        placa: v.placa, 
        recorrido: v.tipoRecorrido, 
        estacion: v.estacion,
        conductor: v.conductor, 
        km_inicial: v.odoInicial,  // <-- Nombres correctos para la BD
        km_final: v.odoFinal, 
        galones: v.galones,
        costo_total: v.galones * v.precioGalon, 
        gasto_peaje: v.gasto_peaje || 0, 
        observaciones: v.observaciones, 
        registrado_por: this.userId
      };

      if (!this.isEditing) {
        const { error: err1 } = await this.supabase.from('control_combustible').insert(payloadCombustible);
        if (err1) throw err1;

        if (v.gasto_peaje > 0) {
          const payloadPeaje = {
            fecha: v.fecha, placa: v.placa, ruta: v.tipoRecorrido, monto: v.gasto_peaje,
            observaciones: v.observaciones, registrado_por: this.userId
          };
          await this.supabase.from('control_peajes').insert(payloadPeaje);
        }
        
        this.toast.showSuccess('Ticket registrado en Supabase');
        localStorage.removeItem('combustible_draft');
      } else {
        const { error } = await this.supabase.from('control_combustible').update(payloadCombustible).eq('id', this.currentEditId);
        if (error) throw error;
        this.toast.showInfo('Registro actualizado');
      }

      await this.cargarTickets();
      this.closeForm();
    } catch (error: any) {
      this.toast.showError('Error al guardar: ' + error.message);
    } finally {
      this.isLoading = false;
    }
  }

  async executeDelete() {
    if (!this.recordToDelete) return;
    this.isLoading = true;
    try {
      const { error } = await this.supabase.from('control_combustible').delete().eq('id', this.recordToDelete);
      if (error) throw error;
      
      this.toast.showSuccess('Ticket eliminado permanentemente');
      this.cancelDelete();
      await this.cargarTickets();
    } catch (error: any) {
      this.toast.showError('Error al eliminar');
    } finally {
      this.isLoading = false;
    }
  }

  openForm(registro?: any) {
    if (registro) {
      this.isEditing = true; this.currentEditId = registro.id;
      
      // Calculamos el precio por galón a la inversa (Costo Total / Galones)
      const precioCalculado = (registro.galones && registro.galones > 0) 
        ? (registro.costo_total / registro.galones) 
        : 0;

      this.combustibleForm.patchValue({
        fecha: registro.fecha, 
        placa: registro.placa, 
        tipoRecorrido: registro.recorrido || 'Primera vuelta', 
        estacion: registro.estacion || '', 
        conductor: registro.conductor || '',
        odoInicial: registro.km_inicial || 0, 
        odoFinal: registro.km_final || 0, 
        galones: registro.galones || 0, 
        precioGalon: parseFloat(precioCalculado.toFixed(2)), 
        gasto_peaje: registro.gasto_peaje || 0, 
        observaciones: registro.observaciones || ''
      }, { emitEvent: false });
    } else {
      this.isEditing = false; this.currentEditId = null;
      this.combustibleForm.reset({ tipoRecorrido: 'Primera vuelta', odoInicial: 0, odoFinal: 0, galones: 0, precioGalon: 0, gasto_peaje: 0 }, { emitEvent: false });
      const draft = localStorage.getItem('combustible_draft');
      if (draft) this.combustibleForm.patchValue(JSON.parse(draft), { emitEvent: false });
    }
    this.showForm = true;
  }

  closeForm() { this.showForm = false; }
  confirmDelete(id: string) { this.recordToDelete = id; this.showDeleteConfirm = true; }
  cancelDelete() { this.showDeleteConfirm = false; this.recordToDelete = null; }

  async getBase64ImageFromUrl(url: string): Promise<string> {
    const res = await fetch(url); const blob = await res.blob();
    return new Promise((res, rej) => { const r = new FileReader(); r.onloadend = () => res(r.result as string); r.onerror = rej; r.readAsDataURL(blob); });
  }

  async exportarExcel() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Control Kilometraje');
    const logoBase64 = await this.getBase64ImageFromUrl('/rsc.png');
    const chartBase64 = this.chartInstance!.toBase64Image();
    const logoId = workbook.addImage({ base64: logoBase64, extension: 'png' });
    const chartId = workbook.addImage({ base64: chartBase64, extension: 'png' });

    worksheet.addImage(logoId, { tl: { col: 0.5, row: 0.5 }, ext: { width: 140, height: 50 } });
    worksheet.mergeCells('D1:K2'); const titleCell = worksheet.getCell('D1');
    titleCell.value = 'CONTROL DE KILOMETRAJE Y PEAJES - SERVICIOS RSC DIESEL'; titleCell.font = { bold: true, size: 14 }; titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

    const headerFont: Partial<ExcelJS.Font> = { bold: true, size: 10 };
    const border: Partial<ExcelJS.Borders> = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

    worksheet.getRow(7).values = ['ITEM', 'FECHA', 'PLACA', 'ESTACIÓN', 'CONDUCTOR', 'KM INICIAL', 'KM FINAL', 'RECORRIDO', 'GALONES', 'RENDIMIENTO', 'COSTO COMB.', 'PEAJE S/.', 'OBSERVACIÓN'];
    worksheet.getRow(7).eachCell((cell) => { cell.font = headerFont; cell.border = border; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } }; });
    worksheet.columns = [{ width: 6 }, { width: 12 }, { width: 10 }, { width: 25 }, { width: 20 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 10 }, { width: 15 }, { width: 14 }, { width: 12 }, { width: 30 }];

    this.registros.forEach((reg, index) => {
      const rec = (reg.km_final || 0) - (reg.km_inicial || 0);
      const rend = (reg.galones && reg.galones > 0) ? parseFloat((rec / reg.galones).toFixed(2)) : 0;
      const costo = parseFloat((reg.costo_total || 0).toFixed(2));
      const peaje = reg.gasto_peaje ? parseFloat(reg.gasto_peaje.toString()).toFixed(2) : '0.00';
      const row = worksheet.addRow([ index + 1, reg.fecha, reg.placa, reg.estacion || 'No registrado', reg.conductor || 'No registrado', reg.km_inicial, reg.km_final, rec, reg.galones, rend, costo, peaje, reg.observaciones || '' ]);
      row.eachCell(c => { c.border = border; });
    });

    worksheet.addImage(chartId, { tl: { col: 1, row: worksheet.rowCount + 2 }, ext: { width: 800, height: 350 } });
    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `GO-F-010_Combustible_Peaje_${new Date().getTime()}.xlsx`);
  }
}