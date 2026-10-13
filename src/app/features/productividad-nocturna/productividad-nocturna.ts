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
  selector: 'app-productividad-nocturna',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule, ConfirmModalComponent],
  templateUrl: './productividad-nocturna.html'
})
export class ProductividadNocturna implements OnInit, AfterViewInit {
  @ViewChild('distribucionChart') distribucionChartRef!: ElementRef;
  chartInstance: Chart | null = null;

  private fb = inject(FormBuilder);
  private toast = inject(ToastService);
  private supabase = inject(SupabaseService); // <-- CONEXIÓN A SUPABASE
  
  showForm = false;
  isEditing = false;
  isLoading = false; // <-- ESTADO DE CARGA
  currentEditId: string | null = null;
  showDeleteConfirm = false;
  recordToDelete: string | null = null;
  
  // KPIs
  kpiUtilizacionGlobal = 0;
  kpiCargaTotal = 0;

  nocturnaForm: FormGroup = this.fb.group({
    fechaTurno: ['', Validators.required],
    promotor: ['', Validators.required],
    placa: ['', Validators.required],
    grupoOperativo: ['INDUSTRIAS', Validators.required],
    capacidadProgramada: [0, [Validators.required, Validators.min(1)]],
    cargaRealizada: [0, [Validators.required, Validators.min(0)]],
    observaciones: ['']
  });

  registros: any[] = []; // <-- ARREGLO LIMPIO PARA SUPABASE

  async ngOnInit() {
    await this.cargarTurnos(); // <-- CARGA INICIAL DESDE LA NUBE
    
    this.nocturnaForm.valueChanges.subscribe(val => {
      if (!this.isEditing && this.showForm) {
        localStorage.setItem('nocturna_draft', JSON.stringify(val));
      }
    });
  }

  ngAfterViewInit() { setTimeout(() => this.actualizarGrafico(), 100); }

  // Cálculo de cumplimiento en tiempo real para el modal
  get cumplimientoFormulario() {
    const v = this.nocturnaForm.value;
    if (v.capacidadProgramada > 0) {
      return (v.cargaRealizada / v.capacidadProgramada) * 100;
    }
    return 0;
  }

  // --- 1. FUNCIÓN PARA CARGAR DATOS REALES DE SUPABASE ---
  async cargarTurnos() {
    this.isLoading = true;
    try {
      const { data, error } = await this.supabase.from('productividad_nocturna')
        .select('*')
        .order('fecha', { ascending: false });

      if (error) throw error;
      
      // Mapeamos los datos de Supabase a la estructura que ya usaba tu HTML
      this.registros = (data || []).map(t => ({
        id: t.id,
        fecha: t.fecha,
        promotor: t.promotor,
        placa: t.placa,
        grupo: t.grupo_operativo,
        cap: t.cap_prog_kg || 0,
        carga: t.real_kg || 0,
        obs: t.observaciones || ''
      }));

      this.calcularKPIs();
      this.actualizarGrafico();
    } catch (error: any) {
      this.toast.showError('Error al cargar productividad');
    } finally {
      this.isLoading = false;
    }
  }

  calcularKPIs() {
    let totalCarga = 0;
    let totalCapacidad = 0;
    
    this.registros.forEach(r => {
      totalCarga += r.carga;
      totalCapacidad += r.cap;
    });

    this.kpiCargaTotal = totalCarga;
    this.kpiUtilizacionGlobal = totalCapacidad > 0 ? (totalCarga / totalCapacidad) * 100 : 0;
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

  actualizarGrafico() {
    if (!this.distribucionChartRef) return;
    if (this.chartInstance) this.chartInstance.destroy();

    const cargaIndustrias = this.registros.filter(r => r.grupo === 'INDUSTRIAS').reduce((acc, r) => acc + r.carga, 0);
    const cargaDistribuidores = this.registros.filter(r => r.grupo === 'DISTRIBUIDORES').reduce((acc, r) => acc + r.carga, 0);

    this.chartInstance = new Chart(this.distribucionChartRef.nativeElement, {
      type: 'doughnut',
      data: {
        labels: ['Industrias', 'Distribuidores'],
        datasets: [{
          data: [cargaIndustrias, cargaDistribuidores],
          backgroundColor: ['#4F46E5', '#D946EF'], // Azul y Fucsia
          borderWidth: 0,
          hoverOffset: 4
        }]
      },
      options: { 
        responsive: true, 
        maintainAspectRatio: false,
        cutout: '70%',
        plugins: { legend: { position: 'right' } }
      }
    });
  }

  openForm(registro?: any) {
    if (registro) {
      this.isEditing = true; this.currentEditId = registro.id;
      this.nocturnaForm.patchValue({
        fechaTurno: registro.fecha, promotor: registro.promotor, placa: registro.placa, grupoOperativo: registro.grupo,
        capacidadProgramada: registro.cap, cargaRealizada: registro.carga, observaciones: registro.obs
      }, { emitEvent: false });
    } else {
      this.isEditing = false; this.currentEditId = null;
      this.nocturnaForm.reset({ grupoOperativo: 'INDUSTRIAS', capacidadProgramada: 0, cargaRealizada: 0 }, { emitEvent: false });
      const draft = localStorage.getItem('nocturna_draft');
      if (draft) this.nocturnaForm.patchValue(JSON.parse(draft), { emitEvent: false });
    }
    this.showForm = true;
  }

  closeForm() { this.showForm = false; }

  confirmDelete(id: string) { this.recordToDelete = id; this.showDeleteConfirm = true; }
  cancelDelete() { this.showDeleteConfirm = false; this.recordToDelete = null; }

  // --- 2. ELIMINAR EN SUPABASE ---
  async executeDelete() {
    if (!this.recordToDelete) return;
    this.isLoading = true;
    try {
      const { error } = await this.supabase.from('productividad_nocturna').delete().eq('id', this.recordToDelete);
      if (error) throw error;
      
      await this.cargarTurnos(); // Recargar de la BD
      this.toast.showSuccess('Registro nocturno eliminado');
      this.cancelDelete();
    } catch (error: any) {
      this.toast.showError('Error al eliminar');
    } finally {
      this.isLoading = false;
    }
  }

  // --- 3. CREAR/EDITAR EN SUPABASE ---
  async onSubmit() {
    if (this.nocturnaForm.valid) {
      this.isLoading = true;
      const v = this.nocturnaForm.value;
      const eficiencia = v.capacidadProgramada > 0 ? (v.cargaRealizada / v.capacidadProgramada) * 100 : 0;

      const payload = {
        fecha: v.fechaTurno,
        promotor: v.promotor,
        placa: v.placa,
        grupo_operativo: v.grupoOperativo,
        cap_prog_kg: v.capacidadProgramada,
        real_kg: v.cargaRealizada,
        eficiencia: eficiencia,
        observaciones: v.observaciones
      };

      try {
        if (!this.isEditing) {
          const { error } = await this.supabase.from('productividad_nocturna').insert(payload);
          if (error) throw error;
          localStorage.removeItem('nocturna_draft');
          this.toast.showSuccess('Productividad registrada correctamente');
        } else {
          const { error } = await this.supabase.from('productividad_nocturna').update(payload).eq('id', this.currentEditId);
          if (error) throw error;
          this.toast.showInfo('Registro actualizado');
        }
        
        await this.cargarTurnos();
        this.closeForm();
      } catch (error: any) {
        this.toast.showError('Error al guardar: ' + error.message);
      } finally {
        this.isLoading = false;
      }
    } else {
      this.nocturnaForm.markAllAsTouched();
      this.toast.showError('Complete todos los campos obligatorios');
    }
  }

  async exportarExcel() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Productividad Nocturna');
    
    const logoBase64 = await this.getBase64ImageFromUrl('/rsc.png');
    const chartBase64 = this.chartInstance!.toBase64Image();

    const logoId = workbook.addImage({ base64: logoBase64, extension: 'png' });
    const chartId = workbook.addImage({ base64: chartBase64, extension: 'png' });

    worksheet.addImage(logoId, { tl: { col: 0.5, row: 0.5 }, ext: { width: 140, height: 50 } });

    worksheet.mergeCells('D1:I2');
    const titleCell = worksheet.getCell('D1');
    titleCell.value = 'REPORTE DE PRODUCTIVIDAD E INCIDENCIAS - TURNO NOCTURNO';
    titleCell.font = { bold: true, size: 14 };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

    worksheet.getCell('A4').value = 'FECHA DE EMISIÓN:'; worksheet.getCell('A4').font = { bold: true };
    worksheet.getCell('C4').value = new Date().toLocaleString('es-PE');
    worksheet.getCell('A5').value = 'RESPONSABLE:'; worksheet.getCell('A5').font = { bold: true };
    worksheet.getCell('C5').value = 'Usuario Actual (DNI Pendiente)';

    const headerFont: Partial<ExcelJS.Font> = { bold: true, size: 10 };
    const border: Partial<ExcelJS.Borders> = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

    worksheet.getRow(7).values = ['ITEM', 'FECHA TURNO', 'PROMOTOR', 'PLACA', 'GRUPO OPERATIVO', 'CAPACIDAD PROG. (KG)', 'CARGA REAL (KG)', '% EFICIENCIA', 'OBSERVACIONES'];
    worksheet.getRow(7).eachCell((cell) => { cell.font = headerFont; cell.border = border; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6E6FA' } }; });

    worksheet.columns = [{ width: 6 }, { width: 15 }, { width: 25 }, { width: 12 }, { width: 20 }, { width: 22 }, { width: 20 }, { width: 15 }, { width: 40 }];

    this.registros.forEach((reg, index) => {
      const eficiencia = reg.cap > 0 ? (reg.carga / reg.cap) : 0;
      const row = worksheet.addRow([ index + 1, reg.fecha, reg.promotor, reg.placa, reg.grupo, reg.cap, reg.carga, eficiencia, reg.obs ]);
      row.getCell(8).numFmt = '0.00%';
      row.eachCell(c => { c.border = border; });
    });

    const chartRowIndex = worksheet.rowCount + 2;
    worksheet.addImage(chartId, { tl: { col: 1, row: chartRowIndex }, ext: { width: 400, height: 400 } });

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `GO-F-029_Productividad_Nocturna_${new Date().getTime()}.xlsx`);
  }
}