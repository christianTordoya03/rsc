import { Component, OnInit, ElementRef, ViewChild } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms'; // Necesario para los filtros
import { Chart, registerables } from 'chart.js';
import saveAs from 'file-saver';
import * as ExcelJS from 'exceljs';
import { SupabaseService } from '../../core/services/supabase';
import { ToastService } from '../../core/services/toast';

Chart.register(...registerables);

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [CommonModule, FormsModule], // Importamos FormsModule aquí
  templateUrl: './dashboard.html'
})
export class Dashboard implements OnInit {
  @ViewChild('rendimientoChart') rendimientoChartRef!: ElementRef;
  chartInstance: Chart | null = null;
  isLoading = false;

  // Datos base
  despachosCrudos: any[] = [];
  despachosFiltrados: any[] = [];
  
  // Catálogos para filtros
  placas = ['BBU-840', 'ACK-826', 'BHH-935', 'BBT-751', 'F4D-855', 'AYJ-784', 'BAN-786', 'B0L-756', 'F4B-804', 'BEL-713'];
  operadores: any[] = [];
  clientes: any[] = [];

  // Modelos de Filtro
  filtroPlaca = '';
  filtroOperador = '';
  filtroCliente = '';
  fechaInicio = '';
  fechaFin = '';

  // KPIs Globales (En Kilogramos)
  kpiProgramadoKg = 0;
  kpiVendidoKg = 0;
  kpiRetornadoKg = 0;
  kpiPorcentajeVenta = 0;
  kpiPorcentajeRetorno = 0;

  constructor(
    private supabase: SupabaseService,
    private toast: ToastService
  ) {}

  async ngOnInit() {
    // Por defecto filtramos el mes actual
    const hoy = new Date();
    this.fechaInicio = new Date(hoy.getFullYear(), hoy.getMonth(), 1).toISOString().split('T')[0];
    this.fechaFin = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0).toISOString().split('T')[0];

    await this.cargarCatalogos();
    await this.cargarDespachos();
  }

  async cargarCatalogos() {
    const { data: ops } = await this.supabase.perfiles.select('id, nombre_completo').eq('rol', 'OPERACIONES');
    this.operadores = ops || [];
    
    const { data: clis } = await this.supabase.clientes.select('id, razon_social');
    this.clientes = clis || [];
  }

  async cargarDespachos() {
    this.isLoading = true;
    try {
      const { data, error } = await this.supabase.despachos
        .select(`
          *,
          clientes ( razon_social, direccion, distrito ),
          perfiles!despachos_operador_id_fkey ( nombre_completo )
        `)
        .eq('estado', 'COMPLETADO') // Solo cuenta lo que el operador ya finalizó
        .order('fecha_despacho', { ascending: false });

      if (error) throw error;
      this.despachosCrudos = data || [];
      this.aplicarFiltros();
    } catch (error: any) {
      this.toast.showError('Error al cargar datos: ' + error.message);
    } finally {
      this.isLoading = false;
    }
  }

  aplicarFiltros() {
    this.despachosFiltrados = this.despachosCrudos.filter(d => {
      let pasa = true;
      if (this.filtroPlaca && d.placa !== this.filtroPlaca) pasa = false;
      if (this.filtroOperador && d.operador_id !== this.filtroOperador) pasa = false;
      if (this.filtroCliente && d.cliente_id !== this.filtroCliente) pasa = false;
      if (this.fechaInicio && d.fecha_despacho < this.fechaInicio) pasa = false;
      if (this.fechaFin && d.fecha_despacho > this.fechaFin) pasa = false;
      return pasa;
    });

    this.calcularMetricasYGraficos();
  }

  calcularMetricasYGraficos() {
    this.kpiProgramadoKg = 0;
    this.kpiVendidoKg = 0;
    this.kpiRetornadoKg = 0;

    this.despachosFiltrados.forEach(r => {
      // Cálculo de pesos multiplicando cantidad x peso del cilindro
      this.kpiProgramadoKg += (r.p05*5) + (r.p10*10) + (r.pm15*15) + (r.p15*15) + (r.p45*45);
      this.kpiVendidoKg += (r.v05*5) + (r.v10*10) + (r.vm15*15) + (r.v15*15) + (r.v45*45);
      this.kpiRetornadoKg += (r.m05*5) + (r.m10*10) + (r.mm15*15) + (r.m15*15) + (r.m45*45);
    });

    this.kpiPorcentajeVenta = this.kpiProgramadoKg > 0 ? (this.kpiVendidoKg / this.kpiProgramadoKg) * 100 : 0;
    this.kpiPorcentajeRetorno = this.kpiProgramadoKg > 0 ? (this.kpiRetornadoKg / this.kpiProgramadoKg) * 100 : 0;

    this.dibujarGrafico();
  }

  dibujarGrafico() {
    if (!this.rendimientoChartRef) return;
    if (this.chartInstance) this.chartInstance.destroy();

    this.chartInstance = new Chart(this.rendimientoChartRef.nativeElement, {
      type: 'bar',
      data: {
        labels: ['Volumen Global (KG)'],
        datasets: [
          { label: 'Programado (Kg)', data: [this.kpiProgramadoKg], backgroundColor: '#FBBF24', borderRadius: 4 },
          { label: 'Vendido (Kg)', data: [this.kpiVendidoKg], backgroundColor: '#10B981', borderRadius: 4 },
          { label: 'Retornado/Merma (Kg)', data: [this.kpiRetornadoKg], backgroundColor: '#EF4444', borderRadius: 4 }
        ]
      },
      options: { responsive: true, maintainAspectRatio: false }
    });
  }

  async exportarExcel() {
    if (this.despachosFiltrados.length === 0) return this.toast.showWarning('No hay datos para exportar');
    
    this.toast.showInfo('Generando consolidado Excel...');
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Consolidado');

    // Estilos basados en tu imagen del formato Excel
    const fontBold = { name: 'Arial', bold: true, size: 9 };
    const borderAll: Partial<ExcelJS.Borders> = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
    
    const fillHeaderGeneral: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
    const fillProg: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC000' } }; // Naranja/Amarillo
    const fillVend: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF92D050' } }; // Verde
    const fillReto: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF8080' } }; // Rojo claro

    // Fila 1: Grupos Principales
    worksheet.mergeCells('A1:D1');
    worksheet.mergeCells('E1:J1'); worksheet.getCell('E1').value = 'PROGRAMADO'; worksheet.getCell('E1').fill = fillProg;
    worksheet.mergeCells('K1:P1'); worksheet.getCell('K1').value = 'VENDIDO'; worksheet.getCell('K1').fill = fillVend;
    worksheet.mergeCells('Q1:X1'); worksheet.getCell('Q1').value = 'RETORNADO'; worksheet.getCell('Q1').fill = fillReto;
    worksheet.getRow(1).alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getRow(1).font = fontBold;

    // Fila 2: Sub Cabeceras
    const headers = [
      'ITEM', 'Fecha', 'UNIDAD (Chofer)', 'PLACA CARGA',
      '05 KG', '10 KG', 'M-15', '15 KG', '45 KG', 'Programado Kilogramo',
      '05 KG', '10 KG', 'M-15', '15 KG', '45 KG', 'Vendido Kilogramo',
      '05 KG', '10 KG', 'M-15', '15 KG', '45 KG', 'Retorno Kilogramo', '% de RETORNO', '% DE VENTA',
      'Punto Final (Cliente)', 'OBSERVACIONES'
    ];
    worksheet.addRow(headers);
    worksheet.getRow(2).eachCell(cell => {
      cell.font = fontBold;
      cell.border = borderAll;
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });

    // Llenado de Datos
    this.despachosFiltrados.forEach((r, idx) => {
      // Cálculos de Peso
      const kgProg = (r.p05*5) + (r.p10*10) + (r.pm15*15) + (r.p15*15) + (r.p45*45);
      const kgVend = (r.v05*5) + (r.v10*10) + (r.vm15*15) + (r.v15*15) + (r.v45*45);
      const kgReto = (r.m05*5) + (r.m10*10) + (r.mm15*15) + (r.m15*15) + (r.m45*45);
      
      const pctRetorno = kgProg > 0 ? (kgReto / kgProg) : 0;
      const pctVenta = kgProg > 0 ? (kgVend / kgProg) : 0;

      const row = worksheet.addRow([
        idx + 1, r.fecha_despacho, r.perfiles?.nombre_completo, r.placa,
        r.p05, r.p10, r.pm15, r.p15, r.p45, kgProg,
        r.v05, r.v10, r.vm15, r.v15, r.v45, kgVend,
        r.m05, r.m10, r.mm15, r.m15, r.m45, kgReto, pctRetorno, pctVenta,
        r.clientes?.razon_social, r.observaciones
      ]);

      // Formato de porcentaje a las columnas correspondientes
      row.getCell(23).numFmt = '0.00%'; // % Retorno
      row.getCell(24).numFmt = '0.00%'; // % Venta
      
      row.eachCell(c => { c.border = borderAll; c.alignment = { vertical: 'middle' }; });
    });

    // Anchos de columna automáticos básicos
    worksheet.columns.forEach(col => { col.width = 12; });
    worksheet.getColumn(3).width = 25; // Chofer
    worksheet.getColumn(25).width = 30; // Cliente
    worksheet.getColumn(26).width = 35; // Obs

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `Consolidado_Gerencial_${new Date().getTime()}.xlsx`);
  }
}