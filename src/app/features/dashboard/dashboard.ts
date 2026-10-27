import { Component, OnInit, ElementRef, ViewChild, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { Chart, registerables } from 'chart.js';
import saveAs from 'file-saver';
import * as ExcelJS from 'exceljs';
import html2canvas from 'html2canvas'; 
import { SupabaseService } from '../../core/services/supabase';
import { ToastService } from '../../core/services/toast';

Chart.register(...registerables);

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './dashboard.html'
})
export class Dashboard implements OnInit, OnDestroy {
  private autoRefreshSub: any;
  @ViewChild('rendimientoChart') rendimientoChartRef!: ElementRef;
  @ViewChild('contenedorGrafico') contenedorGraficoRef!: ElementRef; 
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
    await this.cargarDespachos(); // Primera carga con pantalla de loading
    
    // Encendemos el motor silencioso
    this.iniciarSincronizacionSilenciosa();
  }

  ngOnDestroy() {
    if (this.autoRefreshSub) {
      clearInterval(this.autoRefreshSub);
    }
  }

  iniciarSincronizacionSilenciosa() {
    this.autoRefreshSub = setInterval(async () => {
      if (!this.isLoading) {
        try {
          await this.cargarDespachos(true); // El "true" le dice que lo haga en silencio
        } catch (err) {}
      }
    }, 10000); // 10 segundos
  }

  async cargarCatalogos() {
    const { data: ops } = await this.supabase.perfiles.select('id, nombre_completo').eq('rol', 'OPERACIONES');
    this.operadores = ops || [];
    
    const { data: clis } = await this.supabase.clientes.select('id, razon_social');
    this.clientes = clis || [];
  }

  // Agregamos el parámetro isSilent para no interrumpir al usuario
  async cargarDespachos(isSilent = false) {
    if (!isSilent) this.isLoading = true;
    try {
      const { data, error } = await this.supabase.despachos
        .select(`
          *,
          clientes ( razon_social, direccion, distrito ),
          perfiles!despachos_operador_id_fkey ( nombre_completo )
        `)
        .neq('estado', 'ANULADO') 
        .order('fecha_despacho', { ascending: false });

      if (error) throw error;
      this.despachosCrudos = data || [];
      this.aplicarFiltros();
    } catch (error: any) {
      if (!isSilent) this.toast.showError('Error al cargar datos: ' + error.message);
    } finally {
      if (!isSilent) this.isLoading = false;
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
      this.kpiProgramadoKg += ((r.p05||0)*5) + ((r.p10||0)*10) + ((r.p15||0)*15) + ((r.p45||0)*45);
      this.kpiVendidoKg += ((r.v05||0)*5) + ((r.v10||0)*10) + ((r.v15||0)*15) + ((r.v45||0)*45);
      this.kpiRetornadoKg += ((r.m05||0)*5) + ((r.m10||0)*10) + ((r.m15||0)*15) + ((r.m45||0)*45);
    });

    this.kpiPorcentajeVenta = this.kpiProgramadoKg > 0 ? (this.kpiVendidoKg / this.kpiProgramadoKg) * 100 : 0;
    this.kpiPorcentajeRetorno = this.kpiProgramadoKg > 0 ? (this.kpiRetornadoKg / this.kpiProgramadoKg) * 100 : 0;

    this.dibujarGrafico();
  }

  dibujarGrafico() {
    if (!this.rendimientoChartRef) return;
    
    // MEJORA: En vez de destruir la gráfica cada 10 seg, solo actualizamos sus datos suavemente
    if (this.chartInstance) {
      this.chartInstance.data.datasets[0].data = [this.kpiProgramadoKg];
      this.chartInstance.data.datasets[1].data = [this.kpiVendidoKg];
      this.chartInstance.data.datasets[2].data = [this.kpiRetornadoKg];
      this.chartInstance.update();
    } else {
      // Si no existe, la creamos por primera vez
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
  }

  async exportarExcel() {
    if (this.despachosFiltrados.length === 0) return this.toast.showWarning('No hay datos para exportar');
    
    this.toast.showInfo('Generando consolidado Excel...');
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Consolidado');

    const fontBold = { name: 'Arial', bold: true, size: 9 };
    const borderAll: Partial<ExcelJS.Borders> = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
    
    const fillProg: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC000' } }; 
    const fillVend: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF92D050' } }; 
    const fillReto: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF8080' } }; 

    worksheet.mergeCells('A1:D1');
    worksheet.mergeCells('E1:I1'); worksheet.getCell('E1').value = 'PROGRAMADO'; worksheet.getCell('E1').fill = fillProg;
    worksheet.mergeCells('J1:N1'); worksheet.getCell('J1').value = 'VENDIDO'; worksheet.getCell('J1').fill = fillVend;
    worksheet.mergeCells('O1:U1'); worksheet.getCell('O1').value = 'RETORNADO / MERMA'; worksheet.getCell('O1').fill = fillReto;
    worksheet.getRow(1).alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getRow(1).font = fontBold;

    const headers = [
      'ITEM', 'Fecha', 'UNIDAD (Chofer)', 'PLACA CARGA',
      '05 KG', '10 KG', '15 KG', '45 KG', 'Programado (Kg)',
      '05 KG', '10 KG', '15 KG', '45 KG', 'Vendido (Kg)',
      '05 KG', '10 KG', '15 KG', '45 KG', 'Merma (Kg)', '% de RETORNO', '% DE VENTA',
      'Punto Final (Cliente)', 'OBSERVACIONES'
    ];
    worksheet.addRow(headers);
    worksheet.getRow(2).eachCell(cell => {
      cell.font = fontBold;
      cell.border = borderAll;
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });

    this.despachosFiltrados.forEach((r, idx) => {
      const kgProg = ((r.p05||0)*5) + ((r.p10||0)*10) + ((r.p15||0)*15) + ((r.p45||0)*45);
      const kgVend = ((r.v05||0)*5) + ((r.v10||0)*10) + ((r.v15||0)*15) + ((r.v45||0)*45);
      const kgReto = ((r.m05||0)*5) + ((r.m10||0)*10) + ((r.m15||0)*15) + ((r.m45||0)*45);
      
      const pctRetorno = kgProg > 0 ? (kgReto / kgProg) : 0;
      const pctVenta = kgProg > 0 ? (kgVend / kgProg) : 0;

      const row = worksheet.addRow([
        idx + 1, r.fecha_despacho, r.perfiles?.nombre_completo || 'Sin Asignar', r.placa,
        r.p05||0, r.p10||0, r.p15||0, r.p45||0, kgProg,
        r.v05||0, r.v10||0, r.v15||0, r.v45||0, kgVend,
        r.m05||0, r.m10||0, r.m15||0, r.m45||0, kgReto, pctRetorno, pctVenta,
        r.clientes?.razon_social || 'Desconocido', r.observaciones || ''
      ]);

      row.getCell(20).numFmt = '0.00%'; 
      row.getCell(21).numFmt = '0.00%'; 
      
      row.eachCell(c => { c.border = borderAll; c.alignment = { vertical: 'middle' }; });
    });

    worksheet.columns.forEach(col => { col.width = 12; });
    worksheet.getColumn(3).width = 25; 
    worksheet.getColumn(22).width = 30; 
    worksheet.getColumn(23).width = 35; 

    if (this.contenedorGraficoRef) {
      try {
        const canvas = await html2canvas(this.contenedorGraficoRef.nativeElement, { scale: 2, backgroundColor: '#ffffff' });
        const graficaBase64 = canvas.toDataURL('image/png');
        const imageId = workbook.addImage({ base64: graficaBase64, extension: 'png' });
        worksheet.addImage(imageId, { tl: { col: 1, row: worksheet.rowCount + 2 }, ext: { width: 550, height: 350 } });
      } catch (e) {
        console.error('Error al capturar la gráfica', e);
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `Consolidado_Gerencial_${new Date().getTime()}.xlsx`);
  }
}