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
  placas = ['BBU-840', 'ACK-826', 'BHH-935', 'BBT-751', 'F4D-855', 'AYJ-784', 'BAN-786', 'B0L-756', 'F4B-804', 'BEL-713', 'F6Y-818', 'BSY-931', 'AUH-927', 'F6Z-867', 'BHI-940', 'B7M-872', 'BTB-817', 'AUH-916', 'F4D-707'];
  operadores: any[] = [];
  clientes: any[] = [];

  // Modelos de Filtro
  filtroPlaca: string[] = [];
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

  // Variables de control para la gráfica
  modoUnidad: 'KG' | 'CILINDROS' = 'KG';
  modoAgrupacion: 'GLOBAL' | 'PLACA' = 'GLOBAL';

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
      // SI HAY PLACAS SELECCIONADAS Y LA PLACA DEL DESPACHO NO ESTÁ EN EL ARREGLO, NO PASA
      if (this.filtroPlaca.length > 0 && !this.filtroPlaca.includes(d.placa)) pasa = false;
      
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

    this.actualizarDatosGrafico();
  }

  actualizarDatosGrafico() {
    if (!this.rendimientoChartRef) return;

    let labels: string[] = [];
    let dataProg: number[] = [];
    let dataVend: number[] = [];
    let dataMerma: number[] = [];

    // Función mágica que convierte a KG o lo deja en unidades físicas
    const factor = (qty: number, peso: number) => this.modoUnidad === 'KG' ? (qty * peso) : qty;

    if (this.modoAgrupacion === 'GLOBAL') {
      labels = [this.modoUnidad === 'KG' ? 'Volumen Global (KG)' : 'Unidades Totales (Cilindros)'];
      let prog = 0, vend = 0, merma = 0;
      
      this.despachosFiltrados.forEach(r => {
        prog += factor(r.p05||0, 5) + factor(r.p10||0, 10) + factor(r.p15||0, 15) + factor(r.p45||0, 45);
        vend += factor(r.v05||0, 5) + factor(r.v10||0, 10) + factor(r.v15||0, 15) + factor(r.v45||0, 45);
        merma += factor(r.m05||0, 5) + factor(r.m10||0, 10) + factor(r.m15||0, 15) + factor(r.m45||0, 45);
      });

      dataProg = [prog];
      dataVend = [vend];
      dataMerma = [merma];
    } else {
      // AGRUPACIÓN POR PLACA (Detecta qué placas trabajaron y suma sus valores)
      const placasMap = new Map<string, {prog: number, vend: number, merma: number}>();
      
      this.despachosFiltrados.forEach(r => {
        const placa = r.placa || 'SIN PLACA';
        if (!placasMap.has(placa)) placasMap.set(placa, {prog: 0, vend: 0, merma: 0});
        
        const p = placasMap.get(placa)!;
        p.prog += factor(r.p05||0, 5) + factor(r.p10||0, 10) + factor(r.p15||0, 15) + factor(r.p45||0, 45);
        p.vend += factor(r.v05||0, 5) + factor(r.v10||0, 10) + factor(r.v15||0, 15) + factor(r.v45||0, 45);
        p.merma += factor(r.m05||0, 5) + factor(r.m10||0, 10) + factor(r.m15||0, 15) + factor(r.m45||0, 45);
      });

      labels = Array.from(placasMap.keys()).sort();
      labels.forEach(placa => {
        const p = placasMap.get(placa)!;
        dataProg.push(p.prog);
        dataVend.push(p.vend);
        dataMerma.push(p.merma);
      });
    }

    if (this.chartInstance) {
      this.chartInstance.destroy();
    }

    // 🔥 PLUGIN NATIVO PARA ETIQUETAS DE DATOS 🔥
    const pluginEtiquetas = {
      id: 'etiquetasDatos',
      afterDatasetsDraw(chart: any) {
        const { ctx } = chart;
        chart.data.datasets.forEach((dataset: any, i: number) => {
          const meta = chart.getDatasetMeta(i);
          meta.data.forEach((bar: any, index: number) => {
            const data = dataset.data[index];
            if (data > 0) { // Solo dibuja el número si es mayor a cero
              ctx.font = 'bold 11px sans-serif';
              ctx.fillStyle = '#475569'; // Color de texto slate-600
              ctx.textAlign = 'center';
              ctx.textBaseline = 'bottom';
              
              // Le da formato con separador de miles (Ej: 31,065)
              const textoFormateado = Number(data).toLocaleString(); 
              
              // Dibuja el texto 5 pixeles por encima de la barra
              ctx.fillText(textoFormateado, bar.x, bar.y - 5);
            }
          });
        });
      }
    };

    this.chartInstance = new Chart(this.rendimientoChartRef.nativeElement, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [
          { label: 'Programado', data: dataProg, backgroundColor: '#FBBF24', borderRadius: 4 },
          { label: 'Vendido', data: dataVend, backgroundColor: '#10B981', borderRadius: 4 },
          { label: 'Retornado/Merma', data: dataMerma, backgroundColor: '#EF4444', borderRadius: 4 }
        ]
      },
      options: { 
        responsive: true, 
        maintainAspectRatio: false,
        layout: {
          padding: { top: 25 } // Espacio extra arriba para que los números no se corten
        },
        scales: { 
          y: { 
            title: { 
              display: true, 
              text: this.modoUnidad === 'KG' ? 'Kilogramos (KG)' : 'Unidades Físicas (Cilindros)' 
            } 
          } 
        }
      },
      plugins: [pluginEtiquetas] // <-- INYECTAMOS EL PLUGIN AQUÍ
    });
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

  togglePlaca(placa: string) {
    const index = this.filtroPlaca.indexOf(placa);
    if (index > -1) {
      this.filtroPlaca.splice(index, 1); // Si ya está, la quitamos
    } else {
      this.filtroPlaca.push(placa); // Si no está, la agregamos
    }
    this.aplicarFiltros();
  }
}