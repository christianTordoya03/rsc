import { Component, OnInit, OnDestroy, ElementRef, ViewChild } from '@angular/core';
import { CommonModule, DatePipe } from '@angular/common';
import { FormBuilder, FormGroup, ReactiveFormsModule, Validators, FormsModule } from '@angular/forms';
import * as ExcelJS from 'exceljs';
import { SupabaseService } from '../../core/services/supabase';
import { ToastService } from '../../core/services/toast';
import { ConfirmModalComponent } from '../../shared/components/confirm-modal/confirm-modal';

@Component({
  selector: 'app-programacion-rutas',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule, FormsModule, ConfirmModalComponent],
  providers: [DatePipe], 
  templateUrl: './programacion-rutas.html'
})
export class ProgramacionRutasComponent implements OnInit, OnDestroy {
  private autoRefreshSub: any;
  @ViewChild('fileInput') fileInput!: ElementRef;

  modoVista: 'DISTRIBUIDORES' | 'INDUSTRIAS' = 'DISTRIBUIDORES';

  showForm = false;
  isLoading = false;
  isEditing = false;
  currentEditId: string | null = null;
  showDeleteAllConfirm = false;
  
  rutas: any[] = [];
  clientesActivos: any[] = [];
  clientesNoEncontrados: string[] = [];
  
  // MODAL FECHA IMPORTACION
  showModalFechaImportacion = false;
  fechaImportacion = '';
  
  placas = ['BBU-840', 'ACK-826', 'BHH-935', 'BBT-751', 'F4D-855', 'AYJ-784', 'BAN-786', 'B0L-756', 'F4B-804', 'BEL-713', 'F6Y-818', 'BSY-931', 'AUH-927', 'F6Z-867', 'BHI-940', 'B7M-872', 'BTB-817', 'AUH-916', 'F4D-707'];

  rutaForm: FormGroup;
  filtroCliente = '';
  filtroPlaca = '';
  filtroFecha = new Date().toISOString().split('T')[0]; // POR DEFECTO CARGA "HOY"

  paginaPendientes = 1;
  paginaHistorial = 1;
  itemsPorPagina = 50;

  showFotoModal = false;
  fotoUrlActual = '';
  showDetalleModal = false;
  rutaDetalleActual: any = null;
  historialEntregas: any[] = [];
  editingEntregaId: string | null = null;
  entregaEditForm: any = {};

  confirmDialog = {
    isOpen: false,
    title: '',
    message: '',
    confirmText: '',
    actionType: '',
    payload: null as any
  };

  userName = ''; 
  userRole = ''; 

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
      observacion: [''], 
      p05: [0, Validators.min(0)], p10: [0, Validators.min(0)], p15: [0, Validators.min(0)], p45: [0, Validators.min(0)]
    });
  }

  async ngOnInit() {
    this.cargarErroresGuardados();

    const { data: { session } } = await this.supabase.auth.getSession();
    if (session) {
      const { data } = await this.supabase.perfiles.select('nombre_completo, rol').eq('id', session.user.id).single();
      if (data) {
        this.userName = data.nombre_completo;
        this.userRole = data.rol || '';
      }
    }
    await this.cargarDatosBase();
    await this.cargarRutasDelDia();
    this.iniciarSincronizacionSilenciosa();
  }

  abrirConfirmacion(title: string, message: string, confirmText: string, actionType: string, payload: any) {
    this.confirmDialog = { isOpen: true, title, message, confirmText, actionType, payload };
  }

  cerrarConfirmacion() {
    this.confirmDialog.isOpen = false;
  }

  // ==== LOGICA DEL MODAL DE FECHA DE IMPORTACION ====
  abrirModalFechaImportacion() {
    this.fechaImportacion = this.filtroFecha; 
    this.showModalFechaImportacion = true;
  }

  cerrarModalFechaImportacion() {
    this.showModalFechaImportacion = false;
  }

  continuarImportacion() {
    if (!this.fechaImportacion) return this.toast.showWarning('Debe seleccionar una fecha');
    this.showModalFechaImportacion = false;
    this.fileInput.nativeElement.click(); 
  }
  // =================================================

  ejecutarAccionConfirmada() {
    const { actionType, payload } = this.confirmDialog;
    this.cerrarConfirmacion();
    
    if (actionType === 'ELIMINAR_REPORTE') this.eliminarEntregaParcial(payload);
    else if (actionType === 'EDITAR_REPORTE') this.guardarEdicionEntrega(payload);
    else if (actionType === 'ANULAR_RUTA') this.anularRuta(payload);
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
          const { data, error } = await this.supabase.despachos
            .select(`*, clientes ( razon_social, tipo_cliente ), perfiles!despachos_operador_id_fkey ( nombre_completo )`)
            .order('creado_en', { ascending: false });

          if (!error && data) {
            this.rutas = data; 
          }
        } catch (err) {} 
      }
    }, 10000); 
  }

  trackById(index: number, item: any): string {
    return item.id;
  }

  cambiarModoVista(modo: 'DISTRIBUIDORES' | 'INDUSTRIAS') {
    this.modoVista = modo;
    this.paginaPendientes = 1;
    this.paginaHistorial = 1;
    this.cargarErroresGuardados(); 
  }

  cargarErroresGuardados() {
    // Ahora busca la alerta usando el MODO y la FECHA EXACTA
    const erroresGuardados = localStorage.getItem(`rutasErrores_${this.modoVista}_${this.filtroFecha}`);
    if (erroresGuardados) {
      this.clientesNoEncontrados = JSON.parse(erroresGuardados);
    } else {
      this.clientesNoEncontrados = [];
    }
  }

  cerrarPanelErrores() {
    this.clientesNoEncontrados = [];
    localStorage.removeItem(`rutasErrores_${this.modoVista}_${this.filtroFecha}`);
  }

  async cargarDatosBase() {
    try {
      let todosLosClientes: any[] = [];
      let desde = 0;
      const paso = 500;

      while (true) {
        const { data, error } = await this.supabase.from('clientes')
          .select('id, razon_social, ruc, codigo_cliente, tipo_cliente') 
          .eq('estado', true)
          .order('id', { ascending: true }) 
          .range(desde, desde + paso - 1);

        if (error) throw error;
        if (!data || data.length === 0) break;

        todosLosClientes.push(...data);
        if (data.length < paso) break;
        desde += paso;
      }
      
      todosLosClientes.sort((a, b) => (a.razon_social || '').localeCompare(b.razon_social || ''));
      this.clientesActivos = todosLosClientes;
    } catch (error) {
      this.toast.showError('Error al cargar clientes');
    }
  }

  async cargarRutasDelDia() {
    this.isLoading = true;
    try {
      const { data, error } = await this.supabase.despachos
        .select(`*, clientes ( razon_social, tipo_cliente ), perfiles!despachos_operador_id_fkey ( nombre_completo )`)
        .order('creado_en', { ascending: false });

      if (error) throw error;
      this.rutas = data || [];
    } catch (error: any) {
      this.toast.showError('Error al cargar rutas: ' + error.message);
    } finally {
      this.isLoading = false;
    }
  }

  onSearchChange() {
    this.paginaPendientes = 1;
    this.paginaHistorial = 1;
    this.cargarErroresGuardados(); // <-- MAGIA: Actualiza la alerta roja al cambiar la fecha
  }

  cumpleFiltros(r: any): boolean {
    const matchPlaca = !this.filtroPlaca || r.placa.toLowerCase().includes(this.filtroPlaca.toLowerCase());
    const nombreCliente = r.clientes?.razon_social || '';
    const matchCliente = !this.filtroCliente || nombreCliente.toLowerCase().includes(this.filtroCliente.toLowerCase());
    const matchFecha = r.fecha_despacho === this.filtroFecha; 
    return matchPlaca && matchCliente && matchFecha;
  }

  get rutasPendientes() { 
    const filtradas = this.rutas.filter(r => {
      const origen = r.tipo_negocio || (r.clientes?.tipo_cliente === 'B2B/INDUSTRIAS' ? 'INDUSTRIAS' : 'DISTRIBUIDORES');
      const coincideModo = this.modoVista === origen;
      return coincideModo && (r.estado === 'PENDIENTE' || r.estado === 'EN_RUTA' || r.estado === 'EN_PROCESO') && this.cumpleFiltros(r);
    }); 

    return filtradas.sort((a, b) => {
      const avanceA = (a.v05||0) + (a.v10||0) + (a.v15||0) + (a.v45||0) + (a.m05||0) + (a.m10||0) + (a.m15||0) + (a.m45||0);
      const avanceB = (b.v05||0) + (b.v10||0) + (b.v15||0) + (b.v45||0) + (b.m05||0) + (b.m10||0) + (b.m15||0) + (b.m45||0);

      if (avanceA > 0 && avanceB === 0) return -1;
      if (avanceB > 0 && avanceA === 0) return 1;

      const clienteA = a.clientes?.razon_social || '';
      const clienteB = b.clientes?.razon_social || '';
      return clienteA.localeCompare(clienteB);
    });
  }

  get rutasHistorial() { 
    return this.rutas.filter(r => {
      const origen = r.tipo_negocio || (r.clientes?.tipo_cliente === 'B2B/INDUSTRIAS' ? 'INDUSTRIAS' : 'DISTRIBUIDORES');
      const coincideModo = this.modoVista === origen;
      return coincideModo && (r.estado === 'COMPLETADO' || r.estado === 'ANULADO') && this.cumpleFiltros(r);
    }); 
  }

  get clientesParaFormulario() {
    return this.clientesActivos.filter(c => {
      const tipo = c.tipo_cliente || 'DISTRIBUIDORES';
      if (this.modoVista === 'INDUSTRIAS') {
        return tipo.includes('INDUSTRIAS');
      } else {
        return tipo.includes('DISTRIBUIDORES') || (!tipo.includes('INDUSTRIAS') && !tipo.includes('DISTRIBUIDORES'));
      }
    });
  }

  get totalPaginasPendientes() { return Math.ceil(this.rutasPendientes.length / this.itemsPorPagina) || 1; }
  cambiarPaginaPendientes(delta: number) {
    const nueva = this.paginaPendientes + delta;
    if (nueva >= 1 && nueva <= this.totalPaginasPendientes) this.paginaPendientes = nueva;
  }

  get totalPaginasHistorial() { return Math.ceil(this.rutasHistorial.length / this.itemsPorPagina) || 1; }
  cambiarPaginaHistorial(delta: number) {
    const nueva = this.paginaHistorial + delta;
    if (nueva >= 1 && nueva <= this.totalPaginasHistorial) this.paginaHistorial = nueva;
  }

  getCellValue(cell: any): any {
    if (!cell) return null;
    if (cell.type === ExcelJS.ValueType.Error) return null;
    if (cell.isMerged && cell.master) return cell.master.value;
    return cell.value;
  }

  getCellValueText(cell: any): string {
    return this.getString(this.getCellValue(cell));
  }

  getCellValueNum(cell: any): number {
    return this.getNum(this.getCellValue(cell));
  }

  getNum(val: any): number {
    if (val === null || val === undefined || val === '') return 0;
    if (typeof val === 'number') return val;
    if (typeof val === 'object' && val.result !== undefined && val.result !== null) return Number(val.result) || 0; 
    const parsed = Number(val);
    return isNaN(parsed) ? 0 : parsed;
  }

  getString(val: any): string {
    if (val === null || val === undefined) return '';
    if (typeof val === 'string') return val.trim().toUpperCase();
    if (typeof val === 'number') return String(val).trim().toUpperCase();
    
    if (typeof val === 'object') {
      if (val.result !== undefined && val.result !== null) return this.getString(val.result); 
      if (val.richText) return val.richText.map((rt: any) => rt.text).join('').trim().toUpperCase();
      if (val.text) return this.getString(val.text);
      return ''; 
    }
    return String(val).trim().toUpperCase();
  }

  encontrarPlacaOficial(texto: string): string | null {
    if (!texto) return null;
    const t = texto.toUpperCase();
    const txtLimpio = t.replace(/[^A-Z0-9]/g, '');

    for (const p of this.placas) {
      const pClean = p.replace(/-/g, '');
      if (txtLimpio.includes(pClean)) return p; 
    }
    
    const match = t.match(/([A-Z0-9]{3})\s*-?\s*([A-Z0-9]{3})/);
    if (match) {
       const posiblePlaca = `${match[1]}-${match[2]}`;
       if (/[A-Z]/.test(posiblePlaca) && /[0-9]/.test(posiblePlaca)) {
           if (!this.placas.includes(posiblePlaca)) this.placas.push(posiblePlaca);
           return posiblePlaca;
       }
    }
    return null;
  }

  async importarExcel(event: any) {
    const file = event.target.files[0];
    if (!file) return;

    this.isLoading = true;
    this.toast.showInfo(`Analizando Excel de ${this.modoVista}...`);

    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await file.arrayBuffer());
      
      // PRIORIDAD PARA LA HOJA IMPRIMIR
      const worksheet = workbook.getWorksheet('Imprimir') || workbook.getWorksheet('Programacion') || workbook.getWorksheet('Programación') || workbook.worksheets[0];
      const sheetInicial = workbook.getWorksheet('PROGRAMACION INICIAL'); 

      const normalize = (str: any) => {
          if (str === null || str === undefined) return '';
          return String(str).trim().toUpperCase().replace(/\s+/g, ' ');
      };
      const stripPunct = (str: any) => {
          if (str === null || str === undefined) return '';
          return String(str).replace(/[^A-Z0-9]/g, '');
      };
      
      const mapNombres = new Map<string, string>();
      const mapCleanNombres = new Map<string, string>();
      const mapRucs = new Map<string, string>();
      const mapCodigos = new Map<string, string>();

      this.clientesActivos.forEach((c: any) => {
        if (c.razon_social) {
          const norm = normalize(c.razon_social);
          mapNombres.set(norm, c.id);
          mapCleanNombres.set(stripPunct(norm), c.id);
        }
        if (c.ruc) mapRucs.set(normalize(c.ruc), c.id);
        if (c.codigo_cliente) mapCodigos.set(normalize(c.codigo_cliente), c.id); 
      });

      const mapPedidoADistribuidor = new Map<string, string>();

      if (sheetInicial) {
        sheetInicial.eachRow((row: any) => {
          let codSap = '';
          let pedido = '';
          let nombre = '';

          for (let i = 1; i <= 15; i++) { 
            const val = this.getCellValueText(row.getCell(i));
            if (!val || val === '0') continue;
            const valClean = val.replace(/\s/g, '');

            if (/^4[45]\d{5}/.test(valClean)) pedido = valClean; 
            else if (/^\d{5,10}$/.test(valClean)) codSap = valClean; 
            else if (val.length > 4 && /[A-Z]/.test(val) && !['LIMA LOCAL', 'SUR CHICO', 'NORTE CHICO', 'PEDIDO', 'DISTRIBUIDOR', 'CLIENTE', 'DIRECCION', 'NOMBRE', 'NOMBRE DE CLIENTE'].includes(val)) {
              if (!nombre) nombre = val; 
            }
          }
          if (pedido && nombre) mapPedidoADistribuidor.set(pedido, nombre);
          if (pedido && codSap && !nombre) mapPedidoADistribuidor.set(pedido, codSap); 
        });
      }

      const findClienteId = (texto: string) => {
        if (!texto) return null;
        const t = normalize(texto);
        if (['DISTRIBUIDOR', 'CLIENTE', 'DIRECCION', 'NOMBRE', 'NOMBRE DE CLIENTE', 'DESTINATARIO'].includes(t)) return null;

        if (this.modoVista === 'INDUSTRIAS' && mapCodigos.has('B2B-' + t)) return mapCodigos.get('B2B-' + t);
        if (mapCodigos.has(t)) return mapCodigos.get(t); 
        if (mapRucs.has(t)) return mapRucs.get(t);
        if (mapNombres.has(t)) return mapNombres.get(t);

        const matchNum = t.match(/\b\d{5,10}\b/);
        if (matchNum) {
          const cod = matchNum[0];
          if (this.modoVista === 'INDUSTRIAS' && mapCodigos.has('B2B-' + cod)) return mapCodigos.get('B2B-' + cod);
          if (mapCodigos.has(cod)) return mapCodigos.get(cod);
        }

        const tClean = stripPunct(t);
        if (mapCleanNombres.has(tClean)) return mapCleanNombres.get(tClean);

        const matchPrefijo = t.match(/^(\d+)\s*-/);
        const prefijo = matchPrefijo ? matchPrefijo[1] : null;

        for (const [nameDB, id] of mapNombres.entries()) {
          if (prefijo && nameDB.startsWith(prefijo + ' -')) return id;
          if (t.length > 4 && nameDB.length > 4 && (nameDB.includes(t) || t.includes(nameDB))) return id;
        }

        for (const [nameCleanDB, id] of mapCleanNombres.entries()) {
           if (tClean.length > 4 && nameCleanDB.length > 4 && (nameCleanDB.includes(tClean) || tClean.includes(nameCleanDB))) return id;
        }

        const matchFallback = this.clientesActivos.find(c => {
            const cCod = stripPunct(c.codigo_cliente);
            const cRuc = stripPunct(c.ruc);
            const cNom = stripPunct(c.razon_social);
            
            if (cCod && cCod === tClean) return true;
            if (cRuc && cRuc === tClean) return true;
            if (cNom && cNom === tClean) return true;
            if (cNom && cNom.length > 4 && tClean.length > 4 && (cNom.includes(tClean) || tClean.includes(cNom))) return true;
            return false;
        });

        if (matchFallback) return matchFallback.id;

        return null;
      };

      const nuevasRutas: any[] = [];
      const noEncontrados = new Set<string>();
      
      const fechaProg = this.fechaImportacion; // 🔥 FUERZA LA FECHA ELEGIDA EN EL MODAL

      if (this.modoVista === 'INDUSTRIAS') {
        let cPlaca = 0, cTrans = 0, cCliente = 0, cCant = 0, cTipo = 0, cObs = 0, cEntrega = 0;
        let headerRow = -1;
        
        worksheet.eachRow((row: any, rowNum: number) => {
          if (rowNum <= 10 && headerRow === -1) {
            row.eachCell((cell: any, colNum: number) => {
              const val = this.getCellValueText(cell).toUpperCase();
              if (val === 'UNIDAD') cPlaca = colNum;
              if (val === 'N° TRANSPORTE') cTrans = colNum;
              if (val === 'NOMBRE DE CLIENTE') { cCliente = colNum; headerRow = rowNum; }
              if (val === 'CANT.' || val === 'CANTIDAD') cCant = colNum;
              if (val === 'TIPO DE PRODUCTO') cTipo = colNum;
              if (val === 'OBSERVACION') cObs = colNum;
              if (val === 'N° ENTREGA' || val === 'ENTREGA') cEntrega = colNum;
            });
          }
        });

        if (headerRow === -1) throw new Error('No se detectó el formato de Industrias (Falta columna "NOMBRE DE CLIENTE").');

        const rutasAgrupadas = new Map<string, any>();

        worksheet.eachRow((row: any, rowNumber: number) => {
          if (rowNumber <= headerRow) return;

          const txtPlaca = this.getCellValueText(row.getCell(cPlaca));
          const txtTrans = this.getCellValueText(row.getCell(cTrans));
          const txtCliente = this.getCellValueText(row.getCell(cCliente));
          const cant = this.getCellValueNum(row.getCell(cCant));
          const tipo = this.getCellValueText(row.getCell(cTipo)).toUpperCase();
          const obs = this.getCellValueText(row.getCell(cObs));
          const txtEntrega = this.getCellValueText(row.getCell(cEntrega));

          if (!txtCliente) return;

          let cartelLogistico: string | null = null;
          let esRecojo = false;

          const strCheck = (txtEntrega + ' ' + obs + ' ' + txtCliente).toUpperCase();
          if (strCheck.includes('RECOJO')) {
            esRecojo = true; cartelLogistico = '📦 RECOJO';
          } else if (strCheck.includes('BLOQUEAD') || strCheck.includes('CAMBIO')) {
            esRecojo = true; cartelLogistico = '🚫 ' + (strCheck.includes('CAMBIO') ? 'CAMBIO' : 'BLOQUEADO');
          }

          if (cant === 0 && !esRecojo) return;

          const placaFinal = this.encontrarPlacaOficial(txtPlaca) || this.encontrarPlacaOficial(txtTrans) || 'SIN PLACA';
          const obsFinal = cartelLogistico || obs;

          const clienteId = findClienteId(txtCliente);

          if (!clienteId) {
             const valorFallido = txtCliente || 'CELDA VACÍA';
             noEncontrados.add(`Fila ${rowNumber}: ${valorFallido}`);
             return;
          }

          const mapKey = `${placaFinal}_${clienteId}`;
          if (!rutasAgrupadas.has(mapKey)) {
            rutasAgrupadas.set(mapKey, {
              fecha_despacho: fechaProg, placa: placaFinal, cliente_id: clienteId,
              observacion: obsFinal || '', p05: 0, p10: 0, p15: 0, p45: 0,
              estado: 'PENDIENTE', venta_real: 0, merma: 0, programador_nombre: this.userName,
              tipo_negocio: this.modoVista 
            });
          }

          const ruta = rutasAgrupadas.get(mapKey);
          
          if (tipo.includes('5') && !tipo.includes('15') && !tipo.includes('45')) ruta.p05 += cant;
          if (tipo.includes('10')) ruta.p10 += cant;
          if (tipo.includes('15')) ruta.p15 += cant;
          if (tipo.includes('45')) ruta.p45 += cant;

          if (obsFinal && obsFinal !== ruta.observacion) {
            if (!ruta.observacion.includes(obsFinal)) {
               ruta.observacion = ruta.observacion ? `${ruta.observacion} | ${obsFinal}` : obsFinal;
            }
          }
        });

        nuevasRutas.push(...Array.from(rutasAgrupadas.values()));

      } else {
        let colPlaca = 0, colTransportes = 0, colPedido = 0, colEntrega = 0, colDistribuidor = 0; 
        let colLimite = 999; 
        const cols05: number[] = [], cols10: number[] = [], cols15: number[] = [], cols45: number[] = [];

        worksheet.eachRow((row: any, rowNum: number) => {
          if (rowNum <= 5) { 
            row.eachCell((cell: any, colNum: number) => {
              const val = this.getCellValueText(cell);
              if (val.includes('TOTAL KG') || val === 'TOTAL') {
                  if (colNum < colLimite) colLimite = colNum;
              }
              if (val === 'PLACA' && colPlaca === 0) colPlaca = colNum;
              if (val === 'TRANSPORTES' && colTransportes === 0) colTransportes = colNum;
              if (val === 'PEDIDO' && colPedido === 0) colPedido = colNum;
              if (val === 'ENTREGA' && colEntrega === 0) colEntrega = colNum;
              if ((val.includes('DISTRIBUIDOR') || val.includes('CLIENTE') || val.includes('DESTINATARIO') || val.includes('NOMBRE') || val.includes('RAZON SOCIAL')) && colDistribuidor === 0) {
                  colDistribuidor = colNum;
              }

              if (colNum < colLimite) {
                  const valLimpio = val.replace(/\s+/g, '');
                  // 1. Detección Clásica
                  if (valLimpio.includes('5KG') && !valLimpio.includes('15') && !valLimpio.includes('45')) cols05.push(colNum);
                  if (valLimpio.includes('10KG')) cols10.push(colNum);
                  if (val.includes('15 KG') || val.includes('M-15') || val.includes('M15') || valLimpio.includes('15KG') || valLimpio.includes('M15')) cols15.push(colNum);
                  if (valLimpio.includes('45KG')) cols45.push(colNum);
                  
                  // 2. Detección por Códigos SAP de Material (Hoja Imprimir)
                  if (['41000102', '41000103', '41000101'].includes(valLimpio)) cols05.push(colNum);
                  if (['41000104', '41000105', '41000110', '41000111', '41000108', '41000109'].includes(valLimpio)) cols10.push(colNum);
                  if (['41000115', '41000113', '41000114', '41000116'].includes(valLimpio)) cols15.push(colNum);
                  if (['41000119', '41000121', '41000106', '41000107'].includes(valLimpio)) cols45.push(colNum);
              }
            });
          }
        });

        if (colPlaca === 0) colPlaca = 1;
        if (colPedido === 0) colPedido = 3;
        if (colEntrega === 0) colEntrega = 4;
        if (colDistribuidor === 0) colDistribuidor = 5; 

        worksheet.eachRow((row: any, rowNumber: number) => {
          if (rowNumber < 5) return; 

          let esSubTotal = false;
          for (let i = 1; i <= 10; i++) {
            const valC = this.getCellValueText(row.getCell(i));
            if (valC.includes('SUB TOTAL') || valC.includes('SUBTOTAL') || valC === 'TOTAL') {
               esSubTotal = true; break;
            }
          }
          if (esSubTotal) return; 

          let p05 = 0, p10 = 0, p15 = 0, p45 = 0;
          cols05.forEach((c: number) => p05 += this.getCellValueNum(row.getCell(c)));
          cols10.forEach((c: number) => p10 += this.getCellValueNum(row.getCell(c)));
          cols15.forEach((c: number) => p15 += this.getCellValueNum(row.getCell(c)));
          cols45.forEach((c: number) => p45 += this.getCellValueNum(row.getCell(c)));

          const sumaCilindros = p05 + p10 + p15 + p45;

          const txtColPlaca = this.getCellValueText(row.getCell(colPlaca));
          const txtColTrans = colTransportes > 0 ? this.getCellValueText(row.getCell(colTransportes)) : '';
          const txtPedido = this.getCellValueText(row.getCell(colPedido));
          const txtEntrega = this.getCellValueText(row.getCell(colEntrega));
          let textoBusqueda = this.getCellValueText(row.getCell(colDistribuidor));

          let nombreDelPedido = '';
          if (txtPedido) {
             nombreDelPedido = mapPedidoADistribuidor.get(txtPedido) || '';
          }

          if (!textoBusqueda || textoBusqueda === '0') {
             textoBusqueda = nombreDelPedido || txtPedido;
          }

          let esRecojo = false;
          let cartelLogistico: string | null = null;

          const strCheck = (txtEntrega + ' ' + txtPedido + ' ' + textoBusqueda).toUpperCase();
          if (strCheck.includes('RECOJO')) {
            esRecojo = true; cartelLogistico = '📦 RECOJO';
          } else if (strCheck.includes('BLOQUEAD') || strCheck.includes('CAMBIO')) {
            esRecojo = true; cartelLogistico = '🚫 ' + (strCheck.includes('CAMBIO') ? 'CAMBIO' : 'BLOQUEADO');
          }

          if (sumaCilindros === 0 && !esRecojo) return;

          const combiTexto = txtColPlaca + ' ' + txtColTrans;
          if (combiTexto.includes('2 VUELT') || combiTexto.includes('2DA VUELT')) {
            cartelLogistico = cartelLogistico ? cartelLogistico + ' | 2da Vuelta' : '♻️ 2da Vuelta';
          }

          let placaFinal = this.encontrarPlacaOficial(txtColPlaca) || this.encontrarPlacaOficial(txtColTrans) || 'SIN PLACA'; 
          let obsFinal = cartelLogistico;

          const placaSecuestradora = this.encontrarPlacaOficial(txtEntrega);
          if (placaSecuestradora && !esRecojo) {
            placaFinal = placaSecuestradora; 
            obsFinal = '🚚 Reasignado'; 
          }

          let busquedaAvanzada = textoBusqueda;
          const matchN = textoBusqueda.match(/\b\d{5,10}\b/);
          if (matchN) busquedaAvanzada = matchN[0];

          const clienteId = findClienteId(busquedaAvanzada) || findClienteId(textoBusqueda);
          
          if (!clienteId) {
            const valorMostrar = nombreDelPedido || textoBusqueda || txtPedido || 'CELDA VACÍA';
            noEncontrados.add(`Fila ${rowNumber}: ${valorMostrar}`);
            return;
          }

          if (placaFinal && clienteId && (sumaCilindros > 0 || esRecojo)) {
            nuevasRutas.push({
              fecha_despacho: fechaProg,
              placa: placaFinal, 
              cliente_id: clienteId,
              observacion: obsFinal, 
              p05, p10, p15, p45,
              estado: 'PENDIENTE', venta_real: 0, merma: 0,
              programador_nombre: this.userName,
              tipo_negocio: this.modoVista 
            });
          }
        });
      }

      if (noEncontrados.size > 0) {
        const arrayErrores = Array.from(noEncontrados);
        localStorage.setItem(`rutasErrores_${this.modoVista}_${fechaProg}`, JSON.stringify(arrayErrores));
        // Si estás mirando el mismo día que acabas de subir, muéstralo de inmediato
        if (this.filtroFecha === fechaProg) {
           this.clientesNoEncontrados = arrayErrores;
        }
      } else {
        localStorage.removeItem(`rutasErrores_${this.modoVista}_${fechaProg}`);
        if (this.filtroFecha === fechaProg) {
           this.clientesNoEncontrados = []; 
        }
      }

      const rutasNuevasFiltradas = nuevasRutas.filter(nr => {
        return !this.rutas.some(r => 
           r.fecha_despacho === nr.fecha_despacho && 
           r.placa === nr.placa && 
           r.cliente_id === nr.cliente_id &&
           r.estado !== 'ANULADO'
        );
      });

      if (rutasNuevasFiltradas.length > 0) {
        const { data, error } = await this.supabase.despachos.insert(rutasNuevasFiltradas).select(`*, clientes ( razon_social, tipo_cliente ), perfiles!despachos_operador_id_fkey ( nombre_completo )`);
        if (error) throw error;
        this.rutas = [...(data || []), ...this.rutas];
        this.paginaPendientes = 1;
        this.toast.showSuccess(`Se procesaron ${rutasNuevasFiltradas.length} rutas de ${this.modoVista} correctamente.`);
      } else if (noEncontrados.size === 0) {
        this.toast.showInfo(`Las rutas de este Excel ya estaban programadas. No se agregaron duplicados.`);
      } else {
        this.toast.showWarning(`Revisión completada. Se detectaron fallos en ${noEncontrados.size} clientes.`);
      }

    } catch (error: any) {
      this.toast.showError('Error al procesar Excel: ' + error.message);
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
      const total = v.p05 + v.p10 + v.p15 + v.p45;
      if (total === 0) throw new Error('Programe al menos 1 cilindro');

      if (!this.isEditing) {
        const payload = {
          fecha_despacho: v.fecha_despacho, placa: v.placa, cliente_id: v.cliente_id,
          observacion: v.observacion || null,
          p05: v.p05, p10: v.p10, p15: v.p15, p45: v.p45,
          estado: 'PENDIENTE', venta_real: 0, merma: 0,
          programador_nombre: this.userName,
          tipo_negocio: this.modoVista 
        };

        const { data, error } = await this.supabase.despachos.insert(payload).select(`*, clientes ( razon_social, tipo_cliente ), perfiles!despachos_operador_id_fkey ( nombre_completo )`).single();
        if (error) throw error;
        this.rutas.unshift(data);
        this.toast.showSuccess('Carga asignada');
      } else {
        const payloadUpdate = {
          fecha_despacho: v.fecha_despacho, placa: v.placa, cliente_id: v.cliente_id,
          observacion: v.observacion || null,
          p05: v.p05, p10: v.p10, p15: v.p15, p45: v.p45,
          editado_en: new Date().toISOString(),
          editor_nombre: this.userName,
          tipo_negocio: this.modoVista 
        };

        const { data, error } = await this.supabase.despachos.update(payloadUpdate).eq('id', this.currentEditId).select(`*, clientes ( razon_social, tipo_cliente ), perfiles!despachos_operador_id_fkey ( nombre_completo )`).single();
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

  confirmDeleteAll() { this.showDeleteAllConfirm = true; }
  cancelDeleteAll() { this.showDeleteAllConfirm = false; }

  async executeDeleteAll() {
    this.isLoading = true;
    try {
      const rutasABorrar = this.rutas.filter(r => {
        const origen = r.tipo_negocio || (r.clientes?.tipo_cliente === 'B2B/INDUSTRIAS' ? 'INDUSTRIAS' : 'DISTRIBUIDORES');
        
        // 🔥 EL CANDADO SALVAVIDAS: Solo selecciona las de esta pestaña Y de la fecha que estás mirando
        const coincideModo = origen === this.modoVista;
        const coincideFecha = r.fecha_despacho === this.filtroFecha;
        
        return coincideModo && coincideFecha;
      });

      const ids = rutasABorrar.map(r => r.id);

      if (ids.length === 0) {
        this.toast.showWarning(`No hay rutas de ${this.modoVista} el ${this.filtroFecha} para eliminar.`);
        this.cancelDeleteAll();
        this.isLoading = false;
        return;
      }

      const chunkSize = 300;
      for (let i = 0; i < ids.length; i += chunkSize) {
        const chunk = ids.slice(i, i + chunkSize);
        const { error } = await this.supabase.despachos.delete().in('id', chunk);
        if (error) throw error;
      }
      
      this.rutas = this.rutas.filter(r => !ids.includes(r.id));
      this.paginaPendientes = 1;
      this.paginaHistorial = 1;
      
      this.toast.showSuccess(`Se eliminaron las rutas de ${this.modoVista} del día ${this.filtroFecha}.`);
      this.cancelDeleteAll();
    } catch (error: any) {
      this.toast.showError('Error al vaciar DB: ' + error.message);
    } finally {
      this.isLoading = false;
    }
  }

  openForm(registro?: any) {
    if (registro) {
      this.isEditing = true;
      this.currentEditId = registro.id;
      this.rutaForm.patchValue({
        fecha_despacho: registro.fecha_despacho, placa: registro.placa, cliente_id: registro.cliente_id,
        observacion: registro.observacion || '',
        p05: registro.p05, p10: registro.p10, p15: registro.p15, p45: registro.p45
      });
    } else {
      this.isEditing = false;
      this.currentEditId = null;
      this.rutaForm.reset({ p05: 0, p10: 0, p15: 0, p45: 0, observacion: '', fecha_despacho: new Date().toISOString().split('T')[0] });
    }
    this.showForm = true;
  }
  
  closeForm() { this.showForm = false; }

  async verDetalleRuta(ruta: any) {
    this.rutaDetalleActual = ruta;
    this.showDetalleModal = true;
    this.isLoading = true;
    this.historialEntregas = [];
    try {
      const { data, error } = await this.supabase.from('entregas_parciales')
        .select('*').eq('despacho_id', ruta.id).order('created_at', { ascending: false });
      if (error && error.code !== '42P01') throw error; 
      this.historialEntregas = data || [];
    } catch (error: any) {
      this.toast.showError('Error al cargar el historial.');
    } finally {
      this.isLoading = false;
    }
  }

  cerrarDetalleRuta() {
    this.showDetalleModal = false;
    this.historialEntregas = [];
    this.rutaDetalleActual = null;
  }

  verEvidencia(url: string) {
    this.fotoUrlActual = url;
    this.showFotoModal = true;
  }

  cerrarEvidencia() {
    this.showFotoModal = false;
    this.fotoUrlActual = '';
  }

  async eliminarEntregaParcial(entrega: any) {

    this.isLoading = true;
    try {
      const { error: errDelete } = await this.supabase.from('entregas_parciales').delete().eq('id', entrega.id);
      if (errDelete) throw errDelete;

      const { data: restantes, error: errFetch } = await this.supabase.from('entregas_parciales').select('*').eq('despacho_id', this.rutaDetalleActual.id);
      if (errFetch) throw errFetch;

      let tv05 = 0, tv10 = 0, tv15 = 0, tv45 = 0;
      let tm05 = 0, tm10 = 0, tm15 = 0, tm45 = 0;

      (restantes || []).forEach(e => {
        tv05 += (e.v05 || 0); tv10 += (e.v10 || 0); tv15 += (e.v15 || 0); tv45 += (e.v45 || 0);
        tm05 += (e.m05 || 0); tm10 += (e.m10 || 0); tm15 += (e.m15 || 0); tm45 += (e.m45 || 0);
      });

      const totalEntregado = tv05 + tv10 + tv15 + tv45;
      const totalProgramado = (this.rutaDetalleActual.p05 || 0) + (this.rutaDetalleActual.p10 || 0) + (this.rutaDetalleActual.p15 || 0) + (this.rutaDetalleActual.p45 || 0);
      
      let nuevoEstado = 'PENDIENTE';
      if (totalEntregado > 0) {
        nuevoEstado = totalEntregado >= totalProgramado ? 'COMPLETADO' : 'EN_PROCESO';
      }

      const payloadUpdate = {
        v05: tv05, v10: tv10, v15: tv15, v45: tv45,
        m05: tm05, m10: tm10, m15: tm15, m45: tm45,
        estado: nuevoEstado
      };

      const { data: updatedRuta, error: errUpdate } = await this.supabase.despachos.update(payloadUpdate)
        .eq('id', this.rutaDetalleActual.id)
        .select(`*, clientes ( razon_social, tipo_cliente ), perfiles!despachos_operador_id_fkey ( nombre_completo )`).single();
      if (errUpdate) throw errUpdate;

      this.historialEntregas = restantes || [];
      this.rutaDetalleActual = updatedRuta;
      
      const idx = this.rutas.findIndex(r => r.id === this.rutaDetalleActual.id);
      if (idx !== -1) this.rutas[idx] = updatedRuta;

      this.toast.showSuccess('Reporte anulado y ruta recalculada.');
    } catch (error: any) {
      this.toast.showError('Error al anular entrega: ' + error.message);
    } finally {
      this.isLoading = false;
    }
  }

  iniciarEdicionEntrega(entrega: any) {
    this.editingEntregaId = entrega.id;
    this.entregaEditForm = {
      v05: entrega.v05 || 0, v10: entrega.v10 || 0, v15: entrega.v15 || 0, v45: entrega.v45 || 0,
      m05: entrega.m05 || 0, m10: entrega.m10 || 0, m15: entrega.m15 || 0, m45: entrega.m45 || 0,
      observaciones: entrega.observaciones || ''
    };
  }

  cancelarEdicionEntrega() {
    this.editingEntregaId = null;
  }

  async guardarEdicionEntrega(entregaId: string) {
    this.isLoading = true;
    try {
      const { error: errUpdate } = await this.supabase.from('entregas_parciales').update({ ...this.entregaEditForm }).eq('id', entregaId);
      if (errUpdate) throw errUpdate;

      const { data: restantes, error: errFetch } = await this.supabase.from('entregas_parciales').select('*').eq('despacho_id', this.rutaDetalleActual.id);
      if (errFetch) throw errFetch;

      let tv05 = 0, tv10 = 0, tv15 = 0, tv45 = 0;
      let tm05 = 0, tm10 = 0, tm15 = 0, tm45 = 0;

      (restantes || []).forEach(e => {
        tv05 += (e.v05 || 0); tv10 += (e.v10 || 0); tv15 += (e.v15 || 0); tv45 += (e.v45 || 0);
        tm05 += (e.m05 || 0); tm10 += (e.m10 || 0); tm15 += (e.m15 || 0); tm45 += (e.m45 || 0);
      });

      const totalEntregado = tv05 + tv10 + tv15 + tv45;
      const totalProgramado = (this.rutaDetalleActual.p05 || 0) + (this.rutaDetalleActual.p10 || 0) + (this.rutaDetalleActual.p15 || 0) + (this.rutaDetalleActual.p45 || 0);
      
      let nuevoEstado = 'PENDIENTE';
      if (totalEntregado > 0) {
        nuevoEstado = totalEntregado >= totalProgramado ? 'COMPLETADO' : 'EN_PROCESO';
      }

      const payloadUpdate = {
        v05: tv05, v10: tv10, v15: tv15, v45: tv45,
        m05: tm05, m10: tm10, m15: tm15, m45: tm45,
        estado: nuevoEstado
      };

      const { data: updatedRuta, error: errUpdateDespacho } = await this.supabase.despachos.update(payloadUpdate)
        .eq('id', this.rutaDetalleActual.id)
        .select(`*, clientes ( razon_social, tipo_cliente ), perfiles!despachos_operador_id_fkey ( nombre_completo )`).single();
      if (errUpdateDespacho) throw errUpdateDespacho;

      this.historialEntregas = restantes || [];
      this.rutaDetalleActual = updatedRuta;
      
      const idx = this.rutas.findIndex(r => r.id === this.rutaDetalleActual.id);
      if (idx !== -1) this.rutas[idx] = updatedRuta;

      this.editingEntregaId = null;
      this.toast.showSuccess('Reporte editado y ruta recalculada con éxito.');
    } catch (error: any) {
      this.toast.showError('Error al editar entrega: ' + error.message);
    } finally {
      this.isLoading = false;
    }
  }
}