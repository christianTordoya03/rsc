import { Component, OnInit, ElementRef, ViewChild } from '@angular/core';
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
export class ProgramacionRutasComponent implements OnInit {
  @ViewChild('fileInput') fileInput!: ElementRef;

  showForm = false;
  isLoading = false;
  isEditing = false;
  currentEditId: string | null = null;
  showDeleteAllConfirm = false;
  
  rutas: any[] = [];
  clientesActivos: any[] = [];
  clientesNoEncontrados: string[] = [];
  
  placas = [
    'BBU-840', 'ACK-826', 'BHH-935', 'BBT-751', 'F4D-855', 
    'AYJ-784', 'BAN-786', 'B0L-756', 'F4B-804', 'BEL-713', 'F6Y-818'
  ];

  rutaForm: FormGroup;
  filtroCliente = '';
  filtroPlaca = '';

  paginaPendientes = 1;
  paginaHistorial = 1;
  itemsPorPagina = 50;

  showFotoModal = false;
  fotoUrlActual = '';
  showDetalleModal = false;
  rutaDetalleActual: any = null;
  historialEntregas: any[] = [];

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
    const erroresGuardados = localStorage.getItem('rutasErroresClientes');
    if (erroresGuardados) {
      this.clientesNoEncontrados = JSON.parse(erroresGuardados);
    }

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
  }

  cerrarPanelErrores() {
    this.clientesNoEncontrados = [];
    localStorage.removeItem('rutasErroresClientes');
  }

  async cargarDatosBase() {
    try {
      let todosLosClientes: any[] = [];
      let hayMasDatos = true;
      let desde = 0;
      const limite = 1000;

      while (hayMasDatos) {
        const { data, error } = await this.supabase.clientes
          .select('id, razon_social, ruc, codigo_cliente')
          .eq('estado', true)
          .range(desde, desde + limite - 1);

        if (error) throw error;
        if (data && data.length > 0) {
          todosLosClientes = [...todosLosClientes, ...data];
          desde += limite;
          if (data.length < limite) hayMasDatos = false;
        } else {
          hayMasDatos = false;
        }
      }
      this.clientesActivos = todosLosClientes;
    } catch (error) {
      this.toast.showError('Error al cargar clientes');
    }
  }

  async cargarRutasDelDia() {
    this.isLoading = true;
    try {
      const { data, error } = await this.supabase.despachos
        .select(`*, clientes ( razon_social ), perfiles!despachos_operador_id_fkey ( nombre_completo )`)
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
  }

  cumpleFiltros(r: any): boolean {
    const matchPlaca = !this.filtroPlaca || r.placa.toLowerCase().includes(this.filtroPlaca.toLowerCase());
    const nombreCliente = r.clientes?.razon_social || '';
    const matchCliente = !this.filtroCliente || nombreCliente.toLowerCase().includes(this.filtroCliente.toLowerCase());
    return matchPlaca && matchCliente;
  }

  get rutasPendientes() { return this.rutas.filter(r => (r.estado === 'PENDIENTE' || r.estado === 'EN_RUTA' || r.estado === 'EN_PROCESO') && this.cumpleFiltros(r)); }
  get rutasHistorial() { return this.rutas.filter(r => (r.estado === 'COMPLETADO' || r.estado === 'ANULADO') && this.cumpleFiltros(r)); }

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

  getNombreRealEnFila(row: any, colInicio: number): string {
    const ignorar = ['CAMBIO', 'BLOQUEADO', 'RECOJO', 'ENTREGA', 'PEDIDO', 'LIMA', 'CHICO', 'SUR', 'NORTE', 'ANULADO'];
    for (let i = colInicio; i <= colInicio + 5; i++) {
      const val = this.getCellValueText(row.getCell(i));
      if (val && val.length > 4 && /[A-Z]{3,}/.test(val)) {
        const esInvalido = ignorar.some(ig => val.includes(ig));
        if (!esInvalido) return val; 
      }
    }
    return '';
  }

  async importarExcel(event: any) {
    const file = event.target.files[0];
    if (!file) return;

    this.isLoading = true;
    this.toast.showInfo('Analizando Excel Logístico...');

    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await file.arrayBuffer());
      
      const worksheet = workbook.getWorksheet('Programacion') || workbook.worksheets[0];
      const sheetInicial = workbook.getWorksheet('PROGRAMACION INICIAL'); 

      const normalize = (str: string) => str ? str.trim().toUpperCase().replace(/\s+/g, ' ') : '';
      const stripPunct = (str: string) => str ? str.replace(/[^A-Z0-9]/g, '') : ''; 
      
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
            else if (/^(10|20|30)\d{5}$/.test(valClean)) codSap = valClean;
            else if (val.length > 4 && /[A-Z]/.test(val) && !['LIMA LOCAL', 'SUR CHICO', 'NORTE CHICO', 'PEDIDO', 'DISTRIBUIDOR', 'CLIENTE', 'DIRECCION'].includes(val)) {
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
        if (t === 'DISTRIBUIDOR' || t === 'CLIENTE' || t === 'DIRECCION' || t === 'NOMBRE') return null;

        const matchNum = t.match(/\b(10|20|30)\d{4,6}\b/);
        if (matchNum && mapCodigos.has(matchNum[0])) return mapCodigos.get(matchNum[0]);

        if (mapCodigos.has(t)) return mapCodigos.get(t); 
        if (mapRucs.has(t)) return mapRucs.get(t);
        if (mapNombres.has(t)) return mapNombres.get(t);

        const tClean = stripPunct(t);
        if (mapCleanNombres.has(tClean)) return mapCleanNombres.get(tClean);

        const matchPrefijo = t.match(/^(\d+)\s*-/);
        const prefijo = matchPrefijo ? matchPrefijo[1] : null;

        for (const [nameDB, id] of mapNombres.entries()) {
          if (prefijo && nameDB.startsWith(prefijo + ' -')) return id;
          if (t.length > 5 && nameDB.length > 5 && (nameDB.includes(t) || t.includes(nameDB))) return id;
        }
        return null;
      };

      let colPlaca = 0, colTransportes = 0, colPedido = 0, colEntrega = 0, colDistribuidor = 0; 
      let colLimite = 999; 
      
      const cols05: number[] = [];
      const cols10: number[] = [];
      const cols15: number[] = [];
      const cols45: number[] = [];

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
            if ((val === 'DISTRIBUIDOR' || val === 'CLIENTE' || val === 'DESTINATARIO' || val === 'NOMBRE' || val === 'RAZON SOCIAL') && colDistribuidor === 0) colDistribuidor = colNum;

            if (colNum < colLimite) {
                const valLimpio = val.replace(/\s+/g, '');
                if (valLimpio.includes('5KG') && !valLimpio.includes('15') && !valLimpio.includes('45')) cols05.push(colNum);
                if (valLimpio.includes('10KG')) cols10.push(colNum);
                if (val.includes('15 KG') || val.includes('M-15') || val.includes('M15') || valLimpio.includes('15KG') || valLimpio.includes('M15')) cols15.push(colNum);
                if (valLimpio.includes('45KG')) cols45.push(colNum);
            }
          });
        }
      });

      if (colPlaca === 0) colPlaca = 1;
      if (colPedido === 0) colPedido = 3;
      if (colEntrega === 0) colEntrega = 4;
      if (colDistribuidor === 0) colDistribuidor = 5; 

      let fechaProg = new Date().toISOString().split('T')[0];
      const celdaFechaA = worksheet.getCell('A2').value;
      if (celdaFechaA instanceof Date) fechaProg = celdaFechaA.toISOString().split('T')[0];

      const nuevasRutas: any[] = [];
      const noEncontrados = new Set<string>();
      
      let placaEnMemoria = '';
      let obsEnMemoria: string | null = null; 

      const palabrasIgnoradas = ['CAMBIO', 'BLOQUEADO', 'BLOQUEAD', 'RECOJO', 'VACIO', 'ANULADO'];

      worksheet.eachRow((row: any, rowNumber: number) => {
        if (rowNumber < 5) return; 

        let esSubTotal = false;
        for (let i = 1; i <= 10; i++) {
          const valC = this.getCellValueText(row.getCell(i));
          if (valC.includes('SUB TOTAL') || valC.includes('SUBTOTAL') || valC === 'TOTAL') {
             esSubTotal = true;
             break;
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

        let esRecojo = false;
        let cartelLogistico: string | null = null;

        if (txtEntrega.includes('RECOJO') || txtPedido.includes('RECOJO')) {
          esRecojo = true;
          cartelLogistico = '📦 RECOJO';
        } else if (txtEntrega.includes('BLOQUEAD') || txtPedido.includes('BLOQUEAD') || txtEntrega.includes('CAMBIO') || txtPedido.includes('CAMBIO')) {
          esRecojo = true;
          cartelLogistico = '🚫 ' + (txtEntrega.includes('CAMBIO') ? 'CAMBIO' : 'BLOQUEADO');
        }

        if (sumaCilindros === 0 && !esRecojo) {
          return;
        }

        const combiTexto = txtColPlaca + ' ' + txtColTrans;
        if (combiTexto.includes('2 VUELT') || combiTexto.includes('2DA VUELT')) {
          cartelLogistico = cartelLogistico ? cartelLogistico + ' | 2da Vuelta' : '♻️ 2da Vuelta';
        }

        const placaA = this.encontrarPlacaOficial(txtColPlaca);
        const placaB = this.encontrarPlacaOficial(txtColTrans);

        if (placaA) {
          placaEnMemoria = placaA;
          obsEnMemoria = cartelLogistico; 
        } else if (placaB) {
          placaEnMemoria = placaB;
          obsEnMemoria = cartelLogistico;
        }

        let placaFinal = placaEnMemoria || 'SIN PLACA'; 
        let obsFinal = cartelLogistico || obsEnMemoria;

        const placaSecuestradora = this.encontrarPlacaOficial(txtEntrega);
        if (placaSecuestradora && !esRecojo) {
          placaFinal = placaSecuestradora; 
          obsFinal = '🚚 Reasignado'; 
        }

        let textoBusqueda = this.getCellValueText(row.getCell(colDistribuidor));
        const textoLimpio = textoBusqueda.toUpperCase().trim();
        
        if (palabrasIgnoradas.some(palabra => textoLimpio === palabra || textoLimpio.includes(palabra))) {
          return; 
        }

        if ((!textoBusqueda || textoBusqueda === '0') && txtPedido) {
          textoBusqueda = mapPedidoADistribuidor.get(txtPedido) || '';
        }

        if (!textoBusqueda || textoBusqueda === '0') {
           noEncontrados.add(`Fila ${rowNumber}`);
           return; 
        }

        let busquedaAvanzada = textoBusqueda;
        const matchN = textoBusqueda.match(/\b(10|20|30)\d{4,6}\b/);
        if (matchN) busquedaAvanzada = matchN[0];

        const clienteId = findClienteId(busquedaAvanzada);
        
        // ====== LÓGICA ULTRA LIMPIA: SOLO LA FILA O FILA + NOMBRE REAL ======
        if (!clienteId) {
          const valorExactoDistribuidor = this.getCellValueText(row.getCell(colDistribuidor));
          
          // Verificamos si el valor capturado tiene al menos una letra (no es un código SAP puro)
          if (valorExactoDistribuidor && /[A-Za-z]/.test(valorExactoDistribuidor) && valorExactoDistribuidor !== '0') {
             noEncontrados.add(`Fila ${rowNumber}: ${valorExactoDistribuidor}`);
          } else {
             // Si está vacío, es '0', o es puro número (3031507), mostramos SOLO la fila
             noEncontrados.add(`Fila ${rowNumber}`);
          }
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
            programador_nombre: this.userName
          });
        }
      });

      if (noEncontrados.size > 0) {
        this.clientesNoEncontrados = Array.from(noEncontrados);
        localStorage.setItem('rutasErroresClientes', JSON.stringify(this.clientesNoEncontrados));
      } else {
        this.cerrarPanelErrores(); 
      }

      if (nuevasRutas.length > 0) {
        const { data, error } = await this.supabase.despachos.insert(nuevasRutas).select(`*, clientes ( razon_social ), perfiles!despachos_operador_id_fkey ( nombre_completo )`);
        if (error) throw error;
        this.rutas = [...(data || []), ...this.rutas];
        this.paginaPendientes = 1;
        this.toast.showSuccess(`Se procesaron ${nuevasRutas.length} rutas correctamente. Faltan ${noEncontrados.size} clientes.`);
      } else if (noEncontrados.size === 0) {
        this.toast.showWarning('No se generó ninguna ruta válida. Revisa el Excel.');
      } else {
        this.toast.showWarning(`Revisión completada. Faltan ${noEncontrados.size} clientes, 0 rutas nuevas asignadas.`);
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
          programador_nombre: this.userName
        };

        const { data, error } = await this.supabase.despachos.insert(payload).select(`*, clientes ( razon_social ), perfiles!despachos_operador_id_fkey ( nombre_completo )`).single();
        if (error) throw error;
        this.rutas.unshift(data);
        this.toast.showSuccess('Carga asignada');
      } else {
        const payloadUpdate = {
          fecha_despacho: v.fecha_despacho, placa: v.placa, cliente_id: v.cliente_id,
          observacion: v.observacion || null,
          p05: v.p05, p10: v.p10, p15: v.p15, p45: v.p45,
          editado_en: new Date().toISOString(),
          editor_nombre: this.userName
        };

        const { data, error } = await this.supabase.despachos.update(payloadUpdate).eq('id', this.currentEditId).select(`*, clientes ( razon_social ), perfiles!despachos_operador_id_fkey ( nombre_completo )`).single();
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
    if (!confirm('¿Está seguro de anular esta ruta?')) return;
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
      const { error } = await this.supabase.despachos.delete().not('id', 'is', null);
      if (error) throw error;
      this.rutas = [];
      this.paginaPendientes = 1;
      this.paginaHistorial = 1;
      this.toast.showSuccess('Todas las rutas fueron eliminadas de la base de datos.');
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
}