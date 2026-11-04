import { Component, OnInit, OnDestroy, HostListener } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormBuilder, FormGroup, ReactiveFormsModule, Validators, FormsModule } from '@angular/forms';
import { Router } from '@angular/router';
import { SupabaseService } from '../../core/services/supabase';
import { ToastService } from '../../core/services/toast';

@Component({
  selector: 'app-operador',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule, FormsModule],
  templateUrl: './operador.html'
})
export class OperadorComponent implements OnInit, OnDestroy {
  showModalCheckin = false; 
  showDespachoModal = false;
  showModalManual = false;
  showModalSalir = false;
  isLoading = false;

  jornadaActiva = false;
  datosJornada: any = null;
  userId: string = '';

  rutasAsignadas: any[] = [];
  rutaActual: any = null;
  archivoEvidencia: File | null = null;
  filtroModo: 'TODOS' | 'ACTIVOS' = 'TODOS';
  
  vistaActiva: 'DESPACHOS' | 'COMBUSTIBLE' = 'DESPACHOS';

  placas = ['BBU-840', 'ACK-826', 'BHH-935', 'BBT-751', 'F4D-855', 'AYJ-784', 'BAN-786', 'B0L-756', 'F4B-804', 'BEL-713', 'F6Y-818', 'BSY-931', 'AUH-927', 'F6Z-867', 'BHI-940', 'B7M-872', 'BTB-817', 'AUH-916', 'F4D-707', 'CAE-701'];

  checkinForm: FormGroup;
  despachoForm: FormGroup;
  gastosForm: FormGroup;
  private autoRefreshSub: any;

  private obtenerFechaOperativa() {
    const ahora = new Date();
    ahora.setMinutes(ahora.getMinutes() - ahora.getTimezoneOffset());
    const horas = ahora.getUTCHours();
    const minutos = ahora.getUTCMinutes();
    
    if (horas < 3 || (horas === 3 && minutos < 45)) {
       ahora.setUTCDate(ahora.getUTCDate() - 1);
    }
    return ahora.toISOString().split('T')[0];
  }

  constructor(
    private fb: FormBuilder,
    private supabase: SupabaseService,
    private toast: ToastService,
    private router: Router
  ) {
    this.checkinForm = this.fb.group({
      placa: ['', Validators.required]
    });

    this.despachoForm = this.fb.group({
      v05: [0], v10: [0], v15: [0], v45: [0], 
      m05: [0], m10: [0], m15: [0], m45: [0], 
      observaciones: [''],
      cambio_programacion: [0],
      tipo_movimiento_cilindro: ['NINGUNO'],
      cantidad_cilindros: [0]
    });

    this.gastosForm = this.fb.group({
      kmInicial: ['', Validators.required],
      kmFinal: ['', Validators.required],
      galones: [0, Validators.min(0)],
      precioGalon: [0, Validators.min(0)],
      peajes: [0, Validators.min(0)]
    });
  }

  async ngOnInit() {
    const { data: { session } } = await this.supabase.auth.getSession();
    if (session) {
      this.userId = session.user.id;
      this.verificarJornada();
      this.iniciarSincronizacionSilenciosa();
      this.iniciarAutocalculoRetornos();
      history.pushState(null, '', window.location.href);
    } else {
      this.router.navigate(['/login']);
    }
  }

  @HostListener('window:popstate', ['$event'])
  onPopState(event: any) {
    history.pushState(null, '', window.location.href); 
    this.showModalSalir = true; 
  }

  ngOnDestroy() {
    if (this.autoRefreshSub) clearInterval(this.autoRefreshSub);
  }

  iniciarSincronizacionSilenciosa() {
    this.autoRefreshSub = setInterval(async () => {
      if (!this.isLoading && this.jornadaActiva) {
        try { await this.cargarRutasReales(); } catch (err) {}
      }
    }, 10000); 
  }

  iniciarAutocalculoRetornos() {
    ['05', '10', '15', '45'].forEach(tipo => {
      this.despachoForm.get(`v${tipo}`)?.valueChanges.subscribe(val => {
        if (this.rutaActual) {
          const programado = this.rutaActual[`p${tipo}`] || 0;
          const vendido = val || 0;
          if (programado > 0 && vendido <= programado) {
            this.despachoForm.patchValue({ [`m${tipo}`]: programado - vendido }, { emitEvent: false });
          }
        }
      });
    });
  }

  trackById(index: number, item: any): string { return item.id; }

  async cerrarSesionCompleta() {
    localStorage.removeItem('jornada_rsc'); 
    this.showModalSalir = false;
    await this.supabase.auth.signOut();
    this.router.navigate(['/login']);
  }

  verificarJornada() {
    const jornadaStorage = localStorage.getItem('jornada_rsc');
    if (jornadaStorage) {
      this.datosJornada = JSON.parse(jornadaStorage);
      const hoyReal = this.obtenerFechaOperativa();
      
      if (this.datosJornada.fecha !== hoyReal) {
         localStorage.removeItem('jornada_rsc');
         this.jornadaActiva = false;
         this.showModalCheckin = true;
         this.rutasAsignadas = [];
      } else {
         this.jornadaActiva = true;
         this.cargarRutasReales();
      }
    } else {
      this.jornadaActiva = false;
      this.showModalCheckin = true;
    }
  }

  async iniciarJornada() {
    if (this.checkinForm.invalid) return this.checkinForm.markAllAsTouched();
    
    this.isLoading = true;
    const placaSeleccionada = this.checkinForm.value.placa;
    const hoy = this.obtenerFechaOperativa();

    try {
      const { error: updateError } = await this.supabase.despachos
        .update({ operador_id: this.userId })
        .eq('placa', placaSeleccionada)
        .eq('fecha_despacho', hoy)
        .in('estado', ['PENDIENTE', 'PROXIMO', 'EN_PROCESO'])
        .is('operador_id', null);

      if (updateError) throw updateError;

      this.datosJornada = { placa: placaSeleccionada, fecha: hoy };
      localStorage.setItem('jornada_rsc', JSON.stringify(this.datosJornada));
      
      this.jornadaActiva = true;
      this.showModalCheckin = false;
      await this.cargarRutasReales();
      this.toast.showSuccess(`Unidad ${placaSeleccionada} iniciada exitosamente`);
    } catch(error: any) {
      this.toast.showError('Error al iniciar: ' + error.message);
    } finally {
      this.isLoading = false;
    }
  }

  get rutasVisibles() {
    if (this.filtroModo === 'ACTIVOS') {
      return this.rutasAsignadas.filter(r => r.estado === 'PROXIMO' || r.estado === 'EN_PROCESO');
    }
    return this.rutasAsignadas;
  }

  async cargarRutasReales() {
    if (!this.datosJornada?.placa) return;

    const hoyReal = this.obtenerFechaOperativa();
    if (this.datosJornada.fecha !== hoyReal) {
      localStorage.removeItem('jornada_rsc');
      this.jornadaActiva = false;
      this.showModalCheckin = true;
      this.rutasAsignadas = [];
      this.rutaActual = null;
      this.showDespachoModal = false;
      return;
    }

    const { data, error } = await this.supabase.despachos
      .select(`*, clientes(razon_social, direccion), entregas_parciales(*)`)
      .eq('placa', this.datosJornada.placa)
      .eq('fecha_despacho', this.datosJornada.fecha) 
      .neq('estado', 'COMPLETADO') 
      .neq('estado', 'ANULADO')
      .order('creado_en', { ascending: true });

    if (error) throw error;
    this.rutasAsignadas = this.ordenarRutasPrioritarias(data || []);

    if (this.rutaActual) {
      const rutaActualizada = this.rutasAsignadas.find(r => r.id === this.rutaActual.id);
      if (rutaActualizada) {
        this.rutaActual = rutaActualizada;
      } else {
        this.showDespachoModal = false; 
        this.rutaActual = null;
      }
    }
  }

  async cambiarEstadoRapido(id: string, estadoNuevo: string) {
    try {
      const { error } = await this.supabase.despachos.update({ estado: estadoNuevo }).eq('id', id);
      if (error) throw error;
      
      const ruta = this.rutasAsignadas.find(r => r.id === id);
      if (ruta) ruta.estado = estadoNuevo;
      this.rutasAsignadas = this.ordenarRutasPrioritarias(this.rutasAsignadas);
    } catch(e) {
      this.toast.showError('Error al cambiar estado');
    }
  }

  abrirDespacho(ruta: any) {
    this.rutaActual = ruta;
    this.despachoForm.reset({
      v05: 0, v10: 0, v15: 0, v45: 0,
      m05: 0, m10: 0, m15: 0, m45: 0,
      cambio_programacion: 0, tipo_movimiento_cilindro: 'NINGUNO', cantidad_cilindros: 0
    });
    this.archivoEvidencia = null;
    this.showDespachoModal = true;
  }

  getFaltante(tipo: string): number {
    if (!this.rutaActual) return 0;
    const programado = this.rutaActual[`p${tipo}`] || 0;
    const vendido = this.rutaActual[`v${tipo}`] || 0;
    const merma = this.rutaActual[`m${tipo}`] || 0;
    const faltante = programado - vendido - merma;
    return faltante > 0 ? faltante : 0;
  }

  get tieneMerma() {
    const v = this.despachoForm.value;
    return (v.m05 + v.m10 + v.m15 + v.m45) > 0;
  }

  onFileSelected(event: any) {
    const file = event.target.files[0];
    if (file) {
      this.archivoEvidencia = file;
      this.toast.showInfo('Evidencia fotográfica adjuntada');
    }
  }

  async registrarDespacho() {
    if (this.despachoForm.invalid) return this.despachoForm.markAllAsTouched();
    this.isLoading = true;
    const v = this.despachoForm.value;
    const r = this.rutaActual;
    let urlFoto = null; 

    try {
      if (this.tieneMerma && this.archivoEvidencia) {
        const fileExt = this.archivoEvidencia.name.split('.').pop();
        const fileName = `${new Date().getTime()}_${Math.random().toString(36).substring(2)}.${fileExt}`;
        const { error: uploadError } = await this.supabase.storage.from('evidencias').upload(`mermas/${fileName}`, this.archivoEvidencia);
        if (uploadError) throw uploadError;
        const { data: publicUrlData } = this.supabase.storage.from('evidencias').getPublicUrl(`mermas/${fileName}`);
        urlFoto = publicUrlData.publicUrl;
      }

      const entregaIndividual = {
        despacho_id: r.id,
        v05: v.v05, v10: v.v10, v15: v.v15, v45: v.v45,
        m05: v.m05, m10: v.m10, m15: v.m15, m45: v.m45,
        foto_merma_url: urlFoto, 
        cambio_programacion: v.cambio_programacion,
        tipo_movimiento_cilindro: v.tipo_movimiento_cilindro,
        cantidad_cilindros: v.cantidad_cilindros,
        observaciones: v.observaciones
      };
      const { error: insertError } = await this.supabase.from('entregas_parciales').insert(entregaIndividual);
      if (insertError) throw insertError;

      const newV05 = (r.v05 || 0) + v.v05; const newV10 = (r.v10 || 0) + v.v10;
      const newV15 = (r.v15 || 0) + v.v15; const newV45 = (r.v45 || 0) + v.v45;
      const newM05 = (r.m05 || 0) + v.m05; const newM10 = (r.m10 || 0) + v.m10;
      const newM15 = (r.m15 || 0) + v.m15; const newM45 = (r.m45 || 0) + v.m45;

      const totalProcesado = newV05 + newV10 + newV15 + newV45 + newM05 + newM10 + newM15 + newM45;
      const totalProg = (r.p05||0) + (r.p10||0) + (r.p15||0) + (r.p45||0);
      const nuevoEstado = totalProcesado >= totalProg ? 'COMPLETADO' : 'EN_PROCESO';

      const payloadDespacho = {
        venta_real: newV05 + newV10 + newV15 + newV45,
        merma: newM05 + newM10 + newM15 + newM45,
        v05: newV05, v10: newV10, v15: newV15, v45: newV45,
        m05: newM05, m10: newM10, m15: newM15, m45: newM45,
        estado: nuevoEstado
      };

      const { error: updateError } = await this.supabase.despachos.update(payloadDespacho).eq('id', r.id);
      if (updateError) throw updateError;
      
      this.toast.showSuccess(`Viaje registrado. Estado: ${nuevoEstado}`);
      this.showDespachoModal = false;
      await this.cargarRutasReales(); 
    } catch (error: any) {
      this.toast.showError('Error al registrar: ' + error.message);
    } finally {
      this.isLoading = false;
    }
  }

  ordenarRutasPrioritarias(rutasArray: any[]) {
    return rutasArray.sort((a, b) => {
      if (a.estado === 'COMPLETADO' && b.estado !== 'COMPLETADO') return 1;
      if (b.estado === 'COMPLETADO' && a.estado !== 'COMPLETADO') return -1;
      
      const pesoEstado = (estado: string) => estado === 'EN_PROCESO' ? 3 : estado === 'PROXIMO' ? 2 : 1;
      if (pesoEstado(a.estado) > pesoEstado(b.estado)) return -1;
      if (pesoEstado(b.estado) > pesoEstado(a.estado)) return 1;

      const clienteA = a.clientes?.razon_social || '';
      const clienteB = b.clientes?.razon_social || '';
      return clienteA.localeCompare(clienteB);
    });
  }

  async subirEvidenciaTardia(event: any, entregaId: string, despachoId: string) {
    const file = event.target.files[0];
    if (!file) return;
    this.isLoading = true;
    this.toast.showInfo('Subiendo foto desde galería...');

    try {
      const fileExt = file.name.split('.').pop();
      const fileName = `${new Date().getTime()}_${Math.random().toString(36).substring(2)}.${fileExt}`;
      const { error: uploadError } = await this.supabase.storage.from('evidencias').upload(`mermas/${fileName}`, file);
      if (uploadError) throw uploadError;
      const { data: publicUrlData } = this.supabase.storage.from('evidencias').getPublicUrl(`mermas/${fileName}`);
      const { error: updateError } = await this.supabase.from('entregas_parciales').update({ foto_merma_url: publicUrlData.publicUrl }).eq('id', entregaId);
      if (updateError) throw updateError;
      
      this.toast.showSuccess('Evidencia subida correctamente');
      await this.cargarRutasReales(); 
    } catch (error: any) {
      this.toast.showError('Error al subir imagen: ' + error.message);
    } finally {
      this.isLoading = false;
      event.target.value = ''; 
    }
  }

  liberarPlaca() {
    localStorage.removeItem('jornada_rsc');
    this.jornadaActiva = false;
    this.rutasAsignadas = [];
    this.rutaActual = null;
    this.showDespachoModal = false;
    this.showModalSalir = false;
    this.showModalCheckin = true;
  }

  async guardarGastos() {
    if (this.gastosForm.invalid) return this.toast.showError('Complete el odómetro inicial y final');
    
    const v = this.gastosForm.value;
    if (v.kmFinal <= v.kmInicial) return this.toast.showError('El Km Final debe ser mayor al inicial');
    
    this.isLoading = true;
    try {
      const costoTotal = (v.galones || 0) * (v.precioGalon || 0);
  
      const payload = {
        fecha: this.datosJornada.fecha,
        placa: this.datosJornada.placa,
        km_inicial: v.kmInicial,
        km_final: v.kmFinal,
        galones: v.galones || 0,
        costo_total: costoTotal,
        gasto_peaje: v.peajes || 0,
        registrado_por: this.userId 
      };
      
      const { error } = await this.supabase.from('control_combustible').insert(payload);
      if (error) throw error; 
  
      this.toast.showSuccess('Reporte vehicular enviado correctamente.');
      this.gastosForm.reset(); 
      this.vistaActiva = 'DESPACHOS'; 
    } catch (error) {
      this.toast.showError('Error al guardar reporte en la nube.');
    } finally {
      this.isLoading = false;
    }
  }
}