import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormBuilder, FormGroup, ReactiveFormsModule, Validators } from '@angular/forms';
import { Router } from '@angular/router';
import { SupabaseService } from '../../core/services/supabase';
import { ToastService } from '../../core/services/toast';

@Component({
  selector: 'app-operador',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule],
  templateUrl: './operador.html'
})
export class OperadorComponent implements OnInit, OnDestroy {
  // Estados de Modales
  showModalCheckin = false; 
  showModalCheckout = false;
  showDespachoModal = false;
  showModalManual = false;
  isLoading = false;

  // Jornada
  jornadaActiva = false;
  datosJornada: any = null;
  userId: string = '';

  rutasAsignadas: any[] = [];
  rutaActual: any = null;
  archivoEvidencia: File | null = null;

  placas = ['BBU-840', 'ACK-826', 'BHH-935', 'BBT-751', 'F4D-855', 'AYJ-784', 'BAN-786', 'B0L-756', 'F4B-804', 'BEL-713', 'F6Y-818', 'BSY-931', 'AUH-927', 'F6Z-867', 'BHI-940', 'B7M-872', 'BTB-817'];

  checkinForm: FormGroup;
  checkoutForm: FormGroup;
  despachoForm: FormGroup;

  // Motor de actualización en tiempo real
  private autoRefreshSub: any;

  constructor(
    private fb: FormBuilder,
    private supabase: SupabaseService,
    private toast: ToastService,
    private router: Router
  ) {
    this.checkinForm = this.fb.group({
      placa: ['', Validators.required],
      kilometraje: ['', [Validators.required, Validators.min(0)]]
    });

    this.checkoutForm = this.fb.group({
      kmFinal: ['', [Validators.required, Validators.min(0)]],
      galones: [0, Validators.min(0)],
      montoCombustible: [0, Validators.min(0)],
      peajes: [0, Validators.min(0)]
    });

    // LIMPIO Y ESTANDARIZADO
    this.despachoForm = this.fb.group({
      v05: [0], v10: [0], v15: [0], v45: [0], 
      m05: [0], m10: [0], m15: [0], m45: [0], 
      observaciones: [''],
      cambio_programacion: [0],
      tipo_movimiento_cilindro: ['NINGUNO'],
      cantidad_cilindros: [0]
    });
  }

  async ngOnInit() {
    const { data: { session } } = await this.supabase.auth.getSession();
    if (session) {
      this.userId = session.user.id;
      this.verificarJornada();
      this.iniciarSincronizacionSilenciosa();
    } else {
      this.router.navigate(['/login']);
    }
  }

  ngOnDestroy() {
    if (this.autoRefreshSub) {
      clearInterval(this.autoRefreshSub);
    }
  }

  // ==== SINCRONIZADOR EN SEGUNDO PLANO ====
  iniciarSincronizacionSilenciosa() {
    this.autoRefreshSub = setInterval(async () => {
      // Solo busca actualizaciones si el chofer está en turno y no está guardando algo
      if (!this.isLoading && this.jornadaActiva) {
        try {
          await this.cargarRutasReales();
        } catch (err) {
           // Fallo silencioso (ignora errores de red pasajeros)
        }
      }
    }, 10000); // 10 segundos
  }

  // ==== ANTI-PARPADEO ====
  trackById(index: number, item: any): string {
    return item.id;
  }

  async cerrarSesion() {
    await this.supabase.auth.signOut();
    this.router.navigate(['/login']);
  }

  // --- CONTROL DE JORNADA ---
  verificarJornada() {
    const jornadaStorage = localStorage.getItem('jornada_rsc');
    if (jornadaStorage) {
      this.datosJornada = JSON.parse(jornadaStorage);
      this.jornadaActiva = true;
      this.cargarRutasReales();
    } else {
      this.jornadaActiva = false;
      this.showModalCheckin = true;
    }
  }

  // --- PASO 1: EL CHOFER ASUME LA UNIDAD (CHECK-IN) ---
  async iniciarJornada() {
    if (this.checkinForm.invalid) return this.checkinForm.markAllAsTouched();
    
    this.isLoading = true;
    const placaSeleccionada = this.checkinForm.value.placa;
    const kmInicial = this.checkinForm.value.kilometraje;
    const hoy = new Date().toISOString().split('T')[0];

    try {
      const { error: updateError } = await this.supabase.despachos
        .update({ operador_id: this.userId })
        .eq('placa', placaSeleccionada)
        .eq('fecha_despacho', hoy)
        .in('estado', ['PENDIENTE', 'EN_PROCESO'])
        .is('operador_id', null);

      if (updateError) throw updateError;

      this.datosJornada = { placa: placaSeleccionada, kmInicial: kmInicial, fecha: hoy };
      localStorage.setItem('jornada_rsc', JSON.stringify(this.datosJornada));
      
      this.jornadaActiva = true;
      this.showModalCheckin = false;
      
      await this.cargarRutasReales();

      this.toast.showSuccess(`Unidad ${placaSeleccionada} asignada exitosamente`);
    } catch(error: any) {
      this.toast.showError('Error al asignar unidad: ' + error.message);
    } finally {
      this.isLoading = false;
    }
  }

  // --- PASO 3: EL CHOFER CIERRA LA UNIDAD (CHECK-OUT) ---
  async finalizarJornada() {
    if (this.checkoutForm.invalid) return this.toast.showError('Complete el odómetro final');
    if (this.checkoutForm.value.kmFinal <= this.datosJornada.kmInicial) {
      return this.toast.showError('El Km Final debe ser mayor al inicial');
    }
    
    this.isLoading = true;
    try {
      // INYECTAMOS EL USUARIO PARA LA CONEXIÓN PERFECTA
      const payload = {
        fecha: this.datosJornada.fecha,
        placa: this.datosJornada.placa,
        km_inicial: this.datosJornada.kmInicial,
        km_final: this.checkoutForm.value.kmFinal,
        galones: this.checkoutForm.value.galones,
        costo_total: this.checkoutForm.value.montoCombustible,
        gasto_peaje: this.checkoutForm.value.peajes,
        registrado_por: this.userId 
      };
      
      const { error } = await this.supabase.from('control_combustible').insert(payload);
      if (error) console.error("Error guardando combustible:", error); 

      localStorage.removeItem('jornada_rsc');
      this.jornadaActiva = false;
      this.showModalCheckout = false;
      this.showModalCheckin = true;
      this.checkoutForm.reset();
      this.checkinForm.reset();
      this.rutasAsignadas = [];
      this.toast.showSuccess('Jornada finalizada correctamente');
    } catch (error) {
      this.toast.showError('Error al cerrar jornada');
    } finally {
      this.isLoading = false;
    }
  }

  async cargarRutasReales() {
    if (!this.datosJornada?.placa) return;

    const hoyReal = new Date().toISOString().split('T')[0];

    if (this.datosJornada.fecha !== hoyReal) {
      this.rutasAsignadas = []; // Vaciamos la pantalla a la fuerza
      this.rutaActual = null;
      this.showDespachoModal = false;
      return; // Cortamos la búsqueda, así Logística no haya anulado las rutas, el chofer no las verá.
    }

    const { data, error } = await this.supabase.despachos
      .select(`*, clientes(razon_social, direccion), entregas_parciales(*)`)
      .eq('placa', this.datosJornada.placa)
      .eq('fecha_despacho', this.datosJornada.fecha) 
      .neq('estado', 'COMPLETADO') 
      .neq('estado', 'ANULADO')
      .order('creado_en', { ascending: true });

    if (error) throw error;
    
    // 1. Aplicamos el ordenador
    this.rutasAsignadas = this.ordenarRutasPrioritarias(data || []);

    // 2. MAGIA DE SINCRONIZACIÓN: Actualizamos el modal del operador en vivo
    if (this.rutaActual) {
      const rutaActualizada = this.rutasAsignadas.find(r => r.id === this.rutaActual.id);
      if (rutaActualizada) {
        this.rutaActual = rutaActualizada; // Refresca los números si logística editó
      } else {
        // Si logística anuló o completó la ruta, se le cierra el modal automáticamente
        this.showDespachoModal = false; 
        this.rutaActual = null;
      }
    }
  }

  // --- SEMÁFORO LÓGICA (Corregido a las variables puras p15 y v15) ---
  getEstadoSemaforo(ruta: any) {
    const totalProg = (ruta.p05||0) + (ruta.p10||0) + (ruta.p15||0) + (ruta.p45||0);
    const totalVendido = (ruta.v05||0) + (ruta.v10||0) + (ruta.v15||0) + (ruta.v45||0) + (ruta.m05||0) + (ruta.m10||0) + (ruta.m15||0) + (ruta.m45||0);

    if (totalVendido === 0) return { color: 'bg-red-500', texto: 'Pendiente', pct: 0 };
    if (totalVendido > 0 && totalVendido < totalProg) return { color: 'bg-yellow-400', texto: 'En Proceso', pct: Math.round((totalVendido/totalProg)*100) };
    return { color: 'bg-green-500', texto: 'Completado', pct: 100 };
  }

  // --- PASO 2: DESPACHO (ENTREGA PARCIAL O TOTAL) ---
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
    // if (this.tieneMerma && !this.archivoEvidencia) return this.toast.showWarning('Adjunte foto de merma obligatoria');

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

      // GUARDAR EL VIAJE INDIVIDUAL
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

      // ACUMULADOR MAESTRO
      const newV05 = (r.v05 || 0) + v.v05;
      const newV10 = (r.v10 || 0) + v.v10;
      const newV15 = (r.v15 || 0) + v.v15;
      const newV45 = (r.v45 || 0) + v.v45;

      const newM05 = (r.m05 || 0) + v.m05;
      const newM10 = (r.m10 || 0) + v.m10;
      const newM15 = (r.m15 || 0) + v.m15;
      const newM45 = (r.m45 || 0) + v.m45;

      const totalVendidoHistorico = newV05 + newV10 + newV15 + newV45;
      const totalMermasHistorico = newM05 + newM10 + newM15 + newM45;
      
      const totalProcesado = totalVendidoHistorico + totalMermasHistorico;
      const totalProg = (r.p05||0) + (r.p10||0) + (r.p15||0) + (r.p45||0);
      
      const nuevoEstado = totalProcesado >= totalProg ? 'COMPLETADO' : 'EN_PROCESO';

      const payloadDespacho = {
        venta_real: totalVendidoHistorico,
        merma: totalMermasHistorico,
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
      // 1. Si la ruta ya está 100% COMPLETADA, la mandamos al fondo del todo
      if (a.estado === 'COMPLETADO' && b.estado !== 'COMPLETADO') return 1;
      if (b.estado === 'COMPLETADO' && a.estado !== 'COMPLETADO') return -1;

      // 2. Calculamos si la ruta ya tiene algún avance (cilindros o mermas)
      const avanceA = (a.v05||0) + (a.v10||0) + (a.v15||0) + (a.v45||0) + (a.m05||0) + (a.m10||0) + (a.m15||0) + (a.m45||0);
      const avanceB = (b.v05||0) + (b.v10||0) + (b.v15||0) + (b.v45||0) + (b.m05||0) + (b.m10||0) + (b.m15||0) + (b.m45||0);

      // 3. Regla de Oro: Los que tienen avance > 0 suben a la cima automáticamente
      if (avanceA > 0 && avanceB === 0) return -1;
      if (avanceB > 0 && avanceA === 0) return 1;

      // 4. Si están empatados (ambos en 0 o ambos empezados), ordenamos alfabéticamente por cliente
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
      
      // Actualizamos la entrega individual con la nueva foto
      const { error: updateError } = await this.supabase.from('entregas_parciales')
        .update({ foto_merma_url: publicUrlData.publicUrl })
        .eq('id', entregaId);
        
      if (updateError) throw updateError;
      
      this.toast.showSuccess('Evidencia subida correctamente');
      await this.cargarRutasReales(); // Recargamos para actualizar la vista
    } catch (error: any) {
      this.toast.showError('Error al subir imagen: ' + error.message);
    } finally {
      this.isLoading = false;
      event.target.value = ''; // Limpiamos el input
    }
  }
}