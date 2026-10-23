import { Component, OnInit } from '@angular/core';
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
export class OperadorComponent implements OnInit {
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

  placas = ['BBU-840', 'ACK-826', 'BHH-935', 'BBT-751', 'F4D-855', 'AYJ-784', 'BAN-786', 'B0L-756', 'F4B-804', 'BEL-713', 'F6Y-818', 'BSY-931', 'AUH-927', 'F6Z-867', 'BHI-940'];

  checkinForm: FormGroup;
  checkoutForm: FormGroup;
  despachoForm: FormGroup;

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
    } else {
      this.router.navigate(['/login']);
    }
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
    const { data, error } = await this.supabase.despachos
      .select(`*, clientes(razon_social, direccion)`)
      .eq('placa', this.datosJornada.placa)
      .neq('estado', 'COMPLETADO') 
      .neq('estado', 'ANULADO')
      .order('creado_en', { ascending: true });

    if (error) throw error;
    this.rutasAsignadas = data || [];
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
    if (this.tieneMerma && !this.archivoEvidencia) return this.toast.showWarning('Adjunte foto de merma obligatoria');

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
}