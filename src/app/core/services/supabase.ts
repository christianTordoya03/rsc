import { Injectable } from '@angular/core';
import { createClient, SupabaseClient } from '@supabase/supabase-js';

@Injectable({
  providedIn: 'root'
})
export class SupabaseService {
  private supabase: SupabaseClient;

  constructor() {
    this.supabase = createClient(
      'https://gnjamtgovdbruntkvpuc.supabase.co', 
      'sb_publishable_KkZLdS_gODu0pnV26nEQ8g_wfHaY27L'
    );
  }

  // Accesos directos a tablas recurrentes
  get clientes() { return this.supabase.from('clientes'); }
  get perfiles() { return this.supabase.from('perfiles'); }
  get despachos() { return this.supabase.from('despachos'); }
  
  // Accesos a los módulos de Supabase
  get auth() { return this.supabase.auth; }
  get storage() { return this.supabase.storage; } // <- Soluciona el error TS2339 de storage

  // Método genérico para consultar cualquier tabla dinámica
  from(tableName: string) {
    return this.supabase.from(tableName); // <- Soluciona el error TS2339 de .from()
  }
}