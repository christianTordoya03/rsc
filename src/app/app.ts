import { Component, signal, OnInit, inject } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { SwUpdate } from '@angular/service-worker';

@Component({
  selector: 'app-root',
  imports: [RouterOutlet],
  templateUrl: './app.html',
  styleUrl: './app.scss'
})
export class App implements OnInit {
  protected readonly title = signal('rsc-erp');
  
  // Inyectamos el motor de actualizaciones
  private swUpdate = inject(SwUpdate);

  ngOnInit() {
    // 1. Verificamos si el Service Worker está activo (solo corre en producción/Vercel)
    if (this.swUpdate.isEnabled) {
      // 2. Obliga al celular a revisar si subiste nuevo código a Vercel
      this.swUpdate.checkForUpdate();

      // 3. Si el radar detecta que la nueva versión ya se descargó...
      this.swUpdate.versionUpdates.subscribe((event: any) => {
        if (event.type === 'VERSION_READY') {
          // ...Refresca la pantalla automáticamente al instante
          window.location.reload();
        }
      });
    }
  }
}