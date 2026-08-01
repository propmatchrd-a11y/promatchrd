/**
 * ============================================================================
 *  FONDO PREMIUM — mesh gradient + aurora + noise + partículas interactivas
 *  + mouse parallax. Se incluye con una sola línea en cualquier página:
 *
 *      <script src="fondo-premium.js"></script>
 *
 *  No requiere React, build, ni npm — funciona con una etiqueta <script>
 *  normal. Prioriza rendimiento real en celulares económicos (comunes en
 *  RD) por encima del efecto visual máximo:
 *  - Respeta "reducir movimiento" del sistema operativo — si está
 *    activado, no se anima ni se cargan partículas (accesibilidad real,
 *    no solo un detalle técnico).
 *  - Las partículas (lo más costoso en CPU/batería) solo se cargan en
 *    pantallas anchas (escritorio) — en celular, el mesh + aurora + noise
 *    en CSS puro ya dan la sensación "premium" sin gastar batería extra
 *    justo en el dispositivo donde más se usa el sistema.
 *  - Todo vive detrás del contenido (z-index negativo) y con
 *    pointer-events:none — nunca interfiere con ningún clic o formulario.
 * ============================================================================
 */
(function () {
  const prefiereMovimientoReducido = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const esPantallaAncha = window.matchMedia('(min-width: 900px)').matches;

  // 1) Mesh gradient — 4 blobs de marca, ya animados por CSS (flotarN)
  const mesh = document.createElement('div');
  mesh.id = 'fondo-mesh';
  mesh.innerHTML = '<div class="blob"></div><div class="blob"></div><div class="blob"></div><div class="blob"></div>';
  document.body.prepend(mesh);

  // 2) Aurora — capa de degradado cónico girando, ya animada por CSS
  const aurora = document.createElement('div');
  aurora.id = 'fondo-aurora';
  document.body.prepend(aurora);

  // 3) Noise texture — textura sutil de grano, sensación "premium/analógica"
  const noise = document.createElement('div');
  noise.id = 'fondo-noise';
  document.body.appendChild(noise);

  if (prefiereMovimientoReducido) return; // nada de parallax ni partículas

  // 4) Mouse parallax sobre los blobs — sutil, con inercia suave (CSS
  //    transition ya definida en .blob), throttled con requestAnimationFrame
  //    para no sobrecargar el hilo principal.
  let mouseX = 0, mouseY = 0, rafPendiente = false;
  window.addEventListener('mousemove', function (e) {
    mouseX = (e.clientX / window.innerWidth - 0.5) * 2; // rango -1 a 1
    mouseY = (e.clientY / window.innerHeight - 0.5) * 2;
    if (rafPendiente) return;
    rafPendiente = true;
    requestAnimationFrame(function () {
      document.querySelectorAll('#fondo-mesh .blob').forEach(function (blob, i) {
        const intensidad = 12 + i * 5; // cada blob se mueve un poco distinto — sensación de profundidad
        blob.style.transform = 'translate(' + (mouseX * intensidad) + 'px, ' + (mouseY * intensidad) + 'px)';
      });
      rafPendiente = false;
    });
  }, { passive: true });

  // 5) Partículas interactivas (tsParticles vía CDN, sin instalar nada) —
  //    SOLO en pantalla ancha, por rendimiento. Conteo bajo y sin líneas de
  //    conexión pesadas: prioriza que se mantenga a 60 FPS siempre, por
  //    encima de la densidad visual.
  if (!esPantallaAncha) return;

  const contenedorParticulas = document.createElement('div');
  contenedorParticulas.id = 'fondo-particulas';
  document.body.appendChild(contenedorParticulas);

  const script = document.createElement('script');
  script.src = 'https://cdn.jsdelivr.net/npm/tsparticles-slim@2/tsparticles.slim.bundle.min.js';
  script.onload = function () {
    if (typeof tsParticles === 'undefined') return; // si el CDN falla, el resto del fondo (mesh/aurora/noise) sigue funcionando igual
    tsParticles.load({
      id: 'fondo-particulas',
      options: {
        fpsLimit: 60,
        particles: {
          number: { value: 26, density: { enable: true, area: 900 } },
          color: { value: ['#0f7a6b', '#e4572e', '#10231c'] },
          opacity: { value: 0.25 },
          size: { value: { min: 1, max: 3 } },
          move: { enable: true, speed: 0.4, direction: 'none', random: true, outModes: { default: 'out' } },
          links: { enable: true, distance: 140, color: '#0f7a6b', opacity: 0.12, width: 1 }
        },
        interactivity: {
          events: { onHover: { enable: true, mode: 'grab' }, resize: true },
          modes: { grab: { distance: 160, links: { opacity: 0.25 } } }
        },
        detectRetina: true
      }
    });
  };
  document.body.appendChild(script);
})();
