/**
 * ============================================================================
 *  FONDO PREMIUM — mesh gradient + aurora + noise + partículas interactivas
 *  + parallax (mouse en escritorio, toque en celular) + ondas al tocar.
 *  Se incluye con una sola línea en cualquier página:
 *
 *      <script src="fondo-premium.js"></script>
 *
 *  No requiere React, build, ni npm — funciona con una etiqueta <script>
 *  normal. Prioriza rendimiento real en celulares económicos (comunes en
 *  RD) por encima del efecto visual máximo:
 *  - Respeta "reducir movimiento" del sistema operativo — si está
 *    activado, no se anima ni se cargan partículas (accesibilidad real,
 *    no solo un detalle técnico).
 *  - En celular, las partículas se muestran con un conteo mucho más bajo
 *    (sin líneas de conexión, que son lo más costoso), para mantenerse
 *    fluido incluso en equipos modestos.
 *  - Todo vive detrás del contenido (z-index negativo) y con
 *    pointer-events:none — nunca interfiere con ningún clic, toque o
 *    formulario.
 * ============================================================================
 */
(function () {
  const prefiereMovimientoReducido = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const esPantallaAncha = window.matchMedia('(min-width: 900px)').matches;
  const esTactil = 'ontouchstart' in window || navigator.maxTouchPoints > 0;

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

  if (prefiereMovimientoReducido) return; // nada de parallax, ondas ni partículas

  // 4) Parallax de los blobs — con el mouse en escritorio, arrastrando el
  //    dedo en celular. Mismo mecanismo, distinta fuente de entrada.
  function moverBlobs(xRelativo, yRelativo) {
    document.querySelectorAll('#fondo-mesh .blob').forEach(function (blob, i) {
      const intensidad = 12 + i * 5; // cada blob se mueve un poco distinto — sensación de profundidad
      blob.style.transform = 'translate(' + (xRelativo * intensidad) + 'px, ' + (yRelativo * intensidad) + 'px)';
    });
  }

  let rafPendiente = false;
  function programarMovimiento(x, y) {
    if (rafPendiente) return;
    rafPendiente = true;
    requestAnimationFrame(function () {
      moverBlobs(x, y);
      rafPendiente = false;
    });
  }

  window.addEventListener('mousemove', function (e) {
    programarMovimiento((e.clientX / window.innerWidth - 0.5) * 2, (e.clientY / window.innerHeight - 0.5) * 2);
  }, { passive: true });

  window.addEventListener('touchmove', function (e) {
    if (!e.touches || !e.touches[0]) return;
    const t = e.touches[0];
    programarMovimiento((t.clientX / window.innerWidth - 0.5) * 2, (t.clientY / window.innerHeight - 0.5) * 2);
  }, { passive: true });

  // 5) Onda al tocar — feedback táctil satisfactorio y visual: un círculo
  //    que se expande y se desvanece desde el punto exacto donde tocaste.
  //    Costo mínimo (una sola animación CSS por toque, se autodestruye).
  if (esTactil) {
    window.addEventListener('touchstart', function (e) {
      if (!e.touches || !e.touches[0]) return;
      const t = e.touches[0];
      const onda = document.createElement('div');
      onda.className = 'onda-toque';
      onda.style.left = t.clientX + 'px';
      onda.style.top = t.clientY + 'px';
      document.body.appendChild(onda);
      onda.addEventListener('animationend', function () { onda.remove(); });
    }, { passive: true });
  }

  // 6) Partículas interactivas (tsParticles vía CDN, sin instalar nada) —
  //    en celular se cargan igual, pero con un conteo mucho más bajo y sin
  //    líneas de conexión (lo más costoso en CPU), para mantenerse fluido
  //    incluso en equipos modestos. En escritorio, versión completa con
  //    conexiones y reacción al pasar el mouse.
  const contenedorParticulas = document.createElement('div');
  contenedorParticulas.id = 'fondo-particulas';
  document.body.appendChild(contenedorParticulas);

  const script = document.createElement('script');
  script.src = 'https://cdn.jsdelivr.net/npm/tsparticles-slim@2/tsparticles.slim.bundle.min.js';
  script.onload = function () {
    if (typeof tsParticles === 'undefined') return; // si el CDN falla, el resto del fondo (mesh/aurora/noise) sigue funcionando igual

    const opcionesEscritorio = {
      fpsLimit: 60,
      particles: {
        number: { value: 26, density: { enable: true, area: 900 } },
        color: { value: ['#0f7a6b', '#e4572e', '#10231c'] },
        opacity: { value: 0.4 },
        size: { value: { min: 1, max: 3 } },
        move: { enable: true, speed: 0.4, direction: 'none', random: true, outModes: { default: 'out' } },
        links: { enable: true, distance: 140, color: '#0f7a6b', opacity: 0.12, width: 1 }
      },
      interactivity: {
        events: { onHover: { enable: true, mode: 'grab' }, resize: true },
        modes: { grab: { distance: 160, links: { opacity: 0.25 } } }
      },
      detectRetina: true
    };

    const opcionesCelular = {
      fpsLimit: 60,
      particles: {
        number: { value: 14, density: { enable: true, area: 700 } },
        color: { value: ['#0f7a6b', '#e4572e'] },
        opacity: { value: 0.42 },
        size: { value: { min: 1, max: 2.5 } },
        move: { enable: true, speed: 0.35, direction: 'none', random: true, outModes: { default: 'out' } },
        links: { enable: false }
      },
      interactivity: {
        events: { onClick: { enable: true, mode: 'bubble' }, resize: true },
        modes: { bubble: { distance: 120, size: 4, duration: 0.4, opacity: 0.5 } }
      },
      detectRetina: true
    };

    tsParticles.load({ id: 'fondo-particulas', options: esPantallaAncha ? opcionesEscritorio : opcionesCelular });
  };
  document.body.appendChild(script);
})();
