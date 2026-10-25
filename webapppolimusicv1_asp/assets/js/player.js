/* ======================================================================
   PoliMusic · reproductor de audio y utilidades de la lista de canciones
   Depende de window.PM_QUEUE (definido en songs.asp) con los objetos:
   {id, title, artists, cover, src, available, plays}
   ====================================================================== */
(function () {
  "use strict";

  var queue = Array.isArray(window.PM_QUEUE) ? window.PM_QUEUE.slice() : [];
  var audio = new Audio();
  audio.preload = "metadata";
  audio.volume = 0.8;
  var current = -1;
  var toastTimer = null;

  function $(id) { return document.getElementById(id); }

  var el = {
    player: $("player"),
    cover: $("pm-cover"),
    title: $("pm-title"),
    artists: $("pm-artists"),
    toggle: $("pm-toggle"),
    prev: $("pm-prev"),
    next: $("pm-next"),
    seek: $("pm-seek"),
    cur: $("pm-cur"),
    dur: $("pm-dur"),
    vol: $("pm-vol"),
    mute: $("pm-mute"),
    count: $("pm-count"),
    toast: $("pm-toast"),
    search: $("song-search"),
    rows: Array.prototype.slice.call(document.querySelectorAll(".song-row")),
    noResults: $("no-results"),
    resultCount: $("result-count")
  };

  function fmt(sec) {
    if (!isFinite(sec) || sec < 0) sec = 0;
    var m = Math.floor(sec / 60);
    var s = Math.floor(sec % 60);
    return m + ":" + (s < 10 ? "0" : "") + s;
  }

  function toast(msg) {
    if (!el.toast) return;
    el.toast.textContent = msg;
    el.toast.classList.add("show");
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { el.toast.classList.remove("show"); }, 3000);
  }

  function rowOf(i) {
    for (var k = 0; k < el.rows.length; k++) {
      if (Number(el.rows[k].getAttribute("data-index")) === i) return el.rows[k];
    }
    return null;
  }

  function firstAvailable() {
    for (var k = 0; k < queue.length; k++) {
      if (queue[k].available && queue[k].src) return k;
    }
    return -1;
  }

  function markRows() {
    var on = !audio.paused && !audio.ended;
    el.rows.forEach(function (r) {
      var isCur = Number(r.getAttribute("data-index")) === current;
      r.classList.toggle("playing", isCur);
      r.classList.toggle("is-playing", isCur && on);
    });
  }

  /* Informa al servidor para sumar una reproducción (PLAYS) */
  function registerPlay(id) {
    if (!id) return;
    fetch("api_play.asp", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: "id=" + encodeURIComponent(id)
    }).then(function (res) {
      if (!res.ok) throw new Error("status " + res.status);
      return res.text();
    }).then(function (txt) {
      if (txt.indexOf("OK") !== 0) return;
      var row = rowOf(current);
      if (!row) return;
      var cell = row.querySelector(".td-plays");
      if (!cell) return;
      var n = Number(cell.getAttribute("data-plays") || cell.textContent) + 1;
      cell.setAttribute("data-plays", String(n));
      cell.textContent = String(n);
      if (queue[current]) queue[current].plays = n;
    }).catch(function () { /* el conteo visual no es crítico */ });
  }

  function setPlayingUI(on) {
    if (el.player) el.player.classList.toggle("playing", on);
    if (el.toggle) el.toggle.setAttribute("aria-label", on ? "Pausar" : "Reproducir");
  }

  function load(i, autoplay) {
    if (i < 0 || i >= queue.length) return;
    var s = queue[i];
    if (!s.available || !s.src) {
      toast("El audio de «" + s.title + "» no está disponible en los recursos.");
      return;
    }
    current = i;
    el.cover.src = s.cover || "";
    el.cover.alt = "Portada de " + s.title;
    el.title.textContent = s.title;
    el.artists.textContent = s.artists || "—";
    el.seek.value = 0;
    el.cur.textContent = "0:00";
    el.dur.textContent = "0:00";
    audio.src = s.src;
    markRows();
    registerPlay(s.id);
    if (autoplay) {
      var p = audio.play();
      if (p && p.catch) p.catch(function () { toast("No se pudo iniciar la reproducción."); });
    }
  }

  function toggle() {
    if (current === -1) {
      var first = firstAvailable();
      if (first === -1) { toast("No hay canciones con audio disponible."); return; }
      load(first, true);
      return;
    }
    if (audio.paused) {
      var p = audio.play();
      if (p && p.catch) p.catch(function () { toast("No se pudo iniciar la reproducción."); });
    } else {
      audio.pause();
    }
  }

  function step(delta) {
    if (queue.length === 0) return;
    if (current === -1) { var f = firstAvailable(); if (f !== -1) load(f, true); return; }
    var idx = current;
    for (var n = 1; n <= queue.length; n++) {
      idx = (idx + delta + queue.length) % queue.length;
      if (queue[idx].available && queue[idx].src) { load(idx, true); return; }
    }
    toast("No hay más canciones con audio disponible.");
  }

  /* ---------------- Eventos del elemento de audio ---------------- */
  audio.addEventListener("timeupdate", function () {
    if (audio.duration && el.seek) {
      el.seek.value = String(Math.round((audio.currentTime / audio.duration) * 1000));
      el.cur.textContent = fmt(audio.currentTime);
    }
  });
  audio.addEventListener("loadedmetadata", function () {
    if (el.dur) el.dur.textContent = fmt(audio.duration);
  });
  audio.addEventListener("play", function () { setPlayingUI(true); markRows(); });
  audio.addEventListener("pause", function () { setPlayingUI(false); markRows(); });
  audio.addEventListener("ended", function () { step(1); });
  audio.addEventListener("error", function () {
    if (current !== -1) toast("No se pudo cargar el audio de «" + queue[current].title + "».");
    setPlayingUI(false);
  });

  /* ---------------- Controles del reproductor ---------------- */
  if (el.player && el.toggle) {
    el.toggle.addEventListener("click", toggle);
    el.prev.addEventListener("click", function () { step(-1); });
    el.next.addEventListener("click", function () { step(1); });
    el.seek.addEventListener("input", function () {
      if (audio.duration) audio.currentTime = (Number(el.seek.value) / 1000) * audio.duration;
    });
    el.vol.addEventListener("input", function () {
      audio.volume = Number(el.vol.value) / 100;
      audio.muted = audio.volume === 0;
      el.player.classList.toggle("muted", audio.muted);
    });
    el.mute.addEventListener("click", function () {
      audio.muted = !audio.muted;
      el.player.classList.toggle("muted", audio.muted);
    });
    if (el.count) {
      el.count.textContent = queue.length + (queue.length === 1 ? " canción en cola" : " canciones en cola");
    }

    /* Teclado: barra espaciadora reproduce o pausa */
    document.addEventListener("keydown", function (ev) {
      if (ev.code !== "Space" && ev.key !== " ") return;
      var t = ev.target;
      if (t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA" || t.tagName === "SELECT" || t.tagName === "BUTTON")) return;
      ev.preventDefault();
      toggle();
    });
  }

  /* ---------------- Tabla: reproducir por fila o botón ---------------- */
  if (el.rows.length) {
    Array.prototype.forEach.call(document.querySelectorAll(".btn-play"), function (btn) {
      btn.addEventListener("click", function (ev) {
        ev.stopPropagation();
        load(Number(btn.getAttribute("data-index")), true);
      });
    });
    el.rows.forEach(function (row) {
      row.addEventListener("click", function (ev) {
        if (ev.target.closest("a, button, input")) return;
        load(Number(row.getAttribute("data-index")), true);
      });
    });
  }

  /* ---------------- Búsqueda en vivo ---------------- */
  if (el.search) {
    el.search.addEventListener("input", function () {
      var q = el.search.value.trim().toLowerCase();
      var visible = 0;
      el.rows.forEach(function (row) {
        var match = row.textContent.toLowerCase().indexOf(q) !== -1;
        row.hidden = !match;
        if (match) visible++;
      });
      if (el.noResults) el.noResults.hidden = visible !== 0;
      if (el.resultCount) {
        el.resultCount.textContent = q
          ? visible + " de " + el.rows.length + " canciones"
          : el.rows.length + " canciones";
      }
    });
    el.search.dispatchEvent(new Event("input"));
  }
})();
