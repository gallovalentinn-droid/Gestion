import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabase";

/* ═══════════════════════════════════════════════════════════
   GESTIÓN DE MOROSOS
   Sistema de diseño + aplicación.

   Escala tipográfica: 7 pasos. Espaciado: múltiplos de 4.
   Radios: 2 (6px general, 4px controles chicos).
   El rojo del escudo es identidad y criticidad; no decora.
   ═══════════════════════════════════════════════════════════ */

const ESCUDO = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAIYAAACgBAMAAAA7ogmJAAAAJFBMVEX///////v//v7+/v/+/v73//zSz8/uNDXuMjbwMDfsMjeEMza/T5KIAAAKJ0lEQVR42q2a348bVxXHP3dsp03ZZO/1Oi0VpDu7AQSq2sx2lyfa7KRseCjQTMUutG+oaoIQEn8APPHIC1qJVkhNVYrEPjRJiyMRJEiaeqvkIVI38RLlgaqxx7sU8aPxXDtbCcWeuTzY+9tje7zch90Hz3x0zrnnnnPu1ybNXlc6bV3aK+Nta89m8J3/AyOdhrNi4NeNQJ9qpqHx1uA2vHB1DtKQntWD+3HcBQuag5vRJAfsLaZp4Rb3yPCbOWNvY8ht//pYyxPLrvBhPdWbBaxnAZrnYabg9sNw9f4wxSYjcxqqGnj4NKbcF4LygfmQTYbYNwcLaJqHPBq/6s+XsWuhYHNfzOfy4LU/y/cZkmb7hfWYVs9ARoOExpv95JwfFa65ltnKENMgFmHNg+P9GPFkamgWvdUXzGKjZVcD3u8DURFy3zRqmy9APnMSUegzNw7Xw9ka7GTAEObLTTL9MCz5x2m9m6GByH0X+smOCq9E9k5GxuTxiCDPUB+MUeEW2W1Hg0zB8vqsJeZqaO30pekCQqyS7+fURcqaF0rutEM2pslQafRlhdC3qcldvkiG4WTG9BPQ0AzNG/TueOSbeQ56DdlPCRm5FmpsOuUHkcyzefbiXWFWgN+Z4XloDV33JoLVxWlLKbOTIT3WfPbxSW6zLnVGpJZuOLORZrcvGrhEplHQ0LXvWcxMVtx7cstj64zMhgN56Hrw7pTklAm3Y1tRWmwxjKGHHdGRmpjWBLsZQJpj0HCBrhXAYiIXlVa3darN9qKbkEc0wDWxVqxI/cR0ZdLUOthxohUKDW63eFiOP3ZdjKLlbsYlmyFfwkVlrpE5qeIYKmRoXgvU7piK0x7fP572G+jsD+FlE5tft565ZnaHCGh+ZQ4yLywXQJzyIPNyjDdWYZTd08qWXM/4jXfdQjtb4lb2mlv2gx3HZ9/FhbVPH5aAKVNwxRgApbjTJg850gTZDees4vHP0gDpf58FmJGeNq8DMBPTHT9b/jUB2R2TTOvvzMahmYlPrxt179DXoxUj4+PRc6WeXxbuLTWepZMdfS5h54zTKf0TTIEryi13KFBJGEqmw3GzBzv8K/6RIa+2pYwmZghnxGG/qRAkt0PJ5mUE+DhiSFYnyDpJGerc+UNz8lKB0Q8PFDPTI52n5e478c7pjAcvv/vG8W8FT+RqMjlDXvixB5D5gfXWz87WpoURSRmL5kcenCPjMTdHI1eTvpIJGd943KNx/k34zylg37G/Hh3TCe1QF34OZ38/A2cPeTAbOJGQnUt97GqcgvsLLnDiI4Bi3MNdGH/y4LcegDmYB9Ol5cSuU3Bf/6H11Jlu7a8Lw4Nzsj0WnQThJGd8AKy190FfBm4mZ8xB4wJXKL0/dNk/ocGJG23i9rYpJWA++DPA6fMPamkSxyO95kHkTufhXD7z0rSPSO7LB8B5zGs0Xn8tz0HT7Zbbo37kYEYVvJQMEDKpHe12+SRcwQU7qCa3Y3NePiJcjG9TSxqPpmz1vQjsHNQOwNgAebq+Lh+XhLU9xBQyf4GqVl1H1r66U3ZvdvBpPSuVGNQOCRK0DBB3faLEua4BoooFKVGGkcdY1kntKABYowaMtiESgZ04TxuAp2oZAAVCKDVQPDK+zGEiX4MgLCZmHAPEhxUXkbUVRlM/nJgxlIfMsycUZMckd4f1uPKTMtY08Fj75FXrN9Ugdf0c8FIYVIMgKGlxVFKzEzOeB4QIhAap7qQwsUkWz9BvASrAVEtoS2Bin+2yt+cAMaZVeVwwIdEycZ7iP5cH1DATJd9XmFqYtL+w9M3WPJgNRFYqCORdOykjLcTTWgKKbEv1HUm8t08Ucy0BsSW7lAepH/cmXIugDTFBaZA6ZufQygRGSExQ9ScHYQjXKMSBJcuuODeYDFRyX5pX0YEfpiYnAsnIpMFPyKgcVJl5DKQwwbipSkTsYYnzZXTVvh1ZWRQI1epvdtJ9ieTH8/X+70YdGavyk9DfI2O0OZty+md0ikdJjLyU5Lq52w5TnRyfupOEsdsOIai5QZzKV6Tt5RJMxjPkp1PE3HaoqLblkepmR3PEciNpdMfOKIZFC25tynQd4lEQufipZyy7fpNS2WyMHZFFevHVLupYxaiaDZSNEHZnRqpsu7dNTerhmG0cg2EojY7F7m1ELZLz1UDIzgNtJJRSxjfDKLU5B2xnWM1jI4uG0dhsMkBWtgqB0Z1jmr7JrO4iGPo+CONzD8ydrXboLb5MLh7Tq/Ea7N22Crsh/fi7fTF8r15z4u1IAYLRSQGs2u26Ym2XoP0Lx40M4jXpUWkAS2MQRzvFo3KgOPEK2EE3NRoYMaEENmYra1OSi5yCW3G1YjyeEYJuj61saFzb4pENp0wPOQYgWk5tU5asrYLiUxlX+70ZVmtvnM45dhXlBD0qjoZQIba8upUx3JwvEnRVpA9CgHgKA4TRTkbkRxPXQ6fXlw2ydbna9p3aBuOWrDN/C1Rf5TcqdvLFjdSikbpnAdZAVuqtqbnBMPeGZvuQ2iDUCCz05qi5GQ8yzyzbsg8rAKpO6+xsYfj49S/mdL+9rb11/jaGgy0OuasExZ7vF6VeLz92W3a2ACqiImWu6qC6+rIEgnBEgxmvAUUzXNxgjHK4JtyVHmMG40EAjgnWvRnV0tlgRAUV5ao9o6FU0D6yASEgVGrTF+tp66BrmZ7JcVOBtRwCNy2NIFzaZPgprhvZc2oJbCWRtqMxtpJg6lv25bFac752c0z29EWBJYuYohASxtWRDUY4wn6k3TvNFYA6Au3vKtXwCoDlQqEwJHWkejI27DRPrSebDtv9VjgZ19RkT8bNiRYqiERgWp3OFoB1CSaHckE/KT5BeKNUCihrfScolUraCCnadsjFV/v7qYRJTWjgqFV1agrupBxMO6bN2ajUzziqQUMtsBApS2uEbXwEpGHuFy8aVb67+53sqM6u319ECeziROTDQS3Lw1SpWuGNyRDy6WnElLgxqZY6HvJi6wD4snXWlwDK6xpm5AMOCO8LXrVqxdTRdt5FdRl1GKgDQIzxy/SaRlVjciM8Wm+Pm2G9wyMKjEWjItLWRcx74x0OnDlS3Rj9TK2D1CdkaUqy4Avcxz3omB+yrFpTqhHG6tiIFTR+6ouZ8PZC3GZWRTUbWWCq9ybinlnQOrXip2txl8YHP8kK/bdH9H9XvvZgnCjwu4cLqcg07/1r9asx9Xv58/989IGH9j/waIzw9M6ZYwUEcKJR6PjETzzjy7BeO2xZsvFcZ8rs+VZxTbsxcvBFgurK8Fggxlj4e4yoSffvrkde9EwgjAiysnGyMZhe+OwZBGiE5Kw3oOZoHsqjtDKKRuVK79YZo9OJi1RFMM7bHxcG1T7Nc3lkYEFFMiiDtTNY0uZ+9/bV4zd9s494wG9u+EnvyFvWe595cN/4DO4Lznfz8MYef5P5pX2XLk71eCbV4/P6t/9xtekn1wy23ZjTH/W09X9SK7MZD7AC1wAAAABJRU5ErkJggg==";
const MESES_ALERTA = 3;          // más de esto dispara la alerta de atraso
const DIAS_ENTRE_RECLAMOS = 7;   // no reaparece en la cola antes de esto
const POR_PAGINA = 50;

const ALIAS_ACTIVIDAD = { GIMNASIA: "ACROBÁTICA" };
const ORDEN_ACTIVIDAD = ["RUGBY", "HOCKEY", "ACROBÁTICA", "TENIS", "MORENO FANS", "SIN ACTIVIDAD"];
const ordenAct = (a) => {
  const i = ORDEN_ACTIVIDAD.indexOf(a);
  return i === -1 ? ORDEN_ACTIVIDAD.length : i;
};
const CORTE_ANIO_ACTUAL = 202601;

/* Estados: cada uno declara su tono semántico y si cierra la gestión. */
const ESTADOS = [
  { id: "RECLAMAR",      label: "A reclamar",    tono: "critico", cierra: false },
  { id: "RECLAMADO",     label: "Reclamado",     tono: "info",    cierra: false },
  { id: "PAGO",          label: "Pagó",          tono: "exito",   cierra: true  },
  { id: "PAGO PARCIAL",  label: "Pago parcial",  tono: "aviso",   cierra: false },
  { id: "AL DIA",        label: "Al día",        tono: "exito",   cierra: true  },
  { id: "PLAN DE PAGOS", label: "Plan de pagos", tono: "info",    cierra: true  },
  { id: "PEDIR NUMERO",  label: "Pedir número",  tono: "neutro",  cierra: false },
  { id: "MIRAR CUENTA",  label: "Mirar cuenta",  tono: "neutro",  cierra: false },
  { id: "BAJA/RENUNCIA", label: "Baja",          tono: "apagado", cierra: true  },
  { id: "NO RECLAMAR",   label: "No reclamar",   tono: "apagado", cierra: true  },
];
const ESTADO_MAP = Object.fromEntries(ESTADOS.map((e) => [e.id, e]));

/* ── Formato ─────────────────────────────────────────────── */

const money = (n) =>
  new Intl.NumberFormat("es-AR", { style: "currency", currency: "ARS", maximumFractionDigits: 0 }).format(n || 0);

const moneyCorto = (n) => {
  const v = Math.round(n || 0);
  if (Math.abs(v) >= 1000000) return `$${(v / 1000000).toFixed(1)}M`;
  if (Math.abs(v) >= 1000) return `$${Math.round(v / 1000)}k`;
  return `$${v}`;
};

const MESES_CORTOS = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
const periodoLabel = (p) => `${MESES_CORTOS[parseInt(String(p).slice(4, 6), 10) - 1]} ${String(p).slice(0, 4)}`;

const fechaLarga = (iso) => {
  if (!iso) return "";
  const d = new Date(iso);
  return `${String(d.getDate()).padStart(2, "0")}/${MESES_CORTOS[d.getMonth()]}/${d.getFullYear()}`;
};

const diasDesde = (iso) => (iso ? Math.floor((Date.now() - new Date(iso).getTime()) / 86400000) : null);
const fechaAccion = (e) => fechaLarga(e?.log?.[0]?.fecha || e?.ultimoReclamo || null);

/* ═══════════════════════════════════════════════════════════
   SISTEMA DE DISEÑO
   ═══════════════════════════════════════════════════════════ */

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Bodoni+Moda:opsz,wght@6..96,600&family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@400;450;500;600;700&display=swap');

.ds{
  /* — Neutrales — */
  --n0:#ffffff; --n50:#fafafa; --n100:#f4f5f5; --n200:#eaebec;
  --n300:#dcdee0; --n400:#b0b4b8; --n500:#6a6f74; --n700:#42464a; --n900:#1e2124;
  /* — Marca y semánticos — */
  --marca:#ec3237;            /* identidad: escudo e indicador de sección */
  --critico:#c0272c; --critico-bg:#fdeced; --critico-bd:#f6cfd0;
  --aviso:#8a5a06;   --aviso-bg:#fdf3e0;   --aviso-bd:#f0dcb4;
  --exito:#146b45;   --exito-bg:#e8f3ed;   --exito-bd:#c7e2d5;
  --info:#2f5c8c;    --info-bg:#ecf2f8;    --info-bd:#cfdeed;
  --foco:#2f5c8c;             /* anillo de foco: azul, para no confundir con error */
  /* — Superficies — */
  --fondo:var(--n100); --sup:var(--n0); --sup2:var(--n50); --borde:var(--n200); --borde2:var(--n300);
  /* — Espaciado, múltiplos de 4 — */
  --s1:4px; --s2:8px; --s3:12px; --s4:16px; --s5:20px; --s6:24px; --s8:32px; --s10:40px; --s12:48px;
  /* — Radios y elevación — */
  --r-sm:4px; --r:6px; --r-pill:999px;
  --el1:0 1px 2px rgba(30,33,36,.06);
  --el2:0 8px 24px rgba(30,33,36,.12), 0 1px 3px rgba(30,33,36,.06);
  --alto-barra:56px; --alto-nav:44px;

  font-family:'IBM Plex Sans',system-ui,-apple-system,sans-serif;
  font-size:14px; line-height:1.5; font-weight:450;
  color:var(--n900); background:var(--fondo); min-height:100vh;
  -webkit-font-smoothing:antialiased;
}
.ds *{box-sizing:border-box}
.ds ::selection{background:#fddfe0}

/* ── Escala tipográfica (7 pasos) ─────────────────────────── */
.ds .t-display{font-family:'Bodoni Moda',Georgia,serif;font-weight:600;font-size:28px;line-height:1.15;letter-spacing:.005em}
.ds .t-h1{font-size:20px;line-height:1.25;font-weight:600;letter-spacing:-.01em}
.ds .t-h2{font-size:15px;line-height:1.35;font-weight:600;letter-spacing:-.005em}
.ds .t-body{font-size:14px;line-height:1.5;font-weight:450}
.ds .t-strong{font-size:14px;line-height:1.5;font-weight:600}
.ds .t-label{font-size:12px;line-height:1.35;font-weight:500;color:var(--n700)}
.ds .t-cap{font-size:11px;line-height:1.35;font-weight:500;color:var(--n500)}
.ds .t-eyebrow{font-size:11px;line-height:1;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:var(--n500)}
.ds .num{font-family:'IBM Plex Mono',monospace;font-variant-numeric:tabular-nums;font-weight:500;letter-spacing:-.01em}
.ds .num-xl{font-family:'IBM Plex Mono',monospace;font-variant-numeric:tabular-nums;font-weight:600;font-size:24px;line-height:1.15;letter-spacing:-.03em}
.ds .apagado{color:var(--n500)}

/* ── Foco visible, uno solo para todo ─────────────────────── */
.ds :focus-visible{outline:2px solid var(--foco);outline-offset:2px;border-radius:var(--r-sm)}
.ds :focus:not(:focus-visible){outline:none}

/* ── Botones ──────────────────────────────────────────────── */
.ds .btn{
  display:inline-flex;align-items:center;justify-content:center;gap:var(--s2);
  font:inherit;font-size:13px;font-weight:550;line-height:1;
  min-height:36px;padding:0 var(--s3);border-radius:var(--r);
  border:1px solid var(--borde2);background:var(--sup);color:var(--n900);
  cursor:pointer;white-space:nowrap;transition:background .12s,border-color .12s,box-shadow .12s;
}
.ds .btn:hover:not(:disabled){background:var(--n100);border-color:var(--n400)}
.ds .btn:active:not(:disabled){background:var(--n200)}
.ds .btn:disabled{opacity:.45;cursor:not-allowed}
.ds .btn-1{background:var(--n900);border-color:var(--n900);color:#fff;box-shadow:var(--el1)}
.ds .btn-1:hover:not(:disabled){background:#000;border-color:#000}
.ds .btn-1:active:not(:disabled){background:#000}
.ds .btn-3{border-color:transparent;background:transparent;color:var(--n700)}
.ds .btn-3:hover:not(:disabled){background:var(--n200);border-color:transparent}
.ds .btn-peligro{color:var(--critico);border-color:var(--critico-bd)}
.ds .btn-peligro:hover:not(:disabled){background:var(--critico-bg);border-color:var(--critico)}
.ds .btn-peligro-1{background:var(--critico);border-color:var(--critico);color:#fff}
.ds .btn-peligro-1:hover:not(:disabled){background:#a91f24;border-color:#a91f24}
.ds .btn-sm{min-height:30px;font-size:12px;padding:0 var(--s2)}
.ds .btn-icono{min-width:36px;padding:0 var(--s2)}
.ds .btn svg{width:15px;height:15px;flex:none}
.ds .spin{width:13px;height:13px;border:2px solid currentColor;border-right-color:transparent;border-radius:50%;animation:gira .6s linear infinite}
@keyframes gira{to{transform:rotate(360deg)}}

/* ── Insignias de estado ──────────────────────────────────── */
.ds .badge{display:inline-flex;align-items:center;gap:5px;font-size:11px;font-weight:550;line-height:1;
  padding:var(--s1) var(--s2);border-radius:var(--r-pill);border:1px solid transparent;white-space:nowrap}
.ds .badge::before{content:"";width:5px;height:5px;border-radius:50%;background:currentColor;flex:none}
.ds .badge-critico{background:var(--critico-bg);color:var(--critico);border-color:var(--critico-bd)}
.ds .badge-aviso{background:var(--aviso-bg);color:var(--aviso);border-color:var(--aviso-bd)}
.ds .badge-exito{background:var(--exito-bg);color:var(--exito);border-color:var(--exito-bd)}
.ds .badge-info{background:var(--info-bg);color:var(--info);border-color:var(--info-bd)}
.ds .badge-neutro{background:var(--n100);color:var(--n700);border-color:var(--borde)}
.ds .badge-apagado{background:var(--n100);color:var(--n500);border-color:var(--borde)}
.ds .badge-vacio{background:transparent;color:var(--n500);border-color:var(--borde2);border-style:dashed}
.ds .badge-vacio::before{display:none}
.ds .cuenta{display:inline-flex;align-items:center;justify-content:center;min-width:20px;height:18px;padding:0 var(--s2);
  border-radius:var(--r-pill);background:var(--n200);color:var(--n700);font-size:11px;font-weight:600;line-height:1}

/* ── Barra superior ───────────────────────────────────────── */
.ds .barra{position:sticky;top:0;z-index:40;background:var(--sup);border-bottom:1px solid var(--borde)}
.ds .barra-in{max-width:1240px;margin:0 auto;height:var(--alto-barra);padding:0 var(--s6);
  display:flex;align-items:center;gap:var(--s4)}
.ds .escudo{height:32px;width:auto;display:block;flex:none}
.ds .marca{display:flex;align-items:center;gap:var(--s3);flex:none}
.ds .marca .titulo{font-family:'Bodoni Moda',Georgia,serif;font-weight:600;font-size:20px;line-height:1.1;letter-spacing:.005em}
.ds .marca .sep{width:1px;height:24px;background:var(--borde2)}

/* ── Navegación de secciones ──────────────────────────────── */
.ds .nav{position:sticky;top:var(--alto-barra);z-index:39;background:var(--sup);border-bottom:1px solid var(--borde)}
.ds .nav-in{max-width:1240px;margin:0 auto;padding:0 var(--s6);display:flex;gap:var(--s1);height:var(--alto-nav);align-items:stretch;overflow-x:auto;scrollbar-width:none}
.ds .nav-in::-webkit-scrollbar{display:none}
.ds .nav-in button{display:inline-flex;align-items:center;gap:var(--s2);font:inherit;font-size:13px;font-weight:500;
  padding:0 var(--s3);border:0;border-bottom:2px solid transparent;background:none;color:var(--n500);cursor:pointer;white-space:nowrap}
.ds .nav-in button:hover{color:var(--n900)}
.ds .nav-in button[aria-current="page"]{color:var(--n900);font-weight:600;border-bottom-color:var(--marca)}

/* ── Cuerpo ───────────────────────────────────────────────── */
.ds .lienzo{max-width:1240px;margin:0 auto;padding:var(--s6) var(--s6) var(--s12)}
.ds .cabecera-vista{display:flex;gap:var(--s4);align-items:flex-start;margin-bottom:var(--s5);flex-wrap:wrap}
.ds .cabecera-vista .txt{flex:1 1 320px;min-width:0}
.ds .cabecera-vista p{margin:var(--s1) 0 0;color:var(--n500);font-size:13px;max-width:62ch}
.ds .cabecera-vista .cta{display:flex;gap:var(--s2);align-items:center;flex-wrap:wrap}

/* ── Filtro de actividad (segmentado) ─────────────────────── */
.ds .filtro-act{display:flex;gap:var(--s1);flex-wrap:wrap;margin-bottom:var(--s5);
  padding:var(--s1);background:var(--n200);border-radius:var(--r)}
.ds .filtro-act button{display:inline-flex;align-items:center;gap:var(--s2);font:inherit;font-size:13px;font-weight:500;
  min-height:32px;padding:0 var(--s3);border:0;border-radius:var(--r-sm);background:transparent;color:var(--n700);cursor:pointer}
.ds .filtro-act button:hover{background:rgba(255,255,255,.6)}
.ds .filtro-act button[aria-pressed="true"]{background:var(--sup);color:var(--n900);font-weight:600;box-shadow:var(--el1)}

/* ── Resumen de contexto ──────────────────────────────────── */
.ds .contexto{display:flex;gap:var(--s6);align-items:baseline;flex-wrap:wrap;
  padding:var(--s4) 0 var(--s5);border-bottom:1px solid var(--borde);margin-bottom:var(--s5)}
.ds .contexto .dato .k{font-size:11px;font-weight:500;color:var(--n500);letter-spacing:.04em;text-transform:uppercase}
.ds .contexto .dato .v{margin-top:2px}
.ds .contexto .primario .v{font-size:28px}

/* ── Tarjetas KPI ─────────────────────────────────────────── */
.ds .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(196px,1fr));gap:var(--s3);margin-bottom:var(--s6)}
.ds .kpi{background:var(--sup);border:1px solid var(--borde);border-radius:var(--r);padding:var(--s4)}
.ds .kpi .k{font-size:11px;font-weight:500;color:var(--n500);letter-spacing:.04em;text-transform:uppercase}
.ds .kpi .v{margin-top:var(--s2)}
.ds .kpi .s{margin-top:var(--s1);font-size:12px;color:var(--n500)}
.ds .kpi-1{grid-column:span 1}
.ds .kpi-1 .v{font-size:28px}

/* ── Lista de trabajo (cola) ──────────────────────────────── */
.ds .cola{display:flex;flex-direction:column;gap:var(--s3)}
.ds .fila-fam{background:var(--sup);border:1px solid var(--borde);border-radius:var(--r);
  transition:border-color .12s,box-shadow .12s}
.ds .fila-fam:hover{border-color:var(--borde2);box-shadow:var(--el1)}
.ds .fam-cuerpo{display:flex;gap:var(--s5);padding:var(--s4) var(--s5);align-items:flex-start}
.ds .fam-info{flex:1 1 auto;min-width:0}
.ds .fam-tit{display:flex;align-items:center;gap:var(--s2);flex-wrap:wrap}
.ds .fam-tit .nom{font-size:15px;font-weight:600;letter-spacing:-.005em}
.ds .fam-sub{margin-top:var(--s1);font-size:12px;color:var(--n500);display:flex;gap:var(--s2);flex-wrap:wrap;align-items:center}
.ds .fam-sub .punto{color:var(--n400)}
.ds .hermanos{margin-top:var(--s3);padding-left:var(--s3);border-left:2px solid var(--borde2);
  display:flex;flex-direction:column;gap:var(--s1)}
.ds .hermano{display:flex;gap:var(--s2);align-items:baseline;font-size:13px;color:var(--n700)}
.ds .hermano .h-nom{font-weight:500;color:var(--n900);background:none;border:0;padding:0;font:inherit;font-weight:500;cursor:pointer;text-align:left}
.ds .hermano .h-nom:hover{text-decoration:underline;text-underline-offset:2px}
.ds .hermano .h-imp{margin-left:auto;padding-left:var(--s3);font-size:12px}
.ds .fam-plata{flex:none;text-align:right;min-width:132px}
.ds .fam-plata .sub{margin-top:var(--s1);font-size:11px;color:var(--n500)}
.ds .fam-acc{display:flex;gap:var(--s2);align-items:center;padding:var(--s3) var(--s5);
  border-top:1px solid var(--borde);background:var(--sup2);border-radius:0 0 var(--r) var(--r);flex-wrap:wrap}
.ds .fam-acc .der{margin-left:auto;display:flex;gap:var(--s2);align-items:center}

/* ── Menú desplegable ─────────────────────────────────────── */
.ds .menu-cont{position:relative}
.ds .menu{position:absolute;right:0;top:calc(100% + 6px);min-width:212px;background:var(--sup);
  border:1px solid var(--borde);border-radius:var(--r);box-shadow:var(--el2);padding:var(--s1);z-index:60}
.ds .menu.izq{right:auto;left:0}
.ds .menu button{display:flex;width:100%;align-items:center;gap:var(--s2);font:inherit;font-size:13px;font-weight:450;
  padding:var(--s2) var(--s3);border:0;border-radius:var(--r-sm);background:none;color:var(--n900);cursor:pointer;text-align:left;min-height:34px}
.ds .menu button:hover{background:var(--n100)}
.ds .menu button.peligro{color:var(--critico)}
.ds .menu button.peligro:hover{background:var(--critico-bg)}
.ds .menu hr{border:0;border-top:1px solid var(--borde);margin:var(--s1) 0}
.ds .menu .titulo-menu{padding:var(--s2) var(--s3) var(--s1);font-size:11px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:var(--n500)}

/* ── Buscador ─────────────────────────────────────────────── */
.ds .buscador{position:relative;flex:1 1 300px;max-width:420px}
.ds .campo-busca{display:flex;align-items:center;gap:var(--s2);height:36px;padding:0 var(--s3);
  border:1px solid var(--borde2);border-radius:var(--r);background:var(--sup2);transition:background .12s,border-color .12s}
.ds .campo-busca:focus-within{background:var(--sup);border-color:var(--n400)}
.ds .campo-busca svg{width:15px;height:15px;color:var(--n500);flex:none}
.ds .campo-busca input{flex:1;min-width:0;border:0;background:none;font:inherit;font-size:13px;color:var(--n900);outline:none}
.ds .campo-busca input::placeholder{color:var(--n500)}
.ds .campo-busca kbd{font-family:'IBM Plex Mono';font-size:11px;color:var(--n500);border:1px solid var(--borde2);
  border-radius:var(--r-sm);padding:2px var(--s1);background:var(--sup);flex:none}
.ds .resultados{position:absolute;top:calc(100% + 6px);left:0;right:0;background:var(--sup);border:1px solid var(--borde);
  border-radius:var(--r);box-shadow:var(--el2);max-height:342px;overflow:auto;z-index:60;padding:var(--s1)}
.ds .resultados button{display:flex;width:100%;gap:var(--s3);align-items:center;text-align:left;font:inherit;
  background:none;border:0;border-radius:var(--r-sm);padding:var(--s2) var(--s3);cursor:pointer;min-height:40px}
.ds .resultados button:hover,.ds .resultados button.activo{background:var(--n100)}
.ds .resultados .r-nom{font-size:13px;font-weight:500}
.ds .resultados .r-meta{font-size:11px;color:var(--n500);margin-top:1px}
.ds .resultados .r-imp{margin-left:auto;text-align:right;font-size:13px}
.ds .resultados .r-vacio{padding:var(--s4);font-size:13px;color:var(--n500)}

/* ── Tabla ────────────────────────────────────────────────── */
.ds .caja-tabla{background:var(--sup);border:1px solid var(--borde);border-radius:var(--r)}
.ds .caja-tabla > table:first-child thead th:first-child{border-top-left-radius:var(--r)}
.ds .caja-tabla > table:first-child thead th:last-child{border-top-right-radius:var(--r)}
.ds .caja-tabla .vacio{border:0;border-radius:0;background:transparent}
.ds .barra-tabla{display:flex;gap:var(--s2);align-items:center;flex-wrap:wrap;padding:var(--s3) var(--s4);border-bottom:1px solid var(--borde)}
.ds table{width:100%;border-collapse:collapse}
.ds thead th{position:sticky;top:calc(var(--alto-barra) + var(--alto-nav));z-index:5;background:var(--sup2);
  font-size:11px;font-weight:600;letter-spacing:.05em;text-transform:uppercase;color:var(--n500);
  text-align:left;padding:var(--s2) var(--s4);border-bottom:1px solid var(--borde);white-space:nowrap}
.ds thead th button{display:inline-flex;align-items:center;gap:4px;font:inherit;background:none;border:0;padding:0;cursor:pointer;color:inherit}
.ds thead th button:hover{color:var(--n900)}
.ds thead th .flecha{font-size:10px;color:var(--n900)}
.ds tbody td{padding:var(--s3) var(--s4);border-bottom:1px solid var(--n100);font-size:13px;vertical-align:middle}
.ds tbody tr{cursor:pointer}
.ds tbody tr:hover td{background:var(--sup2)}
.ds tbody tr:last-child td{border-bottom:0}
.ds .celda-num{font-family:'IBM Plex Mono',monospace;font-variant-numeric:tabular-nums;text-align:right;white-space:nowrap}
.ds .vacio-celda{color:var(--n400)}
.ds .paginacion{display:flex;gap:var(--s3);align-items:center;justify-content:space-between;padding:var(--s3) var(--s4);border-top:1px solid var(--borde);flex-wrap:wrap}

/* ── Estados vacíos ───────────────────────────────────────── */
.ds .vacio{padding:var(--s12) var(--s6);text-align:center;background:var(--sup);border:1px solid var(--borde);border-radius:var(--r)}
.ds .vacio .icono{width:40px;height:40px;margin:0 auto var(--s3);color:var(--n400)}
.ds .vacio h3{margin:0 0 var(--s2);font-size:15px;font-weight:600}
.ds .vacio p{margin:0 auto;max-width:44ch;color:var(--n500);font-size:13px}
.ds .vacio .acc{margin-top:var(--s5);display:flex;gap:var(--s2);justify-content:center;flex-wrap:wrap}

/* ── Pantalla de bienvenida ───────────────────────────────── */
.ds .inicio{max-width:560px;margin:var(--s12) auto;text-align:center;position:relative}
.ds .inicio .agua{width:110px;margin:0 auto var(--s6);display:block;opacity:.9}
.ds .inicio h2{margin:0 0 var(--s3)}
.ds .inicio p{margin:0 auto var(--s6);color:var(--n500);font-size:14px;max-width:46ch}
.ds .pasos{text-align:left;background:var(--sup);border:1px solid var(--borde);border-radius:var(--r);
  padding:var(--s5);margin-top:var(--s8);display:flex;flex-direction:column;gap:var(--s4)}
.ds .paso{display:flex;gap:var(--s3);align-items:flex-start}
.ds .paso .n{flex:none;width:22px;height:22px;border-radius:50%;background:var(--n900);color:#fff;
  font-size:11px;font-weight:600;display:flex;align-items:center;justify-content:center;margin-top:1px}
.ds .paso .d{font-size:13px;color:var(--n500);margin-top:2px}

/* ── Panel lateral ────────────────────────────────────────── */
.ds .velo{position:fixed;inset:0;background:rgba(30,33,36,.42);z-index:70;animation:fundir .15s ease-out}
@keyframes fundir{from{opacity:0}to{opacity:1}}
.ds .panel{position:fixed;top:0;right:0;bottom:0;width:min(580px,100%);background:var(--sup);z-index:71;
  overflow:auto;display:flex;flex-direction:column;animation:entrar .18s ease-out}
@keyframes entrar{from{transform:translateX(16px);opacity:.6}to{transform:none;opacity:1}}
.ds .panel-barra{position:sticky;top:0;background:var(--sup);border-bottom:1px solid var(--borde);
  padding:var(--s3) var(--s5);display:flex;align-items:center;gap:var(--s3);z-index:2}
.ds .panel-barra .ruta{font-size:12px;color:var(--n500);flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ds .panel-sec{padding:var(--s5);border-bottom:1px solid var(--borde)}
.ds .panel-sec:last-child{border-bottom:0}
.ds .panel-sec h4{margin:0 0 var(--s3);font-size:11px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:var(--n500)}
.ds .resumen-socio{display:flex;gap:var(--s5);align-items:flex-start;flex-wrap:wrap}
.ds .detalle{width:100%;border-collapse:collapse;font-size:13px}
.ds .detalle td{padding:var(--s2) 0;border-bottom:1px solid var(--n100)}
.ds .detalle .per{font-size:12px;font-weight:600;color:var(--n700);padding-top:var(--s3)}
.ds .detalle .imp{font-family:'IBM Plex Mono';font-variant-numeric:tabular-nums;text-align:right;white-space:nowrap}
.ds .linea-log{display:flex;gap:var(--s3);align-items:baseline;font-size:13px;padding:var(--s2) 0;border-bottom:1px solid var(--n100)}
.ds .linea-log .f{font-family:'IBM Plex Mono';font-size:11px;color:var(--n500);white-space:nowrap;min-width:82px}
.ds .comentario{padding:var(--s2) 0 var(--s2) var(--s3);border-left:2px solid var(--info-bd);margin-bottom:var(--s3);font-size:13px}
.ds .comentario .meta{font-size:11px;color:var(--n500);margin-bottom:2px}

/* ── Formularios ──────────────────────────────────────────── */
.ds .campo{display:flex;flex-direction:column;gap:var(--s1)}
.ds .campo > label{font-size:12px;font-weight:550;color:var(--n700)}
.ds .campo .ayuda{font-size:11px;color:var(--n500)}
.ds .campo .error{font-size:11px;color:var(--critico);display:flex;gap:5px;align-items:center}
.ds input[type=checkbox]{width:15px;height:15px;min-height:auto;accent-color:var(--n900)}
.ds input[type=text],.ds input[type=date],.ds select,.ds textarea{font:inherit;font-size:13px;min-height:36px;padding:var(--s2) var(--s3);
  border:1px solid var(--borde2);border-radius:var(--r);background:var(--sup);color:var(--n900);width:100%}
.ds input[type=text]:hover,.ds select:hover{border-color:var(--n400)}
.ds textarea{min-height:72px;resize:vertical;line-height:1.5}
.ds .campo-plata{display:flex;align-items:center;border:1px solid var(--borde2);border-radius:var(--r);background:var(--sup);overflow:hidden}
.ds .campo-plata:focus-within{outline:2px solid var(--foco);outline-offset:2px}
.ds .campo-plata .sig{padding:0 var(--s1) 0 var(--s3);color:var(--n500);font-family:'IBM Plex Mono';font-size:13px}
.ds .campo-plata input{border:0;outline:none;text-align:right;font-family:'IBM Plex Mono';font-variant-numeric:tabular-nums;
  font-size:14px;font-weight:500;padding:var(--s2) var(--s3) var(--s2) 0;border-radius:0}
.ds .campo-plata.mal{border-color:var(--critico)}
.ds .fila-cobro{display:flex;gap:var(--s3);align-items:flex-end;padding:var(--s3) 0;border-bottom:1px solid var(--n100);flex-wrap:wrap}
.ds .fila-cobro .quien{flex:1 1 170px;min-width:0}
.ds .fila-cobro .campo{flex:0 0 168px}

/* ── Avisos y notificaciones ──────────────────────────────── */
.ds .aviso{display:flex;gap:var(--s3);align-items:flex-start;padding:var(--s3) var(--s4);border-radius:var(--r);
  border:1px solid var(--info-bd);background:var(--info-bg);color:var(--info);font-size:13px;line-height:1.5}
.ds .aviso.warn{border-color:var(--aviso-bd);background:var(--aviso-bg);color:var(--aviso)}
.ds .aviso.err{border-color:var(--critico-bd);background:var(--critico-bg);color:var(--critico)}
.ds .aviso svg{width:16px;height:16px;flex:none;margin-top:1px}
.ds .avisos{position:fixed;left:50%;transform:translateX(-50%);bottom:var(--s6);z-index:90;
  display:flex;flex-direction:column;gap:var(--s2);width:min(520px,calc(100vw - 32px))}
.ds .toast{display:flex;gap:var(--s3);align-items:center;background:var(--n900);color:#fff;border-radius:var(--r);
  padding:var(--s3) var(--s3) var(--s3) var(--s4);box-shadow:var(--el2);font-size:13px;animation:sube .18s ease-out}
@keyframes sube{from{transform:translateY(8px);opacity:0}to{transform:none;opacity:1}}
.ds .toast .txt{flex:1;line-height:1.45}
.ds .toast.err{background:var(--critico)}
.ds .toast .btn{background:transparent;border-color:rgba(255,255,255,.32);color:#fff}
.ds .toast .btn:hover{background:rgba(255,255,255,.14);border-color:rgba(255,255,255,.5)}

/* ── Modal ────────────────────────────────────────────────── */
.ds .modal-velo{position:fixed;inset:0;background:rgba(30,33,36,.5);z-index:100;display:flex;align-items:center;justify-content:center;padding:var(--s4)}
.ds .modal{background:var(--sup);border-radius:var(--r);box-shadow:var(--el2);width:min(460px,100%);padding:var(--s6);animation:sube .16s ease-out}
.ds .modal h3{margin:0 0 var(--s2);font-size:20px;font-weight:600}
.ds .modal > div{color:var(--n500);font-size:13px;line-height:1.55}
.ds .modal .acc{display:flex;gap:var(--s2);justify-content:flex-end;margin-top:var(--s6);flex-wrap:wrap}

/* ── Responsive ───────────────────────────────────────────── */
@media (max-width:900px){
  .ds .lienzo{padding:var(--s5) var(--s4) var(--s10)}
  .ds .barra-in,.ds .nav-in{padding:0 var(--s4)}
  .ds .oculta-md{display:none!important}
}
@media (max-width:680px){
  .ds{font-size:14px;--alto-barra:52px}
  .ds .lienzo{padding:var(--s4) var(--s4) var(--s10)}
  .ds .marca .titulo,.ds .marca .sep{display:none}
  .ds .buscador{max-width:none}
  .ds .oculta-sm{display:none!important}
  .ds .fam-cuerpo{flex-direction:column;gap:var(--s3);padding:var(--s4)}
  .ds .fam-plata{text-align:left;min-width:0}
  .ds .fam-acc{padding:var(--s3) var(--s4)}
  .ds .fam-acc .btn{flex:1 1 auto}
  .ds .fam-acc .der{margin-left:0;width:100%}
  .ds thead th{position:static}
  .ds tbody td{padding:var(--s3)}
  .ds .kpis{grid-template-columns:1fr 1fr}
  .ds .kpi-1{grid-column:1 / -1}
  .ds .contexto{gap:var(--s5)}
  .ds .btn{min-height:40px}
  .ds .nav-in button{min-height:44px}
  .ds .avisos{bottom:var(--s3)}
  .ds .panel{width:100%}
}
@media (prefers-reduced-motion:reduce){.ds *{animation:none!important;transition:none!important}}
`;

/* ═══════════════════════════════════════════════════════════
   COMPONENTES BASE
   ═══════════════════════════════════════════════════════════ */

const Ico = {
  busca: <><circle cx="7" cy="7" r="5" /><path d="M11 11l4 4" strokeLinecap="round" /></>,
  cerrar: <path d="M4 4l8 8M12 4l-8 8" strokeLinecap="round" />,
  mas: <><circle cx="3" cy="8" r="1.4" fill="currentColor" stroke="none" /><circle cx="8" cy="8" r="1.4" fill="currentColor" stroke="none" /><circle cx="13" cy="8" r="1.4" fill="currentColor" stroke="none" /></>,
  alerta: <><path d="M8 2.5L15 14H1L8 2.5z" strokeLinejoin="round" /><path d="M8 6.8v3.1M8 11.8v.1" strokeLinecap="round" /></>,
  ok: <path d="M3 8.4l3.2 3.2L13 5" strokeLinecap="round" strokeLinejoin="round" />,
  info: <><circle cx="8" cy="8" r="6.4" /><path d="M8 7.4v4M8 4.9v.1" strokeLinecap="round" /></>,
  carpeta: <path d="M2 4.2h4.2l1.4 1.8H14v7.6H2V4.2z" strokeLinejoin="round" />,
  listo: <><circle cx="8" cy="8" r="6.4" /><path d="M5.2 8.2l1.9 1.9L11 6.2" strokeLinecap="round" strokeLinejoin="round" /></>,
  atras: <path d="M9.5 3.5L5 8l4.5 4.5" strokeLinecap="round" strokeLinejoin="round" />,
  descarga: <><path d="M8 2.5v8" strokeLinecap="round" /><path d="M4.6 7.4L8 10.8l3.4-3.4" strokeLinecap="round" strokeLinejoin="round" /><path d="M2.6 13h10.8" strokeLinecap="round" /></>,
  sube: <><path d="M8 13.5v-8" strokeLinecap="round" /><path d="M4.6 8.6L8 5.2l3.4 3.4" strokeLinecap="round" strokeLinejoin="round" /><path d="M2.6 3h10.8" strokeLinecap="round" /></>,
  deshacer: <><path d="M3 8h7.2a3 3 0 010 6H7" strokeLinecap="round" strokeLinejoin="round" /><path d="M5.6 5.4L3 8l2.6 2.6" strokeLinecap="round" strokeLinejoin="round" /></>,
};

function Svg({ d, size = 16, ...rest }) {
  return (
    <svg viewBox="0 0 16 16" width={size} height={size} fill="none" stroke="currentColor" strokeWidth="1.6" aria-hidden="true" {...rest}>
      {d}
    </svg>
  );
}

function Badge({ tono = "neutro", children, ...rest }) {
  return <span className={`badge badge-${tono}`} {...rest}>{children}</span>;
}

/* Menú accesible: cierra con Escape, clic afuera o al elegir una opción. */
function Menu({ etiqueta, children, alineacion = "der", disparador }) {
  const [abierto, setAbierto] = useState(false);
  const cont = useRef(null);

  useEffect(() => {
    if (!abierto) return;
    const fuera = (e) => { if (cont.current && !cont.current.contains(e.target)) setAbierto(false); };
    const esc = (e) => { if (e.key === "Escape") setAbierto(false); };
    document.addEventListener("mousedown", fuera);
    document.addEventListener("keydown", esc);
    return () => { document.removeEventListener("mousedown", fuera); document.removeEventListener("keydown", esc); };
  }, [abierto]);

  return (
    <div className="menu-cont" ref={cont}>
      {disparador ? (
        disparador(() => setAbierto((v) => !v), abierto)
      ) : (
        <button className="btn btn-icono" aria-haspopup="menu" aria-expanded={abierto} aria-label={etiqueta} onClick={() => setAbierto((v) => !v)}>
          <Svg d={Ico.mas} />
        </button>
      )}
      {abierto && (
        <div className={alineacion === "izq" ? "menu izq" : "menu"} role="menu" onClick={() => setAbierto(false)}>
          {children}
        </div>
      )}
    </div>
  );
}

function Modal({ titulo, children, onCerrar, acciones }) {
  useEffect(() => {
    const esc = (e) => e.key === "Escape" && onCerrar();
    document.addEventListener("keydown", esc);
    return () => document.removeEventListener("keydown", esc);
  }, [onCerrar]);
  return (
    <div className="modal-velo" onClick={onCerrar}>
      <div className="modal" role="dialog" aria-modal="true" aria-label={titulo} onClick={(e) => e.stopPropagation()}>
        <h3>{titulo}</h3>
        <div>{children}</div>
        <div className="acc">{acciones}</div>
      </div>
    </div>
  );
}

/* Notificaciones efímeras. La reversión vive acá, junto a la acción que la produjo. */
function Avisos({ lista, quitar }) {
  return (
    <div className="avisos" aria-live="polite" aria-atomic="false">
      {lista.map((t) => (
        <div key={t.id} className={t.tono === "error" ? "toast err" : "toast"}>
          <span className="txt">{t.texto}</span>
          {t.accion && (
            <button className="btn btn-sm" onClick={() => { t.accion.fn(); quitar(t.id); }}>
              <Svg d={Ico.deshacer} size={13} /> {t.accion.label}
            </button>
          )}
          <button className="btn btn-sm btn-icono" aria-label="Cerrar aviso" onClick={() => quitar(t.id)}>
            <Svg d={Ico.cerrar} size={12} />
          </button>
        </div>
      ))}
    </div>
  );
}

function BotonArchivo({ onFile, children, variante = "", cargando }) {
  const ref = useRef(null);
  return (
    <>
      <button className={`btn ${variante}`} disabled={cargando} onClick={() => ref.current?.click()}>
        {cargando ? <><span className="spin" /> Procesando…</> : children}
      </button>
      <input ref={ref} type="file" accept=".xls,.xlsx" style={{ display: "none" }}
        onChange={(e) => { if (e.target.files[0]) onFile(e.target.files[0]); e.target.value = ""; }} />
    </>
  );
}

function Vacio({ icono, titulo, texto, acciones }) {
  return (
    <div className="vacio">
      {icono && <Svg d={icono} size={40} className="icono" strokeWidth="1.2" />}
      <h3>{titulo}</h3>
      <p>{texto}</p>
      {acciones && <div className="acc">{acciones}</div>}
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════
   DATOS
   ═══════════════════════════════════════════════════════════ */

function construirDatos(filas) {
  const porSocio = new Map();
  for (const f of filas) {
    const socio = Number(f.SOCIO);
    if (!socio) continue;
    if (!porSocio.has(socio)) {
      porSocio.set(socio, {
        socio, nombre: String(f.NOMBRE || "").trim(),
        division: String(f.DIVISION || "").trim(),
        categoria: String(f.CATEGORIA_SOCIAL || "").trim(),
        jefefam: Number(f.JEFEFAM) || socio,
        actividad: "SIN ACTIVIDAD", _dep: {}, _div: new Set(),
        deuda: 0, vieja: 0, corriente: 0, meses: 0, tienePlan: false, detalle: [],
      });
    }
    const s = porSocio.get(socio);
    const dep = String(f.DEPORTE || "").trim();
    const div = String(f.DIVISION || "").trim();
    if (dep) s._dep[dep] = (s._dep[dep] || 0) + 1;
    if (div) s._div.add(div.toUpperCase());
    const d = Number(f.DEUDA) || 0;
    const per = Number(f.PERIODO) || 0;
    s.deuda += d;
    if (per >= CORTE_ANIO_ACTUAL) s.corriente += d; else s.vieja += d;
    if (String(f.DESCRI_CONCEPTO_LIQ || "").toUpperCase().includes("PLAN DE PAGOS")) s.tienePlan = true;
    s.detalle.push({ periodo: per, concepto: String(f.DESCRI_CONCEPTO_LIQ || "").trim(), deuda: d });
  }

  const socios = [...porSocio.values()];
  for (const s of socios) {
    s.detalle.sort((a, b) => b.periodo - a.periodo);
    s.meses = new Set(s.detalle.filter((d) => d.deuda > 0).map((d) => d.periodo)).size;
    if (s._div.has("MORENO FANS")) s.actividad = "MORENO FANS";
    else {
      const top = Object.entries(s._dep).sort((a, b) => b[1] - a[1])[0];
      s.actividad = top ? ALIAS_ACTIVIDAD[top[0]] || top[0] : "SIN ACTIVIDAD";
    }
    delete s._dep; delete s._div;
  }

  const porFam = new Map();
  for (const s of socios) {
    if (!porFam.has(s.jefefam)) porFam.set(s.jefefam, { jefefam: s.jefefam, miembros: [], deuda: 0, vieja: 0, corriente: 0 });
    const f = porFam.get(s.jefefam);
    f.miembros.push(s); f.deuda += s.deuda; f.vieja += s.vieja; f.corriente += s.corriente;
  }
  const familias = [...porFam.values()];
  for (const f of familias) {
    f.miembros.sort((a, b) => b.deuda - a.deuda);
    f.titular = f.miembros.find((m) => m.socio === f.jefefam) || f.miembros[0];
    f.meses = Math.max(...f.miembros.map((m) => m.meses));
    f.actividades = [...new Set(f.miembros.map((m) => m.actividad))].sort((a, b) => ordenAct(a) - ordenAct(b));
  }
  return { socios, familias };
}

const norm = (t) => String(t).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/[^A-Z0-9]/g, "");

function hallarCol(cab, candidatas) {
  return cab.find((c) => candidatas.some((k) => norm(c).includes(norm(k))));
}

function leerCobranzas(filas) {
  if (!filas.length) return { error: "El archivo no tiene ninguna fila de datos." };
  const cab = Object.keys(filas[0]);
  const colSocio = hallarCol(cab, ["SOCIO", "NROSOCIO", "NUMEROSOCIO", "LEGAJO"]);
  const colImporte = hallarCol(cab, ["COBRADO", "IMPORTE", "MONTO", "PAGO", "ABONADO", "TOTAL"]);
  if (!colSocio || !colImporte) {
    const falta = !colSocio && !colImporte ? "la del socio y la del importe" : !colSocio ? "la del número de socio" : "la del importe";
    return { error: `Falta ${falta}. El archivo trae estas columnas: ${cab.join(", ")}. Renombrá la columna correspondiente a SOCIO o IMPORTE y volvé a subirlo.` };
  }
  const pagos = new Map();
  for (const f of filas) {
    const socio = Number(String(f[colSocio]).replace(/\D/g, ""));
    const bruto = String(f[colImporte] ?? "").replace(/[^\d,.-]/g, "").replace(/\.(?=\d{3}\b)/g, "").replace(",", ".");
    const importe = parseFloat(bruto);
    if (!socio || !importe || importe <= 0) continue;
    pagos.set(socio, (pagos.get(socio) || 0) + importe);
  }
  if (!pagos.size) return { error: `Encontré las columnas "${colSocio}" e "${colImporte}" pero ninguna fila con un importe válido mayor a cero.` };
  return { pagos, colSocio, colImporte };
}

function exportarCobranzas(datos, gestion) {
  const fila = (s) => {
    const e = gestion[s.socio] || {};
    return {
      SOCIO: s.socio, NOMBRE: s.nombre, ACTIVIDAD: s.actividad, DIVISION: s.division,
      MESES_ADEUDADOS: s.meses, DEUDA: s.deuda, ESTADO: e.estado || "",
      FECHA_ACCION: fechaAccion(e), ULTIMO_RECLAMO: fechaLarga(e.ultimoReclamo), COBRADO: Number(e.pago) || 0,
    };
  };
  const cobranzas = datos.socios.filter((s) => Number((gestion[s.socio] || {}).pago) > 0).map(fila);
  const gestionados = datos.socios.filter((s) => (gestion[s.socio] || {}).estado).map(fila);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cobranzas), "COBRANZAS");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(gestionados), "GESTION");
  XLSX.writeFile(wb, `cobranzas_${new Date().toISOString().slice(0, 10)}.xlsx`);
  return cobranzas.length;
}

/* ═══════════════════════════════════════════════════════════
   BUSCADOR GLOBAL
   ═══════════════════════════════════════════════════════════ */

function Buscador({ socios, abrir }) {
  const [q, setQ] = useState("");
  const [foco, setFoco] = useState(0);
  const input = useRef(null);

  useEffect(() => {
    const atajo = (e) => {
      if (e.key === "/" && !["INPUT", "TEXTAREA"].includes(e.target?.tagName)) { e.preventDefault(); input.current?.focus(); }
    };
    window.addEventListener("keydown", atajo);
    return () => window.removeEventListener("keydown", atajo);
  }, []);

  const res = useMemo(() => {
    const t = q.trim().toLowerCase();
    if (t.length < 2) return [];
    const palabras = t.split(/\s+/);
    return socios
      .filter((s) => palabras.every((p) => s.nombre.toLowerCase().includes(p) || String(s.socio).startsWith(p)))
      .sort((a, b) => b.deuda - a.deuda).slice(0, 12);
  }, [q, socios]);

  const elegir = (socio) => { abrir(socio); setQ(""); setFoco(0); input.current?.blur(); };

  return (
    <div className="buscador">
      <div className="campo-busca">
        <Svg d={Ico.busca} />
        <input
          ref={input} type="text" value={q}
          placeholder="Buscar socio por nombre, apellido o número"
          aria-label="Buscar socio por nombre, apellido o número"
          onChange={(e) => { setQ(e.target.value); setFoco(0); }}
          onKeyDown={(e) => {
            if (e.key === "Escape") { setQ(""); e.currentTarget.blur(); return; }
            if (!res.length) return;
            if (e.key === "ArrowDown") { e.preventDefault(); setFoco((f) => Math.min(f + 1, res.length - 1)); }
            if (e.key === "ArrowUp") { e.preventDefault(); setFoco((f) => Math.max(f - 1, 0)); }
            if (e.key === "Enter") { e.preventDefault(); elegir(res[foco].socio); }
          }}
        />
        {q ? (
          <button className="btn btn-3 btn-sm btn-icono" aria-label="Limpiar búsqueda" onClick={() => { setQ(""); setFoco(0); }}>
            <Svg d={Ico.cerrar} size={12} />
          </button>
        ) : (
          <kbd className="oculta-sm">/</kbd>
        )}
      </div>

      {q.trim().length >= 2 && (
        <div className="resultados" role="listbox">
          {!res.length ? (
            <div className="r-vacio">Ningún socio con saldo coincide con «{q.trim()}». Probá con el apellido solo o con el número de socio.</div>
          ) : res.map((s, i) => (
            <button key={s.socio} role="option" aria-selected={i === foco} className={i === foco ? "activo" : ""}
              onMouseEnter={() => setFoco(i)} onClick={() => elegir(s.socio)}>
              <span style={{ minWidth: 0 }}>
                <span className="r-nom">{s.nombre}</span>
                <span className="r-meta">Socio {s.socio} · {s.actividad}</span>
              </span>
              <span className="r-imp num">{money(s.deuda)}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════
   CARGA DE IMPORTE
   ═══════════════════════════════════════════════════════════ */

/* ── Línea de historial editable ──────────────────────────────
   Secretaría puede corregir o borrar un registro. Todos los demás
   sólo lo ven. El borrado pide confirmación en el lugar, sin abrir
   un modal aparte, porque es una corrección puntual y frecuente. */

function LineaLog({ l, esSecretaria, onEditar, onEliminar }) {
  const [editando, setEditando] = useState(false);
  const [confirmando, setConfirmando] = useState(false);
  const [estado, setEstado] = useState(l.estado || "");
  const [monto, setMonto] = useState(l.monto ? String(l.monto) : "");
  const [fecha, setFecha] = useState(l.fecha ? l.fecha.slice(0, 10) : "");

  if (editando) {
    return (
      <div className="linea-log" style={{ flexWrap: "wrap", gap: "var(--s2)" }}>
        <input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} style={{ width: 140, minHeight: 30 }} />
        <select value={estado} onChange={(e) => setEstado(e.target.value)} style={{ width: 150, minHeight: 30 }}>
          {ESTADOS.map((e) => <option key={e.id} value={e.id}>{e.label}</option>)}
        </select>
        <span className="campo-plata" style={{ width: 120 }}>
          <span className="sig">$</span>
          <input type="text" inputMode="decimal" value={monto} placeholder="0"
            onChange={(e) => setMonto(e.target.value.replace(/[^\d.,]/g, "").replace(",", "."))} />
        </span>
        <button className="btn btn-sm btn-1" onClick={() => {
          onEditar(l.id, { estado, monto: parseFloat(monto) || null, fecha: new Date(`${fecha}T12:00:00`).toISOString() });
          setEditando(false);
        }}>Guardar</button>
        <button className="btn btn-sm" onClick={() => setEditando(false)}>Cancelar</button>
      </div>
    );
  }

  return (
    <div className="linea-log">
      <span className="f">{fechaLarga(l.fecha)}</span>
      <span>
        {ESTADO_MAP[l.estado]?.label || l.estado}
        {l.monto > 0 && <span className="num" style={{ color: "var(--exito)", marginLeft: 6 }}>{money(l.monto)}</span>}
      </span>
      <span className="t-cap" style={{ marginLeft: "auto" }}>{l.por}</span>
      {esSecretaria && (
        confirmando ? (
          <span style={{ display: "flex", gap: 4 }}>
            <span className="t-cap" style={{ color: "var(--critico)" }}>¿Eliminar?</span>
            <button className="btn btn-sm btn-peligro" onClick={() => onEliminar(l.id)}>Sí</button>
            <button className="btn btn-sm" onClick={() => setConfirmando(false)}>No</button>
          </span>
        ) : (
          <span style={{ display: "flex", gap: 2 }}>
            <button className="btn btn-sm btn-3 btn-icono" aria-label="Corregir este registro" onClick={() => setEditando(true)}>✎</button>
            <button className="btn btn-sm btn-3 btn-icono" aria-label="Eliminar este registro" onClick={() => setConfirmando(true)}>✕</button>
          </span>
        )
      )}
    </div>
  );
}

/* ── Carga de un reclamo o pago anterior al uso de la página ── */

function FormularioLogManual({ onGuardar, onCancelar }) {
  const [estado, setEstado] = useState("RECLAMADO");
  const [monto, setMonto] = useState("");
  const [fecha, setFecha] = useState(() => new Date().toISOString().slice(0, 10));
  const [actualizarEstado, setActualizarEstado] = useState(false);

  return (
    <div style={{ background: "var(--sup2)", borderRadius: "var(--r)", padding: "var(--s3)", marginTop: "var(--s2)" }}>
      <div style={{ display: "flex", gap: "var(--s2)", flexWrap: "wrap", alignItems: "flex-end" }}>
        <div className="campo" style={{ flex: "1 1 130px" }}>
          <label>Fecha</label>
          <input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} />
        </div>
        <div className="campo" style={{ flex: "1 1 150px" }}>
          <label>Qué pasó</label>
          <select value={estado} onChange={(e) => setEstado(e.target.value)}>
            {ESTADOS.map((e) => <option key={e.id} value={e.id}>{e.label}</option>)}
          </select>
        </div>
        <div className="campo" style={{ flex: "0 0 130px" }}>
          <label>Importe (si aplica)</label>
          <div className="campo-plata">
            <span className="sig">$</span>
            <input type="text" inputMode="decimal" value={monto} placeholder="0"
              onChange={(e) => setMonto(e.target.value.replace(/[^\d.,]/g, "").replace(",", "."))} />
          </div>
        </div>
      </div>
      <label style={{ display: "flex", gap: 6, alignItems: "center", marginTop: "var(--s3)", fontSize: 12.5, color: "var(--n700)", cursor: "pointer" }}>
        <input type="checkbox" checked={actualizarEstado} onChange={(e) => setActualizarEstado(e.target.checked)} style={{ minHeight: "auto", width: "auto" }} />
        Además, dejar esto como el estado actual de la cuenta
      </label>
      <div style={{ display: "flex", gap: "var(--s2)", marginTop: "var(--s3)" }}>
        <button className="btn btn-1 btn-sm" onClick={() => onGuardar({ estado, monto: parseFloat(monto) || null, fecha, actualizarEstado })}>
          Agregar al historial
        </button>
        <button className="btn btn-sm" onClick={onCancelar}>Cancelar</button>
      </div>
    </div>
  );
}

function CargaCobro({ miembros, onGuardar, onCancelar }) {
  const [montos, setMontos] = useState(() => Object.fromEntries(miembros.map((m) => [m.socio, ""])));
  const [tocado, setTocado] = useState({});
  const valor = (s) => parseFloat(montos[s]) || 0;
  const total = miembros.reduce((a, m) => a + valor(m.socio), 0);
  const errorDe = (m) => (tocado[m.socio] && valor(m.socio) > m.deuda ? `Supera la deuda de ${money(m.deuda)}. Si el socio pagó de más, marcalo igual y revisá la cuenta después.` : "");

  return (
    <div className="panel-sec" style={{ background: "var(--sup2)" }}>
      <h4>Registrar pago</h4>
      {miembros.map((m) => (
        <div className="fila-cobro" key={m.socio}>
          <div className="quien">
            <div className="t-strong">{m.nombre}</div>
            <div className="t-cap">Debe {money(m.deuda)} · {m.meses} {m.meses === 1 ? "mes" : "meses"}</div>
          </div>
          <div className="campo">
            <label htmlFor={`imp-${m.socio}`}>Importe abonado</label>
            <div className={errorDe(m) ? "campo-plata mal" : "campo-plata"}>
              <span className="sig">$</span>
              <input id={`imp-${m.socio}`} type="text" inputMode="decimal" value={montos[m.socio]} placeholder="0"
                onBlur={() => setTocado((t) => ({ ...t, [m.socio]: true }))}
                onChange={(e) => {
                  const v = e.target.value.replace(/[^\d.,]/g, "").replace(",", ".");
                  setMontos((p) => ({ ...p, [m.socio]: v }));
                }} />
            </div>
            {errorDe(m)
              ? <span className="error"><Svg d={Ico.alerta} size={12} /> {errorDe(m)}</span>
              : <button type="button" className="ayuda" style={{ background: "none", border: 0, padding: 0, textAlign: "left", cursor: "pointer", textDecoration: "underline", textUnderlineOffset: 2 }}
                  onClick={() => setMontos((p) => ({ ...p, [m.socio]: String(m.deuda) }))}>
                  Completar con el total adeudado
                </button>}
          </div>
        </div>
      ))}
      <div style={{ display: "flex", gap: "var(--s2)", alignItems: "center", marginTop: "var(--s4)", flexWrap: "wrap" }}>
        <button className="btn btn-1" disabled={total <= 0}
          onClick={() => onGuardar(Object.fromEntries(miembros.map((m) => [m.socio, valor(m.socio)])))}>
          Guardar pago de {money(total)}
        </button>
        <button className="btn" onClick={onCancelar}>Cancelar</button>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════
   VISTA: COLA DE RECLAMOS
   ═══════════════════════════════════════════════════════════ */

function FilaFamilia({ fam, actividad, marcar, g, abrir }) {
  const [cobrando, setCobrando] = useState(false);
  const dias = diasDesde(fam.ultimoReclamo);
  const socios = fam.miembros.map((m) => m.socio);
  const otras = actividad ? fam.actividades.filter((a) => a !== actividad) : [];
  const varios = fam.miembros.length > 1;
  const cobrado = fam.miembros.reduce((a, m) => a + (Number(g(m.socio).pago) || 0), 0);

  return (
    <article className="fila-fam">
      <div className="fam-cuerpo">
        <div className="fam-info">
          <div className="fam-tit">
            <button className="nom" style={{ background: "none", border: 0, padding: 0, font: "inherit", fontSize: 15, fontWeight: 600, cursor: "pointer" }}
              onClick={() => abrir(fam.titular.socio)}>
              {varios ? `Familia ${fam.titular.nombre.split(" ")[0]}` : fam.titular.nombre}
            </button>
            {fam.meses > MESES_ALERTA && <Badge tono="critico">{fam.meses} meses sin pagar</Badge>}
            {cobrado > 0 && <Badge tono="exito">{money(cobrado)} cobrado</Badge>}
          </div>
          <div className="fam-sub">
            <span>{varios ? `${fam.miembros.length} socios · un solo reclamo` : `Socio ${fam.titular.socio} · ${fam.titular.division}`}</span>
            <span className="punto">·</span>
            <span>{dias === null ? "Nunca reclamado" : `Último reclamo hace ${dias} ${dias === 1 ? "día" : "días"}`}</span>
            {otras.length > 0 && (<><span className="punto">·</span><span style={{ color: "var(--aviso)" }}>También en {otras.join(" y ")}</span></>)}
          </div>

          {varios && (
            <div className="hermanos">
              {fam.miembros.map((m) => {
                const est = g(m.socio).estado;
                return (
                  <div className="hermano" key={m.socio}>
                    <button className="h-nom" onClick={() => abrir(m.socio)}>{m.nombre}</button>
                    <span className="t-cap">{m.actividad} · {m.division}</span>
                    {est && <Badge tono={ESTADO_MAP[est].tono}>{ESTADO_MAP[est].label}</Badge>}
                    <span className="h-imp num apagado">{money(m.deuda)}</span>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        <div className="fam-plata">
          <div className="num-xl">{money(fam.deuda)}</div>
          <div className="sub">
            {fam.vieja > 0 ? <>{moneyCorto(fam.corriente)} de 2026 · {moneyCorto(fam.vieja)} de años anteriores</> : "Todo del año en curso"}
          </div>
        </div>
      </div>

      {cobrando ? (
        <CargaCobro miembros={fam.miembros} onCancelar={() => setCobrando(false)}
          onGuardar={(montos) => {
            setCobrando(false);
            const conPago = fam.miembros.filter((m) => montos[m.socio] > 0).map((m) => m.socio);
            marcar(conPago, null, (s) => {
              const m = fam.miembros.find((x) => x.socio === s);
              return { pago: montos[s], saldoAlPagar: m ? m.deuda : 0, estadoAuto: montos[s] >= (m?.deuda || 0) ? "PAGO" : "PAGO PARCIAL" };
            });
          }} />
      ) : (
        <div className="fam-acc">
          <button className="btn btn-1" onClick={() => marcar(socios, "RECLAMADO")}>Registrar reclamo</button>
          <button className="btn" onClick={() => setCobrando(true)}>Registrar pago</button>
          <div className="der">
            <button className="btn btn-3" onClick={() => abrir(fam.titular.socio)}>Ver cuenta</button>
            <Menu etiqueta="Más acciones para esta familia">
              <div className="titulo-menu">Sacar de la cola</div>
              <button role="menuitem" onClick={() => marcar(socios, "AL DIA")}>Está al día</button>
              <button role="menuitem" onClick={() => marcar(socios, "PLAN DE PAGOS")}>Tiene plan de pagos</button>
              <button role="menuitem" onClick={() => marcar(socios, "BAJA/RENUNCIA")}>Dio de baja</button>
              <button role="menuitem" onClick={() => marcar(socios, "NO RECLAMAR")}>No reclamar</button>
              <hr />
              <button role="menuitem" onClick={() => marcar(socios, "PEDIR NUMERO")}>Falta el teléfono</button>
              <button role="menuitem" onClick={() => marcar(socios, "MIRAR CUENTA")}>Revisar la cuenta</button>
            </Menu>
          </div>
        </div>
      )}
    </article>
  );
}

function VistaCola({ cola, actividad, marcar, g, abrir, hayFiltro, limpiar }) {
  if (!cola.length) {
    return (
      <Vacio icono={Ico.listo}
        titulo={hayFiltro ? `No queda nadie por reclamar en ${actividad}` : "No queda nadie por reclamar"}
        texto={hayFiltro
          ? "Ya gestionaste todos los grupos de esta actividad. Mirá otra o sacá el filtro para ver el resto."
          : `Gestionaste todos los grupos con deuda. Los reclamados vuelven a aparecer cuando pasen ${DIAS_ENTRE_RECLAMOS} días.`}
        acciones={hayFiltro && <button className="btn" onClick={limpiar}>Ver todas las actividades</button>} />
    );
  }
  return (
    <div className="cola">
      {cola.map((f) => <FilaFamilia key={f.jefefam} fam={f} actividad={actividad} marcar={marcar} g={g} abrir={abrir} />)}
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════
   VISTA: SOCIOS
   ═══════════════════════════════════════════════════════════ */

const COLUMNAS = [
  { id: "socio",  label: "Socio",       ancho: 78,  orden: (a, b) => a.socio - b.socio, clase: "oculta-sm" },
  { id: "nombre", label: "Nombre",      orden: (a, b) => a.nombre.localeCompare(b.nombre) },
  { id: "estado", label: "Estado" },
  { id: "meses",  label: "Meses",       orden: (a, b) => a.meses - b.meses, centro: true },
  { id: "fecha",  label: "Últ. acción", clase: "oculta-md" },
  { id: "deuda",  label: "Deuda",       orden: (a, b) => a.deuda - b.deuda, num: true },
];

function VistaSocios({ lista, g, abrir, busqueda, setBusqueda, filtroEstado, setFiltroEstado, limpiarTodo }) {
  const [orden, setOrden] = useState({ col: "deuda", desc: true });
  const [pagina, setPagina] = useState(0);

  useEffect(() => setPagina(0), [lista.length, orden]);

  const ordenada = useMemo(() => {
    const c = COLUMNAS.find((x) => x.id === orden.col);
    if (!c?.orden) return lista;
    const arr = [...lista].sort(c.orden);
    return orden.desc ? arr.reverse() : arr;
  }, [lista, orden]);

  const paginas = Math.max(1, Math.ceil(ordenada.length / POR_PAGINA));
  const visible = ordenada.slice(pagina * POR_PAGINA, (pagina + 1) * POR_PAGINA);

  const ordenar = (id) => setOrden((o) => (o.col === id ? { col: id, desc: !o.desc } : { col: id, desc: true }));

  return (
    <div className="caja-tabla">
      <div className="barra-tabla">
        <div style={{ flex: "1 1 240px", maxWidth: 320 }}>
          <div className="campo-busca">
            <Svg d={Ico.busca} />
            <input type="text" value={busqueda} placeholder="Filtrar esta lista" aria-label="Filtrar la lista de socios"
              onChange={(e) => setBusqueda(e.target.value)} />
          </div>
        </div>
        <select value={filtroEstado} onChange={(e) => setFiltroEstado(e.target.value)} aria-label="Filtrar por estado" style={{ width: "auto", minWidth: 160 }}>
          <option value="">Todos los estados</option>
          <option value="__sin">Sin gestionar</option>
          {ESTADOS.map((e) => <option key={e.id} value={e.id}>{e.label}</option>)}
        </select>
        <span className="t-cap" style={{ marginLeft: "auto" }}>
          {ordenada.length === 1 ? "1 socio" : `${ordenada.length} socios`}
        </span>
      </div>

      {!ordenada.length ? (
        <Vacio icono={Ico.busca} titulo="Ningún socio coincide con estos filtros"
          texto="Probá con menos filtros o revisá cómo escribiste el nombre."
          acciones={<button className="btn" onClick={limpiarTodo}>Limpiar filtros</button>} />
      ) : (
        <>
          <table>
            <thead>
              <tr>
                {COLUMNAS.map((c) => (
                  <th key={c.id} className={c.clase} style={{ width: c.ancho, textAlign: c.num ? "right" : c.centro ? "center" : "left" }}
                    aria-sort={orden.col === c.id ? (orden.desc ? "descending" : "ascending") : undefined}>
                    {c.orden ? (
                      <button onClick={() => ordenar(c.id)}>
                        {c.label}
                        {orden.col === c.id && <span className="flecha">{orden.desc ? "▼" : "▲"}</span>}
                      </button>
                    ) : c.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {visible.map((s) => {
                const e = g(s.socio);
                const est = e.estado;
                return (
                  <tr key={s.socio} onClick={() => abrir(s.socio)} tabIndex={0}
                    onKeyDown={(ev) => { if (ev.key === "Enter") abrir(s.socio); }}>
                    <td className="oculta-sm num apagado" style={{ fontSize: 12 }}>{s.socio}</td>
                    <td>
                      <div className="t-strong">{s.nombre}</div>
                      <div className="t-cap">{s.actividad} · {s.division}{s.tienePlan && " · plan de pagos"}</div>
                    </td>
                    <td>{est ? <Badge tono={ESTADO_MAP[est].tono}>{ESTADO_MAP[est].label}</Badge> : <Badge tono="vacio">Sin gestionar</Badge>}</td>
                    <td style={{ textAlign: "center" }}>
                      {s.meses > MESES_ALERTA
                        ? <Badge tono="critico" title={`Adeuda ${s.meses} períodos`}>{s.meses}</Badge>
                        : <span className="num apagado">{s.meses}</span>}
                    </td>
                    <td className="oculta-md num" style={{ fontSize: 12 }}>
                      {fechaAccion(e) ? <span>{fechaAccion(e)}</span> : <span className="vacio-celda">—</span>}
                    </td>
                    <td className="celda-num">
                      <div style={{ fontWeight: 600 }}>{money(s.deuda)}</div>
                      {e.pago > 0 && <div className="t-cap" style={{ color: "var(--exito)" }}>{money(e.pago)} cobrado</div>}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>

          {paginas > 1 && (
            <div className="paginacion">
              <span className="t-cap">
                {pagina * POR_PAGINA + 1}–{Math.min((pagina + 1) * POR_PAGINA, ordenada.length)} de {ordenada.length}
              </span>
              <div style={{ display: "flex", gap: "var(--s2)", alignItems: "center" }}>
                <button className="btn btn-sm" disabled={pagina === 0} onClick={() => setPagina((p) => p - 1)}>Anterior</button>
                <span className="t-cap">Página {pagina + 1} de {paginas}</span>
                <button className="btn btn-sm" disabled={pagina >= paginas - 1} onClick={() => setPagina((p) => p + 1)}>Siguiente</button>
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════
   VISTA: RESUMEN
   ═══════════════════════════════════════════════════════════ */

function VistaResumen({ filas, totales }) {
  const cols = ["RECLAMAR", "RECLAMADO", "PAGO", "PAGO PARCIAL", "AL DIA", "BAJA/RENUNCIA", "__sin"];
  const etiqueta = (c) => (c === "__sin" ? "Sin gestionar" : ESTADO_MAP[c].label);
  return (
    <>
      <div className="kpis">
        <div className="kpi kpi-1">
          <div className="k">Deuda total</div>
          <div className="v num-xl">{money(totales.deuda)}</div>
          <div className="s">{totales.socios} socios · {totales.familias} grupos familiares</div>
        </div>
        <div className="kpi">
          <div className="k">Cobrado</div>
          <div className="v num-xl" style={{ color: totales.cobrado ? "var(--exito)" : undefined }}>{money(totales.cobrado)}</div>
          <div className="s">{totales.deuda ? ((totales.cobrado / totales.deuda) * 100).toFixed(1) : 0}% de la deuda</div>
        </div>
        <div className="kpi">
          <div className="k">Socios con más de 3 meses de deuda</div>
          <div className="v num-xl">{totales.atrasados}</div>
          <div className="s">de un total de {totales.socios}</div>
        </div>
        <div className="kpi">
          <div className="k">Gestionados</div>
          <div className="v num-xl">{totales.gestionados}</div>
          <div className="s">de {totales.socios} socios</div>
        </div>
      </div>

      <div className="caja-tabla">
        <table>
          <thead>
            <tr>
              <th>Actividad</th>
              <th style={{ textAlign: "right" }}>Socios</th>
              <th style={{ textAlign: "right" }}>Deuda</th>
              {cols.map((c) => <th key={c} className="oculta-md" style={{ textAlign: "right" }}>{etiqueta(c)}</th>)}
              <th style={{ textAlign: "right" }}>Cobrado</th>
            </tr>
          </thead>
          <tbody>
            {filas.map((f) => (
              <tr key={f.actividad} style={{ cursor: "default" }}>
                <td className="t-strong">{f.actividad}</td>
                <td className="celda-num apagado">{f.socios}</td>
                <td className="celda-num" style={{ fontWeight: 600 }}>{money(f.deuda)}</td>
                {cols.map((c) => (
                  <td key={c} className="oculta-md celda-num">
                    {f.conteo[c] ? f.conteo[c] : <span className="vacio-celda">—</span>}
                  </td>
                ))}
                <td className="celda-num" style={{ color: f.cobrado ? "var(--exito)" : undefined }}>
                  {f.cobrado ? money(f.cobrado) : <span className="vacio-celda">—</span>}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  );
}

/* ═══════════════════════════════════════════════════════════
   PANEL DE CUENTA
   ═══════════════════════════════════════════════════════════ */

function Panel({ socio, familia, estado, esSecretaria, marcar, comentar, editarLog, eliminarLog, agregarLogManual, cerrar }) {
  const [texto, setTexto] = useState("");
  const [cobrando, setCobrando] = useState(false);
  const [cargandoManual, setCargandoManual] = useState(false);
  const panel = useRef(null);

  useEffect(() => {
    panel.current?.focus();
    const esc = (e) => e.key === "Escape" && cerrar();
    document.addEventListener("keydown", esc);
    return () => document.removeEventListener("keydown", esc);
  }, [cerrar]);

  const porPeriodo = useMemo(() => {
    const m = new Map();
    for (const d of socio.detalle) {
      if (!m.has(d.periodo)) m.set(d.periodo, []);
      m.get(d.periodo).push(d);
    }
    return [...m.entries()].sort((a, b) => b[0] - a[0]);
  }, [socio]);

  const est = estado.estado;

  return (
    <>
      <div className="velo" onClick={cerrar} />
      <aside className="panel" ref={panel} tabIndex={-1} role="dialog" aria-modal="true" aria-label={`Cuenta de ${socio.nombre}`}>
        <div className="panel-barra">
          <button className="btn btn-3 btn-icono" onClick={cerrar} aria-label="Volver al listado"><Svg d={Ico.atras} /></button>
          <span className="ruta">Socios / {socio.nombre}</span>
          <button className="btn btn-3 btn-icono" onClick={cerrar} aria-label="Cerrar panel"><Svg d={Ico.cerrar} /></button>
        </div>

        <div className="panel-sec">
          <div className="t-h1">{socio.nombre}</div>
          <div className="t-cap" style={{ marginTop: 2 }}>Socio {socio.socio} · {socio.actividad} · {socio.division} · {socio.categoria}</div>

          <div className="resumen-socio" style={{ marginTop: "var(--s5)" }}>
            <div>
              <div className="t-cap">Deuda</div>
              <div className="num-xl">{money(socio.deuda)}</div>
            </div>
            <div>
              <div className="t-cap">Períodos impagos</div>
              <div className="num-xl">{socio.meses}</div>
            </div>
            {estado.pago > 0 && (
              <div>
                <div className="t-cap">Cobrado</div>
                <div className="num-xl" style={{ color: "var(--exito)" }}>{money(estado.pago)}</div>
              </div>
            )}
          </div>

          <div style={{ display: "flex", gap: "var(--s2)", marginTop: "var(--s4)", flexWrap: "wrap" }}>
            {est && <Badge tono={ESTADO_MAP[est].tono}>{ESTADO_MAP[est].label}</Badge>}
            {socio.meses > MESES_ALERTA && <Badge tono="critico">Más de 3 meses de deuda</Badge>}
            {socio.tienePlan && <Badge tono="info">Plan de pagos</Badge>}
            {fechaAccion(estado) && <span className="t-cap">Última acción el {fechaAccion(estado)}</span>}
          </div>

          {familia && familia.miembros.length > 1 && (
            <div className="aviso warn" style={{ marginTop: "var(--s4)" }}>
              <Svg d={Ico.info} />
              <span>
                Comparte grupo familiar con {familia.miembros.filter((m) => m.socio !== socio.socio).map((m) => m.nombre).join(", ")}.
                El reclamo va una sola vez a la familia.
              </span>
            </div>
          )}
        </div>

        {esSecretaria && (cobrando ? (
          <CargaCobro miembros={[socio]} onCancelar={() => setCobrando(false)}
            onGuardar={(m) => {
              setCobrando(false);
              marcar([socio.socio], null, () => ({
                pago: m[socio.socio], saldoAlPagar: socio.deuda,
                estadoAuto: m[socio.socio] >= socio.deuda ? "PAGO" : "PAGO PARCIAL",
              }));
            }} />
        ) : (
          <div className="panel-sec">
            <h4>Acciones</h4>
            <div style={{ display: "flex", gap: "var(--s2)", flexWrap: "wrap" }}>
              <button className="btn btn-1" onClick={() => marcar([socio.socio], "RECLAMADO")}>Registrar reclamo</button>
              <button className="btn" onClick={() => setCobrando(true)}>Registrar pago</button>
              <Menu etiqueta="Cambiar el estado de esta cuenta">
                <div className="titulo-menu">Marcar como</div>
                {ESTADOS.filter((e) => !["RECLAMADO", "PAGO", "PAGO PARCIAL"].includes(e.id)).map((e) => (
                  <button key={e.id} role="menuitem" onClick={() => marcar([socio.socio], e.id)}
                    style={est === e.id ? { fontWeight: 600 } : undefined}>
                    {e.label}{est === e.id && " ✓"}
                  </button>
                ))}
              </Menu>
            </div>
          </div>
        ))}

        <div className="panel-sec">
          <h4>Saldo detallado</h4>
          <table className="detalle">
            <tbody>
              {porPeriodo.map(([per, items]) => (
                <React.Fragment key={per}>
                  <tr>
                    <td className="per">{periodoLabel(per)}</td>
                    <td className="per imp">{money(items.reduce((a, i) => a + i.deuda, 0))}</td>
                  </tr>
                  {items.map((i, k) => (
                    <tr key={k}>
                      <td className="apagado" style={{ paddingLeft: "var(--s3)" }}>{i.concepto}</td>
                      <td className="imp apagado">{money(i.deuda)}</td>
                    </tr>
                  ))}
                </React.Fragment>
              ))}
            </tbody>
          </table>
        </div>

        <div className="panel-sec">
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "var(--s3)" }}>
            <h4 style={{ margin: 0 }}>Historial</h4>
            {esSecretaria && !cargandoManual && (
              <button className="btn btn-sm btn-3" onClick={() => setCargandoManual(true)}>+ Reclamo o pago anterior</button>
            )}
          </div>

          {!estado.log?.length ? (
            <p className="t-cap" style={{ margin: 0 }}>Todavía no se registró ninguna acción sobre esta cuenta.</p>
          ) : estado.log.map((l) => (
            <LineaLog key={l.id ?? l.fecha} l={l} esSecretaria={esSecretaria} onEditar={editarLog} onEliminar={(id) => eliminarLog(id, socio.socio)} />
          ))}

          {cargandoManual && (
            <FormularioLogManual
              onCancelar={() => setCargandoManual(false)}
              onGuardar={(datos) => { agregarLogManual(socio.socio, datos); setCargandoManual(false); }}
            />
          )}
        </div>

        <div className="panel-sec">
          <h4>Comentarios</h4>
          {estado.comentarios?.map((c, i) => (
            <div className="comentario" key={i}>
              <div className="meta">{c.por} · {fechaLarga(c.fecha)}</div>
              {c.texto}
            </div>
          ))}
          <div className="campo" style={{ marginTop: "var(--s3)" }}>
            <label htmlFor="nuevo-com">Nuevo comentario</label>
            <textarea id="nuevo-com" value={texto} placeholder="Ej: llamé al padre, dijo que paga el viernes"
              onChange={(e) => setTexto(e.target.value)} />
            <span className="ayuda">Lo ve todo el equipo, incluida la comisión.</span>
          </div>
          <button className="btn" style={{ marginTop: "var(--s3)" }} disabled={!texto.trim()}
            onClick={() => { comentar(socio.socio, texto.trim()); setTexto(""); }}>
            Agregar comentario
          </button>
        </div>
      </aside>
    </>
  );
}

/* ═══════════════════════════════════════════════════════════
   APLICACIÓN
   ═══════════════════════════════════════════════════════════ */

const VISTAS = {
  cola:    { label: "Cola de reclamos", h1: "Cola de reclamos", desc: "Grupos familiares con deuda, ordenados por importe. Un reclamo por familia." },
  socios:  { label: "Socios",           h1: "Socios",           desc: "Todas las cuentas con saldo. Tocá una fila para ver el detalle por período." },
  resumen: { label: "Resumen",          h1: "Resumen",          desc: "Pagos y estado de las cuentas corrientes por actividad, para la Comisión Directiva." },
};

export default function App() {
  const [datos, setDatos] = useState(null);
  const [gestion, setGestion] = useState({});
  const [avisos, setAvisos] = useState([]);
  const [listo, setListo] = useState(false);
  const [procesando, setProcesando] = useState("");
  const [errorInicio, setErrorInicio] = useState("");
  const [rol, setRol] = useState("secretaria");
  const [vista, setVista] = useState("cola");
  const [actividad, setActividad] = useState("");
  const [abierto, setAbierto] = useState(null);
  const [busqueda, setBusqueda] = useState("");
  const [filtroEstado, setFiltroEstado] = useState("");
  const [confirmar, setConfirmar] = useState(null);

  const esSecretaria = rol === "secretaria";

  /* ── Avisos ── */
  const avisar = useCallback((texto, opts = {}) => {
    const id = Math.random().toString(36).slice(2);
    setAvisos((a) => [...a, { id, texto, ...opts }]);
    if (!opts.persistente) setTimeout(() => setAvisos((a) => a.filter((x) => x.id !== id)), opts.accion ? 9000 : 5000);
  }, []);
  const quitarAviso = useCallback((id) => setAvisos((a) => a.filter((x) => x.id !== id)), []);

  const cargarSaldosDesdeSupabase = useCallback(async () => {
    const filas = [];
    let desde = 0;
    // Supabase pagina de a 1000 filas por default; con ~3500 renglones hacen falta varias vueltas.
    while (true) {
      const { data, error } = await supabase.from("saldos").select("*").range(desde, desde + 999);
      if (error) { avisar("No pude traer los saldos guardados.", { tono: "error" }); return; }
      filas.push(...data);
      if (data.length < 1000) break;
      desde += 1000;
    }
    if (!filas.length) return;
    const crudo = filas.map((f) => ({
      SOCIO: f.socio, NOMBRE: f.nombre, DIVISION: f.division, CATEGORIA_SOCIAL: f.categoria_social,
      JEFEFAM: f.jefefam, DEPORTE: f.deporte, PERIODO: f.periodo, DESCRI_CONCEPTO_LIQ: f.descri_concepto_liq, DEUDA: f.deuda,
    }));
    setDatos(construirDatos(crudo));
  }, [avisar]);

  /* ── Persistencia: todo vive en Supabase, no en este navegador ──
     "gestion" trae el estado actual de cada socio. "gestion_log" y
     "comentarios" son historiales que se arman por socio en memoria
     al leerlos. Un canal de tiempo real avisa cuando alguien más
     (otra persona, en otra pestaña) cambia algo, y se refresca sola. */

  const cargarGestion = useCallback(async () => {
    const [{ data: filas, error: e1 }, { data: logs, error: e2 }, { data: coms, error: e3 }] = await Promise.all([
      supabase.from("gestion").select("*"),
      supabase.from("gestion_log").select("*").order("fecha", { ascending: false }).limit(2000),
      supabase.from("comentarios").select("*").order("fecha", { ascending: false }).limit(1000),
    ]);
    if (e1 || e2 || e3) { avisar("No pude conectar con la base de datos. Revisá tu conexión y recargá la página.", { tono: "error", persistente: true }); return; }

    const g = {};
    for (const f of filas || []) {
      g[f.socio] = { estado: f.estado, ultimoReclamo: f.ultimo_reclamo, pago: Number(f.pago) || 0, saldoAlPagar: Number(f.saldo_al_pagar) || 0, log: [], comentarios: [] };
    }
    for (const l of logs || []) {
      (g[l.socio] ??= { estado: null, ultimoReclamo: null, pago: 0, saldoAlPagar: 0, log: [], comentarios: [] })
        .log.push({ id: l.id, fecha: l.fecha, estado: l.estado, por: l.por, monto: l.monto ? Number(l.monto) : undefined });
    }
    for (const c of coms || []) {
      (g[c.socio] ??= { estado: null, ultimoReclamo: null, pago: 0, saldoAlPagar: 0, log: [], comentarios: [] })
        .comentarios.push({ fecha: c.fecha, por: c.por, texto: c.texto });
    }
    for (const socio of Object.keys(g)) g[socio].log.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
    setGestion(g);
  }, [avisar]);

  useEffect(() => {
    (async () => { await cargarGestion(); await cargarSaldosDesdeSupabase(); setListo(true); })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Tiempo real: cualquier cambio en estas tablas (propio o ajeno) recarga la gestión.
  useEffect(() => {
    const canal = supabase.channel("gestion-en-vivo")
      .on("postgres_changes", { event: "*", schema: "public", table: "gestion" }, cargarGestion)
      .on("postgres_changes", { event: "*", schema: "public", table: "gestion_log" }, cargarGestion)
      .on("postgres_changes", { event: "*", schema: "public", table: "comentarios" }, cargarGestion)
      .on("postgres_changes", { event: "*", schema: "public", table: "saldos" }, () => cargarSaldosDesdeSupabase())
      .subscribe();
    return () => supabase.removeChannel(canal);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [cargarGestion]);

  const g = useCallback((s) => gestion[s] || { estado: null, ultimoReclamo: null, pago: 0, log: [], comentarios: [] }, [gestion]);

  const marcar = useCallback(async (socios, estado, extra = {}) => {
    if (!socios.length) return;
    const ahora = new Date().toISOString();
    const quien = esSecretaria ? "Secretaría" : "Comisión";
    let ultimoEstado = estado;
    const previos = socios.map((s) => ({ socio: s, antes: g(s) })); // para poder deshacer sólo estos socios

    for (const s of socios) {
      const prev = g(s);
      const campos = typeof extra === "function" ? extra(s) : extra;
      const nuevoEstado = campos.estadoAuto || estado || prev.estado;
      ultimoEstado = nuevoEstado;
      const { estadoAuto, ...limpio } = campos;
      const pago = limpio.pago !== undefined ? limpio.pago : prev.pago;
      const saldoAlPagar = limpio.saldoAlPagar !== undefined ? limpio.saldoAlPagar : prev.saldoAlPagar;

      const { error: eG } = await supabase.from("gestion").upsert({
        socio: s, estado: nuevoEstado,
        ultimo_reclamo: nuevoEstado === "RECLAMADO" ? ahora : prev.ultimoReclamo,
        pago, saldo_al_pagar: saldoAlPagar, actualizado_en: ahora, actualizado_por: quien,
      });
      if (eG) { avisar("No se pudo guardar el cambio. Probá de nuevo en unos segundos.", { tono: "error" }); return; }

      await supabase.from("gestion_log").insert({ socio: s, fecha: ahora, estado: nuevoEstado, por: quien, monto: limpio.pago || null });
    }

    await cargarGestion();
    const que = ESTADO_MAP[ultimoEstado]?.label || ultimoEstado;
    const cuantos = socios.length > 1 ? ` en ${socios.length} socios` : "";
    avisar(`Se marcó «${que}»${cuantos}.`, {
      accion: {
        label: "Deshacer",
        fn: async () => {
          for (const { socio, antes } of previos) {
            await supabase.from("gestion").upsert({
              socio, estado: antes.estado, ultimo_reclamo: antes.ultimoReclamo,
              pago: antes.pago, saldo_al_pagar: antes.saldoAlPagar, actualizado_en: new Date().toISOString(), actualizado_por: quien,
            });
            await supabase.from("gestion_log").insert({ socio, fecha: new Date().toISOString(), estado: antes.estado || "REVERTIDO", por: `${quien} (deshacer)`, monto: null });
          }
          await cargarGestion();
          avisar(`Se revirtió «${que}»${cuantos}.`);
        },
      },
    });
  }, [g, cargarGestion, avisar, esSecretaria]);

  const comentar = useCallback(async (socio, texto) => {
    const quien = esSecretaria ? "Secretaría" : "Comisión";
    const { error } = await supabase.from("comentarios").insert({ socio, fecha: new Date().toISOString(), por: quien, texto });
    if (error) { avisar("No se pudo guardar el comentario.", { tono: "error" }); return; }
    await cargarGestion();
    avisar("Comentario agregado.");
  }, [cargarGestion, avisar, esSecretaria]);

  /* ── Edición del historial ──────────────────────────────────
     Para corregir un reclamo mal cargado o sumar reclamos/pagos de
     antes de usar la página. Sólo secretaría; queda igual para todos
     apenas se guarda, porque el historial vive en la base, no en cada
     navegador. */

  const editarLog = useCallback(async (id, cambios) => {
    const { error } = await supabase.from("gestion_log").update(cambios).eq("id", id);
    if (error) { avisar("No se pudo guardar la corrección.", { tono: "error" }); return; }
    await cargarGestion();
    avisar("Registro corregido.");
  }, [cargarGestion, avisar]);

  const eliminarLog = useCallback(async (id, socio) => {
    const { error } = await supabase.from("gestion_log").delete().eq("id", id);
    if (error) { avisar("No se pudo eliminar el registro.", { tono: "error" }); return; }

    // Si no queda ningún registro para este socio, ya no tiene sentido que
    // siga marcado como "Reclamado" (ni con ningún otro estado): se limpia.
    const { count, error: eCount } = await supabase
      .from("gestion_log").select("id", { count: "exact", head: true }).eq("socio", socio);
    if (!eCount && !count) {
      await supabase.from("gestion").delete().eq("socio", socio);
    }

    await cargarGestion();
    avisar(!eCount && !count ? "Se eliminó el último registro: el socio volvió a quedar sin gestionar." : "Registro eliminado del historial.");
  }, [cargarGestion, avisar]);

  const agregarLogManual = useCallback(async (socio, { estado, monto, fecha, actualizarEstado }) => {
    const quien = esSecretaria ? "Secretaría" : "Comisión";
    const fechaIso = new Date(`${fecha}T12:00:00`).toISOString();
    const { error: eLog } = await supabase.from("gestion_log")
      .insert({ socio, fecha: fechaIso, estado, por: `${quien} (carga manual)`, monto: monto || null });
    if (eLog) { avisar("No se pudo agregar el registro.", { tono: "error" }); return; }

    if (actualizarEstado) {
      const prev = g(socio);
      const { error: eG } = await supabase.from("gestion").upsert({
        socio, estado, ultimo_reclamo: estado === "RECLAMADO" ? fechaIso : prev.ultimoReclamo,
        pago: monto ? (Number(prev.pago) || 0) + monto : prev.pago,
        saldo_al_pagar: prev.saldoAlPagar, actualizado_en: fechaIso, actualizado_por: `${quien} (carga manual)`,
      });
      if (eG) { avisar("El registro se agregó, pero no pude actualizar el estado actual de la cuenta.", { tono: "error" }); }
    }
    await cargarGestion();
    avisar("Se agregó el registro al historial.");
  }, [g, cargarGestion, avisar, esSecretaria]);

  /* ── Importaciones ── */
  // Trae los saldos que subió cualquiera (vos u otra persona) desde Supabase.

  async function importarSaldos(file) {
    setProcesando("saldos"); setErrorInicio("");
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const filas = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      if (!filas.length || !("SOCIO" in filas[0]) || !("DEUDA" in filas[0])) {
        setErrorInicio("Ese archivo no tiene las columnas del export de saldos. Hacen falta SOCIO, NOMBRE, DEUDA, PERIODO y DESCRI_CONCEPTO_LIQ.");
        setProcesando(""); return;
      }
      const filasDb = filas.map((f) => ({
        socio: Number(f.SOCIO) || 0, nombre: String(f.NOMBRE || ""), division: String(f.DIVISION || ""),
        categoria_social: String(f.CATEGORIA_SOCIAL || ""), jefefam: Number(f.JEFEFAM) || Number(f.SOCIO) || 0,
        deporte: String(f.DEPORTE || ""), periodo: Number(f.PERIODO) || 0,
        descri_concepto_liq: String(f.DESCRI_CONCEPTO_LIQ || ""), deuda: Number(f.DEUDA) || 0,
      })).filter((f) => f.socio > 0);

      // Reemplaza la tabla entera: se borra todo y se sube el archivo nuevo, en tandas de 500.
      const { error: eDel } = await supabase.from("saldos").delete().gt("id", 0);
      if (eDel) throw eDel;
      for (let i = 0; i < filasDb.length; i += 500) {
        const { error: eIns } = await supabase.from("saldos").insert(filasDb.slice(i, i + 500));
        if (eIns) throw eIns;
      }

      const d = construirDatos(filas);
      setDatos(d); setVista("cola");
      avisar(`Cargué ${d.socios.length} socios con saldo desde ${file.name}. Ya lo ven todos.`);
    } catch (err) {
      setErrorInicio("No pude subir el archivo a la base de datos. Revisá tu conexión y probá de nuevo — si el archivo es válido, no debería perderse nada.");
    }
    setProcesando("");
  }

  async function importarCobranzas(file) {
    setProcesando("cobranzas");
    const quien = esSecretaria ? "Secretaría" : "Comisión";
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const { pagos, error } = leerCobranzas(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]));
      if (error) { avisar(error, { tono: "error", persistente: true }); setProcesando(""); return; }

      const ahora = new Date().toISOString();
      let aplicados = 0, totales = 0, parciales = 0; const sinMatch = [];
      for (const [nro, importe] of pagos) {
        const soc = datos.socios.find((x) => x.socio === nro);
        if (!soc) { sinMatch.push(nro); continue; }
        const prev = g(nro);
        const acum = (Number(prev.pago) || 0) + importe;
        const estado = acum >= soc.deuda ? "PAGO" : "PAGO PARCIAL";
        estado === "PAGO" ? totales++ : parciales++; aplicados++;
        await supabase.from("gestion").upsert({ socio: nro, estado, pago: acum, saldo_al_pagar: soc.deuda, actualizado_en: ahora, actualizado_por: `Importación (${quien})` });
        await supabase.from("gestion_log").insert({ socio: nro, fecha: ahora, estado, por: "Importación de cobranzas", monto: importe });
      }
      await cargarGestion();
      avisar(`${aplicados} cobranzas aplicadas: ${totales} cancelaron la deuda, ${parciales} quedaron parciales.` +
        (sinMatch.length ? ` ${sinMatch.length} socio(s) no estaban en los saldos.` : ""));
      if (sinMatch.length) avisar(`No encontré estos socios en los saldos cargados: ${sinMatch.slice(0, 10).join(", ")}${sinMatch.length > 10 ? "…" : ""}. Revisá los números o volvé a importar los saldos.`, { tono: "error", persistente: true });
    } catch {
      avisar("No pude leer ese archivo. Tiene que ser un .xls o .xlsx con una fila por cobranza.", { tono: "error" });
    }
    setProcesando("");
  }

  /* ── Derivados ── */
  const actividades = useMemo(() => {
    if (!datos) return [];
    const m = new Map();
    for (const s of datos.socios) {
      if (!m.has(s.actividad)) m.set(s.actividad, { id: s.actividad, socios: 0, deuda: 0 });
      const a = m.get(s.actividad); a.socios++; a.deuda += s.deuda;
    }
    return [...m.values()].sort((a, b) => ordenAct(a.id) - ordenAct(b.id));
  }, [datos]);

  const cola = useMemo(() => {
    if (!datos) return [];
    return datos.familias
      .map((f) => {
        const estados = f.miembros.map((m) => g(m.socio).estado);
        const fechas = f.miembros.map((m) => g(m.socio).ultimoReclamo).filter(Boolean).sort();
        return { ...f, estados, ultimoReclamo: fechas.length ? fechas[fechas.length - 1] : null };
      })
      .filter((f) => {
        if (f.deuda < 1000) return false;
        if (f.miembros.some((m) => m.tienePlan)) return false;
        if (f.estados.every((e) => e && ESTADO_MAP[e]?.cierra)) return false;
        const d = diasDesde(f.ultimoReclamo);
        if (d !== null && d < DIAS_ENTRE_RECLAMOS) return false;
        if (actividad && !f.actividades.includes(actividad)) return false;
        return true;
      })
      .sort((a, b) => b.deuda - a.deuda);
  }, [datos, gestion, actividad, g]);

  const listaSocios = useMemo(() => {
    if (!datos) return [];
    const q = busqueda.trim().toLowerCase();
    return datos.socios.filter((s) => {
      if (q && !(s.nombre.toLowerCase().includes(q) || String(s.socio).includes(q))) return false;
      if (actividad && s.actividad !== actividad) return false;
      if (filtroEstado) {
        const e = g(s.socio).estado;
        if (filtroEstado === "__sin" ? e : e !== filtroEstado) return false;
      }
      return true;
    });
  }, [datos, busqueda, actividad, filtroEstado, gestion, g]);

  const resumen = useMemo(() => {
    if (!datos) return [];
    const m = new Map();
    for (const s of datos.socios) {
      if (!m.has(s.actividad)) m.set(s.actividad, { actividad: s.actividad, socios: 0, deuda: 0, cobrado: 0, conteo: {} });
      const r = m.get(s.actividad); r.socios++; r.deuda += s.deuda;
      const e = g(s.socio).estado || "__sin";
      r.conteo[e] = (r.conteo[e] || 0) + 1;
      r.cobrado += Number(g(s.socio).pago) || 0;
    }
    return [...m.values()].sort((a, b) => ordenAct(a.actividad) - ordenAct(b.actividad));
  }, [datos, gestion, g]);

  const totales = useMemo(() => {
    if (!datos) return null;
    const rel = actividad ? datos.socios.filter((s) => s.actividad === actividad) : datos.socios;
    return {
      deuda: rel.reduce((a, s) => a + s.deuda, 0),
      cobrado: rel.reduce((a, s) => a + (Number(g(s.socio).pago) || 0), 0),
      gestionados: rel.filter((s) => g(s.socio).estado).length,
      atrasados: rel.filter((s) => s.meses > MESES_ALERTA).length,
      socios: rel.length,
      familias: datos.familias.length,
    };
  }, [datos, gestion, actividad, g]);

  const socioAbierto = useMemo(() => (datos && abierto ? datos.socios.find((s) => s.socio === abierto) : null), [datos, abierto]);
  const secciones = esSecretaria ? ["cola", "socios", "resumen"] : ["socios", "resumen"];

  if (!listo) return <div className="ds"><style>{CSS}</style></div>;

  /* ── Pantalla de bienvenida ── */
  if (!datos) {
    return (
      <div className="ds">
        <style>{CSS}</style>
        <div className="lienzo">
          <div className="inicio">
            <img className="agua" src={ESCUDO} alt="Escudo del club" />
            <h2 className="t-display">Saldos de cuentas corrientes</h2>
            <p>Los saldos entran tal cual salen del sistema del club. La gestión que cargues encima se guarda aparte, así no se pierde cuando volvés a importar.</p>
            <BotonArchivo onFile={importarSaldos} variante="btn-1" cargando={procesando === "saldos"}>
              <Svg d={Ico.sube} /> Subir archivo de saldos
            </BotonArchivo>
            {errorInicio && (
              <div className="aviso err" style={{ marginTop: "var(--s5)", textAlign: "left" }}>
                <Svg d={Ico.alerta} /><span>{errorInicio}</span>
              </div>
            )}
            <div className="pasos">
              {[["Subís el export de saldos", "El archivo .xls que baja el sistema, sin tocar nada."],
                ["Reclamás desde la cola", "Ordenada por deuda y agrupada por familia, para no reclamar dos veces a la misma casa."],
                ["La comisión mira el resumen", "Ve los saldos y el avance de la cobranza, sin poder alterar la gestión."]].map(([t, d], i) => (
                <div className="paso" key={i}>
                  <span className="n">{i + 1}</span>
                  <div><div className="t-strong">{t}</div><div className="d">{d}</div></div>
                </div>
              ))}
            </div>
          </div>
        </div>
        <Avisos lista={avisos} quitar={quitarAviso} />
      </div>
    );
  }

  const cfg = VISTAS[vista];

  return (
    <div className="ds">
      <style>{CSS}</style>

      {/* ── Barra superior: identidad, búsqueda, cuenta ── */}
      <header className="barra">
        <div className="barra-in">
          <div className="marca">
            <img className="escudo" src={ESCUDO} alt="Escudo del club" />
            <span className="sep" aria-hidden="true" />
            <span className="titulo">Saldos de cuentas corrientes</span>
          </div>

          <Buscador socios={datos.socios} abrir={setAbierto} />

          <div style={{ marginLeft: "auto", display: "flex", gap: "var(--s2)", alignItems: "center" }}>
            <Menu etiqueta="Cambiar de rol" disparador={(toggle, ab) => (
              <button className="btn" aria-haspopup="menu" aria-expanded={ab} onClick={toggle}>
                {esSecretaria ? "Secretaría" : "Comisión"}
              </button>
            )}>
              <div className="titulo-menu">Ver la página como</div>
              <button role="menuitem" onClick={() => setRol("secretaria")}>Secretaría {esSecretaria && "✓"}</button>
              <button role="menuitem" onClick={() => { setRol("comision"); if (vista === "cola") setVista("socios"); }}>
                Comisión directiva {!esSecretaria && "✓"}
              </button>
            </Menu>

            <Menu etiqueta="Datos y archivos">
              <div className="titulo-menu">Datos</div>
              <button role="menuitem" onClick={() => { const n = exportarCobranzas(datos, gestion); avisar(n ? `Descargué el listado con ${n} cobranza(s).` : "El listado salió sin cobranzas porque todavía no cargaste ninguna."); }}>
                Exportar cobranzas
              </button>
              {esSecretaria && <button role="menuitem" onClick={() => setConfirmar("cobranzas")}>Importar cobranzas…</button>}
              <button role="menuitem" onClick={() => setConfirmar("saldos")}>Cargar saldos de otro mes…</button>
              {esSecretaria && (<><hr />
                <button role="menuitem" className="peligro" disabled={!Object.keys(gestion).length} onClick={() => setConfirmar("borrar")}>
                  Borrar toda la gestión…
                </button></>)}
            </Menu>
          </div>
        </div>
      </header>

      {/* ── Navegación de secciones ── */}
      <nav className="nav" aria-label="Secciones">
        <div className="nav-in">
          {secciones.map((v) => (
            <button key={v} aria-current={vista === v ? "page" : undefined} onClick={() => setVista(v)}>
              {VISTAS[v].label}
              {v === "cola" && <span className="cuenta">{cola.length}</span>}
            </button>
          ))}
        </div>
      </nav>

      <main className="lienzo">
        {/* ── Cabecera de vista: dónde estoy, qué hago acá ── */}
        <div className="cabecera-vista">
          <div className="txt">
            <h1 className="t-h1">{cfg.h1}{actividad && <span className="apagado" style={{ fontWeight: 450 }}> · {actividad}</span>}</h1>
            <p>{cfg.desc}</p>
          </div>
          {esSecretaria && vista !== "resumen" && (
            <div className="cta">
              <BotonArchivo onFile={importarCobranzas} cargando={procesando === "cobranzas"}>
                <Svg d={Ico.descarga} /> Importar cobranzas
              </BotonArchivo>
            </div>
          )}
        </div>

        {/* ── Filtro de actividad ── */}
        <div className="filtro-act" role="group" aria-label="Filtrar por actividad">
          <button aria-pressed={actividad === ""} onClick={() => setActividad("")}>Todas</button>
          {actividades.map((a) => (
            <button key={a.id} aria-pressed={actividad === a.id} onClick={() => setActividad(a.id)}>
              {a.id}<span className="cuenta">{a.socios}</span>
            </button>
          ))}
        </div>

        {/* ── Contexto: sólo las dos cifras que importan en esta vista ── */}
        {vista !== "resumen" && (
          <div className="contexto">
            <div className="dato primario">
              <div className="k">{actividad ? `Deuda de ${actividad}` : "Deuda total"}</div>
              <div className="v num-xl">{money(totales.deuda)}</div>
            </div>
            <div className="dato">
              <div className="k">Cobrado</div>
              <div className="v num-xl" style={{ color: totales.cobrado ? "var(--exito)" : undefined }}>{money(totales.cobrado)}</div>
            </div>
            <div className="dato oculta-sm">
              <div className="k">+3 meses de deuda</div>
              <div className="v num-xl">{totales.atrasados}</div>
            </div>
          </div>
        )}

        {vista === "cola" && esSecretaria && (
          <VistaCola cola={cola} actividad={actividad} marcar={marcar} g={g} abrir={setAbierto}
            hayFiltro={!!actividad} limpiar={() => setActividad("")} />
        )}
        {vista === "socios" && (
          <VistaSocios lista={listaSocios} g={g} abrir={setAbierto}
            busqueda={busqueda} setBusqueda={setBusqueda}
            filtroEstado={filtroEstado} setFiltroEstado={setFiltroEstado}
            limpiarTodo={() => { setBusqueda(""); setFiltroEstado(""); setActividad(""); }} />
        )}
        {vista === "resumen" && <VistaResumen filas={resumen} totales={totales} />}
      </main>

      {socioAbierto && (
        <Panel socio={socioAbierto} familia={datos.familias.find((f) => f.jefefam === socioAbierto.jefefam)}
          estado={g(socioAbierto.socio)} esSecretaria={esSecretaria}
          marcar={marcar} comentar={comentar} editarLog={editarLog} eliminarLog={eliminarLog}
          agregarLogManual={agregarLogManual} cerrar={() => setAbierto(null)} />
      )}

      {confirmar === "borrar" && (
        <Modal titulo="¿Borrar toda la gestión cargada?" onCerrar={() => setConfirmar(null)}
          acciones={<>
            <button className="btn" onClick={() => setConfirmar(null)}>Cancelar</button>
            <button className="btn btn-peligro-1" onClick={async () => {
              const n = Object.keys(gestion).length;
              setConfirmar(null);
              const [{ error: e1 }, { error: e2 }, { error: e3 }] = await Promise.all([
                supabase.from("gestion").delete().gt("socio", 0),
                supabase.from("gestion_log").delete().gt("id", 0),
                supabase.from("comentarios").delete().gt("id", 0),
              ]);
              if (e1 || e2 || e3) { avisar("No pude borrar todo. Probá de nuevo.", { tono: "error" }); return; }
              await cargarGestion();
              avisar(`Borré la gestión de ${n} socios. Esto no se puede deshacer.`);
            }}>Sí, borrar {Object.keys(gestion).length} registros</button>
          </>}>
          Se borran, para todos los que usan la página, los estados, los pagos, los comentarios y el historial
          de {Object.keys(gestion).length} socios. Los saldos importados quedan intactos.
          <strong style={{ display: "block", marginTop: 8, color: "var(--critico)" }}>
            Esta acción no se puede deshacer.
          </strong>
        </Modal>
      )}

      {confirmar === "saldos" && (
        <Modal titulo="Cargar los saldos de otro mes" onCerrar={() => setConfirmar(null)}
          acciones={<>
            <button className="btn" onClick={() => setConfirmar(null)}>Cancelar</button>
            <BotonArchivo variante="btn-1" cargando={procesando === "saldos"}
              onFile={(f) => { setConfirmar(null); importarSaldos(f); }}>Elegir archivo de septiembre (o el que corresponda)</BotonArchivo>
          </>}>
          Se reemplazan los saldos por los del archivo nuevo — es lo mismo que hacer un corte de mes.
          Los reclamos, pagos y comentarios que ya cargaste no se tocan: quedan guardados por número de
          socio y se vuelven a cruzar solos con el saldo actualizado de cada uno.
        </Modal>
      )}

      {confirmar === "cobranzas" && (
        <Modal titulo="Importar cobranzas" onCerrar={() => setConfirmar(null)}
          acciones={<>
            <button className="btn" onClick={() => setConfirmar(null)}>Cancelar</button>
            <BotonArchivo variante="btn-1" cargando={procesando === "cobranzas"}
              onFile={(f) => { setConfirmar(null); importarCobranzas(f); }}>Elegir archivo</BotonArchivo>
          </>}>
          El archivo necesita una columna con el número de socio y otra con el importe. Cada cobranza se marca como
          «Pagó» o «Pago parcial» según el saldo del socio. Si te equivocás de archivo, podés deshacerlo.
        </Modal>
      )}

      <Avisos lista={avisos} quitar={quitarAviso} />
    </div>
  );
}
