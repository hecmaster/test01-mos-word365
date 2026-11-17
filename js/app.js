// Portal Inversiones Inteligentes: progreso de cuestionarios y acceso al examen final.
function abrirRetoFinal() { window.open('examen_final.html', '_blank', 'noopener'); }
function actualizarProgresoQuiz() {
  const datos = (window.progresoQuiz && window.progresoQuiz()) || {};
  const tit = { oro: '🏆', plata: '🥈', bronce: '🥉', base: '💪' };
  let hechos = 0, suma = 0, total = 0;
  document.querySelectorAll('.card-theory').forEach(card => {
    const m = (card.getAttribute('onclick') || '').match(/abrirCuestionarioTeorico\((\d+)\)/); if (!m) return; total++;
    const d = datos[m[1]]; let el = card.querySelector('.quiz-best'); if (!el) { el = document.createElement('div'); el.className = 'quiz-best'; card.querySelector('p').after(el); }
    if (d) { hechos++; suma += d.best; el.innerHTML = `${tit[d.medalla] || ''} Mejor: <b>${d.best} %</b> · ${d.intentos} intento${d.intentos > 1 ? 's' : ''}`; card.classList.toggle('hecho', d.best >= 70); } else { el.textContent = 'Aún sin intentos'; }
  });
  const r = document.getElementById('quizResumen');
  if (r) r.innerHTML = total ? `<b>${hechos}/${total}</b> cuestionarios intentados${hechos ? ` · promedio de mejores resultados <b>${Math.round(suma / hechos)} %</b>` : ' · ¡empieza por el primero!'}<i><s style="width:${hechos / total * 100}%"></s></i>` : '';
}
document.addEventListener('DOMContentLoaded', actualizarProgresoQuiz);
