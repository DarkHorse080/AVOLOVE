const track = document.getElementById('divcarrusel');
const btnIzq = document.getElementById('btn-izq');
const btnDer = document.getElementById('btn-der');

// Calculamos el ancho de desplazamiento (ancho de la tarjeta + el espacio/gap)
const tarjetaWidth = document.querySelector('.tarjeta').offsetWidth + 20;
const desplazamiento= tarjetaWidth*4
btnDer.addEventListener('click', () => {
  track.scrollLeft += desplazamiento;
});

btnIzq.addEventListener('click', () => {
  track.scrollLeft -= desplazamiento;
});