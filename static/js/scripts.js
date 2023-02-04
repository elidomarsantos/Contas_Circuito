// Elemento com o Texto
var elemento = document.getElementById('teste').innerHTML;
// Escrevendo em outro Elemento
var texto = document.getElementById('texto');
texto.innerHTML = "Texto Copiado: " + elemento;