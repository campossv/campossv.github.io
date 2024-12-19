function copyCode(blockId) {
  const codeBlock = document.getElementById(blockId);
  const confirmation = document.getElementById(
    `copyConfirmation${blockId.slice(-1)}`
  );

  // Crear un elemento de texto temporal
  const textArea = document.createElement("textarea");
  textArea.value = codeBlock.textContent.trim();

  // Añadir el elemento al DOM
  document.body.appendChild(textArea);

  // Seleccionar y copiar el texto
  textArea.select();
  document.execCommand("copy");

  // Eliminar el elemento temporal
  document.body.removeChild(textArea);

  // Mostrar mensaje de confirmación
  confirmation.classList.add("show");
  setTimeout(() => {
    confirmation.classList.remove("show");
  }, 2000);
}
