function copyCode(button) {
    const codeBlock = button.parentElement.querySelector('code');
    const confirmation = button.parentElement.querySelector('.copy-confirmation');
    
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
function copyCodeToClipboard(codeBlockId) {
  var codeBlock = document.getElementById(codeBlockId);
  var code = codeBlock.innerText;
  navigator.clipboard.writeText(code).then(() => {
    Toastify({
      text: "Código copiado",
      duration: 2000,
      backgroundColor: "#4CAF50",
      className: "toastify"
    }).showToast();
  });
}
