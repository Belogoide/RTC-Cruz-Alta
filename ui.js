(function () {
  function criarModal() {
    let modal = document.getElementById("uiModal");

    if (modal) return modal;

    modal = document.createElement("div");
    modal.id = "uiModal";
    modal.className = "ui-modal-overlay";
    modal.setAttribute("role", "dialog");
    modal.setAttribute("aria-modal", "true");
    modal.innerHTML = `
      <div class="ui-modal" role="document">
        <h2 id="uiModalTitulo" class="ui-modal-titulo"></h2>
        <p id="uiModalMensagem" class="ui-modal-mensagem"></p>
        <input id="uiModalEntrada" class="ui-modal-entrada" type="text" autocomplete="email" />
        <div class="ui-modal-acoes">
          <button id="uiModalCancelar" class="botao botao-secundario" type="button">Cancelar</button>
          <button id="uiModalConfirmar" class="botao" type="button">OK</button>
        </div>
      </div>
    `;

    document.body.appendChild(modal);
    return modal;
  }

  function abrirModal({ titulo, mensagem, entrada, confirmar, cancelar }) {
    const modal = criarModal();
    const tituloElemento = document.getElementById("uiModalTitulo");
    const mensagemElemento = document.getElementById("uiModalMensagem");
    const entradaElemento = document.getElementById("uiModalEntrada");
    const confirmarElemento = document.getElementById("uiModalConfirmar");
    const cancelarElemento = document.getElementById("uiModalCancelar");

    tituloElemento.textContent = titulo || "RTC Cruz Alta";
    mensagemElemento.textContent = mensagem || "";
    entradaElemento.value = entrada || "";
    entradaElemento.placeholder = "";
    entradaElemento.style.display = entrada === undefined ? "none" : "block";
    confirmarElemento.textContent = cancelar ? "Confirmar" : "OK";
    cancelarElemento.style.display = cancelar ? "block" : "none";
    modal.style.display = "flex";

    const fechar = (resultado) => {
      modal.style.display = "none";
      document.removeEventListener("keydown", teclaEscape);
      resultadoFinal(resultado);
    };

    let resultadoFinal = () => {};
    const teclaEscape = (event) => {
      if (event.key === "Escape") fechar(cancelar ? false : null);
    };

    document.addEventListener("keydown", teclaEscape);
    confirmarElemento.onclick = () => fechar(entrada === undefined ? true : entradaElemento.value.trim());
    cancelarElemento.onclick = () => fechar(cancelar ? false : null);

    if (entrada !== undefined) {
      entradaElemento.focus();
      entradaElemento.onkeydown = (event) => {
        if (event.key === "Enter") fechar(entradaElemento.value.trim());
      };
    } else {
      confirmarElemento.focus();
    }

    return new Promise((resolve) => {
      resultadoFinal = resolve;
    });
  }

  window.uiAlert = (mensagem, titulo = "Aviso") =>
    abrirModal({ titulo, mensagem });

  window.uiConfirm = (mensagem, titulo = "Confirmar ação") =>
    abrirModal({ titulo, mensagem, cancelar: true });

  window.uiPrompt = (mensagem, valorInicial = "", titulo = "Informação") =>
    abrirModal({ titulo, mensagem, entrada: valorInicial });
})();
