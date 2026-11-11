(() => {
  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("fileInput");
  const browseBtn = document.getElementById("browseBtn");
  const addMoreBtn = document.getElementById("addMoreBtn");
  const sortAzBtn = document.getElementById("sortAzBtn");
  const reverseBtn = document.getElementById("reverseBtn");
  const clearAllBtn = document.getElementById("clearAllBtn");
  const fileListEl = document.getElementById("fileList");
  const emptyStateEl = document.getElementById("emptyState");
  const fileCountBadge = document.getElementById("fileCountBadge");
  const outputNameInput = document.getElementById("outputName");
  const convertBtn = document.getElementById("convertBtn");

  const progressCard = document.getElementById("progressCard");
  const progressTitle = document.getElementById("progressTitle");
  const progressDetail = document.getElementById("progressDetail");
  const progressBarFill = document.getElementById("progressBarFill");

  const resultCard = document.getElementById("resultCard");
  const resultTitle = document.getElementById("resultTitle");
  const resultSummary = document.getElementById("resultSummary");
  const previewPdfBtn = document.getElementById("previewPdfBtn");
  const downloadPdfBtn = document.getElementById("downloadPdfBtn");

  const errorCard = document.getElementById("errorCard");
  const errorMessage = document.getElementById("errorMessage");

  // Estado: lista ordenada de arquivos { id, file, name, size }
  let queue = [];
  let dragSrcIndex = null;
  let currentBlobUrl = null;

  const ALLOWED_EXTENSIONS = [".docx", ".doc", ".rtf", ".odt", ".pdf"];

  function formatBytes(bytes) {
    if (bytes === 0) return "0 KB";
    const kb = bytes / 1024;
    if (kb < 1024) return `${kb.toFixed(1)} KB`;
    return `${(kb / 1024).toFixed(2)} MB`;
  }

  function getExtension(filename) {
    const idx = filename.lastIndexOf(".");
    return idx !== -1 ? filename.slice(idx).toLowerCase() : "";
  }

  function showError(msg) {
    errorMessage.textContent = msg;
    errorCard.classList.remove("hidden");
  }

  function hideAlerts() {
    errorCard.classList.add("hidden");
    resultCard.classList.add("hidden");
  }

  function addFiles(fileList) {
    hideAlerts();
    const incoming = Array.from(fileList);
    const rejected = [];

    for (const file of incoming) {
      const ext = getExtension(file.name);
      if (!ALLOWED_EXTENSIONS.includes(ext)) {
        rejected.push(file.name);
        continue;
      }
      queue.push({
        id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random()}`,
        file,
        name: file.name,
        size: file.size
      });
    }

    if (rejected.length > 0) {
      showError(
        `Formato não suportado ignorado: ${rejected.join(", ")}. Use arquivos Word (.docx, .doc, .rtf) ou .pdf.`
      );
    }

    renderQueue();
  }

  function renderQueue() {
    fileCountBadge.textContent = queue.length;
    convertBtn.disabled = queue.length === 0;

    if (queue.length === 0) {
      emptyStateEl.classList.remove("hidden");
      fileListEl.innerHTML = "";
      return;
    }

    emptyStateEl.classList.add("hidden");
    fileListEl.innerHTML = "";

    queue.forEach((item, index) => {
      const li = document.createElement("li");
      li.className = "file-item";
      li.draggable = true;
      li.dataset.index = index;

      li.innerHTML = `
        <div class="file-left">
          <span class="drag-handle" title="Arraste para reordenar">☰</span>
          <span class="order-badge">${index + 1}</span>
          <div class="file-meta">
            <div class="file-name" title="${escapeHtml(item.name)}">${escapeHtml(item.name)}</div>
            <div class="file-size">${formatBytes(item.size)}</div>
          </div>
        </div>
        <div class="file-actions">
          <button type="button" class="icon-btn move-up" title="Mover para cima" ${index === 0 ? "disabled" : ""}>↑</button>
          <button type="button" class="icon-btn move-down" title="Mover para baixo" ${index === queue.length - 1 ? "disabled" : ""}>↓</button>
          <button type="button" class="icon-btn remove-btn" title="Remover arquivo">✕</button>
        </div>
      `;

      // Botões de mover / remover
      li.querySelector(".move-up").addEventListener("click", () => moveItem(index, index - 1));
      li.querySelector(".move-down").addEventListener("click", () => moveItem(index, index + 1));
      li.querySelector(".remove-btn").addEventListener("click", () => removeItem(index));

      // Eventos de Drag and Drop para reordenar
      li.addEventListener("dragstart", (e) => {
        dragSrcIndex = index;
        li.classList.add("dragging");
        e.dataTransfer.effectAllowed = "move";
      });

      li.addEventListener("dragend", () => {
        li.classList.remove("dragging");
        document.querySelectorAll(".file-item").forEach((el) => el.classList.remove("drag-over"));
      });

      li.addEventListener("dragover", (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
        li.classList.add("drag-over");
      });

      li.addEventListener("dragleave", () => {
        li.classList.remove("drag-over");
      });

      li.addEventListener("drop", (e) => {
        e.preventDefault();
        li.classList.remove("drag-over");
        if (dragSrcIndex !== null && dragSrcIndex !== index) {
          moveItem(dragSrcIndex, index);
        }
        dragSrcIndex = null;
      });

      fileListEl.appendChild(li);
    });
  }

  function moveItem(fromIndex, toIndex) {
    if (toIndex < 0 || toIndex >= queue.length) return;
    const [moved] = queue.splice(fromIndex, 1);
    queue.splice(toIndex, 0, moved);
    renderQueue();
  }

  function removeItem(index) {
    queue.splice(index, 1);
    renderQueue();
  }

  function escapeHtml(str) {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function fileToBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = reader.result;
        const base64 = result.split(",")[1];
        resolve(base64);
      };
      reader.onerror = (err) => reject(err);
      reader.readAsDataURL(file);
    });
  }

  function base64ToUint8Array(base64) {
    const binary = atob(base64);
    const len = binary.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  }

  // Eventos de Upload
  browseBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    fileInput.click();
  });

  addMoreBtn.addEventListener("click", () => fileInput.click());

  dropzone.addEventListener("click", () => fileInput.click());

  fileInput.addEventListener("change", () => {
    if (fileInput.files.length > 0) {
      addFiles(fileInput.files);
      fileInput.value = "";
    }
  });

  dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("drag-active");
  });

  dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("drag-active");
  });

  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("drag-active");
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      addFiles(e.dataTransfer.files);
    }
  });

  // Botões de Ordenação Rápida
  sortAzBtn.addEventListener("click", () => {
    if (queue.length < 2) return;
    queue.sort((a, b) => a.name.localeCompare(b.name, "pt-BR", { numeric: true }));
    renderQueue();
  });

  reverseBtn.addEventListener("click", () => {
    if (queue.length < 2) return;
    queue.reverse();
    renderQueue();
  });

  clearAllBtn.addEventListener("click", () => {
    queue = [];
    hideAlerts();
    renderQueue();
  });

  // Conversão e Unificação
  convertBtn.addEventListener("click", async () => {
    if (queue.length === 0) return;

    hideAlerts();
    convertBtn.disabled = true;
    progressCard.classList.remove("hidden");
    progressTitle.textContent = "Preparando arquivos para o Microsoft Word...";
    progressDetail.textContent = `Lendo ${queue.length} arquivo(s) na ordem escolhida...`;
    progressBarFill.style.width = "15%";

    try {
      const payloadFiles = [];
      for (let i = 0; i < queue.length; i++) {
        const item = queue[i];
        const b64 = await fileToBase64(item.file);
        payloadFiles.push({
          name: item.name,
          data: b64
        });
      }

      const hasPdfLib = typeof window.PDFLib !== "undefined";

      progressTitle.textContent = "Convertendo documentos no Microsoft Word...";
      progressDetail.textContent = `Exportando ${queue.length} documento(s) com fidelidade total de formatação...`;
      progressBarFill.style.width = "50%";

      const response = await fetch("/api/convert", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          files: payloadFiles,
          mergeInWord: !hasPdfLib
        })
      });

      const result = await response.json();
      if (!response.ok || !result.ok) {
        throw new Error(result.error || "Falha ao converter arquivos no servidor.");
      }

      let finalPdfBytes = null;
      let totalPages = result.totalPages || 0;

      if (result.mergedPdfBase64) {
        progressTitle.textContent = "Finalizando PDF unificado pelo Python...";
        progressBarFill.style.width = "90%";
        finalPdfBytes = base64ToUint8Array(result.mergedPdfBase64);
      } else if (hasPdfLib && result.files) {
        progressTitle.textContent = "Unificando páginas em um único PDF...";
        progressDetail.textContent = "Juntando todos os documentos na sequência definida...";
        progressBarFill.style.width = "85%";

        const { PDFDocument } = window.PDFLib;
        const mergedPdf = await PDFDocument.create();

        for (const converted of result.files) {
          const pdfBytes = base64ToUint8Array(converted.pdfBase64);
          const sourcePdf = await PDFDocument.load(pdfBytes);
          const copiedPages = await mergedPdf.copyPages(sourcePdf, sourcePdf.getPageIndices());
          copiedPages.forEach((page) => {
            mergedPdf.addPage(page);
            totalPages++;
          });
        }

        finalPdfBytes = await mergedPdf.save();
      } else {
        throw new Error("Resposta inesperada do servidor Python.");
      }

      progressBarFill.style.width = "100%";

      if (currentBlobUrl) {
        URL.revokeObjectURL(currentBlobUrl);
      }

      const blob = new Blob([finalPdfBytes], { type: "application/pdf" });
      currentBlobUrl = URL.createObjectURL(blob);

      let rawOutputName = (outputNameInput.value || "Documento_Unificado").trim();
      if (rawOutputName.toLowerCase().endsWith(".pdf")) {
        rawOutputName = rawOutputName.slice(0, -4);
      }
      const finalFileName = `${rawOutputName || "Documento_Unificado"}.pdf`;

      previewPdfBtn.href = currentBlobUrl;
      downloadPdfBtn.href = currentBlobUrl;
      downloadPdfBtn.download = finalFileName;

      progressCard.classList.add("hidden");
      resultTitle.textContent = `PDF Unificado Pronto: ${finalFileName}`;
      resultSummary.textContent = totalPages > 0
        ? `${queue.length} arquivo(s) combinados em ${totalPages} página(s) ao todo.`
        : `${queue.length} arquivo(s) combinados com sucesso.`;
      resultCard.classList.remove("hidden");

      // Dispara download automático para facilitar o fluxo
      downloadPdfBtn.click();
    } catch (err) {
      progressCard.classList.add("hidden");
      showError(err.message || "Erro inesperado durante a conversão.");
    } finally {
      convertBtn.disabled = queue.length === 0;
    }
  });
})();
