/**
 * SuperCopy Pro - Complemento para Script Lab
 * Exporta selecciones de Excel con formato adaptado para WhatsApp, Imagen y Correo.
 * Respeta estilos de fondo y color de fuente del encabezado de Excel.
 */

// Variables de estado del complemento (declaradas al inicio para evitar TDZ)
let cachedSelectionData = null;
let currentModalText = "";
let currentModalAddress = "";
let currentModalDataUrl = "";
let currentModalHtml = "";
let toastTimeout;

/**
 * Determina de forma segura si el entorno actual permite invocar navigator.clipboard.writeText / write
 * sin disparar una violación de Permissions Policy (crbug.com/414348233).
 * En Excel Online, el iframe de Script Lab no posee el atributo allow="clipboard-write".
 */
function isClipboardWriteAllowed() {
  if (typeof navigator === "undefined" || !navigator.clipboard) {
    return false;
  }
  const doc = document;

  // 1. Verificación estándar mediante Permissions Policy API (Chromium 88+)
  if (doc.permissionsPolicy && typeof doc.permissionsPolicy.allowsFeature === "function") {
    try {
      if (!doc.permissionsPolicy.allowsFeature("clipboard-write")) {
        return false;
      }
    } catch (_) {
      return false;
    }
  }

  // 2. Verificación de respaldo mediante Feature Policy API
  if (doc.featurePolicy && typeof doc.featurePolicy.allowsFeature === "function") {
    try {
      if (!doc.featurePolicy.allowsFeature("clipboard-write")) {
        return false;
      }
    } catch (_) {
      return false;
    }
  }

  // 3. Verificación de entorno iframe:
  // Si estamos en un iframe anidado (Script Lab en Excel Web) y la política no confirmó afirmativamente
  // el permiso de clipboard-write, omitimos la llamada asíncrona para prevenir DOMException.
  try {
    if (window.self !== window.top) {
      if (doc.permissionsPolicy && typeof doc.permissionsPolicy.allowsFeature === "function") {
        return doc.permissionsPolicy.allowsFeature("clipboard-write");
      }
      return false;
    }
  } catch (_) {
    // Si window.top lanza SecurityError por origen cruzado, es un iframe restringido
    return false;
  }
  return true;
}

// Inicialización de eventos al cargar el script
document.addEventListener("DOMContentLoaded", () => {
  setupEventListeners();
  loadCurrentSelection();
});

// Respaldo directo si el DOM ya estuviera listo en el runner de Script Lab
if (document.readyState === "complete" || document.readyState === "interactive") {
  setupEventListeners();
  loadCurrentSelection();
}
function setupEventListeners() {
  const btnRefresh = document.getElementById("btn-refresh");
  const btnChat = document.getElementById("btn-copy-chat");
  const btnImage = document.getElementById("btn-copy-image");
  const btnEmail = document.getElementById("btn-copy-email");
  btnRefresh?.addEventListener("click", () => loadCurrentSelection());
  btnChat?.addEventListener("click", () => handleCopy("chat"));
  btnImage?.addEventListener("click", () => handleCopy("image"));
  btnEmail?.addEventListener("click", () => handleCopy("email"));

  // Eventos del Modal Universal de Resultados
  const btnCloseModal = document.getElementById("btn-close-result-modal");
  const modalOverlay = document.getElementById("result-modal");
  btnCloseModal?.addEventListener("click", () => closeResultModal());
  modalOverlay?.addEventListener("click", e => {
    if (e.target === modalOverlay) closeResultModal();
  });

  // Botón Copiar Selección en Modal Texto
  const btnCopySelected = document.getElementById("btn-copy-selected-text");
  btnCopySelected?.addEventListener("click", () => {
    const textarea = document.getElementById("result-textarea");
    if (textarea) {
      textarea.focus();
      textarea.select();
      try {
        const ok = document.execCommand("copy");
        if (ok) {
          showToast("¡Texto copiado al portapapeles!");
        } else {
          showToast("Presiona Ctrl + C para copiar");
        }
      } catch (e) {
        showToast("Presiona Ctrl + C para copiar");
      }
    }
  });

  // Botón Abrir en WhatsApp Web
  const btnOpenWa = document.getElementById("btn-open-wa");
  btnOpenWa?.addEventListener("click", () => {
    if (currentModalText) {
      const url = "https://web.whatsapp.com/send?text=" + encodeURIComponent(currentModalText);
      window.open(url, "_blank");
    }
  });

  // Botón Descargar .txt
  const btnDownloadTxt = document.getElementById("btn-download-txt");
  btnDownloadTxt?.addEventListener("click", () => {
    if (currentModalText) {
      downloadTextFile(currentModalText, currentModalAddress);
      showToast("Archivo .txt guardado");
    }
  });

  // Botón Volver a descargar PNG
  const btnReDownloadImg = document.getElementById("btn-re-download-img");
  btnReDownloadImg?.addEventListener("click", () => {
    if (currentModalDataUrl) {
      const link = document.createElement("a");
      const cleanAddress = currentModalAddress.replace(/[^a-zA-Z0-9_-]/g, "_");
      link.download = `SuperCopy_${cleanAddress}.png`;
      link.href = currentModalDataUrl;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      showToast("Imagen descargada");
    }
  });

  // Botón Seleccionar HTML completo
  const btnSelectHtml = document.getElementById("btn-select-html");
  btnSelectHtml?.addEventListener("click", () => {
    const container = document.getElementById("result-html-container");
    if (container) {
      const range = document.createRange();
      range.selectNodeContents(container);
      const sel = window.getSelection();
      sel?.removeAllRanges();
      sel?.addRange(range);
      try {
        const ok = document.execCommand("copy");
        if (ok) {
          showToast("¡Tabla copiada al portapapeles!");
        } else {
          showToast("Tabla seleccionada: presiona Ctrl + C");
        }
      } catch (e) {
        showToast("Tabla seleccionada: presiona Ctrl + C");
      }
    }
  });

  // Suscribirse al cambio de selección activo en Excel si Office está disponible
  if (typeof Office !== "undefined") {
    Office.onReady(() => {
      loadCurrentSelection();
      Excel.run(async context => {
        context.workbook.onSelectionChanged.add(onSelectionChange);
        await context.sync();
      }).catch(() => {
        // En caso de que el evento ya esté suscrito o no sea soportado en el entorno
      });
    });
  }
}
async function onSelectionChange() {
  await loadCurrentSelection();
}

/**
 * Lee la selección actual de Excel con sus valores formateados y estilos de encabezado
 */
async function getSelectionData() {
  try {
    return await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      range.load(["address", "rowCount", "columnCount"]);
      await context.sync();
      if (range.rowCount === 0 || range.columnCount === 0) {
        return null;
      }

      // Protección contra selección de columnas enteras (ej. 1M filas en Excel Web)
      if (range.rowCount > 500) {
        showToast("Rango muy grande (máx 500 filas para exportar)");
        return null;
      }
      if (range.columnCount > 40) {
        showToast("Demasiadas columnas (máx 40 columnas)");
        return null;
      }
      range.load(["text", "values"]);

      // Cargar formato de las celdas del header (fila 0 de la selección) con protección
      const headerCells = [];
      for (let c = 0; c < range.columnCount; c++) {
        const cell = range.getCell(0, c);
        cell.load(["format/fill/color", "format/font/color", "format/font/bold"]);
        headerCells.push(cell);
      }
      try {
        await context.sync();
      } catch (styleErr) {
        console.warn("No se pudieron leer estilos del encabezado en Excel Web, usando por defecto:", styleErr);
      }

      // 'text' contiene el texto ya formateado (fechas, monedas, etc.)
      const matrix = range.text || [];
      const rawMatrix = range.values || [];
      if (!matrix || matrix.length === 0) {
        return null;
      }
      const headers = matrix[0].map(h => String(h || ""));
      const rows = matrix.slice(1).map(row => row.map(cell => String(cell || "")));

      // Extraer y normalizar colores de fondo y texto del encabezado
      const headerStyles = headers.map((_, idx) => {
        const cell = headerCells[idx];
        const rawBg = cell?.format?.fill?.color;
        const rawFont = cell?.format?.font?.color;
        const isBold = cell?.format?.font?.bold ?? true;
        let bg = "#0f172a"; // Color ejecutivo predeterminado
        if (rawBg && rawBg !== "" && rawBg.toLowerCase() !== "#00000000") {
          bg = rawBg;
        }
        let color = "#ffffff";
        if (rawFont && rawFont !== "") {
          color = rawFont;
        } else {
          color = isLightColor(bg) ? "#0f172a" : "#ffffff";
        }
        return {
          bg,
          color,
          bold: isBold
        };
      });
      return {
        address: range.address,
        rowCount: range.rowCount,
        colCount: range.columnCount,
        headers,
        headerStyles,
        rows,
        rawValues: rawMatrix
      };
    });
  } catch (error) {
    console.error("Error al leer la selección:", error);
    return null;
  }
}

/**
 * Actualiza los indicadores visuales en la interfaz del panel
 */
async function loadCurrentSelection() {
  const addressElem = document.getElementById("range-address");
  const dimElem = document.getElementById("range-dimensions");
  if (!addressElem || !dimElem) return;
  const data = await getSelectionData();
  cachedSelectionData = data;
  if (!data) {
    addressElem.textContent = "Sin selección";
    dimElem.textContent = "0 filas × 0 columnas";
    return;
  }
  addressElem.textContent = data.address.split("!").pop() || data.address;
  dimElem.textContent = `${data.rowCount} fila${data.rowCount > 1 ? "s" : ""} × ${data.colCount} col${data.colCount > 1 ? "s" : ""}`;
  renderPreview(data);
}

/**
 * Manejador principal para cada una de las 3 opciones de copiado
 */
async function handleCopy(mode) {
  // Asegurar foco en el iframe del complemento (previene "Document is not focused")
  window.focus();

  // Priorizar datos en memoria para no perder el User Gesture con llamadas asíncronas lentas a Excel Web
  let data = cachedSelectionData;
  if (!data || data.rowCount === 0) {
    data = await getSelectionData();
    cachedSelectionData = data;
  }
  if (!data || data.rowCount === 0) {
    showToast("Selecciona un rango con datos en Excel primero");
    return;
  }
  try {
    if (mode === "chat") {
      const isDirectCopied = await copyForChat(data);
      if (isDirectCopied) {
        showToast("Copiado para WhatsApp / Chat");
      } else {
        showToast("Texto listo: presiona Ctrl + C o abre WhatsApp");
      }
    } else if (mode === "image") {
      const isDirectCopied = await copyAsImage(data);
      if (isDirectCopied) {
        showToast("Copiado como Imagen al portapapeles");
      } else {
        showToast("Imagen descargada en tu equipo");
      }
    } else if (mode === "email") {
      const isDirectCopied = await copyForEmail(data);
      if (isDirectCopied) {
        showToast("Copiado como Tabla para Correo / HTML");
      } else {
        showToast("Tabla lista en pantalla para copiar");
      }
    }
  } catch (error) {
    console.error("Error al copiar:", error);
    showToast("Error al copiar. Revisa la selección de Excel.");
  }
}

/**
 * 1. FORMATO WHATSAPP / CHAT
 * WhatsApp es texto plano. Destacamos el encabezado mediante:
 * - Indicador temático según el color de fondo del header en Excel (🟦, 🟩, ⬛, etc.)
 * - Separador de doble línea (===+===) que resalta visualmente la cabecera en el bloque monoespaciado.
 */
async function copyForChat(data) {
  const allRows = [data.headers, ...data.rows];
  const colWidths = [];
  for (let c = 0; c < data.colCount; c++) {
    let max = 0;
    for (let r = 0; r < allRows.length; r++) {
      const len = (allRows[r][c] || "").length;
      if (len > max) max = len;
    }
    colWidths.push(Math.max(max, 3));
  }
  const formatRow = row => {
    return row.map((cell, idx) => {
      const width = colWidths[idx];
      const val = cell || "";
      const isNum = isNumeric(val);
      return isNum ? val.padStart(width, " ") : val.padEnd(width, " ");
    }).join(" | ");
  };
  const cleanAddress = data.address.split("!").pop() || data.address;
  const primaryBg = data.headerStyles[0]?.bg || "#0f172a";
  const badge = getColorBadge(primaryBg);
  const title = `${badge} *TABLA EXCEL (${cleanAddress})*`;
  const headerLine = formatRow(data.headers);
  const separatorLine = colWidths.map(w => "=".repeat(w)).join("=+=");
  const dataLines = data.rows.map(row => formatRow(row));
  const textTable = [headerLine, separatorLine, ...dataLines].join("\n");
  const output = `${title}\n\`\`\`\n${textTable}\n\`\`\``;

  // Intento 1: API moderna del portapapeles (solo si está permitida por Permissions Policy)
  if (isClipboardWriteAllowed() && navigator.clipboard && navigator.clipboard.writeText) {
    try {
      await navigator.clipboard.writeText(output);
      return true;
    } catch (_) {
      // Si falla, continuar al fallback sin romper ejecución
    }
  }

  // Intento 2: Fallback por execCommand con textarea temporal en memoria
  const isFallbackOk = copyTextFallback(output);
  if (isFallbackOk) {
    return true;
  }

  // Intento 3: Visor de resultados con selección automática y enlace a WhatsApp Web (infalible en Excel Web)
  showTextResultModal(output, cleanAddress);
  return false;
}

/**
 * 2. FORMATO IMAGEN EJECUTIVA
 * Renderiza la tabla sobre un Canvas de alta resolución respetando el fondo y fuente del header de Excel.
 */
async function copyAsImage(data) {
  const colCount = data.colCount;

  // Medidas del canvas
  const paddingX = 32;
  const paddingY = 28;
  const rowHeight = 36;
  const headerHeight = 42;
  const fontBody = "13px -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif";
  const fontHeader = "600 13px -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif";

  // Canvas auxiliar para calcular anchos de texto
  const tempCanvas = document.createElement("canvas");
  const tempCtx = tempCanvas.getContext("2d");
  if (!tempCtx) throw new Error("No se pudo iniciar el contexto 2D");
  tempCtx.font = fontHeader;
  const colWidths = [];
  for (let c = 0; c < colCount; c++) {
    let max = tempCtx.measureText(data.headers[c] || "").width;
    tempCtx.font = fontBody;
    for (let r = 0; r < data.rows.length; r++) {
      const w = tempCtx.measureText(data.rows[r][c] || "").width;
      if (w > max) max = w;
    }
    colWidths.push(Math.ceil(max) + 28); // margen interno de columna
  }
  const tableWidth = colWidths.reduce((a, b) => a + b, 0);
  const totalWidth = tableWidth + paddingX * 2;
  const totalHeight = paddingY * 2 + headerHeight + data.rows.length * rowHeight;

  // Escala para alta densidad (Retina 2x)
  const scale = 2;
  const canvas = document.createElement("canvas");
  canvas.width = totalWidth * scale;
  canvas.height = totalHeight * scale;
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("Error de contexto en canvas final");
  ctx.scale(scale, scale);

  // Fondo exterior elegante
  ctx.fillStyle = "#f8fafc";
  ctx.fillRect(0, 0, totalWidth, totalHeight);

  // Contenedor tipo tarjeta redondeada
  const cardX = paddingX;
  const cardY = paddingY;
  const cardW = tableWidth;
  const cardH = headerHeight + data.rows.length * rowHeight;
  const radius = 10;
  drawRoundedRect(ctx, cardX, cardY, cardW, cardH, radius);
  ctx.fillStyle = "#ffffff";
  ctx.fill();
  ctx.strokeStyle = "#e2e8f0";
  ctx.lineWidth = 1;
  ctx.stroke();

  // Clip para que las esquinas de filas respeten el radio
  ctx.save();
  drawRoundedRect(ctx, cardX, cardY, cardW, cardH, radius);
  ctx.clip();

  // 1. Cabecera (Fondo dinámico según cada celda del header de Excel)
  let curX = cardX;
  for (let c = 0; c < colCount; c++) {
    const w = colWidths[c];
    const style = data.headerStyles[c] || {
      bg: "#0f172a",
      color: "#ffffff",
      bold: true
    };
    ctx.fillStyle = style.bg;
    ctx.fillRect(curX, cardY, w, headerHeight);
    curX += w;
  }
  ctx.font = fontHeader;
  ctx.textBaseline = "middle";

  // Texto del encabezado con el color de fuente configurado
  curX = cardX;
  for (let c = 0; c < colCount; c++) {
    const w = colWidths[c];
    const text = data.headers[c] || "";
    const style = data.headerStyles[c] || {
      bg: "#0f172a",
      color: "#ffffff",
      bold: true
    };
    const isNum = data.rows.length > 0 && isNumeric(data.rows[0][c]);
    ctx.fillStyle = style.color;
    if (isNum) {
      ctx.textAlign = "right";
      ctx.fillText(text, curX + w - 14, cardY + headerHeight / 2);
    } else {
      ctx.textAlign = "left";
      ctx.fillText(text, curX + 14, cardY + headerHeight / 2);
    }
    curX += w;
  }

  // 2. Filas de datos (Cebra alternada con tipografía legible)
  ctx.font = fontBody;
  for (let r = 0; r < data.rows.length; r++) {
    const rowY = cardY + headerHeight + r * rowHeight;

    // Fondo alternado (zebra striping)
    ctx.fillStyle = r % 2 === 1 ? "#f8fafc" : "#ffffff";
    ctx.fillRect(cardX, rowY, cardW, rowHeight);

    // Borde inferior sutil
    ctx.fillStyle = "#f1f5f9";
    ctx.fillRect(cardX, rowY + rowHeight - 1, cardW, 1);

    // Valores de celdas
    ctx.fillStyle = "#1e293b";
    curX = cardX;
    for (let c = 0; c < colCount; c++) {
      const w = colWidths[c];
      const val = data.rows[r][c] || "";
      const isNum = isNumeric(val);
      if (isNum) {
        ctx.textAlign = "right";
        ctx.fillText(val, curX + w - 14, rowY + rowHeight / 2);
      } else {
        ctx.textAlign = "left";
        ctx.fillText(val, curX + 14, rowY + rowHeight / 2);
      }
      curX += w;
    }
  }
  ctx.restore();

  // Generar Blob con Promise
  const blob = await new Promise(resolve => canvas.toBlob(resolve, "image/png"));
  if (!blob) {
    throw new Error("No se pudo generar el blob de la imagen");
  }
  const cleanAddress = data.address.split("!").pop() || data.address;
  const dataUrl = canvas.toDataURL("image/png");

  // Intento 1: API moderna del portapapeles (solo si está permitida por Permissions Policy)
  if (isClipboardWriteAllowed() && navigator.clipboard && typeof ClipboardItem !== "undefined") {
    try {
      const item = new ClipboardItem({
        "image/png": blob
      });
      await navigator.clipboard.write([item]);
      return true; // Copiado directamente al portapapeles
    } catch (_) {
      // Continúa al flujo de visor y descarga web
    }
  }

  // Fallback 1: Copiar la imagen como HTML enriquecido en el portapapeles con dataURL
  const imgHtml = `<img src="${dataUrl}" alt="SuperCopy Excel Table" />`;
  const plainText = [data.headers.join("\t"), ...data.rows.map(r => r.join("\t"))].join("\n");
  copyHtmlFallback(imgHtml, plainText);

  // Fallback 2: Descarga directa del archivo PNG (solución infalible en Excel Web e iframes)
  downloadImage(canvas, data.address);

  // Fallback 3: Mostrar visor de imagen para clic derecho > "Copiar imagen" (ideal para WhatsApp Web)
  showImageResultModal(dataUrl, cleanAddress);
  return false; // Indica que se usó el flujo de visor/descarga web
}

/**
 * 3. FORMATO CORREO / OUTLOOK / HTML
 * Genera una tabla con estilos CSS inline respetando fielmente el fondo y color de texto del header de Excel.
 */
async function copyForEmail(data) {
  const tableStyle = "border-collapse: collapse; width: 100%; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; font-size: 13px; color: #1e293b; margin: 8px 0;";
  let html = `<table style="${tableStyle}"><thead><tr>`;
  data.headers.forEach((h, idx) => {
    const style = data.headerStyles[idx] || {
      bg: "#0f172a",
      color: "#ffffff",
      bold: true
    };
    const isNum = data.rows.length > 0 && isNumeric(data.rows[0][idx]);
    const align = isNum ? "right" : "left";
    const fontWeight = style.bold ? "700" : "600";
    const thStyle = `background-color: ${style.bg}; color: ${style.color}; font-weight: ${fontWeight}; padding: 10px 14px; text-align: ${align}; border: 1px solid ${style.bg};`;
    html += `<th style="${thStyle}">${escapeHtml(h)}</th>`;
  });
  html += `</tr></thead><tbody>`;
  data.rows.forEach((row, rIdx) => {
    const bg = rIdx % 2 === 1 ? "#f8fafc" : "#ffffff";
    html += `<tr style="background-color: ${bg};">`;
    row.forEach(cell => {
      const isNum = isNumeric(cell);
      const align = isNum ? "right" : "left";
      const tdStyle = `padding: 8px 14px; border: 1px solid #e2e8f0; text-align: ${align};`;
      html += `<td style="${tdStyle}">${escapeHtml(cell)}</td>`;
    });
    html += `</tr>`;
  });
  html += `</tbody></table>`;

  // Fallback a texto plano
  const plainText = [data.headers.join("\t"), ...data.rows.map(r => r.join("\t"))].join("\n");
  const cleanAddress = data.address.split("!").pop() || data.address;

  // Intento 1: API moderna del portapapeles con ClipboardItem (solo si está permitida por Permissions Policy)
  if (isClipboardWriteAllowed() && navigator.clipboard && typeof ClipboardItem !== "undefined") {
    try {
      const blobHtml = new Blob([html], {
        type: "text/html"
      });
      const blobText = new Blob([plainText], {
        type: "text/plain"
      });
      const item = new ClipboardItem({
        "text/html": blobHtml,
        "text/plain": blobText
      });
      await navigator.clipboard.write([item]);
      return true;
    } catch (_) {
      // Continúa al fallback sin error
    }
  }

  // Intento 2: Interceptor de evento 'copy'
  const isCopied = copyHtmlFallback(html, plainText);
  if (isCopied) {
    return true;
  }

  // Intento 3: Visor de resultados HTML en el panel
  showHtmlResultModal(html, plainText, cleanAddress);
  return false;
}

/**
 * Fallback universal para texto plano usando execCommand
 */
function copyTextFallback(text) {
  try {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.style.position = "fixed";
    textarea.style.top = "0";
    textarea.style.left = "0";
    textarea.style.width = "2em";
    textarea.style.height = "2em";
    textarea.style.padding = "0";
    textarea.style.border = "none";
    textarea.style.outline = "none";
    textarea.style.boxShadow = "none";
    textarea.style.background = "transparent";
    textarea.style.opacity = "0.01";
    document.body.appendChild(textarea);
    textarea.focus();
    textarea.select();
    textarea.setSelectionRange(0, textarea.value.length);
    const successful = document.execCommand("copy");
    document.body.removeChild(textarea);
    return !!successful;
  } catch (err) {
    return false;
  }
}

/**
 * Fallback universal para HTML usando el evento 'copy' y execCommand
 */
function copyHtmlFallback(html, plainText) {
  let success = false;
  const onCopy = e => {
    e.preventDefault();
    if (e.clipboardData) {
      e.clipboardData.setData("text/html", html);
      e.clipboardData.setData("text/plain", plainText);
      success = true;
    }
  };
  document.addEventListener("copy", onCopy);
  try {
    success = document.execCommand("copy") && success;
  } catch (err) {
    success = false;
  } finally {
    document.removeEventListener("copy", onCopy);
  }
  return success;
}

/**
 * Descarga automática de la imagen del Canvas como archivo PNG
 */
function downloadImage(canvas, address) {
  const cleanAddress = (address.split("!").pop() || address).replace(/[^a-zA-Z0-9_-]/g, "_");
  const link = document.createElement("a");
  link.download = `SuperCopy_${cleanAddress}.png`;
  link.href = canvas.toDataURL("image/png");
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

/**
 * Muestra el visor de texto para WhatsApp con selección automática
 */
function showTextResultModal(text, address) {
  currentModalText = text;
  currentModalAddress = address;
  const modal = document.getElementById("result-modal");
  const title = document.getElementById("result-modal-title");
  const viewText = document.getElementById("view-result-text");
  const viewImg = document.getElementById("view-result-image");
  const viewHtml = document.getElementById("view-result-html");
  const textarea = document.getElementById("result-textarea");
  if (!modal || !viewText || !textarea) return;
  if (title) title.textContent = `💬 WhatsApp / Chat (${address})`;
  viewText.classList.remove("hidden");
  viewImg?.classList.add("hidden");
  viewHtml?.classList.add("hidden");
  textarea.value = text;
  modal.classList.remove("hidden");

  // Auto-seleccionar texto inmediatamente para pulsar Ctrl + C
  setTimeout(() => {
    textarea.focus();
    textarea.select();
  }, 50);
}

/**
 * Muestra el visor de imagen con alta resolución y opciones de descarga/copiado
 */
function showImageResultModal(dataUrl, address) {
  currentModalDataUrl = dataUrl;
  currentModalAddress = address;
  const modal = document.getElementById("result-modal");
  const title = document.getElementById("result-modal-title");
  const viewText = document.getElementById("view-result-text");
  const viewImg = document.getElementById("view-result-image");
  const viewHtml = document.getElementById("view-result-html");
  const img = document.getElementById("result-img");
  if (!modal || !viewImg || !img) return;
  if (title) title.textContent = `🖼️ Imagen Ejecutiva (${address})`;
  viewImg.classList.remove("hidden");
  viewText?.classList.add("hidden");
  viewHtml?.classList.add("hidden");
  img.src = dataUrl;
  modal.classList.remove("hidden");
}

/**
 * Muestra el visor de tabla HTML para correo
 */
function showHtmlResultModal(html, plainText, address) {
  currentModalHtml = html;
  currentModalText = plainText;
  currentModalAddress = address;
  const modal = document.getElementById("result-modal");
  const title = document.getElementById("result-modal-title");
  const viewText = document.getElementById("view-result-text");
  const viewImg = document.getElementById("view-result-image");
  const viewHtml = document.getElementById("view-result-html");
  const container = document.getElementById("result-html-container");
  if (!modal || !viewHtml || !container) return;
  if (title) title.textContent = `📧 Tabla para Correo (${address})`;
  viewHtml.classList.remove("hidden");
  viewText?.classList.add("hidden");
  viewImg?.classList.add("hidden");
  container.innerHTML = html;
  modal.classList.remove("hidden");
}

/**
 * Cierra el modal universal de resultados
 */
function closeResultModal() {
  const modal = document.getElementById("result-modal");
  if (!modal) return;
  modal.classList.add("hidden");
}

/**
 * Descarga texto como archivo .txt
 */
function downloadTextFile(text, address) {
  const cleanAddress = address.replace(/[^a-zA-Z0-9_-]/g, "_");
  const blob = new Blob([text], {
    type: "text/plain;charset=utf-8"
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `SuperCopy_${cleanAddress}.txt`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * Renderiza la vista previa en el panel respetando los colores del header
 */
function renderPreview(data) {
  const container = document.getElementById("preview-container");
  if (!container) return;
  if (data.rowCount === 0) {
    container.innerHTML = `<p class="preview-placeholder">Selecciona un rango en Excel para ver la previsualización.</p>`;
    return;
  }
  let html = `<table class="preview-table"><thead><tr>`;
  data.headers.forEach((h, idx) => {
    const style = data.headerStyles[idx] || {
      bg: "#0f172a",
      color: "#ffffff",
      bold: true
    };
    html += `<th style="background-color: ${style.bg}; color: ${style.color};">${escapeHtml(h)}</th>`;
  });
  html += `</tr></thead><tbody>`;
  data.rows.slice(0, 5).forEach(r => {
    html += `<tr>`;
    r.forEach(c => {
      html += `<td>${escapeHtml(c)}</td>`;
    });
    html += `</tr>`;
  });
  if (data.rows.length > 5) {
    html += `<tr><td colspan="${data.colCount}" style="text-align: center; color: #94a3b8; font-style: italic;">... y ${data.rows.length - 5} filas más</td></tr>`;
  }
  html += `</tbody></table>`;
  container.innerHTML = html;
}

/**
 * Dibuja un rectángulo con esquinas redondeadas en Canvas
 */
function drawRoundedRect(ctx, x, y, w, h, radius) {
  ctx.beginPath();
  ctx.moveTo(x + radius, y);
  ctx.lineTo(x + w - radius, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + radius);
  ctx.lineTo(x + w, y + h - radius);
  ctx.quadraticCurveTo(x + w, y + h, x + w - radius, y + h);
  ctx.lineTo(x + radius, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - radius);
  ctx.lineTo(x, y + radius);
  ctx.quadraticCurveTo(x, y, x + radius, y);
  ctx.closePath();
}

/**
 * Muestra notificación flotante (Toast)
 */
function showToast(message) {
  const toast = document.getElementById("toast");
  const toastMsg = document.getElementById("toast-message");
  if (!toast || !toastMsg) return;
  toastMsg.textContent = message;
  toast.classList.remove("hidden");
  if (toastTimeout) clearTimeout(toastTimeout);
  toastTimeout = window.setTimeout(() => {
    toast.classList.add("hidden");
  }, 2500);
}

/**
 * Determina si un color hexadecimal es claro para ajustar el contraste de fuente
 */
function isLightColor(hexColor) {
  if (!hexColor) return false;
  const hex = hexColor.replace("#", "");
  if (hex.length !== 6) return false;
  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance > 0.65;
}

/**
 * Mapea el color de fondo dominante de Excel a un badge discreto para WhatsApp
 */
function getColorBadge(hexColor) {
  if (!hexColor) return "⬛";
  const hex = hexColor.replace("#", "").toUpperCase();
  if (hex.length !== 6) return "⬛";
  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);
  if (r < 60 && g < 60 && b < 60) return "⬛"; // Oscuro / Negro
  if (r > 200 && g > 200 && b > 200) return "⬜"; // Claro / Blanco
  if (g > r + 25 && g > b) return "🟩"; // Verde
  if (b > r + 20 && b > g) return "🟦"; // Azul
  if (r > 180 && g > 120 && b < 80) return "🟧"; // Naranja / Ámbar
  if (r > 170 && g < 80 && b < 80) return "🟥"; // Rojo
  if (r > 120 && b > 120 && g < 100) return "🟪"; // Morado
  return "⬛";
}

/**
 * Detecta si una cadena representa un valor numérico, moneda o porcentaje
 */
function isNumeric(val) {
  if (!val) return false;
  const clean = val.trim().replace(/[\$,€£%\s]/g, "").replace(/,/g, "");
  return !isNaN(Number(clean)) && clean !== "";
}
function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}