# SuperCopy Pro - Paquete de Despliegue de Excel Add-in
Generado con ScriptLab to Add-in Studio (INTEF - Inteligencia Eficiente)

🌐 **URL Pública configurada (GitHub Pages / Hosting):**
- **Sitio Web del Complemento:** [https://github.com/juannbr/Excel-Super-Copy/](https://github.com/juannbr/Excel-Super-Copy/)
- **Punto de Entrada (Taskpane):** [https://github.com/juannbr/Excel-Super-Copy/index.html](https://github.com/juannbr/Excel-Super-Copy/index.html)

---

## 📁 Contenido del paquete:
- `index.html`: Página web del panel con el SDK de Office.js integrado.
- `style.css`: Estilos visuales del complemento.
- `app.js`: Lógica en JavaScript nativo listo para el navegador.
- `manifest.xml`: Manifiesto configurado con tus URLs para registrar el botón en la cinta de Excel.
- `assets/`: Carpeta con los íconos oficiales del complemento (`icon-16.png`, `icon-32.png`, `icon-80.png`).

---

## 🚀 Pasos para publicar y usar:

### 1. Publicar los archivos web:
Sube `index.html`, `style.css`, `app.js` y la carpeta `assets/` a la raíz de tu repositorio de GitHub Pages o hosting:
- Enlace público: **https://github.com/juannbr/Excel-Super-Copy/**
- Verifica que abra correctamente en tu navegador con candado seguro HTTPS.

### 2. Cargar en Excel Desktop (Catálogo Compartido):
1. Guarda tu archivo `manifest.xml` en una carpeta local de tu equipo (o compártela en red).
2. En Excel ve a: **Archivo** ➔ **Opciones** ➔ **Centro de confianza** ➔ **Configuración del Centro de confianza** ➔ **Catálogos de complementos de confianza**.
3. Escribe la ruta de tu carpeta, haz clic en **Agregar catálogo** y activa la casilla **Mostrar en el menú**.
4. Reinicia Excel.
5. Ve a la pestaña **Insertar** ➔ **Mis complementos** ➔ pestaña **Carpeta compartida** y haz clic en **SuperCopy Pro**.

### 3. Cargar en Excel para la Web (Sideloading rápido):
1. Abre cualquier libro en [Excel para la Web (office.com)](https://excel.new).
2. Ve a la pestaña **Insertar** ➔ **Complementos** ➔ **Gestionar mis complementos** ➔ **Cargar mi complemento**.
3. Selecciona el archivo `manifest.xml` de tu equipo y el botón aparecerá en tu cinta de opciones.
