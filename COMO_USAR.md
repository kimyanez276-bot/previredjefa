# Guía de Uso: Conta Marlene (Sistema de Gestión F29)

¡Bienvenida a tu nuevo sistema tributario y contable! Diseñado especialmente para Marlene, para ti y para su equipo de 3 personas.

---

## 🚀 ¿Cómo abrir y usar el programa?

1. Ve a la carpeta del proyecto: `c:\Users\Acer\Documents\testing\Conta Marlene`
2. Haz **doble clic** sobre el archivo **`Abrir_Conta_Marlene.bat`** o sobre **`index.html`** (se abrirá inmediatamente en Google Chrome, Microsoft Edge o tu navegador favorito).
3. ¡Listo! Funciona de forma 100% local, rápida y segura, sin necesidad de instalar programas difíciles ni escribir comandos.

---

## 👤 ¿Por qué no aparecía Don Guido y cómo verlo?

Si tenías la página abierta desde antes, el navegador guardó temporalmente una copia previa vacía y había un detalle en el archivo de respaldo que impedía que se mostrara.

**Ya lo hemos solucionado completamente:**
1. Al abrir la página o recargar con **F5**, la base de datos se actualizará sola a la versión **v3**, y **Don Guido Sepúlveda Méndez aparecerá de inmediato** en la lista principal con su RUT `12.790.291-7`.
2. Además, en la barra superior añadimos un botón amarillo que dice **"Ver Don Guido"**. Si en cualquier momento quieres restablecer su ficha o datos de prueba, solo haz clic ahí.

---

## 📑 Informe de Compras y Ventas (RCV) y Borrado de Facturas

Cuando hagas clic en el botón azul **"F29"** de cualquier cliente (por ejemplo, en Don Guido Sepúlveda):

1. **Dos Pestañas de Trabajo:**
   * **Pestaña 1 ("Cálculo & Resumen F29"):** Aquí está la planilla matemática exacta que usas habitualmente (Boletas, Transbank, Facturas, Notas de Crédito, Remanente UTM, PPM y Retenciones de Honorarios).
   * **Pestaña 2 ("Informe Detalle Documentos - RCV"):** Aquí verás el listado de cada una de las facturas, notas de crédito, boletas y comprobantes que componen ese mes.
2. **Cómo borrar facturas que no correspondan:**
   * En la tabla de documentos, al final de cada fila verás un **icono rojo de basurero**.
   * Si una factura no corresponde a ese período, o vino duplicada, o el cliente no la reconoce, simplemente haz clic en el icono de basurero.
   * El sistema te pedirá confirmación y **¡eliminará el documento recalculando al instante el Débito, Crédito, Remanente e Impuesto Líquido a Pagar del F29!**
3. **Agregar documentos manualmente o cargar RCV del SII:**
   * Puedes presionar **"Agregar Documento"** para incorporar cualquier factura o boleta puntual.
   * Puedes subir el archivo Excel o CSV del RCV descargado del SII y se cargarán los documentos automáticamente.
4. **Descargar e Imprimir el RCV:**
   * Tienes botones directos para **"Descargar Excel"** con la nómina detallada y para **"Imprimir Informe PDF"** con membrete formal de conciliación.

---

## 🛡️ Montos Exentos y Valor Otro Impuesto

Atendiendo tu solicitud, se conservan y respetan de manera estricta todas las líneas especiales:
* **Ventas Exentas:** Para ventas de libros, transporte de pasajeros u operaciones no gravadas.
* **Valor Otro Impuesto Venta:** Para impuestos adicionales como ILA (bebidas analcohólicas/alcohólicas, licores) u otros gravámenes.
* **Compras Exentas:** Compras de insumos o servicios exentos de IVA.
* **Valor Otro Impuesto Compra:** Crédito especial por impuestos específicos de compras.

Estas líneas se muestran en:
* La planilla de cálculo del F29.
* El informe detallado de documentos (RCV).
* La descarga general de Excel de la nómina del mes.
* El comprobante imprimible en PDF para entregar al cliente.

---

## 📂 Carpeta Tributaria Histórica y Gráficos (Por Cliente)

Al lado de cada cliente verás el botón **"Carpeta"** (o puedes hacer clic directo sobre el nombre del contribuyente):
* Se abrirá su **Carpeta Contable Digital**.
* **Gráficos interactivos:**
  1. *Ventas vs Compras:* Para comparar cómo creció la empresa mes a mes.
  2. *Impuesto Pagado vs Crédito a Favor:* Para ver claramente qué meses aportó al SII y qué meses acumuló remanente.
* **Selector de Años:** Puedes filtrar por año (ej: 2026, 2025 o Histórico Total).
* **Historial Cronológico:** Tabla con todos los períodos pasados y botones para **reabrir el F29** o **reimprimir cualquier comprobante antiguo en PDF**.

---

## 👥 ¿Cómo trabajar las 3 personas del equipo (oficina y casa)?

* **Al finalizar el día o al registrar nuevos clientes:**
  * Presiona el botón **"Respaldar"** arriba a la derecha.
  * Se descargará un archivo liviano con la fecha del día (ej: `Respaldo_Conta_Marlene_2026-08_2026-09-23.json`).
  * Compártelo por correo, WhatsApp o déjalo en Google Drive / OneDrive.
* **Para abrirlo en otra computadora (en casa o en la oficina):**
  * Abre la aplicación en esa computadora, presiona **"Cargar"** y selecciona el archivo.
  * En menos de un segundo, las 3 personas tendrán sincronizada la cartera de clientes, facturas y cálculos al 100%.

---

## ⚡ Automatizaciones para tus ~100 Clientes

1. **Actualizador Masivo de UTM:**
   * Ingresa la UTM del mes (ej: `71649`) en la barra superior y presiona **"Aplicar UTM"**.
   * Todos los clientes con crédito remanente actualizan su saldo en pesos en 1 segundo.
2. **Semáforo y Alertas:**
   * Control automático de quiénes vencen el **Día 12** y quiénes el **Día 20**.
   * Alerta del **Acuse de Recibo** para los días 30 o 31 del mes.
3. **Mensajes de WhatsApp Inteligentes:**
   * Si paga: Mensaje formal con el monto exacto y la fecha de vencimiento.
   * Si tiene saldo a favor: Mensaje cordial felicitándolo e informándole cuánto crédito le queda en pesos y en UTM para el mes siguiente.
4. **Dar de baja sin perder información:**
   * Puedes archivar contribuyentes inactivos para que no aparezcan en la vista diaria, pero su historial y declaraciones de años pasados jamás se perderán.
