# Visa Expenses — Google Apps Script

🌐 [Read in English](README.md)

Lee automáticamente los correos de notificación de compras con tarjeta Visa desde Gmail y registra cada transacción en una hoja de cálculo de Google Sheets, organizada por mes de facturación.

> **Importante:** Este script está diseñado para un formato de correo específico (ver abajo). Para que funcione, tu banco debe enviar correos de notificación de transacciones Visa cuyo cuerpo en texto plano contenga campos con esta estructura:
>
> ```
> Fecha       05/04/2026
> Comercio    NOMBRE TIENDA
> Monto       $123.456          ← nacional (CLP)
> Monto       USD 99.50         ← internacional (USD)
> Cuotas      3
> Número tarjeta crédito ****1234
> ```
>
> Si tus correos de notificación usan un formato o nombres de campo diferentes, será necesario ajustar las expresiones regulares en el script.

## Requisitos Previos

| Requisito | Detalles |
|---|---|
| **Cuenta de Google** | Acceso a Gmail + Google Sheets |
| **Etiqueta de Gmail** | Una etiqueta llamada **`Visa`** aplicada a los correos de notificación de tu banco. Si quieres usar un nombre de etiqueta diferente, actualiza el string `'Visa'` en la llamada `GmailApp.getUserLabelByName('Visa')` en `script.js` con el nombre real de tu etiqueta. |
| **Google Apps Script** | El script debe estar vinculado (o tener acceso) a la hoja de cálculo destino |

## Cómo Funciona

### 1. Búsqueda de Correos

La función `registerVisaExpenses()` busca la etiqueta de Gmail **Visa** e itera sobre cada mensaje **no leído** en esos hilos.

### 2. Extracción de Datos

El cuerpo de cada correo se analiza con expresiones regulares para extraer:

| Campo | Patrón Regex | Ejemplo |
|---|---|---|
| Monto nacional (CLP) | `Monto $123.456` | `123456` (puntos eliminados, parseado como entero) |
| Monto internacional (USD) | `Monto USD 99.50` | `99.50` (parseado como decimal) |
| Fecha de compra | `Fecha 05/04/2026` | `05/04/2026` |
| Nombre del comercio | `Comercio NOMBRE TIENDA` | `NOMBRE TIENDA` |
| Últimos 4 dígitos de la tarjeta | `Número tarjeta crédito ****1234` | `1234` |
| Cuotas | `Cuotas 3` | `3` |

Si el correo contiene la palabra **"anulación"**, ambos montos se invierten (negativos) para que se resten de los totales.

### 3. Asignación de Hoja por Mes de Facturación

Las transacciones se colocan en una hoja cuyo nombre sigue el patrón:

```
Facturacion <Mes> <Año>
```

El mes de facturación se determina por la **fecha de compra**:

- Compras a partir del **día 20** de un mes se asignan al ciclo de facturación del **mes siguiente**.
- Compras antes del día 20 se mantienen en el ciclo de facturación del **mes actual**.

**Ejemplo:** Una compra con fecha `22/03/2026` va a la hoja `Facturacion Abril 2026`.

### 4. Creación Automática de Hoja

Cuando la hoja de un mes de facturación aún no existe, el script:

1. Crea la hoja con el nombre calculado.
2. Escribe encabezados en negrita en la primera fila:
   `Fecha | Comercio | Nacional | Internacional | Cuotas | Visa | Total Nacional | Total Internacional`
3. Congela la fila de encabezados.
4. Inserta fórmulas `=SUM()` en las celdas **G2** y **H2** para sumar automáticamente los montos nacionales e internacionales.
5. Aplica formato numérico a las celdas de totales: `$#,##0` para CLP y `USD #,##0.00` para USD.

### 5. Inserción de Filas

El script encuentra la primera fila vacía en la columna A (a partir de la fila 2) y escribe los datos extraídos ahí. El monto nacional (columna C) y el monto internacional (columna D) de cada fila se formatean como `$#,##0` y `USD #,##0.00` respectivamente. Las filas de anulación se resaltan con un fondo rojo claro (`#fff0f0`).

Después de escribir, el correo se **marca como leído** para no procesarlo de nuevo en la siguiente ejecución.

### 6. Manejo de Errores

Cualquier error al procesar un correo individual se captura y se registra en la consola de Apps Script sin detener el procesamiento de los correos restantes.

## Ejemplos de Correos

![Ejemplo 1](assets/example.png)

![Ejemplo 2](assets/example2.png)

## Configuración

1. Abre tu hoja de cálculo de Google Sheets destino.
2. Ve a **Extensiones → Apps Script**.
3. Pega el contenido de `script.js` en el editor.
4. Guarda el proyecto.
5. Crea una etiqueta en Gmail llamada **Visa** y aplícala a los correos de notificación de tu banco (puedes usar un filtro de Gmail para hacerlo automáticamente).

### Ejecución Manual

En el editor de Apps Script, selecciona `registerVisaExpenses` en el menú desplegable de funciones y haz clic en **Ejecutar**. Otorga los permisos de Gmail y Sheets cuando se soliciten.

### Ejecución Automática (Trigger)

1. En el editor de Apps Script ve a **Activadores** (ícono de reloj en la barra lateral).
2. Haz clic en **+ Añadir activador**.
3. Configura:
   - **Función:** `registerVisaExpenses`
   - **Fuente del evento:** Temporizador
   - **Tipo:** Temporizador en horas → **Cada hora** (o elige el intervalo que prefieras)
4. Guarda el activador.

## Estructura de la Hoja

![Ejemplo de estructura](assets/example-sheet-layout.png)

| Columna | Encabezado | Contenido |
|---|---|---|
| A | Fecha | Fecha de compra (`dd/mm/yyyy`) |
| B | Comercio | Nombre del comercio |
| C | Nacional | Monto nacional (CLP) |
| D | Internacional | Monto internacional (USD) |
| E | Cuotas | Número de cuotas |
| F | Visa | Últimos 4 dígitos de la tarjeta |
| G | Total Nacional | Total nacional (fórmula) |
| H | Total Internacional | Total internacional (fórmula) |

## Permisos Requeridos

- **Gmail** — leer mensajes y marcar como leídos (`GmailApp`)
- **Spreadsheet** — crear hojas, leer/escribir celdas (`SpreadsheetApp`)

## Contribuidores ⭐

<table>
  <tr>
    <td align="center"><a href="https://github.com/Phoebe-WD"><img src="https://avatars.githubusercontent.com/u/68600680?v=4" width="100px;" alt=""/><br /><sub><b>Phoebe Sttefi Wilckens Díaz</b></sub></a><br />💻 📖</td>
    <td align="center"><a href="https://github.com/leonardo-astete"><img src="https://avatars.githubusercontent.com/u/85591356?v=4" width="100px;" alt=""/><br /><sub><b>Leonardo Astete</b></sub></a><br />💻</td>
  </tr>
</table>

## Licencia

Este proyecto se proporciona tal cual para uso personal.
