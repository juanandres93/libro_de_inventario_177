# Libro de Inventarios 2026 — AdaptaPro

**Método:** Promedio Ponderado Periódico Mensual  
**Versión:** 1.1.0  
**Moneda:** Bs.  
**Licencia:** Open Source  

Creado por OpenCode bajo instrucciones del Equipo de desarrollo AdaptaPro  
🌐 [adaptaproerp.com/blog_openconde](https://adaptaproerp.com/blog_openconde/)

---

## Descarga

| Archivo | Tipo | Uso |
|---------|------|-----|
| `LIBRO_INVENTARIOS_2026.xlsm` | Excel con macros | Funciones automáticas completas |
| `LIBRO_INVENTARIOS_2026_FORMULAS.xlsx` | Excel solo fórmulas | Sin macros, 100% compatible |

---

## Modalidad CON macros (`.xlsm`)

**Archivo:** `LIBRO_INVENTARIOS_2026.xlsm`

### Beneficios
- **Apertura automática de meses** — botón que crea y abre el siguiente mes con un clic
- **Cierre de mes con validación** — verifica saldos, fórmulas y consistencia antes de cerrar
- **Bitácora de auditoría** — registro automático de cada acción (apertura, cierre, respaldo)
- **Respaldo instantáneo** — genera copia de seguridad del libro con un clic
- **Resumen anual automático** — consolida los 12 meses en una sola vista
- **Validación de datos** — rechaza automáticamente entradas negativas, texto en campos numéricos o códigos duplicados
- **Alta de materiales asistida** — formulario guiado para agregar nuevos artículos al catálogo

### Requisitos
- Microsoft Excel 2016 o superior (Windows)
- Hacer clic en **"Habilitar contenido"** la primera vez que se abre el libro

### Flujo típico
1. Abrir libro → Habilitar contenido
2. **CONFIGURACION**: capturar año y datos de la empresa
3. **CATALOGO**: registrar materiales (código, descripción, unidad, costo)
4. **BtnAbrirMesSiguiente**: crear y abrir el mes (ej: 2026-ENE)
5. **Movimiento**: registrar entradas y salidas en el mes abierto
6. **BtnValidarMes**: verificar consistencia
7. **BtnCerrarMes**: cerrar el mes (queda bloqueado)
8. **BtnResumenAnual**: ver consolidado

---

## Modalidad SIN macros (`.xlsx`)

**Archivo:** `LIBRO_INVENTARIOS_2026_FORMULAS.xlsx`

### Beneficios
- **Cero configuración** — se abre directamente sin barras de advertencia
- **Compatible con cualquier Excel** — funciona en Excel 2010+, LibreOffice, Google Sheets (importación)
- **Sin riesgo de seguridad** — no contiene código VBA ejecutable
- **Fórmulas transparentes** — toda la lógica de cálculo es visible en las celdas
- **Portable** — se puede compartir por correo, USB o la nube sin restricciones
- **Auditable** — cualquier usuario puede verificar las fórmulas manualmente

### Qué funciona sin macros
- **Cálculo de saldos** — columnas K (saldo cantidad), L (saldo valorizado), O (costo promedio)
- **Promedio ponderado periódico mensual** — fórmula nativa: `=(D+F)/(C+E)`
- **Inventario valorizado** — cálculo automático en cada fila
- **Resumen anual** — consolidación por fórmulas

### Qué requiere atención manual
- **Apertura de meses** — copiar la estructura del mes anterior y renombrar
- **Cierre de meses** — proteger manualmente la hoja (Revisar → Proteger hoja)
- **Validación** — revisar manualmente que no haya saldos negativos
- **Respaldo** — copiar el archivo manualmente

---

## Estructura del libro

| Hoja | Función |
|------|---------|
| `CONFIGURACION` | Año, datos de empresa, logo AdaptaPro |
| `CATALOGO` | Maestro de materiales |
| `2026-ENE` … `2026-DIC` | Movimientos mensuales |
| `RESUMEN_ANUAL` | Consolidado de los 12 meses |
| `DETALLE_OTROS` | Materiales sin movimiento |
| `BITACORA` | Registro de acciones (solo con macros) |
| `AYUDA` | Instrucciones de uso |
| `_ESTADOS` | Control interno de estados |
| `_INSTANTANEAS` | Copias internas (solo con macros) |

---

## Estructura de columnas (meses)

| Columna | Contenido |
|---------|-----------|
| A | Código del material |
| B | Descripción |
| C | Cantidad entrada |
| D | Costo entrada |
| E | Cantidad salida |
| F | Costo salida |
| G | Cantidad devolución |
| H | Costo devolución |
| I | Cantidad traslado |
| J | Costo traslado |
| K | **Saldo cantidad** (fórmula) |
| L | **Saldo valorizado** (fórmula) |
| M | Unidad |
| N | Observación |
| O | **Costo promedio** (fórmula) |

### Fórmulas clave
```
K7 = (C7+E7)-G7-I7          ← saldo cantidad
O7 = IF((C7+E7)>0,(D7+F7)/(C7+E7),0)  ← costo promedio ponderado
L7 = K7*O7                   ← saldo valorizado
```

---

## Entorno probado

| Componente | Valor |
|------------|-------|
| Sistema operativo | Windows 11 Home 10.0.26200 |
| Excel | Microsoft 365 Build 19127.20800 (x86 / 32 bits) |
| Región | es-Venezuela (decimal coma, miles punto) |
| VBA | Excel 16.0 VBA7 (sin API de Windows) |

> **Nota:** Probado en Excel 32 bits. En 64 bits debe funcionar igualmente (no se usa API de Windows).

---

## Archivos incluidos

```
entrega/
├── LIBRO_INVENTARIOS_2026.xlsm          ← versión con macros
├── LIBRO_INVENTARIOS_2026_FORMULAS.xlsx ← versión sin macros
├── LIBRO_INVENTARIOS_2026_PRUEBAS.xlsm  ← libro de pruebas (suite T01-T28)
├── codigo/                               ← código VBA exportado (17 módulos)
├── capturas/                             ← capturas de pantalla de validación
├── BOTONES_MACROS_ENTREGA.txt            ← inventario de botones (entrega)
├── BOTONES_MACROS_PRUEBAS.txt            ← inventario de botones (pruebas)
├── COMPILACION_ENTREGA.txt               ← resultado de compilación
├── COMPILACION_PRUEBAS.txt               ← resultado de compilación (pruebas)
├── ENTORNO_PROBADO.txt                   ← entorno de pruebas
├── INFORME_AUTOCONTROL_T01_T28.txt       ← suite de pruebas completa
├── INFORME_PRUEBAS.txt                   ← informe de pruebas
├── LISTA_REVISION_EN_FRIO.txt            ← checklist de revisión
├── MANUAL_OPERACION.txt                  ← manual de operación
├── MANIFIESTO_SHA256.txt                 ← hashes de integridad
├── PROCEDIMIENTO_CONSTRUCCION.txt        ← cómo se construyó el libro
├── README.md                             ← este archivo
└── SMOKE.txt                             ← prueba de humo
```

---

## Verificación de integridad

Para verificar que el archivo no fue alterado:

```powershell
Get-FileHash -Path "LIBRO_INVENTARIOS_2026.xlsm" -Algorithm SHA256
```

Comparar el resultado con el hash registrado en `MANIFIESTO_SHA256.txt`.

---

## Contribuir

1. Hacer un **fork** del repositorio
2. Crear una rama: `git checkout -b mi-mejora`
3. Commit: `git commit -m "Descripción del cambio"`
4. Push: `git push origin mi-mejora`
5. Abrir un **Pull Request**

> Al hacer fork, tu copia es independiente del original. Puedes modificarla sin afectar el repositorio base.

---

## Contacto

- **Web:** [adaptaproerp.com](https://adaptaproerp.com/blog_openconde/)
- **Equipo:** Desarrollo AdaptaPro
- **Creado por:** OpenCode
