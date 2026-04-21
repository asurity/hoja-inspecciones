# Instrucciones: Agregar Controles de Inspección Recurrente

## Fecha: 21/04/2026
## Módulo: frmChecklistVirtual
## Responsable: Desarrollador VBA

---

## 🎯 Objetivo
Agregar controles visuales en frmChecklistVirtual para capturar datos de inspecciones recurrentes (2da, 3ra, etc.).

---

## 📋 PASO 1: Abrir el diseñador VBA

1. En Excel VBA Editor (Alt+F11)
2. En explorador de proyecto, buscar `frmChecklistVirtual`
3. Doble clic para abrir el diseñador visual
4. Si no se ve en modo diseño, clic derecho → "View Code" → luego clic derecho en la pestaña → "View Object"

---

## 🔧 PASO 2: Agregar controles en fraCabecera

**UBICACIÓN:** Dentro del Frame `fraCabecera` (columna izquierda), después del último control existente (cboLugar)

### Control 1: Frame contenedor
- **Tipo:** Frame
- **Nombre:** `fraRecurrentInspection`
- **Caption:** ` Inspección Recurrente `
- Posición: Se configura por código (no importar posicionamiento manual)
- Tamaño: Se configura por código (no importar tamaño manual)

### Control 2: CheckBox principal
- **Tipo:** CheckBox
- **Nombre:** `chkEsRecurrente`
- **Caption:** `Esta NO es la primera inspección de este personal`
- Ubicación: **DENTRO** de `fraRecurrentInspection`

### Control 3: Botón búsqueda
- **Tipo:** CommandButton
- **Nombre:** `btnBuscarHistorico`
- **Caption:** `🔍 Buscar historial`
- Ubicación: **DENTRO** de `fraRecurrentInspection`

### Control 4: Label info histórico
- **Tipo:** Label
- **Nombre:** `lblInfoHistorico`
- **Caption:** `(Información de inspecciones previas aparecerá aquí)`
- Ubicación: **DENTRO** de `fraRecurrentInspection`

### Control 5: Label número inspección
- **Tipo:** Label
- **Nombre:** `lblNumeroInspeccion`
- **Caption:** `Esta es la inspección número:`
- Ubicación: **DENTRO** de `fraRecurrentInspection`

### Control 6: TextBox número inspección
- **Tipo:** TextBox
- **Nombre:** `txtNumeroInspeccion`
- **Text:** `(vacío)`
- Ubicación: **DENTRO** de `fraRecurrentInspection`

### Control 7: Label RPN Anterior
- **Tipo:** Label
- **Nombre:** `lblRPNAnterior`
- **Caption:** `RPN anterior:`
- Ubicación: **DENTRO** de `fraRecurrentInspection`

### Control 8: TextBox RPN Anterior Automático
- **Tipo:** TextBox
- **Nombre:** `txtRPNAnteriorAuto`
- **Text:** `(vacío)`
- Ubicación: **DENTRO** de `fraRecurrentInspection`

### Control 9: TextBox RPN Anterior Manual
- **Tipo:** TextBox
- **Nombre:** `txtRPNAnteriorManual`
- **Text:** `(vacío)`
- Ubicación: **DENTRO** de `fraRecurrentInspection`

### Control 10: Label estado modo
- **Tipo:** Label
- **Nombre:** `lblModoRPN`
- **Caption:** `[Modo RPN: no determinado]`
- Ubicación: **DENTRO** de `fraRecurrentInspection`

---

## ✅ VERIFICACIÓN

Después de agregar los controles en el diseñador:

- [ ] `fraRecurrentInspection` existe y es tipo Frame
- [ ] `chkEsRecurrente` existe y es tipo CheckBox
- [ ] `btnBuscarHistorico` existe y es tipo CommandButton
- [ ] `lblInfoHistorico` existe y es tipo Label
- [ ] `lblNumeroInspeccion` existe y es tipo Label
- [ ] `txtNumeroInspeccion` existe y es tipo TextBox
- [ ] `lblRPNAnterior` existe y es tipo Label
- [ ] `txtRPNAnteriorAuto` existe y es tipo TextBox
- [ ] `txtRPNAnteriorManual` existe y es tipo TextBox
- [ ] `lblModoRPN` existe y es tipo Label

**NO ES NECESARIO** configurar propiedades como tamaño, posición, colores, fuentes, etc.  
Todo eso se configura **automáticamente por código** en `ConfigurarCabecera()`.

---

## 🔄 SIGUIENTE PASO

Una vez agregados los controles en el diseñador:

1. **Guardar el formulario** (Ctrl+S)
2. **Aplicar el código VBA** que configura posicionamiento, estilos y eventos
3. **Probar** el formulario para verificar que los controles aparecen correctamente

---

## ⚠️ NOTAS IMPORTANTES

- Los controles se posicionan **dentro** de `fraRecurrentInspection`
- El frame `fraRecurrentInspection` se posiciona **dentro** de `fraCabecera`
- La jerarquía es: `fraCabecera` → `fraRecurrentInspection` → [controles]
- No importa el orden visual inicial, el código los reorganiza
- La visibilidad inicial de algunos controles se gestiona por código

---

## 📞 CONTACTO

Si hay dudas sobre la ubicación de los controles, revisar:
- `docs/TODO_RPN_AJUSTADO.md` - Sección 2.1: Diseño de controles nuevos
- `Configuration2.bas` - Documentación de columnas 32-40
