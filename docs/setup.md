# Setup del sistema RSVP

## 1. Crear Google Sheet

Crea un Google Sheet con 3 pestanas:

### Pestana "Invitados"
Columnas (primera fila como headers exactos):
```
code | familia | nombre | es_nino | confirmado | menu | notas | tag | actualizado
```

### Pestana "Codigos"
```
code | familia | enviado | fecha_envio
```

### Pestana "Resumen"
Formulas sugeridas (ajustar rangos):
- Total invitados: `=COUNTA(Invitados!C2:C)`
- Confirmados (si): `=COUNTIF(Invitados!E2:E, TRUE)`
- Confirmados (no): `=COUNTIF(Invitados!E2:E, FALSE)`
- Sin responder: `=COUNTBLANK(Invitados!E2:E)`
- Menu normal: `=COUNTIF(Invitados!F2:F, "normal")`
- Menu vegetariano: `=COUNTIF(Invitados!F2:F, "vegetariano")`
- Menu infantil: `=COUNTIF(Invitados!F2:F, "infantil")`

## 2. Importar datos

Copia los datos del CSV existente al tab "Invitados".
Asegurate de que la columna "familia" tenga el mismo valor para todos los miembros de una familia.

## 3. Configurar Apps Script

1. En el Sheet, ve a **Extensiones > Apps Script**
2. Borra el contenido de Code.gs
3. Pega el contenido de `docs/apps-script.js`
4. Guarda (Ctrl+S)
5. Refresca el Sheet, aparecera el menu **RSVP**
6. Click en **RSVP > Generar codigos** para asignar codigos a cada familia

## 4. Desplegar como Web App

1. En Apps Script, click en **Implementar > Nueva implementacion**
2. Tipo: **Aplicacion web**
3. Ejecutar como: **Yo**
4. Quien tiene acceso: **Cualquier persona**
5. Click en **Implementar**
6. Copia la URL generada

## 5. Conectar el frontend

En `rsvp.html`, reemplaza `YOUR_APPS_SCRIPT_URL_HERE` con la URL del paso anterior.

## 6. Probar

Abre: `https://maoaiz.github.io/wedding/rsvp.html?code=CODIGO_DE_PRUEBA`

## 7. Generar mensajes de WhatsApp

En el Sheet, click en **RSVP > Generar mensajes WhatsApp**.
Se creara una pestana "Mensajes" con los textos listos para copiar y enviar.
