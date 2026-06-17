# 🧾 Facturas del Exterior — ESSENTTA / SERPRESS

App para procesar facturas de proveedores del exterior, calcular IGV no domiciliado y retención de renta, y generar el Excel para SUNAT.

## Funcionalidades

- ✅ Lee PDFs de invoices automáticamente con IA (Claude)
- ✅ Carga el tipo de cambio SUNAT desde PDF
- ✅ Calcula IGV (18%), retención renta (30%), totales en S/. y USD
- ✅ Cruza constancias NPS con invoices para llenar N° de orden
- ✅ Genera Excel con formato exacto, una pestaña por empresa
- ✅ Soporta múltiples empresas (ESSENTTA, SERPRESS, etc.)

## Instalación y despliegue

### Paso 1 — Subir a GitHub

1. Crea un repositorio nuevo en GitHub (puede ser privado)
2. Sube todos estos archivos:
   - `app.py`
   - `requirements.txt`
   - `.streamlit/secrets.toml.example`

### Paso 2 — Configurar en Streamlit Cloud

1. Ve a [share.streamlit.io](https://share.streamlit.io)
2. Haz clic en **New app**
3. Conecta tu repositorio de GitHub
4. En **Advanced settings → Secrets**, agrega:

```toml
ANTHROPIC_API_KEY = "sk-ant-TU_API_KEY_AQUI"
```

5. Haz clic en **Deploy**

### Paso 3 — Compartir con el equipo

Una vez desplegada, copia el link (ej: `https://tuapp.streamlit.app`) y compártelo con tu equipo. No necesitan cuenta de ningún tipo.

## Cómo obtener tu API key de Anthropic

1. Ve a [console.anthropic.com](https://console.anthropic.com)
2. Inicia sesión o crea cuenta
3. Ve a **API Keys** → **Create Key**
4. Copia la key y pégala en Streamlit Secrets

## Uso mensual

1. **Configurar** empresa y mes en el panel lateral
2. **Subir PDF** de tipo de cambio SUNAT (se descarga de sunat.gob.pe)
3. **Subir PDFs** de todos los invoices del mes
4. **Revisar** la tabla y corregir si es necesario
5. **Subir NPS** de IGV para llenar los números de orden
6. **Descargar Excel** final

Para procesar otra empresa, cambia el nombre en el panel lateral y repite el proceso — los datos de ambas empresas se acumulan en el mismo Excel.
