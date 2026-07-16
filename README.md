# Analisis de Pedidos

Dashboard web para analizar ventas por asesor, generar informes con IA y exportar reportes en Word/Excel.

## Ejecutar en local

```bash
pip install -r requirements.txt
```

Crear archivo `.env` con tu API key de Gemini:
```
GEMINI_API_KEY=tu_api_key_aqui
```

Ejecutar:
```bash
uvicorn app:app --reload --port 8000
```

Abrir http://127.0.0.1:8000

## Deploy en Render

1. Crear cuenta en [render.com](https://render.com)
2. Conectar el repositorio de GitHub
3. Configurar la variable de entorno `GEMINI_API_KEY` en el dashboard de Render
4. Deploy automatico al hacer push

## Funcionalidades

- Carga archivos Excel de pedidos
- Pivot table automatica por canal de distribucion
- Metricas por asesor: pedida, pendiente, comprometida, backlog, valor, margen
- Generacion de informes con Gemini AI
- Exportacion a Word (.docx) y Excel (.xlsx)
