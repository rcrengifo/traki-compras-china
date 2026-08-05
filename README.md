# Traki · Compras China

Herramienta para gestionar cotizaciones e importaciones desde China.

- **Subir cotización** (Excel de China) → lee productos y **fotos** automáticamente.
- **Aprobar / quitar** productos (lo que decide el jefe).
- **Tablero** con estado de aprobación y etapa de importación.
- **Buscar** en el histórico con suma de cantidades.
- **Calculadora de contenedor** (por volumen y por peso).

## Correr en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Sin configuración usa una base SQLite local (`compras.db`). Es solo para pruebas.

## Base de datos en la nube (Supabase)

1. En Supabase: **SQL Editor** → pega y ejecuta el contenido de [`supabase_schema.sql`](supabase_schema.sql).
2. **Project Settings → Database → Connection string → URI**, copia la cadena y reemplaza `[YOUR-PASSWORD]`.
3. Local: copia `.streamlit/secrets.toml.example` a `.streamlit/secrets.toml` y pega la cadena en `DATABASE_URL`.

## Desplegar en Streamlit Community Cloud

1. Sube este proyecto a un repo de GitHub.
2. En share.streamlit.io → **New app** → elige el repo y `app.py`.
3. En **Advanced settings → Secrets** pega:
   ```toml
   DATABASE_URL = "postgresql://postgres:TU-PASSWORD@db.xxxx.supabase.co:5432/postgres"
   ```
4. Deploy. Para que sea privada: **Settings → Sharing** → solo el correo de tu esposa.
