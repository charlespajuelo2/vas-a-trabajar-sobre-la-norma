# E060 Acero — Usuario

Aplicación educativa en Python + Streamlit para consultar el libro `E060_Acero_Usuario.xlsx`.

## Fuente única

La aplicación lee exclusivamente `E060_Acero_Usuario.xlsx`. No se copian valores normativos al código: factores, unidades, artículos, páginas, estados y fórmulas se leen de las hojas `FUENTES`, `BARRAS`, `CALCULOS`, `ENTRADAS` y `VALIDACION`.

La lámina técnica se carga desde `outputs/lamina_detalle_acero_refuerzo.svg`.

## Ejecutar

```powershell
python -m pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud

- **Main file:** `app.py`
- **Dependencies:** `requirements.txt`
- **Excel source:** `E060_Acero_Usuario.xlsx` at the repository root.
- **Technical plate:** `outputs/lamina_detalle_acero_refuerzo.svg`.

The app uses relative paths based on `app.py`, so it does not depend on a local Windows path. No secrets or credentials are required. If secrets are added later, keep them only in Streamlit Cloud Secrets; `.streamlit/secrets.toml` is ignored by Git.

## Pruebas

```powershell
python tests.py
```

Las pruebas reproducen los seis casos de la hoja `VALIDACION` usando sus entradas y resultados esperados; no incluyen valores normativos escritos en `tests.py`.

## Advertencia

La app muestra siempre fórmula, unidades, referencia normativa y estado de trazabilidad. `NO VERIFICADO` se conserva y se presenta en rojo. Es una herramienta educativa, no apta para diseño estructural.
