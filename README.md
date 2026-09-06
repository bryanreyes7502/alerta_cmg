# Sistema de Alertas Automáticas – GitHub Actions

Este proyecto reemplaza la ejecución permanente en el notebook por una ejecución automática en GitHub Actions.

## Qué hace

- Consulta el CSV de energía del Coordinador Eléctrico Nacional.
- Aplica la misma lógica de selección de centrales y zonas del script original.
- Detecta novedades mediante hashes.
- Genera el Excel con el mismo formato general.
- Envía el correo con el Excel adjunto.
- Conserva los hashes en `state/state.json` para no repetir alertas entre ejecuciones.
- Se ejecuta cada 5 minutos y no requiere que el notebook esté encendido.

## Configuración en GitHub

### 1. Crear el repositorio

Sube todos estos archivos a un repositorio de GitHub.

Para mantener la ejecución cada 5 minutos sin consumir la cuota de minutos de Actions, la opción más conveniente es un repositorio **público**: GitHub indica que los runners estándar hospedados por GitHub son gratuitos en repositorios públicos. En repositorios privados, GitHub Free incluye 2.000 minutos/mes, lo que no alcanza para una ejecución cada 5 minutos de forma continua. citeturn851906search0turn851906search3

El repositorio público no expone la contraseña del correo: la contraseña se guarda como Secret de GitHub y no se incluye en el código.

### 2. Guardar la contraseña del correo

No pongas la contraseña real dentro de `main.py`.

En GitHub abre:

`Settings → Secrets and variables → Actions → New repository secret`

Crea:

- **Name:** `EMAIL_APP_PASS`
- **Secret:** la contraseña real de `alertas.automaticas@rasopower.cl`

GitHub permite almacenar secretos cifrados y utilizarlos dentro de los workflows sin escribirlos en el código. citeturn752971search2turn752971search4

### 3. Subir los archivos

La estructura debe quedar así:

```text
alertas_pm_github/
├── .github/
│   └── workflows/
│       └── alertas.yml
├── descargas_csv/
│   └── .gitkeep
├── reportes/
│   └── .gitkeep
├── state/
│   └── state.json
├── .gitignore
├── main.py
├── requirements.txt
└── README.md
```

### 4. Ejecutar una primera prueba

Ve a:

`Actions → Alertas operación de centrales → Run workflow`

La ejecución manual permite comprobar la conexión con el Coordinador y el envío de correo antes de dejar el proceso automático.

### 5. Dejarlo funcionando

El workflow queda programado cada 5 minutos. GitHub documenta 5 minutos como el intervalo mínimo para workflows `schedule`. También pueden producirse retrasos ocasionales en la ejecución programada. citeturn752971search0turn752971search5

## Por qué ya no uso Selenium

El código original abría Chrome únicamente para acceder al mismo endpoint que entrega el CSV. En GitHub no hace falta mantener un navegador abierto: esta versión descarga directamente el recurso mediante HTTP.

Se eliminaron:

- Selenium
- ChromeDriver
- webdriver-manager
- la ventana de Chrome
- la espera de descarga del navegador
- los pitidos de Windows
- el `while True` permanente

El servidor de GitHub ejecuta una instancia del programa, termina y vuelve a ejecutarla cinco minutos después.

## Cómo se mantiene el control de novedades

En el notebook, `prev_hashes` permanecía en memoria durante toda la ejecución.

En GitHub cada ejecución comienza desde cero, por lo que `state/state.json` guarda los hashes de la última alerta enviada. Después de una alerta exitosa, GitHub Actions hace commit de ese archivo al repositorio.

Si el envío del correo falla, el estado **no** se actualiza. De esta forma, la misma alerta puede reintentarse en la próxima ejecución.

## Zona horaria

El programa usa `America/Santiago`. Así, la fecha usada para consultar el Coordinador corresponde a Chile y no a la hora UTC del servidor.

## Seguridad

Nunca subas la contraseña real al repositorio. Si alguna vez la contraseña queda expuesta en un commit, cambia esa contraseña inmediatamente.
