# Conciliación Bancaria — Etchepare y Asociados

## Archivos del proyecto
- `app.py` — servidor Flask con toda la lógica
- `bd_arrendatarios.json` — base de datos de arrendatarios (actualizar mensualmente)
- `templates/login.html` — pantalla de acceso
- `templates/index.html` — app principal
- `requirements.txt` — dependencias Python
- `Procfile` — comando de inicio para Railway

## Cómo subir a Railway

### 1. Crear cuenta en GitHub (si no tienes)
- Ve a github.com y crea una cuenta gratis

### 2. Crear repositorio en GitHub
- Clic en "New repository"
- Nombre: `conciliacion-bancaria`
- Privado (Private)
- Clic en "Create repository"

### 3. Subir archivos
- Arrastra todos los archivos de esta carpeta al repositorio de GitHub

### 4. Crear cuenta en Railway
- Ve a railway.app
- Clic en "Start a New Project"
- Conecta con tu cuenta de GitHub

### 5. Crear proyecto
- "Deploy from GitHub repo"
- Selecciona `conciliacion-bancaria`
- Railway detecta automáticamente que es Python

### 6. Configurar variables de entorno
En Railway → tu proyecto → Variables, agrega:
```
SECRET_KEY = cualquier-texto-secreto-largo
USER1_NAME = tu-usuario
USER1_PASS = tu-contraseña
USER2_NAME = nombre-compañera
USER2_PASS = contraseña-compañera
```

### 7. Obtener URL
- Railway te da una URL tipo: `conciliacion-bancaria.railway.app`
- Compártela con tu compañera

## Actualizar BD mensualmente
1. Abre `bd_arrendatarios.json` con cualquier editor de texto
2. Reemplaza el contenido con el nuevo JSON exportado
3. Sube el archivo actualizado a GitHub
4. Railway se actualiza automáticamente en 2 minutos

## Flujo diario
1. Entrar a la URL con usuario y contraseña
2. Subir el `Historial_2026_Actualizado.xlsx` desde Drive
3. Subir la cartola del banco del día
4. Clic en "Procesar"
5. Revisar los nuevos depósitos en pantalla
6. Clic en "Descargar Excel" — sale con formato completo
7. Subir el Excel descargado a Drive (reemplaza el anterior)
