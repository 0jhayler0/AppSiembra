# App Siembra - Frontend

Aplicación React para generar reportes de siembra desde archivos Excel/PDF.

## Despliegue en Firebase

1. Instalar Firebase CLI:
   ```bash
   npm install -g firebase-tools
   ```

2. Iniciar sesión en Firebase:
   ```bash
   firebase login
   ```

3. Crear proyecto en Firebase Console y obtener el project ID.

4. Inicializar Firebase Hosting:
   ```bash
   firebase init hosting
   ```

5. Construir la aplicación:
   ```bash
   npm run build
   ```

6. Desplegar:
   ```bash
   firebase deploy
   ```

## Variables de entorno

Crear archivo `.env` en la raíz del proyecto:

```
VITE_BACKEND_URL=https://tu-backend-url.onrender.com
```

## Desarrollo local

```bash
npm install
npm run dev
```
