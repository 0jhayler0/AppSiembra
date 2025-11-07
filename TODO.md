 # Plan para Desplegar App Siembra en Render como Full-Stack

## Información Recopilada
- **Backend**: Servidor Express en Node.js (server.js), maneja uploads y genera reportes. Usa puerto 5000 por defecto.
- **Frontend**: App React con Vite, construye a dist/. Actualmente apunta a localhost:5000.
- **Estrategia**: Configurar backend para servir frontend estático después de build, desplegar como un solo Web Service en Render.

## Plan Detallado
- [x] Actualizar package.json del backend para incluir dependencias del frontend y scripts de build.
- [x] Modificar server.js para servir archivos estáticos del frontend y manejar rutas SPA.
- [x] Actualizar CORS en server.js para permitir la URL de producción de Render.
- [x] Probar build localmente.
- [ ] Configurar despliegue en Render (Web Service).

## Archivos Dependientes
- backend/package.json: Agregar dependencias y scripts.
- backend/server.js: Agregar middleware para servir static y SPA fallback.

## Pasos de Seguimiento
- [x] Ejecutar npm install en backend después de cambios.
- [x] Probar npm run build en backend.
- [x] Probar npm start y verificar que sirva frontend.
- [ ] Subir código a GitHub.
- [ ] Crear Web Service en Render apuntando al repo, con comando build y start.
