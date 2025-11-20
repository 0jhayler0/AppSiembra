# Guía de Despliegue en Render

## ✅ Cambios Realizados

1. ✅ Creado `.gitignore` en la raíz del proyecto
2. ✅ Creado `render.yaml` con configuración de despliegue usando Docker
3. ✅ Creado `Dockerfile` para incluir Node.js y Python
4. ✅ Actualizado `backend/package.json` con script compatible con Linux
5. ✅ Verificado que el servidor está configurado para usar `process.env.PORT`

## 🚀 Pasos para Desplegar en Render

### Paso 1: Preparar el Repositorio
```bash
# Asegúrate de que todo esté commiteado en Git
git add .
git commit -m "Preparar para despliegue en Render"
git push origin main
```

### Paso 2: Conectar con Render
1. Ve a [https://render.com](https://render.com)
2. Crea una cuenta o inicia sesión
3. Haz clic en **"New +"** → **"Web Service"**
4. Selecciona **"Deploy an existing repository"**
5. Conecta tu repositorio de GitHub (0jhayler0/AppSiembraupdate)

### Paso 3: Configurar el Servicio
En la página de configuración del nuevo Web Service:

- **Name:** `appsiembra-backend` (o el nombre que prefieras)
- **Environment:** `Node`
- **Region:** Elige la más cercana a ti (ej: "Ohio", "Frankfurt")
- **Branch:** `main`
- **Build Command:** `npm install && npm run build`
- **Start Command:** `npm start`
- **Root Directory:** `backend` ← Importante, debe apuntar a la carpeta backend

### Paso 4: Configurar Variables de Entorno
En la sección **"Environment Variables"**, agrega:

| Key | Value |
|-----|-------|
| `NODE_ENV` | `production` |
| `FRONTEND_URL` | `https://your-app-name.onrender.com` (después de desplegar) |

### Paso 5: Deploy
Haz clic en **"Create Web Service"** y espera a que Render:
1. Clone tu repositorio
2. Instale dependencias
3. Compile el frontend
4. Inicie el servidor

### Paso 6: Obtén tu URL
Después del despliegue exitoso, Render te asignará una URL pública como:
```
https://appsiembra-backend.onrender.com
```

### Paso 7: Actualiza CORS (Opcional)
Si deseas permitir acceso desde otros dominios, puedes actualizar `backend/server.js`:

```javascript
app.use(cors({
  origin: [
    'http://localhost:5173',
    'http://127.0.0.1:5173',
    'https://appsiembra-backend.onrender.com',
    process.env.FRONTEND_URL || ''
  ],
  methods: ['GET','POST','OPTIONS'],
  exposedHeaders: ['Content-Disposition']
}));
```

## 🔧 Configuración Verificada

✅ **Node.js:** `>=18.0.0` (especificado en `package.json`)
✅ **Puerto:** Configurado para usar `process.env.PORT` (5000 por defecto)
✅ **Build Command:** Compatible con Linux
✅ **Start Command:** `npm start` → `node server.js`
✅ **Frontend Build:** Se construye automáticamente en el build command
✅ **Archivos Estáticos:** Servidos desde `backend/dist`

## ⚠️ Notas Importantes

- **Uploads:** Los archivos subidos se guardarán en `backend/output/uploads/` que es volátil en Render
  - Para producción, considera usar S3 o un servicio de almacenamiento
  
- **Python Script:** Si necesitas usar `convertidor.py`, Render debe tener Python instalado
  - Crea un `Procfile` adicional si es necesario
  
- **Tiempo de Build:** El primer despliegue puede tomar 5-10 minutos

## 🐛 Troubleshooting

### "Build failed"
- Revisa los logs en Render
- Asegúrate de que `backend/` está en el directorio correcto
- Verifica que todas las dependencias en `package.json` sean correctas

### "Cannot GET /"
- El frontend no se compiló correctamente
- Verifica que `dist/` existe después del build

### Puerto no responde
- Render puede tomar algunos minutos para inicializar
- Revisa los logs: "🚀 Servidor corriendo en puerto X"
