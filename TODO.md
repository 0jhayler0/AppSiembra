# TODO: Preparar proyecto para despliegue

## 1. Mover ícono de flor
- [ ] Mover flower-icon.png de app-siembra/public/ a app-siembra/src/assets/

## 2. Convertir python-service a API Flask
- [ ] Convertir convertidor.py a servidor Flask con endpoint /convert
- [ ] Crear requirements.txt para python-service con dependencias (Flask, openpyxl, camelot-py, PyMuPDF, pdfplumber)

## 3. Actualizar backend
- [ ] Actualizar server.js: cambiar CORS para Firebase, remover dist serving, cambiar llamada a Python a API POST
- [ ] Remover build-script.js (ya no necesario)

## 4. Actualizar frontend
- [ ] Actualizar App.jsx: cambiar URL de axios a backend Render URL (usar placeholder por ahora)

## 5. Actualizar configuración de despliegue
- [ ] Actualizar render.yaml para servicios separados (backend y python-service)

## 6. Preparar para Firebase
- [ ] Asegurar que frontend se construya por separado para Firebase
