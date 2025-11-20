const { execSync } = require('child_process');
const path = require('path');
const fs = require('fs');

const frontendDir = path.join(__dirname, '../frontend/app-siembra');
const backendDistDir = path.join(__dirname, 'dist');

try {
  // Install frontend dependencies
  console.log('Installing frontend dependencies...');
  execSync('npm install', { cwd: frontendDir, stdio: 'inherit' });
  
  // Build frontend
  console.log('Building frontend...');
  execSync('npm run build', { cwd: frontendDir, stdio: 'inherit' });
  
  // Copy dist folder
  console.log('Copying build files...');
  const frontendDistDir = path.join(frontendDir, 'dist');
  
  if (!fs.existsSync(backendDistDir)) {
    fs.mkdirSync(backendDistDir, { recursive: true });
  }
  
  execSync(`cp -r ${frontendDistDir}/* ${backendDistDir}/`, { stdio: 'inherit' });
  
  console.log('Build completed successfully!');
} catch (error) {
  console.error('Build failed:', error.message);
  process.exit(1);
}