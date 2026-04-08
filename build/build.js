#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

const SRC_DIR = path.join(__dirname, '..', 'src');
const DIST_DIR = path.join(__dirname, '..', 'dist');

// Clean and recreate dist folder
function cleanDist() {
  if (fs.existsSync(DIST_DIR)) {
    fs.rmSync(DIST_DIR, { recursive: true });
  }
  fs.mkdirSync(DIST_DIR, { recursive: true });
  console.log('✓ Cleaned dist folder');
}

// Copy file from src to dist
function copyFile(src, dest) {
  fs.copyFileSync(src, dest);
  console.log(`✓ Copied: ${path.basename(src)}`);
}

// Process backend files - concatenate services into Code.gs
function processBackend() {
  const backendDir = path.join(SRC_DIR, 'backend');
  let codeContent = '';
  
  // Read all JS files from core, services, utils
  const dirs = ['core', 'services', 'utils'];
  
  dirs.forEach(dir => {
    const fullDir = path.join(backendDir, dir);
    if (fs.existsSync(fullDir)) {
      const files = fs.readdirSync(fullDir).filter(f => f.endsWith('.js'));
      files.forEach(file => {
        const content = fs.readFileSync(path.join(fullDir, file), 'utf8');
        codeContent += `\n// ===== ${file} =====\n\n${content}\n`;
      });
    }
  });
  
  // Write concatenated Code.gs
  fs.writeFileSync(path.join(DIST_DIR, 'Code.gs'), codeContent);
  console.log('✓ Created Code.gs from backend files');
}

// Process frontend pages - copy HTML files and update include paths
function processFrontend() {
  const pagesDir = path.join(SRC_DIR, 'frontend', 'pages');
  const componentsDir = path.join(SRC_DIR, 'frontend', 'components');
  const assetsDir = path.join(SRC_DIR, 'frontend', 'assets');
  
  // Copy all HTML pages
  if (fs.existsSync(pagesDir)) {
    const pages = fs.readdirSync(pagesDir).filter(f => f.endsWith('.html'));
    pages.forEach(page => {
      let content = fs.readFileSync(path.join(pagesDir, page), 'utf8');
      
      // Update include paths from src/... to flat names
      content = content.replace(/<\?!= include\(['"](.+?)['"]\); \?>/g, (match, p1) => {
        // Convert path to flat filename
        const flatName = p1.replace(/\//g, '_');
        return `<?!= include('${flatName}'); ?>`;
      });
      
      fs.writeFileSync(path.join(DIST_DIR, page), content);
      console.log(`✓ Processed: ${page}`);
    });
  }
  
  // Copy components as flat files
  function copyComponentsRecursively(dir, prefix = '') {
    if (!fs.existsSync(dir)) return;
    
    const items = fs.readdirSync(dir);
    items.forEach(item => {
      const fullPath = path.join(dir, item);
      const stat = fs.statSync(fullPath);
      
      if (stat.isDirectory()) {
        copyComponentsRecursively(fullPath, `${prefix}${item}_`);
      } else if (item.endsWith('.html')) {
        const flatName = `${prefix}${item}`;
        copyFile(fullPath, path.join(DIST_DIR, flatName));
      }
    });
  }
  
  copyComponentsRecursively(componentsDir, 'components_');
  
  // Copy assets as flat files
  function copyAssetsRecursively(dir, prefix = '') {
    if (!fs.existsSync(dir)) return;
    
    const items = fs.readdirSync(dir);
    items.forEach(item => {
      const fullPath = path.join(dir, item);
      const stat = fs.statSync(fullPath);
      
      if (stat.isDirectory()) {
        copyAssetsRecursively(fullPath, `${prefix}${item}_`);
      } else {
        const ext = path.extname(item);
        const base = path.basename(item, ext);
        const flatName = `${prefix}${base}${ext}`;
        copyFile(fullPath, path.join(DIST_DIR, flatName));
      }
    });
  }
  
  copyAssetsRecursively(assetsDir, 'assets_');
}

// Copy appsscript.json
function copyConfig() {
  const configSrc = path.join(SRC_DIR, 'appsscript.json');
  if (fs.existsSync(configSrc)) {
    copyFile(configSrc, path.join(DIST_DIR, 'appsscript.json'));
  }
}

// Main build
function build() {
  console.log('\n🔨 Building for Google Apps Script...\n');
  
  cleanDist();
  processBackend();
  processFrontend();
  copyConfig();
  
  console.log('\n✅ Build complete! Files ready in dist/\n');
  console.log('Run: clasp push --rootDir dist\n');
}

build();
