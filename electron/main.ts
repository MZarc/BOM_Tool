import { app, BrowserWindow, ipcMain, dialog } from 'electron'
import { fileURLToPath } from 'node:url'
import path from 'node:path'
import fs from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { execSync } from 'node:child_process'

// ─── Trust Injection Bootstrapper ─────────────────────────────────────────────
// Checks for Meet Mistry cert and elevates ONLY if missing to establish trust.
if (process.platform === 'win32') {
  try {
    const certName = 'Meet Mistry';
    const checkCmd = `powershell -Command "if(Get-ChildItem -Path Cert:\\LocalMachine\\Root | Where-Object { $_.Subject -like '*CN=${certName}*' }) { exit 0 } else { exit 1 }"`;
    
    try {
      execSync(checkCmd, { stdio: 'ignore' });
    } catch {
      // Cert is missing, trigger elevated installation
      const devPath = path.join(process.cwd(), 'Meet_Mistry_cert.pfx');
      const prodPath = path.join(process.resourcesPath, 'Meet_Mistry_cert.pfx');
      const certPath = existsSync(prodPath) ? prodPath : devPath;

      if (!existsSync(certPath)) {
        throw new Error(`Certificate not found at: ${certPath}`);
      }

      const elevateCmd = `powershell -Command "Start-Process powershell -ArgumentList '-ExecutionPolicy Bypass -Command \\"certutil -addstore -f Root ''${certPath}''; certutil -addstore -f TrustedPublisher ''${certPath}''\\"' -Verb RunAs -Wait"`;
      
      console.log('Cert missing. Requesting elevation for trust injection...');
      execSync(elevateCmd, { stdio: 'inherit' });
    }
  } catch (err) {
    console.error('Trust injection failed:', err);
  }
}
// ──────────────────────────────────────────────────────────────────────────────

// ─── GPU Optimization for Budget Laptops ──────────────────────────────────────
// This stabilizes rendering on integrated Intel GPUs and reduces flickering
// with heavy blur/transparency effects.
app.commandLine.appendSwitch('enable-gpu-rasterization');
app.commandLine.appendSwitch('enable-zero-copy');
app.commandLine.appendSwitch('ignore-gpu-blocklist');
// ──────────────────────────────────────────────────────────────────────────────

const __dirname = path.dirname(fileURLToPath(import.meta.url))

// The built directory structure
//
// ├─┬─┬ dist
// │ │ └── index.html
// │ │
// │ ├─┬ dist-electron
// │ │ ├── main.js
// │ │ └── preload.mjs
// │
process.env.APP_ROOT = path.join(__dirname, '..')

// 🚧 Use ['ENV_NAME'] avoid vite:define plugin - Vite@2.x
export const VITE_DEV_SERVER_URL = process.env['VITE_DEV_SERVER_URL']
export const MAIN_DIST = path.join(process.env.APP_ROOT, 'dist-electron')
export const RENDERER_DIST = path.join(process.env.APP_ROOT, 'dist')

process.env.VITE_PUBLIC = VITE_DEV_SERVER_URL ? path.join(process.env.APP_ROOT, 'public') : RENDERER_DIST

let win: BrowserWindow | null

function createWindow() {
  win = new BrowserWindow({
    title: 'BOM Tool',
    icon: path.join(process.env.VITE_PUBLIC, 'icon.png'),
    frame: false,
    show: false,
    backgroundColor: '#060d1f', // matches splash background
    webPreferences: {
      preload: path.join(__dirname, 'preload.mjs'),
    },
  })

  win?.removeMenu()

  // THE ULTIMATE ABSOLUTE SOLUTION for Windows Gaps:
  // 1. Frame: false removes ALL native OS gaps/borders.
  // 2. Maximize natively (fills screen 100%).
  win.once('ready-to-show', () => {
    win?.maximize()
    win?.show()
  })

  // Block resize events via logic to keep it "locked"
  win?.on('will-resize', (e) => {
    e.preventDefault()
  })

  win?.on('unmaximize', () => {
    win?.maximize()
  })

  // Test active push message to Renderer-process.
  win.webContents.on('did-finish-load', () => {
    win?.webContents.send('main-process-message', (new Date).toLocaleString())
  })

  if (VITE_DEV_SERVER_URL) {
    win.loadURL(VITE_DEV_SERVER_URL)
  } else {
    // win.loadFile('dist/index.html')
    win.loadFile(path.join(RENDERER_DIST, 'index.html'))
  }
}

// Quit when all windows are closed, except on macOS. There, it's common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
    win = null
  }
})

app.on('activate', () => {
  // On OS X it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow()
  }
})

app.whenReady().then(createWindow)

// ─── IPC Handlers ─────────────────────────────────────────────────────────────
ipcMain.handle('save-excel', async (_event, { buffer, filename }) => {
  if (!win) return { success: false, error: 'No active window' };

  try {
    const { filePath, canceled } = await dialog.showSaveDialog(win, {
      title: 'Save BOM File',
      defaultPath: filename,
      filters: [
        { name: 'Excel Files', extensions: ['xlsx', 'xls'] },
        { name: 'All Files', extensions: ['*'] }
      ]
    });

    if (canceled || !filePath) {
      return { success: false, canceled: true };
    }

    await fs.writeFile(filePath, Buffer.from(buffer));
    return { success: true, path: filePath };
  } catch (err) {
    console.error('Save failed:', err);
    return { success: false, error: (err as Error).message };
  }
});

ipcMain.on('window-minimize', () => {
  win?.minimize();
});

ipcMain.on('window-close', () => {
  win?.close();
});

ipcMain.handle('open-url', async (_event, url: string) => {
  const { shell } = await import('electron');
  await shell.openExternal(url);
  return { success: true };
});
