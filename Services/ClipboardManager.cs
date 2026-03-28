using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;

namespace KalebClipPro.Services
{
    public class ClipboardManager
    {
        // =========================================================================
        // IMPORTACIONES DE WIN32 API
        // =========================================================================
        [DllImport("user32.dll")] private static extern bool AddClipboardFormatListener(IntPtr hwnd);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
        [DllImport("user32.dll")] private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")] private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll")] public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, uint dwExtraInfo);

        // =========================================================================
        // CONSTANTES DE MODIFICADORES Y TECLAS
        // =========================================================================
        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;
        private const uint MOD_WIN = 0x0008;
        private const uint MOD_WIN_ALT = MOD_WIN | MOD_ALT; // Nueva combinación: Win + Alt
        private const byte VK_LWIN = 0x5B; // Código de la tecla Windows izquierda

        // Combinaciones que usaremos
        private const uint MOD_CTRL_SHIFT = MOD_CONTROL | MOD_SHIFT; // Ctrl + Shift
        private const uint MOD_CTRL_ALT = MOD_CONTROL | MOD_ALT;// Ctrl + Alt
        
        private const uint KEYEVENTF_KEYUP = 0x0002;
        
        // Virtual Keys (Hexadecimales de las teclas)
        private const byte VK_CONTROL = 0x11;
        private const byte VK_SHIFT = 0x10;
        private const byte VK_ALT = 0x12; // Necesario para liberar la tecla Alt
        private const byte VK_C = 0x43;
        private const byte VK_V = 0x56;
        private const byte VK_X = 0x58; // Nueva tecla X
        private const byte VK_S = 0x53;

        public const int MSG_HOTKEY = 0x0312;
        public const int MSG_CLIPBOARDUPDATE = 0x031D;

        public bool IgnorandoSiguienteCaptura { get; set; } = false;

        // =========================================================================
        // MÉTODOS PÚBLICOS
        // =========================================================================
        public void IniciarEscucha(IntPtr hwnd)
        {
            AddClipboardFormatListener(hwnd);

            // Bucle para registrar los atajos numéricos (Del 1 al 9)
            for (int i = 1; i <= 9; i++)
            {
                uint keyStandard = (uint)(0x30 + i);    // Teclado Estándar (arriba de las letras)
                uint keyNumpad = (uint)(0x60 + i);      // Teclado Numérico (Numpad)

                // 1. PEGAR SLOT (Ctrl + Shift + 1 al 9) -> IDs: 1 al 9 (Standard) | 101 al 109 (Numpad)
                RegisterHotKey(hwnd, i, MOD_WIN_ALT, keyStandard);
                RegisterHotKey(hwnd, 100 + i, MOD_WIN_ALT, keyNumpad);

                // 2. ASIGNAR SLOT (Ctrl + Alt + 1 al 9) -> IDs: 21 al 29 (Standard) | 121 al 129 (Numpad)
                RegisterHotKey(hwnd, 20 + i, MOD_CTRL_ALT, keyStandard);
                RegisterHotKey(hwnd, 120 + i, MOD_CTRL_ALT, keyNumpad);
            }
            
            // 3. ACUMULAR EN AUTOMÁTICO (Ctrl + Shift + X) -> ID 10
            RegisterHotKey(hwnd, 10, MOD_CTRL_SHIFT, VK_X);

            // 4. PEGADO GLOBAL RECOLECTOR (Ctrl + Alt + V) -> ID 11
            RegisterHotKey(hwnd, 11, MOD_CTRL_ALT, VK_V);

            // 5. CAMBIAR DE SET / WORKFLOW (Ctrl + Shift + S) -> ID 50
            RegisterHotKey(hwnd, 50, MOD_CTRL_SHIFT, VK_S); // <-- AÑADE ESTA LÍNEA
        }

        public void DetenerEscucha(IntPtr hwnd)
        {
            // Desregistrar todos los números (Standard y Numpad)
            for (int i = 1; i <= 9; i++) 
            {
                UnregisterHotKey(hwnd, i);
                UnregisterHotKey(hwnd, 100 + i);
                UnregisterHotKey(hwnd, 20 + i);
                UnregisterHotKey(hwnd, 120 + i);
            }
            
            // Desregistrar especiales
            UnregisterHotKey(hwnd, 10);
            UnregisterHotKey(hwnd, 11);
            UnregisterHotKey(hwnd, 50);
        }

        public void EjecutarPegadoGlobal(string texto)
        {
            // 1. Soltar TODOS los modificadores físicamente para que el Ctrl+V simulado no falle
            keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_ALT, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0); // VITAL: Soltar tecla Windows

            IgnorandoSiguienteCaptura = true;
            Clipboard.SetDataObject(texto, true);
            
            // Subimos a 100ms. Es el tiempo perfecto para que Windows limpie la memoria del teclado
            Thread.Sleep(100); 

            // 2. Simular pulsación de Ctrl + V
            keybd_event(VK_CONTROL, 0, 0, 0);
            keybd_event(VK_V, 0, 0, 0);
            
            // 3. Soltar Ctrl + V
            keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
        }

        public void SimularCopiaGlobal()
        {
            // Soltar modificadores actuales para que no interfieran
            keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_ALT, 0, KEYEVENTF_KEYUP, 0);

            IgnorandoSiguienteCaptura = true;

            // Simular pulsación de Ctrl + C
            keybd_event(VK_CONTROL, 0, 0, 0);
            keybd_event(VK_C, 0, 0, 0);
            
            // Soltar Ctrl + C
            keybd_event(VK_C, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
            
            Thread.Sleep(100); // Tiempo vital para que el OS llene el portapapeles
        }

        public string ObtenerAppActiva()
        {
            try
            {
                IntPtr h = GetForegroundWindow();
                GetWindowThreadProcessId(h, out uint pid);
                return Process.GetProcessById((int)pid).ProcessName;
            }
            catch { return "Sistema"; }
        }
    }
}