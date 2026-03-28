using KalebClipPro.Models;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;

namespace KalebClipPro.Services
{
    public class ClipboardActionService
    {
        private ClipboardManager _sysManager;
        
        // Delegados (Eventos) para comunicarnos de vuelta con MainWindow sin acoplarla
        public Action<string>? OnTextoInyectado { get; set; }
        public Action<bool>? OnEstadoCapturaCambiado { get; set; }
        public Action? OnActualizarContadorTab { get; set; }
        public Action<Color>? OnNotificarCambioTab { get; set; }
        public Action? OnAvanzarSet { get; set; }

        public ClipboardActionService(ClipboardManager sysManager)
        {
            _sysManager = sysManager;
        }

        public void ProcesarAtajoGlobal(int hotkeyId, 
                                        ObservableCollection<ClipData> clipsRecolector, 
                                        ObservableCollection<ClipData> misClips, 
                                        Dispatcher dispatcher)
        {
            // -------------------------------------------------------------
            // CASO 1: Pegar desde un Slot (1-9) o Numpad (101-109)
            // -------------------------------------------------------------
            if ((hotkeyId >= 1 && hotkeyId <= 9) || (hotkeyId >= 101 && hotkeyId <= 109))
            {
                int slotReal = hotkeyId > 100 ? hotkeyId - 100 : hotkeyId;
                
                dispatcher.Invoke(() => {
                    bool pegadoDesdeRecolector = false;
                    int index = slotReal - 1; 

                    if (index >= 0 && index < 9 && !clipsRecolector[index].EsVacio)
                    {
                        string texto = clipsRecolector[index].Contenido_Plano;
                        OnTextoInyectado?.Invoke(texto);
                        _sysManager.EjecutarPegadoGlobal(texto);
                        pegadoDesdeRecolector = true;
                    }

                    if (!pegadoDesdeRecolector)
                    {
                        var clipAsignado = misClips.FirstOrDefault(c => c.HotKeyIndex == slotReal);
                        if (clipAsignado != null)
                        {
                            string texto = clipAsignado.Contenido_Plano;
                            OnTextoInyectado?.Invoke(texto);
                            _sysManager.EjecutarPegadoGlobal(texto);
                        }
                    }
                });
            }
            // -------------------------------------------------------------
            // CASO 2: Hotkey 10 - Captura rápida al Slot 1 del Recolector
            // -------------------------------------------------------------
            else if (hotkeyId == 10)
            {
                OnEstadoCapturaCambiado?.Invoke(true); 
                _sysManager.SimularCopiaGlobal(); 
                
                var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(150) };
                timer.Tick += (s, e) => {
                    timer.Stop();
                    try 
                    {
                        string text = GetClipboardTextSafe();

                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            string app = _sysManager.ObtenerAppActiva();
                            
                            for (int i = 8; i > 0; i--) 
                            {
                                clipsRecolector[i].Contenido_Plano = clipsRecolector[i - 1].Contenido_Plano;
                                clipsRecolector[i].Origen_App = clipsRecolector[i - 1].Origen_App;
                            }
                            
                            clipsRecolector[0].Contenido_Plano = text;
                            clipsRecolector[0].Origen_App = app.ToUpper();
                            
                            OnActualizarContadorTab?.Invoke();
                            OnNotificarCambioTab?.Invoke(Colors.Gold);
                        }
                    } 
                    catch { }

                    dispatcher.BeginInvoke(async () => {
                        await System.Threading.Tasks.Task.Delay(300);
                        OnEstadoCapturaCambiado?.Invoke(false);
                    });
                };
                timer.Start();
            }
            // -------------------------------------------------------------
            // CASO 3: Hotkey 11 - Pegar TODO el Recolector junto
            // -------------------------------------------------------------
            else if (hotkeyId == 11)
            {
                dispatcher.Invoke(() => {
                    var slotsLlenos = clipsRecolector.Where(c => !c.EsVacio).Select(c => c.Contenido_Plano);
                    
                    if (slotsLlenos.Any())
                    {
                        string textoAcumulado = string.Join(Environment.NewLine + Environment.NewLine, slotsLlenos);
                        OnTextoInyectado?.Invoke(textoAcumulado); 
                        _sysManager.EjecutarPegadoGlobal(textoAcumulado);
                        OnNotificarCambioTab?.Invoke(Colors.LimeGreen);
                    }
                });
            }
            // -------------------------------------------------------------
            // CASO 4: Hotkeys 21-29 o 121-129 - Capturar a un Slot Específico
            // -------------------------------------------------------------
            else if ((hotkeyId >= 21 && hotkeyId <= 29) || (hotkeyId >= 121 && hotkeyId <= 129))
            {
                int slotDestino = hotkeyId > 120 ? hotkeyId - 120 : hotkeyId - 20;
                int indiceArray = slotDestino - 1;

                OnEstadoCapturaCambiado?.Invoke(true);
                _sysManager.SimularCopiaGlobal();

                var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(150) };
                timer.Tick += (s, e) => {
                    timer.Stop();
                    try
                    {
                        string text = GetClipboardTextSafe();
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            string app = _sysManager.ObtenerAppActiva();

                            clipsRecolector[indiceArray].Contenido_Plano = text;
                            clipsRecolector[indiceArray].Origen_App = app.ToUpper();

                            OnActualizarContadorTab?.Invoke();
                            OnNotificarCambioTab?.Invoke(Colors.DeepSkyBlue); 
                        }
                    }
                    catch { }

                    dispatcher.BeginInvoke(async () => {
                        await System.Threading.Tasks.Task.Delay(300);
                        OnEstadoCapturaCambiado?.Invoke(false);
                    });
                };
                timer.Start();
            }
            // -------------------------------------------------------------
            // CASO 5: Hotkey 50 - Avanzar al siguiente Set (A -> B -> C...)
            // -------------------------------------------------------------
            else if (hotkeyId == 50)
            {
                dispatcher.Invoke(() => {
                    OnAvanzarSet?.Invoke();
                });
            }
        }

        public string GetClipboardTextSafe()
        {
            for(int i = 0; i < 4; i++) {
                try { 
                    if (Clipboard.ContainsText()) return Clipboard.GetText(); 
                }
                catch { System.Threading.Thread.Sleep(20); }
            }
            return string.Empty;
        }
    }
}