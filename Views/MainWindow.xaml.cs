using KalebClipPro.Models;
using KalebClipPro.Services;
using KalebClipPro.Infrastructure;
using KalebClipPro.Views;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Interop;
using System.Diagnostics;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using System.Windows.Input;
using System.Windows.Controls.Primitives;
using System.ComponentModel; 
using System.Windows.Data; 
using System.Linq; 
using System.Globalization; 
using GongSolutions.Wpf.DragDrop;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Windows.Documents;

namespace KalebClipPro
{
    public partial class MainWindow : Window, IDropTarget
    {
        DatabaseService db = new DatabaseService();
        ClipboardManager _sysManager = new ClipboardManager(); 
        WorkflowService _workflowService = new WorkflowService();
        
        public ObservableCollection<ClipData> MisClips { get; set; } = new ObservableCollection<ClipData>();
        public ObservableCollection<ClipData> ClipsRecolector { get; set; } = new ObservableCollection<ClipData>();
        
        private int _paginaActual = 0;
        private const int _itemsPorPagina = 25;
        private bool _cargandoMas = false;
        private bool _filtrosActivos = false;
        
        private string _setActual = "A"; 
        private bool isSidebarOpen = true;
        private double _anchoSidebarGuardado = 340; 
        private double _currentDragWidth; 
        private DispatcherTimer _autoSaveTimer = new DispatcherTimer();
        
        private Dictionary<string, RichTextBox> _editoresMemoria = new Dictionary<string, RichTextBox>();
        private Dictionary<string, HistorialSesion> _controladoresHistoria = new Dictionary<string, HistorialSesion>();
        private RichTextBox? _editorActual;

        private ClipData? _clipSeleccionadoParaMenu = null;
        private bool _isTimeTraveling = false;
        private string _ultimoTextoInyectado = "";
        private bool _capturandoParaRecolector = false;
        private bool _scrollHorizontalHabilitado = false;
        private bool _mostrarNumerosLinea = false;
        
        private WorkflowData WorkflowActual = new WorkflowData();
        private FontFamily _fuentePreferida = new FontFamily("Consolas");

        public MainWindow()
        {
            try 
            {
                InitializeComponent();

                TextBoxContainer.PreviewKeyDown += (s, e) => { if (_editorActual != null) Helpers.TableSurgeonHelper.SmartEditor_PreviewKeyDown(_editorActual, s, e); };
                TextBoxContainer.PreviewMouseRightButtonUp += (s, e) => { if (_editorActual != null) Helpers.TableSurgeonHelper.MostrarMenuContextual(_editorActual, e, this); };
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error fatal en el diseño:\n{ex.Message}\n\nDetalle:\n{ex.InnerException?.Message}", 
                                "Error de WPF", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
                return;
            }

            // 1. Cargar el motor de datos JSON a través del servicio
            WorkflowActual = _workflowService.CargarWorkflowDesdeDisco();

            // 2. Inicializar BD y Colecciones
            db.InicializarBaseDeDatos();
            ListaClips.ItemsSource = MisClips;
            ListaRecolector.ItemsSource = ClipsRecolector; 
            
            // 3. Sincronizar UI
            SincronizarBotonesConDatos();
            InicializarSlotsVacios();
            CargarDatos();
            
            this.Topmost = false;

            _autoSaveTimer.Interval = TimeSpan.FromMilliseconds(350); 
            _autoSaveTimer.Tick += AutoSaveTimer_Tick; 
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            var hwnd = new WindowInteropHelper(this).Handle;
            
            _sysManager.IniciarEscucha(hwnd);
            HwndSource.FromHwnd(hwnd).AddHook(HwndHook);
        }

        protected override void OnClosed(EventArgs e)
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            _sysManager.DetenerEscucha(hwnd); 
            base.OnClosed(e);
        }

        private IntPtr HwndHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == ClipboardManager.MSG_HOTKEY)
            {
                ManejarAtajoGlobal(wParam.ToInt32());
            }
            else if (msg == ClipboardManager.MSG_CLIPBOARDUPDATE)
            {
                ProcesarNuevoPortapapeles();
            }
            return IntPtr.Zero;
        }

        private void ManejarAtajoGlobal(int hotkeyId)
        {
            if ((hotkeyId >= 1 && hotkeyId <= 9) || (hotkeyId >= 101 && hotkeyId <= 109))
            {
                int slotReal = hotkeyId > 100 ? hotkeyId - 100 : hotkeyId;
                
                Dispatcher.Invoke(() => {
                    bool pegadoDesdeRecolector = false;
                    int index = slotReal - 1; 

                    if (index >= 0 && index < 9 && !ClipsRecolector[index].EsVacio)
                    {
                        _ultimoTextoInyectado = ClipsRecolector[index].Contenido_Plano;
                        _sysManager.EjecutarPegadoGlobal(_ultimoTextoInyectado);
                        pegadoDesdeRecolector = true;
                    }

                    if (!pegadoDesdeRecolector)
                    {
                        var clipAsignado = MisClips.FirstOrDefault(c => c.HotKeyIndex == slotReal);
                        if (clipAsignado != null)
                        {
                            _ultimoTextoInyectado = clipAsignado.Contenido_Plano;
                            _sysManager.EjecutarPegadoGlobal(_ultimoTextoInyectado);
                        }
                    }
                });
            }
            else if (hotkeyId == 10)
            {
                _capturandoParaRecolector = true; 
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
                            TabRecolector.IsChecked = true;
                            
                            for (int i = 8; i > 0; i--) 
                            {
                                ClipsRecolector[i].Contenido_Plano = ClipsRecolector[i - 1].Contenido_Plano;
                                ClipsRecolector[i].Origen_App = ClipsRecolector[i - 1].Origen_App;
                            }
                            
                            ClipsRecolector[0].Contenido_Plano = text;
                            ClipsRecolector[0].Origen_App = app.ToUpper();
                            
                            ActualizarContadorTab();
                            NotificarCambioTab(Colors.Gold);
                        }
                    } 
                    catch { }

                    Dispatcher.BeginInvoke(async () => {
                        await System.Threading.Tasks.Task.Delay(300);
                        _capturandoParaRecolector = false;
                    });
                };
                timer.Start();
            }
            else if (hotkeyId == 11)
            {
                Dispatcher.Invoke(() => {
                    var slotsLlenos = ClipsRecolector.Where(c => !c.EsVacio).Select(c => c.Contenido_Plano);
                    
                    if (slotsLlenos.Any())
                    {
                        string textoAcumulado = string.Join(Environment.NewLine + Environment.NewLine, slotsLlenos);
                        _ultimoTextoInyectado = textoAcumulado; 
                        _sysManager.EjecutarPegadoGlobal(_ultimoTextoInyectado);
                        NotificarCambioTab(Colors.LimeGreen);
                    }
                });
            }
            else if ((hotkeyId >= 21 && hotkeyId <= 29) || (hotkeyId >= 121 && hotkeyId <= 129))
            {
                int slotDestino = hotkeyId > 120 ? hotkeyId - 120 : hotkeyId - 20;
                int indiceArray = slotDestino - 1;

                _capturandoParaRecolector = true;
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
                            TabRecolector.IsChecked = true;

                            ClipsRecolector[indiceArray].Contenido_Plano = text;
                            ClipsRecolector[indiceArray].Origen_App = app.ToUpper();

                            ActualizarContadorTab();
                            NotificarCambioTab(Colors.DeepSkyBlue); 
                        }
                    }
                    catch { }

                    Dispatcher.BeginInvoke(async () => {
                        await System.Threading.Tasks.Task.Delay(300);
                        _capturandoParaRecolector = false;
                    });
                };
                timer.Start();
            }
            else if (hotkeyId == 50)
            {
                Dispatcher.Invoke(() => {
                    GuardarSlotsActuales();

                    var llaves = WorkflowActual.Sets.Keys.OrderBy(k => k).ToList();
                    int indiceActual = llaves.IndexOf(_setActual);
                    int siguienteIndice = (indiceActual + 1) % llaves.Count;
                    
                    _setActual = llaves[siguienteIndice];

                    InicializarSlotsVacios();
                    ActualizarCheckVisualSets();
                    NotificarCambioTab(Colors.Cyan);
                });
            }
        }

        private string GetClipboardTextSafe()
        {
            for(int i = 0; i < 4; i++) {
                try { 
                    if (Clipboard.ContainsText()) return Clipboard.GetText(); 
                }
                catch { System.Threading.Thread.Sleep(20); }
            }
            return string.Empty;
        }

        private void NotificarCambioTab(Color color)
        {
            TabRecolector.Foreground = new SolidColorBrush(color);
            var t2 = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
            t2.Tick += (s2, e2) => { TabRecolector.ClearValue(Control.ForegroundProperty); t2.Stop(); };
            t2.Start();
        }

        private void Tab_Changed(object sender, RoutedEventArgs e)
        {
            if (TextBoxContainer == null || RecolectorContainer == null) return;

            if (TabEditor.IsChecked == true)
            {
                if (TextBoxContainer != null) TextBoxContainer.Visibility = Visibility.Visible;
                if (RecolectorContainer != null) RecolectorContainer.Visibility = Visibility.Collapsed;
                if (TabStripBorder != null) TabStripBorder.Visibility = Visibility.Visible; 
            }
            else
            {
                if (TextBoxContainer != null) TextBoxContainer.Visibility = Visibility.Collapsed;
                if (RecolectorContainer != null) RecolectorContainer.Visibility = Visibility.Visible;
                if (TabStripBorder != null) TabStripBorder.Visibility = Visibility.Collapsed; 
            }
        }

        private void ProcesarNuevoPortapapeles()
        {
            if (_sysManager.IgnorandoSiguienteCaptura) 
            { 
                _sysManager.IgnorandoSiguienteCaptura = false; 
                return;
            }
            
            try 
            {
                if (Clipboard.ContainsText())
                {
                    string text = Clipboard.GetText();
                    if (string.IsNullOrWhiteSpace(text)) return;
                    
                    if (text == _ultimoTextoInyectado) return; 

                    if (_capturandoParaRecolector)
                    {
                        _capturandoParaRecolector = false; 
                        return; 
                    }

                    if (MisClips.Count > 0 && MisClips[0].Contenido_Plano == text) return; 

                    string appGeneral = _sysManager.ObtenerAppActiva();
                    string nuevoId = db.GuardarClipConOrigen(text, appGeneral);
                    var nuevoClip = new ClipData { Guid_Clip = nuevoId, Contenido_Plano = text, Origen_App = appGeneral.ToUpper(), Fecha_Creacion = DateTime.Now };
                    if (!_filtrosActivos) MisClips.Insert(0, nuevoClip);
                }
            }
            catch { }
        }
        
        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (ListaClips.SelectedItem is ClipData currentClip && _editorActual != null)
            {
                GuardarEstadoSeguro(currentClip, _editorActual);
            }
            base.OnClosing(e);
        }

        private void AlwaysOnTop_Changed(object sender, RoutedEventArgs e)
        {
            if (this.IsLoaded && AlwaysOnTopCheck != null) 
            {
                this.Topmost = AlwaysOnTopCheck.IsChecked ?? false;
            }
        }

        private void CmbFuentes_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbFuentes.SelectedItem is FontFamily fuenteSeleccionada && _editorActual != null)
            {
                _fuentePreferida = fuenteSeleccionada;
                
                // Aplica a la selección si hay texto seleccionado, sino, aplicará al texto siguiente
                _editorActual.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, fuenteSeleccionada);
                _editorActual.Focus();
            }
        }

        private void CmbTamano_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbTamano.SelectedItem is ComboBoxItem item && _editorActual != null)
            {
                if (double.TryParse(item.Content.ToString(), out double nuevoTamano))
                {
                    _editorActual.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, nuevoTamano);
                    _editorActual.Focus();
                }
            }
        }

        private void EjecutarFormato(RoutedCommand comando)
        {
            if (_editorActual == null) return;
            comando.Execute(null, _editorActual);
            _editorActual.Focus();
        }

        // --- Eventos para ABRIR los menús ---
        private void BtnMenuAlign_Click(object sender, RoutedEventArgs e) => PopAlign.IsOpen = true;
        private void BtnMenuList_Click(object sender, RoutedEventArgs e) => PopList.IsOpen = true;
        private void BtnMenuInsert_Click(object sender, RoutedEventArgs e) => PopInsert.IsOpen = true;

        private void CerrarPopups()
        {
            if (PopAlign != null) PopAlign.IsOpen = false;
            if (PopList != null) PopList.IsOpen = false;
            if (PopInsert != null) PopInsert.IsOpen = false;
        }

        // --- ACCIONES BÁSICAS ---
        private void BtnMdBold_Click(object sender, RoutedEventArgs e) => EjecutarFormato(EditingCommands.ToggleBold);
        private void BtnMdItalic_Click(object sender, RoutedEventArgs e) => EjecutarFormato(EditingCommands.ToggleItalic);
        private void BtnMdUnderline_Click(object sender, RoutedEventArgs e) => EjecutarFormato(EditingCommands.ToggleUnderline);
        
        private void BtnMdColor_Click(object sender, RoutedEventArgs e) 
        { 
            if (_editorActual == null) return;

            var dialog = new KalebClipPro.Views.PaletaColores();
            dialog.Owner = this; 
            var colorActualObj = _editorActual.Selection.GetPropertyValue(System.Windows.Documents.TextElement.ForegroundProperty);

            // 🌟 CORRECCIÓN 3: Cálculo de posición anti-zoom de Windows
            Button? boton = sender as Button;
            if (boton != null)
            {
                Point screenPos = boton.PointToScreen(new Point(0, boton.ActualHeight));
                PresentationSource source = PresentationSource.FromVisual(this);
                
                if (source != null)
                {
                    dialog.WindowStartupLocation = WindowStartupLocation.Manual;
                    // Dividimos entre el DPI de la pantalla para evitar saltos locos
                    dialog.Left = screenPos.X / source.CompositionTarget.TransformToDevice.M11;
                    dialog.Top = screenPos.Y / source.CompositionTarget.TransformToDevice.M22;
                }
            }

            if (colorActualObj is SolidColorBrush brush)
            {
                // Le pasamos el color de la brocha a la paleta
                dialog.CargarColorDesdeEditor(brush.Color);
            }
            
            dialog.AlSeleccionarColor = (color) =>
            {
                var brush = new SolidColorBrush(color);
                _editorActual.Selection.ApplyPropertyValue(System.Windows.Documents.TextElement.ForegroundProperty, brush);
                _editorActual.Focus();
            };

            dialog.Show();
        }

        // --- ALINEACIÓN ---
        private void BtnMdAlignLeft_Click(object sender, RoutedEventArgs e) { CerrarPopups(); EjecutarFormato(EditingCommands.AlignLeft); }
        private void BtnMdAlignCenter_Click(object sender, RoutedEventArgs e) { CerrarPopups(); EjecutarFormato(EditingCommands.AlignCenter); }
        private void BtnMdAlignRight_Click(object sender, RoutedEventArgs e) { CerrarPopups(); EjecutarFormato(EditingCommands.AlignRight); }

        // --- LISTAS Y SANGRÍAS ---
        private void BtnMdList_Click(object sender, RoutedEventArgs e) { CerrarPopups(); EjecutarFormato(EditingCommands.ToggleBullets); }
        private void BtnMdListNum_Click(object sender, RoutedEventArgs e) { CerrarPopups(); AplicarListaNumeradaInteligente(_editorActual); }
        private void BtnMdIndentInc_Click(object sender, RoutedEventArgs e) { CerrarPopups(); AplicarSangriaPersonalizada(_editorActual, 1); }
        private void BtnMdIndentDec_Click(object sender, RoutedEventArgs e) { CerrarPopups(); AplicarSangriaPersonalizada(_editorActual, -1); }

        // --- INSERCIÓN COMPLEJA ---
        private void BtnMdQuote_Click(object sender, RoutedEventArgs e) 
        { 
            CerrarPopups(); 
            if (_editorActual == null) return;
            
            Paragraph p = new Paragraph(new Run(_editorActual.Selection.Text)) {
                Margin = new Thickness(20, 10, 0, 10),
                FontStyle = FontStyles.Italic,
                Foreground = Brushes.Gray,
                BorderBrush = Brushes.Gray,
                BorderThickness = new Thickness(2, 0, 0, 0),
                Padding = new Thickness(10, 0, 0, 0)
            };
            _editorActual.Document.Blocks.Add(p);
            _editorActual.Focus();
        }

        private void BtnMdCode_Click(object sender, RoutedEventArgs e) 
        { 
            CerrarPopups(); 
            if (_editorActual == null) return;
            
            _editorActual.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, new FontFamily("Consolas"));
            _editorActual.Selection.ApplyPropertyValue(TextElement.BackgroundProperty, new SolidColorBrush(Color.FromRgb(40, 44, 52)));
            _editorActual.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, Brushes.LightGreen);
            _editorActual.Focus();
        }

        private void BtnMdLink_Click(object sender, RoutedEventArgs e) 
        { 
            CerrarPopups(); 
            if (_editorActual == null || _editorActual.Selection.IsEmpty) return;
            
            string url = "https://google.com"; // Placeholder, se puede cambiar por un InputBox
            Hyperlink link = new Hyperlink(_editorActual.Selection.Start, _editorActual.Selection.End);
            link.NavigateUri = new Uri(url);
            link.Cursor = Cursors.Hand;
            _editorActual.Focus();
        }
        
        private void BtnMdMath_Click(object sender, RoutedEventArgs e) 
        { 
            CerrarPopups(); 
            if (_editorActual == null) return;

            TextRange range = new TextRange(_editorActual.Selection.Start, _editorActual.Selection.End);
            range.Text = $"${_editorActual.Selection.Text}$"; 
            range.ApplyPropertyValue(TextElement.ForegroundProperty, Brushes.Orange);
            _editorActual.Focus();
        }
        
        private void BtnMdTable_Click(object sender, RoutedEventArgs e)
        {
            CerrarPopups();
            if (_editorActual == null) return; 

            var dialogoTabla = new Views.InsertarTablaPopup();
            dialogoTabla.Owner = this; 

            if (dialogoTabla.ShowDialog() == true)
            {
                Helpers.TableSurgeonHelper.ConstruirTabla(_editorActual, dialogoTabla);
            }
        }
        
        // =========================================================================================
        // CONTROL DE INTERFAZ GAVETA
        // =========================================================================================
        private void BtnToggleSmartEditor_Click(object sender, RoutedEventArgs e)
        {
            if (SmartEditorDrawer == null) return;

            if (BtnToggleSmartEditor.IsChecked == true)
            {
                SmartEditorDrawer.Visibility = Visibility.Visible;
            }
            else
            {
                SmartEditorDrawer.Visibility = Visibility.Collapsed;
            }
        }

        private void SmartEditorDrawer_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            bool esCompacto = e.NewSize.Width < 600;

            if (GrpAlignFull != null) GrpAlignFull.Visibility = esCompacto ? Visibility.Collapsed : Visibility.Visible;
            if (BtnMenuAlign != null) BtnMenuAlign.Visibility = esCompacto ? Visibility.Visible : Visibility.Collapsed;

            if (GrpListFull != null) GrpListFull.Visibility = esCompacto ? Visibility.Collapsed : Visibility.Visible;
            if (BtnMenuList != null) BtnMenuList.Visibility = esCompacto ? Visibility.Visible : Visibility.Collapsed;

            if (GrpInsertFull != null) GrpInsertFull.Visibility = esCompacto ? Visibility.Collapsed : Visibility.Visible;
            if (BtnMenuInsert != null) BtnMenuInsert.Visibility = esCompacto ? Visibility.Visible : Visibility.Collapsed;
        }
    }
}