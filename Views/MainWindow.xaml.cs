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

                TextBoxContainer.PreviewKeyDown += SmartEditor_PreviewKeyDown;
                TextBoxContainer.PreviewMouseRightButtonUp += SmartEditor_PreviewMouseRightButtonUp;
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
                // 1. Extraemos todo el ADN desde el Pop-up
                int filas = dialogoTabla.FilasSeleccionadas;
                int columnas = dialogoTabla.ColumnasSeleccionadas;
                int estiloBase = dialogoTabla.EstiloSeleccionado; 
                string tipoBorde = dialogoTabla.BordeSeleccionado; 
                double grosorElegido = dialogoTabla.GrosorSeleccionado; 
                
                Color acento = dialogoTabla.ColorAcentoSeleccionado;
                bool enc = dialogoTabla.TieneEncabezado;
                bool tot = dialogoTabla.TieneTotales;
                bool priCol = dialogoTabla.TienePrimeraColumna;
                bool ultCol = dialogoTabla.TieneUltimaColumna;
                
                bool usaBordePers = dialogoTabla.UsaColorBordePersonalizado;
                Color colorBordePers = dialogoTabla.ColorBordeSeleccionado;

                // 2. Preparamos los colores
                var colorHeader = new SolidColorBrush(acento);
                var colorCebra1 = new SolidColorBrush(Color.FromArgb(60, acento.R, acento.G, acento.B));
                var colorCebra2 = new SolidColorBrush(Color.FromArgb(20, acento.R, acento.G, acento.B));
                var colorBordeTematico = new SolidColorBrush(Color.FromArgb(120, acento.R, acento.G, acento.B));
                var colorBordeGris = new SolidColorBrush(Color.FromRgb(58, 68, 77));

                Brush brushBordeAplicar = (estiloBase == 0) ? colorBordeGris : colorBordeTematico;
                if (usaBordePers) brushBordeAplicar = new SolidColorBrush(colorBordePers);

                // 3. Creamos la tabla y le ponemos su "Mochila" de memoria
                Table tablaWpf = new Table();
                tablaWpf.Tag = new object[] { estiloBase, acento, grosorElegido, tipoBorde, enc, tot, priCol, ultCol, usaBordePers, colorBordePers };
                tablaWpf.CellSpacing = 0;
                
                tablaWpf.BorderBrush = brushBordeAplicar;
                if (tipoBorde == "Ninguno" || tipoBorde == "Horizontales") tablaWpf.BorderThickness = new Thickness(0);
                else tablaWpf.BorderThickness = new Thickness(grosorElegido); 

                tablaWpf.Margin = new Thickness(0, 10, 0, 10); 

                for (int i = 0; i < columnas; i++) { tablaWpf.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) }); }

                TableRowGroup grupoFilas = new TableRowGroup();
                tablaWpf.RowGroups.Add(grupoFilas);

                // 4. Construcción celda por celda
                for (int f = 0; f < filas; f++)
                {
                    TableRow filaVisual = new TableRow();
                    grupoFilas.Rows.Add(filaVisual);

                    for (int c = 0; c < columnas; c++)
                    {
                        Paragraph parrafoCelda = new Paragraph(new Run(""));
                        TableCell celda = new TableCell(parrafoCelda);
                        celda.Padding = new Thickness(8, 4, 8, 4);
                        celda.BorderBrush = brushBordeAplicar;

                        // PASO 1: PATRÓN BASE
                        if (estiloBase == 1) celda.Background = (f % 2 == 0) ? colorCebra1 : colorCebra2;
                        else if (estiloBase == 2) celda.Background = (c % 2 == 0) ? colorCebra1 : colorCebra2;
                        else if (estiloBase == 3) celda.Background = ((f + c) % 2 == 0) ? colorCebra1 : colorCebra2;
                        else celda.Background = Brushes.Transparent; 

                        // PASO 2: REALCES DE ESTILO (Checkboxes)
                        bool esZonaRealzada = false;
                        if (enc && f == 0) esZonaRealzada = true;
                        if (tot && f == filas - 1) esZonaRealzada = true;
                        if (priCol && c == 0) esZonaRealzada = true;
                        if (ultCol && c == columnas - 1) esZonaRealzada = true;

                        if (esZonaRealzada)
                        {
                            celda.Background = colorHeader;
                            celda.Foreground = Brushes.White; 
                            parrafoCelda.FontWeight = FontWeights.Bold;
                            if (estiloBase == 0 && !usaBordePers) celda.BorderBrush = colorBordeTematico;
                        }

                        // LÓGICA DE BORDES Y GROSOR INTERNO
                        Thickness grosorFinal = new Thickness(0);
                        if (tipoBorde == "Todos") grosorFinal = new Thickness(grosorElegido);
                        else if (tipoBorde == "Horizontales") grosorFinal = new Thickness(0, 0, 0, grosorElegido);
                        else if (tipoBorde == "Verticales") grosorFinal = new Thickness(0, 0, grosorElegido, 0);
                        
                        celda.BorderThickness = grosorFinal;
                        filaVisual.Cells.Add(celda);
                    }
                }

                // 5. Inserción en el documento
                _editorActual.BeginChange();
                Block bloqueActual = _editorActual.CaretPosition.Paragraph;
                if (bloqueActual != null && bloqueActual.SiblingBlocks != null) bloqueActual.SiblingBlocks.InsertAfter(bloqueActual, tablaWpf);
                else _editorActual.Document.Blocks.Add(tablaWpf);

                var parrafoFinal = new Paragraph();
                if (tablaWpf.SiblingBlocks != null) tablaWpf.SiblingBlocks.InsertAfter(tablaWpf, parrafoFinal);

                _editorActual.EndChange();
                _editorActual.Focus();
                
                // Mover cursor a la primera celda
                if (grupoFilas.Rows.Count > 0 && grupoFilas.Rows[0].Cells.Count > 0)
                {
                    _editorActual.CaretPosition = ((Paragraph)grupoFilas.Rows[0].Cells[0].Blocks.FirstBlock).ContentStart;
                }
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

        // =========================================================================================
        // LÓGICA INTELIGENTE DE TABLAS (CORREGIDA PARA NULOS Y TABULADOR)
        // =========================================================================================

        private void SmartEditor_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Tab && _editorActual != null)
            {
                var cursor = _editorActual.CaretPosition;
                if (cursor.Paragraph != null && cursor.Paragraph.Parent is TableCell celdaActual)
                {
                    TableRow? filaActual = celdaActual.Parent as TableRow;
                    TableRowGroup? grupoFilas = filaActual?.Parent as TableRowGroup;

                    if (grupoFilas != null && filaActual != null)
                    {
                        bool esUltimaFila = grupoFilas.Rows.IndexOf(filaActual) == grupoFilas.Rows.Count - 1;
                        bool esUltimaCelda = filaActual.Cells.IndexOf(celdaActual) == filaActual.Cells.Count - 1;

                        if (esUltimaFila && esUltimaCelda)
                        {
                            e.Handled = true; 
                            InsertarFila(true); // AQUÍ ESTÁ EL ARREGLO DEL ERROR
                        }
                    }
                }
            }
        }
        
        private void InsertarFila(bool abajo)
        {
            if (_editorActual == null) return;
            var cursor = _editorActual.CaretPosition;
            if (cursor.Paragraph != null && cursor.Paragraph.Parent is TableCell celdaActual)
            {
                TableRow? filaActual = celdaActual.Parent as TableRow;
                TableRowGroup? grupoFilas = filaActual?.Parent as TableRowGroup;
                Table? tabla = grupoFilas?.Parent as Table; 

                if (grupoFilas != null && filaActual != null && tabla != null)
                {
                    int columnas = filaActual.Cells.Count;
                    TableRow nuevaFila = new TableRow();

                    int estilo = 1; Color acento = Color.FromRgb(0, 173, 181); double grosorBase = 0.5; string tipoBorde = "Todos";
                    bool memoriaViva = false; bool priCol = false; bool ultCol = false;
                    bool usaBordePers = false; Color colorBordePers = Color.FromRgb(58, 68, 77);
                    
                    if (tabla.Tag is object[] config)
                    {
                        try {
                            if (config.Length >= 1) estilo = Convert.ToInt32(config[0]);
                            if (config.Length >= 2 && config[1] is Color c) acento = c;
                            if (config.Length >= 3) grosorBase = Convert.ToDouble(config[2]);
                            if (config.Length >= 4) tipoBorde = config[3]?.ToString() ?? "Todos";
                            if (config.Length >= 10) { 
                                priCol = (bool)config[6]; ultCol = (bool)config[7];
                                usaBordePers = (bool)config[8]; colorBordePers = (Color)config[9]; 
                            }
                            memoriaViva = true;
                        } catch { } 
                    }

                    // 👇 INICIA CIRUGÍA: Sanación Visual
                    if (!memoriaViva && grupoFilas.Rows.Count > 0 && grupoFilas.Rows[0].Cells.Count > 0)
                    {
                        TableCell celda0 = grupoFilas.Rows[0].Cells[0];
                        int fCount = grupoFilas.Rows.Count;
                        int cCount = grupoFilas.Rows[0].Cells.Count;
                        int fMed = fCount > 2 ? 1 : 0; 
                        bool dEnc = false, dTot = false;

                        double maxTB = Math.Max(celda0.BorderThickness.Top, celda0.BorderThickness.Bottom);
                        double maxLR = Math.Max(celda0.BorderThickness.Left, celda0.BorderThickness.Right);
                        if (maxTB > 0 && maxLR == 0) tipoBorde = "Horizontales";
                        else if (maxLR > 0 && maxTB == 0) tipoBorde = "Verticales";
                        else tipoBorde = "Todos";
                        grosorBase = Math.Max(maxTB, maxLR);
                        if (grosorBase == 0) grosorBase = 0.5;

                        estilo = 0; 
                        foreach (var row in grupoFilas.Rows) {
                            foreach (var cell in row.Cells) {
                                if (cell.Background is SolidColorBrush bg && bg.Color.A > 0 && bg.Color.A < 255) {
                                    estilo = 1; acento = Color.FromRgb(bg.Color.R, bg.Color.G, bg.Color.B); break;
                                }
                            }
                            if (estilo != 0) break;
                        }

                        if (celda0.BorderBrush is SolidColorBrush sb) {
                            if (sb.Color.A == 255 && (sb.Color.R != 58 || sb.Color.G != 68 || sb.Color.B != 77)) {
                                usaBordePers = true; colorBordePers = sb.Color; 
                            } else if (sb.Color.A == 120 && estilo == 0) {
                                acento = Color.FromRgb(sb.Color.R, sb.Color.G, sb.Color.B);
                            }
                        }

                        // 🌟 FIX INTELIGENTE: Diferenciar celda pintada de un Realce de tabla
                        bool CeldaEsRealce(TableCell c) {
                            return c.Background is SolidColorBrush bg && bg.Color.A == 255 && 
                                   c.Blocks.FirstBlock is Paragraph p && p.FontWeight == FontWeights.Bold &&
                                   p.Foreground is SolidColorBrush fg && fg.Color.R == 255 && fg.Color.G == 255 && fg.Color.B == 255;
                        }

                        if (fCount > 0 && CeldaEsRealce(grupoFilas.Rows[0].Cells[0])) dEnc = true;
                        if (fCount > 1 && CeldaEsRealce(grupoFilas.Rows[fCount-1].Cells[0])) dTot = true;
                        if (cCount > 0 && fCount > 1 && CeldaEsRealce(grupoFilas.Rows[fMed].Cells[0])) { 
                            priCol = true; estilo = 4; 
                            var bP = (SolidColorBrush)grupoFilas.Rows[fMed].Cells[0].Background;
                            acento = Color.FromRgb(bP.Color.R, bP.Color.G, bP.Color.B); 
                        }
                        if (cCount > 1 && fCount > 1 && CeldaEsRealce(grupoFilas.Rows[fMed].Cells[cCount-1])) ultCol = true;

                        tabla.Tag = new object[] { estilo, acento, grosorBase, tipoBorde, dEnc, dTot, priCol, ultCol, usaBordePers, colorBordePers }; 
                    }
                    // 👆 TERMINA CIRUGÍA

                    var colorHeader = new SolidColorBrush(acento);
                    var colorCebra1 = new SolidColorBrush(Color.FromArgb(60, acento.R, acento.G, acento.B));
                    var colorCebra2 = new SolidColorBrush(Color.FromArgb(20, acento.R, acento.G, acento.B));
                    
                    Brush bordeUsar = (estilo == 0) ? new SolidColorBrush(Color.FromRgb(58, 68, 77)) : new SolidColorBrush(Color.FromArgb(120, acento.R, acento.G, acento.B));
                    if (usaBordePers) bordeUsar = new SolidColorBrush(colorBordePers);

                    int indexActual = grupoFilas.Rows.IndexOf(filaActual);
                    int indiceNueva = abajo ? indexActual + 1 : indexActual;

                    Thickness grosorLimpio = new Thickness(0);
                    if (tipoBorde == "Todos") grosorLimpio = new Thickness(grosorBase);
                    else if (tipoBorde == "Horizontales") grosorLimpio = new Thickness(0, 0, 0, grosorBase);
                    else if (tipoBorde == "Verticales") grosorLimpio = new Thickness(0, 0, grosorBase, 0);

                    for (int c = 0; c < columnas; c++)
                    {
                        Paragraph p = new Paragraph(new Run(""));
                        TableCell cell = new TableCell(p) { BorderBrush = bordeUsar, BorderThickness = grosorLimpio, Padding = new Thickness(8, 4, 8, 4) };

                        if (estilo == 1) cell.Background = (indiceNueva % 2 == 0) ? colorCebra1 : colorCebra2;
                        else if (estilo == 2) cell.Background = (c % 2 == 0) ? colorCebra1 : colorCebra2;
                        else if (estilo == 3) cell.Background = ((indiceNueva + c) % 2 == 0) ? colorCebra1 : colorCebra2;
                        else cell.Background = Brushes.Transparent;

                        if ((priCol && c == 0) || (ultCol && c == columnas - 1))
                        {
                            cell.Background = colorHeader;
                            p.Foreground = Brushes.White; p.FontWeight = FontWeights.Bold;
                        }

                        nuevaFila.Cells.Add(cell);
                    }
                    grupoFilas.Rows.Insert(indiceNueva, nuevaFila);
                    _editorActual.Focus();
                    if (nuevaFila.Cells.Count > 0 && nuevaFila.Cells[0].Blocks.FirstBlock is Paragraph primerBloque) _editorActual.CaretPosition = primerBloque.ContentStart;
                }
            }
        }

        private void InsertarColumna(bool derecha)
        {
            if (_editorActual == null) return;
            var cursor = _editorActual.CaretPosition;
            if (cursor.Paragraph != null && cursor.Paragraph.Parent is TableCell celdaActual)
            {
                TableRow? filaActual = celdaActual.Parent as TableRow;
                TableRowGroup? grupoFilas = filaActual?.Parent as TableRowGroup;
                Table? tabla = grupoFilas?.Parent as Table;

                if (tabla != null && filaActual != null && grupoFilas != null)
                {
                    int colIndex = filaActual.Cells.IndexOf(celdaActual);
                    int indiceInsercion = derecha ? colIndex + 1 : colIndex;
                    tabla.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });

                    int estilo = 1; Color acento = Color.FromRgb(0, 173, 181); double grosorBase = 0.5; string tipoBorde = "Todos";
                    bool memoriaViva = false; bool enc = true; bool tot = false;
                    bool usaBordePers = false; Color colorBordePers = Color.FromRgb(58, 68, 77);
                    
                    if (tabla.Tag is object[] config)
                    {
                        try {
                            if (config.Length >= 1) estilo = Convert.ToInt32(config[0]);
                            if (config.Length >= 2 && config[1] is Color c) acento = c;
                            if (config.Length >= 3) grosorBase = Convert.ToDouble(config[2]);
                            if (config.Length >= 4) tipoBorde = config[3]?.ToString() ?? "Todos";
                            if (config.Length >= 10) { 
                                enc = (bool)config[4]; tot = (bool)config[5];
                                usaBordePers = (bool)config[8]; colorBordePers = (Color)config[9]; 
                            }
                            memoriaViva = true;
                        } catch { }
                    }

                    if (!memoriaViva && grupoFilas.Rows.Count > 0 && grupoFilas.Rows[0].Cells.Count > 0)
                    {
                        TableCell celda0 = grupoFilas.Rows[0].Cells[0];
                        int fCount = grupoFilas.Rows.Count;
                        int cCount = grupoFilas.Rows[0].Cells.Count;
                        int fMed = fCount > 2 ? 1 : 0;
                        bool dPri = false, dUlt = false;

                        double maxTB = Math.Max(celda0.BorderThickness.Top, celda0.BorderThickness.Bottom);
                        double maxLR = Math.Max(celda0.BorderThickness.Left, celda0.BorderThickness.Right);
                        if (maxTB > 0 && maxLR == 0) tipoBorde = "Horizontales";
                        else if (maxLR > 0 && maxTB == 0) tipoBorde = "Verticales";
                        else tipoBorde = "Todos";
                        grosorBase = Math.Max(maxTB, maxLR);
                        if (grosorBase == 0) grosorBase = 0.5;

                        estilo = 0;
                        foreach (var r in grupoFilas.Rows) {
                            foreach (var c in r.Cells) {
                                if (c.Background is SolidColorBrush bg && bg.Color.A > 0 && bg.Color.A < 255) {
                                    estilo = 1; acento = Color.FromRgb(bg.Color.R, bg.Color.G, bg.Color.B); break;
                                }
                            }
                            if (estilo != 0) break;
                        }

                        if (celda0.BorderBrush is SolidColorBrush sb) {
                            if (sb.Color.A == 255 && (sb.Color.R != 58 || sb.Color.G != 68 || sb.Color.B != 77)) {
                                usaBordePers = true; colorBordePers = sb.Color; 
                            } else if (sb.Color.A == 120 && estilo == 0) {
                                acento = Color.FromRgb(sb.Color.R, sb.Color.G, sb.Color.B);
                            }
                        }

                        // 🌟 FIX INTELIGENTE: Diferenciar celda pintada
                        bool CeldaEsRealce(TableCell c) {
                            return c.Background is SolidColorBrush bg && bg.Color.A == 255 && 
                                   c.Blocks.FirstBlock is Paragraph p && p.FontWeight == FontWeights.Bold &&
                                   p.Foreground is SolidColorBrush fg && fg.Color.R == 255 && fg.Color.G == 255 && fg.Color.B == 255;
                        }

                        enc = (fCount > 0 && CeldaEsRealce(grupoFilas.Rows[0].Cells[0]));
                        tot = (fCount > 1 && CeldaEsRealce(grupoFilas.Rows[fCount-1].Cells[0]));
                        
                        if (cCount > 0 && fCount > 1 && CeldaEsRealce(grupoFilas.Rows[fMed].Cells[0])) { 
                            dPri = true; estilo = 4; 
                            var bP = (SolidColorBrush)grupoFilas.Rows[fMed].Cells[0].Background;
                            acento = Color.FromRgb(bP.Color.R, bP.Color.G, bP.Color.B); 
                        }
                        if (cCount > 1 && fCount > 1 && CeldaEsRealce(grupoFilas.Rows[fMed].Cells[cCount-1])) dUlt = true;

                        tabla.Tag = new object[] { estilo, acento, grosorBase, tipoBorde, enc, tot, dPri, dUlt, usaBordePers, colorBordePers }; 
                    }

                    var colorHeader = new SolidColorBrush(acento);
                    var colorCebra1 = new SolidColorBrush(Color.FromArgb(60, acento.R, acento.G, acento.B));
                    var colorCebra2 = new SolidColorBrush(Color.FromArgb(20, acento.R, acento.G, acento.B));
                    
                    Brush bordeUsar = (estilo == 0) ? new SolidColorBrush(Color.FromRgb(58, 68, 77)) : new SolidColorBrush(Color.FromArgb(120, acento.R, acento.G, acento.B));
                    if (usaBordePers) bordeUsar = new SolidColorBrush(colorBordePers);

                    Thickness grosorLimpio = new Thickness(0);
                    if (tipoBorde == "Todos") grosorLimpio = new Thickness(grosorBase);
                    else if (tipoBorde == "Horizontales") grosorLimpio = new Thickness(0, 0, 0, grosorBase);
                    else if (tipoBorde == "Verticales") grosorLimpio = new Thickness(0, 0, grosorBase, 0);

                    for (int f = 0; f < grupoFilas.Rows.Count; f++)
                    {
                        TableRow fila = grupoFilas.Rows[f];
                        Paragraph p = new Paragraph(new Run(""));
                        TableCell cell = new TableCell(p) { BorderBrush = bordeUsar, BorderThickness = grosorLimpio, Padding = new Thickness(8, 4, 8, 4) };

                        if (estilo == 1) cell.Background = (f % 2 == 0) ? colorCebra1 : colorCebra2;
                        else if (estilo == 2) cell.Background = (indiceInsercion % 2 == 0) ? colorCebra1 : colorCebra2;
                        else if (estilo == 3) cell.Background = ((f + indiceInsercion) % 2 == 0) ? colorCebra1 : colorCebra2;
                        else cell.Background = Brushes.Transparent;

                        if ((enc && f == 0) || (tot && f == grupoFilas.Rows.Count - 1))
                        {
                            cell.Background = colorHeader;
                            p.Foreground = Brushes.White; p.FontWeight = FontWeights.Bold;
                        }

                        fila.Cells.Insert(indiceInsercion, cell);
                    }
                    _editorActual.Focus();
                    if (filaActual.Cells.Count > indiceInsercion && filaActual.Cells[indiceInsercion].Blocks.FirstBlock is Paragraph bloqueCelda) _editorActual.CaretPosition = bloqueCelda.ContentStart;
                }
            }
        }

        private void EliminarFilaSeleccionada(TableCell celda)
        {
            if (_editorActual == null) return;
            
            // Al recibir la celda directamente, nunca perdemos la referencia
            if (celda.Parent is TableRow filaActual && filaActual.Parent is TableRowGroup grupoFilas)
            {
                _editorActual.BeginChange();
                if (grupoFilas.Rows.Count > 1) 
                { 
                    grupoFilas.Rows.Remove(filaActual); 
                }
                else if (grupoFilas.Parent is Table tablaDestinada) 
                { 
                    // Si intentan borrar la última fila, borramos la tabla completa
                    tablaDestinada.SiblingBlocks?.Remove(tablaDestinada); 
                }
                _editorActual.EndChange();
            }
        }

        private void EliminarColumnaSeleccionada(TableCell celda)
        {
            if (_editorActual == null) return;

            if (celda.Parent is TableRow filaActual && filaActual.Parent is TableRowGroup grupoFilas && grupoFilas.Parent is Table tabla)
            {
                _editorActual.BeginChange();
                int colIndex = filaActual.Cells.IndexOf(celda);

                if (filaActual.Cells.Count <= 1)
                {
                    // Si intentan borrar la última columna, borramos la tabla
                    tabla.SiblingBlocks?.Remove(tabla);
                    _editorActual.EndChange();
                    return;
                }

                // Borramos la definición de la columna
                if (tabla.Columns.Count > colIndex) { tabla.Columns.RemoveAt(colIndex); }

                // Borramos la celda correspondiente en CADA fila de la tabla
                foreach (TableRow fila in grupoFilas.Rows)
                {
                    if (fila.Cells.Count > colIndex)
                    {
                        fila.Cells.RemoveAt(colIndex);
                    }
                }
                _editorActual.EndChange();
            }
        }

        // 🌟 FUNCIÓN AUXILIAR PARA CREAR LOS BOTONCITOS DE LA TOOLBAR
        private Button CrearBtnBorde(string icon, string tooltip, string tipo, TableCell celda, ContextMenu menu)
        {
            Button b = new Button { 
                Content = icon, ToolTip = tooltip, Tag = tipo, 
                Background = new SolidColorBrush(Color.FromRgb(37, 43, 50)), 
                Foreground = Brushes.White, 
                BorderThickness = new Thickness(1), 
                BorderBrush = new SolidColorBrush(Color.FromRgb(58, 68, 77)), 
                Padding = new Thickness(10, 5, 10, 5), 
                Margin = new Thickness(2), 
                Cursor = Cursors.Hand 
            };
            // Al hacer clic, aplicamos el borde y cerramos el menú
            b.Click += (s, ev) => { AplicarBordeASEleccion(celda, tipo); menu.IsOpen = false; };
            return b;
        }

        private Button CrearBtnGrosor(string texto, double grosor, TableCell celda, ContextMenu menu)
        {
            Button b = new Button { 
                Content = texto, ToolTip = $"Grosor {grosor}px", 
                Background = new SolidColorBrush(Color.FromRgb(37, 43, 50)), 
                Foreground = Brushes.White, 
                BorderThickness = new Thickness(1), 
                BorderBrush = new SolidColorBrush(Color.FromRgb(58, 68, 77)), 
                Padding = new Thickness(6, 2, 6, 2), 
                Margin = new Thickness(2), 
                Cursor = Cursors.Hand,
                FontSize = 10
            };
            b.Click += (s, ev) => { AplicarGrosorASeleccion(celda, grosor); menu.IsOpen = false; };
            return b;
        }

        private void SmartEditor_PreviewMouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_editorActual == null) return;

            Point posMouse = e.GetPosition(_editorActual);
            TextPointer posicionClic = _editorActual.GetPositionFromPoint(posMouse, true);

            if (posicionClic != null && posicionClic.Paragraph != null && posicionClic.Paragraph.Parent is TableCell celdaClickeada)
            {
                TableRow filaActual = (TableRow)celdaClickeada.Parent;
                TableRowGroup grupoFilas = (TableRowGroup)filaActual.Parent;
                Table tablaPadre = (Table)grupoFilas.Parent;

                if (_editorActual.Selection.IsEmpty)
                {
                    _editorActual.Focus();
                    _editorActual.CaretPosition = posicionClic;
                }

                ContextMenu menuEspecial = new ContextMenu();
                if (FindResource("MenuOscuroStyle") is Style estiloMenu) menuEspecial.Style = estiloMenu;

                var colorTexto = new SolidColorBrush(Color.FromRgb(232, 236, 239)); 
                var colorFondoSub = new SolidColorBrush(Color.FromRgb(37, 43, 50)); 
                var colorPeligro = new SolidColorBrush(Color.FromRgb(255, 100, 100));

                // ==========================================
                // 1. SUBMENÚ DE BORDES 
                // ==========================================
                MenuItem itemBordes = new MenuItem { Header = "📏 Bordes", Foreground = colorTexto };
                itemBordes.Style = (Style)FindResource("MenuItemSubmenuOscuro"); // Aplica el fondo oscuro sin marcos blancos
                
                MenuItem optSinBorde = new MenuItem { Header = " Sin borde", Background = colorFondoSub, Foreground = colorTexto, Icon = new TextBlock { Text = "🔲", Foreground = Brushes.White } };
                optSinBorde.Click += (s, ev) => AplicarBordeASEleccion(celdaClickeada, "Ninguno");
                
                MenuItem optTodosBordes = new MenuItem { Header = " Todos los bordes", Background = colorFondoSub, Foreground = colorTexto, Icon = new TextBlock { Text = "▦", Foreground = Brushes.White } };
                optTodosBordes.Click += (s, ev) => AplicarBordeASEleccion(celdaClickeada, "Todos");

                MenuItem optBordesExternos = new MenuItem { Header = " Bordes externos", Background = colorFondoSub, Foreground = colorTexto, Icon = new TextBlock { Text = "⬜", Foreground = Brushes.White } };
                optBordesExternos.Click += (s, ev) => AplicarBordeASEleccion(celdaClickeada, "Externos");

                MenuItem optBordesInternos = new MenuItem { Header = " Bordes internos", Background = colorFondoSub, Foreground = colorTexto, Icon = new TextBlock { Text = "➕", Foreground = Brushes.White } };
                optBordesInternos.Click += (s, ev) => AplicarBordeASEleccion(celdaClickeada, "Internos");

                MenuItem optBordeSuperior = new MenuItem { Header = " Borde superior", Background = colorFondoSub, Foreground = colorTexto, Icon = new TextBlock { Text = "▔", Foreground = Brushes.White } };
                optBordeSuperior.Click += (s, ev) => AplicarBordeASEleccion(celdaClickeada, "Superior");

                MenuItem optBordeInferior = new MenuItem { Header = " Borde inferior", Background = colorFondoSub, Foreground = colorTexto, Icon = new TextBlock { Text = " ", Foreground = Brushes.White } };
                optBordeInferior.Click += (s, ev) => AplicarBordeASEleccion(celdaClickeada, "Inferior");

                MenuItem optBordeIzquierdo = new MenuItem { Header = " Borde izquierdo", Background = colorFondoSub, Foreground = colorTexto, Icon = new TextBlock { Text = "▏", Foreground = Brushes.White } };
                optBordeIzquierdo.Click += (s, ev) => AplicarBordeASEleccion(celdaClickeada, "Izquierdo");

                MenuItem optBordeDerecho = new MenuItem { Header = " Borde derecho", Background = colorFondoSub, Foreground = colorTexto, Icon = new TextBlock { Text = "▕", Foreground = Brushes.White } };
                optBordeDerecho.Click += (s, ev) => AplicarBordeASEleccion(celdaClickeada, "Derecho");

                itemBordes.Items.Add(optSinBorde);
                itemBordes.Items.Add(optTodosBordes);
                itemBordes.Items.Add(new Separator { Background = new SolidColorBrush(Color.FromRgb(58, 68, 77)) });
                itemBordes.Items.Add(optBordesExternos);
                itemBordes.Items.Add(optBordesInternos);
                itemBordes.Items.Add(new Separator { Background = new SolidColorBrush(Color.FromRgb(58, 68, 77)) });
                itemBordes.Items.Add(optBordeSuperior);
                itemBordes.Items.Add(optBordeInferior);
                itemBordes.Items.Add(optBordeIzquierdo);
                itemBordes.Items.Add(optBordeDerecho);

                // ==========================================
                // 2. NUEVO: SUBMENÚ DE GROSOR 
                // ==========================================
                MenuItem itemGrosor = new MenuItem { Header = "➖ Grosor de línea", Foreground = colorTexto };
                itemGrosor.Style = (Style)FindResource("MenuItemSubmenuOscuro"); // Aplica el mismo estilo hermoso sin bordes blancos

                MenuItem optGrosorFino = new MenuItem { Header = " Fino (0.5 px)", Background = colorFondoSub, Foreground = colorTexto };
                optGrosorFino.Click += (s, ev) => AplicarGrosorASeleccion(celdaClickeada, 0.5);

                MenuItem optGrosorNormal = new MenuItem { Header = " Normal (1.0 px)", Background = colorFondoSub, Foreground = colorTexto };
                optGrosorNormal.Click += (s, ev) => AplicarGrosorASeleccion(celdaClickeada, 1.0);

                MenuItem optGrosorGrueso = new MenuItem { Header = " Grueso (2.0 px)", Background = colorFondoSub, Foreground = colorTexto };
                optGrosorGrueso.Click += (s, ev) => AplicarGrosorASeleccion(celdaClickeada, 2.0);

                MenuItem optGrosorExtra = new MenuItem { Header = " Extra Grueso (3.0 px)", Background = colorFondoSub, Foreground = colorTexto };
                optGrosorExtra.Click += (s, ev) => AplicarGrosorASeleccion(celdaClickeada, 3.0);

                itemGrosor.Items.Add(optGrosorFino);
                itemGrosor.Items.Add(optGrosorNormal);
                itemGrosor.Items.Add(optGrosorGrueso);
                itemGrosor.Items.Add(optGrosorExtra);

                // ==========================================
                // 3. RESTO DEL MENÚ (CELDAS, COLORES Y TABLA)
                // ==========================================
                MenuItem itemColorCelda = new MenuItem { Header = "🎨 Pintar fondo de celda", Foreground = colorTexto };
                itemColorCelda.Click += (s, ev) => PintarFondoSeleccion(celdaClickeada, tablaPadre);

                MenuItem itemQuitarColor = new MenuItem { Header = "🧽 Quitar color de fondo", Foreground = colorTexto };
                itemQuitarColor.Click += (s, ev) => QuitarFondoSeleccion(celdaClickeada, tablaPadre);

                MenuItem itemColorBorde = new MenuItem { Header = "🖌️ Cambiar color de borde", Foreground = colorTexto };
                itemColorBorde.Click += (s, ev) => PintarBordeSeleccion(celdaClickeada, tablaPadre);

                MenuItem itemTemaTabla = new MenuItem { Header = "🌈 Cambiar paleta de la tabla", Foreground = colorTexto };
                itemTemaTabla.Click += (s, ev) => CambiarTemaTabla(celdaClickeada, tablaPadre);

                MenuItem itemFilaArriba = new MenuItem { Header = "⬆️ Insertar fila arriba", Foreground = colorTexto };
                itemFilaArriba.Click += (s, ev) => InsertarFila(false);

                MenuItem itemFilaAbajo = new MenuItem { Header = "⬇️ Insertar fila debajo", Foreground = colorTexto };
                itemFilaAbajo.Click += (s, ev) => InsertarFila(true);

                MenuItem itemColIzq = new MenuItem { Header = "⬅️ Insertar columna a la izquierda", Foreground = colorTexto };
                itemColIzq.Click += (s, ev) => InsertarColumna(false);

                MenuItem itemColDer = new MenuItem { Header = "➡️ Insertar columna a la derecha", Foreground = colorTexto };
                itemColDer.Click += (s, ev) => InsertarColumna(true);

                MenuItem itemEliminarFila = new MenuItem { Header = "❌ Eliminar esta fila", Foreground = colorPeligro };
                itemEliminarFila.Click += (s, ev) => EliminarFilaSeleccionada(celdaClickeada);

                MenuItem itemEliminarColumna = new MenuItem { Header = "❌ Eliminar esta columna", Foreground = colorPeligro };
                itemEliminarColumna.Click += (s, ev) => EliminarColumnaSeleccionada(celdaClickeada);

                // --- CONSTRUCCIÓN FINAL DEL MENÚ PRINCIPAL ---
                menuEspecial.Items.Add(itemBordes); 
                menuEspecial.Items.Add(itemGrosor); 
                menuEspecial.Items.Add(new Separator { Background = new SolidColorBrush(Color.FromRgb(58, 68, 77)) });
                
                // Nuevos botones de color
                menuEspecial.Items.Add(itemColorCelda);
                menuEspecial.Items.Add(itemQuitarColor);
                menuEspecial.Items.Add(itemColorBorde);
                menuEspecial.Items.Add(itemTemaTabla);
                
                menuEspecial.Items.Add(new Separator { Background = new SolidColorBrush(Color.FromRgb(58, 68, 77)) });
                menuEspecial.Items.Add(itemFilaArriba);
                menuEspecial.Items.Add(itemFilaAbajo);
                menuEspecial.Items.Add(new Separator { Background = new SolidColorBrush(Color.FromRgb(58, 68, 77)) });
                menuEspecial.Items.Add(itemColIzq);
                menuEspecial.Items.Add(itemColDer);
                menuEspecial.Items.Add(new Separator { Background = new SolidColorBrush(Color.FromRgb(58, 68, 77)) });
                menuEspecial.Items.Add(itemEliminarFila);
                menuEspecial.Items.Add(itemEliminarColumna);

                e.Handled = true; 
                menuEspecial.IsOpen = true;
            }
        }

        private void PintarFondoSeleccion(TableCell celdaClickeada, Table tablaPadre)
        {
            if (_editorActual == null) return;
            
            // Obtenemos el grupo de filas para poder navegar por las coordenadas (R, C)
            TableRowGroup grupoFilas = (TableRowGroup)((TableRow)celdaClickeada.Parent).Parent;

            var dialog = new KalebClipPro.Views.PaletaColores();
            dialog.Owner = Window.GetWindow(this);
            dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            
            dialog.AlSeleccionarColor = (colorNuevo) =>
            {
                var selectionRange = _editorActual.Selection;
                var brushNuevo = new SolidColorBrush(colorNuevo);

                _editorActual.BeginChange();
                
                // Si no hay selección, pintamos solo la celda donde se hizo clic
                if (selectionRange.IsEmpty)
                {
                    celdaClickeada.Background = brushNuevo;
                }
                else
                {
                    // 🌟 SOLUCIÓN: Aplicamos la lógica de "Caja Delimitadora" (Grid Bounding Box)
                    TableCell? celdaInicio = null, celdaFin = null;
                    TextPointer pInicio = selectionRange.Start;
                    TextPointer pFin = selectionRange.End.GetNextInsertionPosition(LogicalDirection.Backward) ?? selectionRange.End;
                    
                    if (pInicio.CompareTo(pFin) > 0) pFin = selectionRange.End;

                    // 1. Detectamos exactamente en qué celda empieza y termina la selección visual
                    foreach (var r in grupoFilas.Rows) {
                        foreach (var c in r.Cells) {
                            if (pInicio.CompareTo(c.ElementStart) >= 0 && pInicio.CompareTo(c.ElementEnd) < 0) celdaInicio = c;
                            if (pFin.CompareTo(c.ElementStart) >= 0 && pFin.CompareTo(c.ElementEnd) < 0) celdaFin = c;
                        }
                    }

                    celdaInicio ??= celdaClickeada; 
                    celdaFin ??= celdaClickeada;

                    // 2. Extraemos las coordenadas X, Y (Filas y Columnas) de esos dos puntos
                    int r1 = grupoFilas.Rows.IndexOf((TableRow)celdaInicio.Parent);
                    int c1 = ((TableRow)celdaInicio.Parent).Cells.IndexOf(celdaInicio);
                    int r2 = grupoFilas.Rows.IndexOf((TableRow)celdaFin.Parent);
                    int c2 = ((TableRow)celdaFin.Parent).Cells.IndexOf(celdaFin);

                    // 3. Calculamos los límites de la caja seleccionada
                    int minFila = Math.Min(r1, r2), maxFila = Math.Max(r1, r2);
                    int minCol = Math.Min(c1, c2), maxCol = Math.Max(c1, c2);

                    // 4. Pintamos de forma quirúrgica solo las celdas dentro de los límites
                    for (int r = minFila; r <= maxFila; r++) 
                    {
                        for (int c = minCol; c <= maxCol; c++) 
                        {
                            if (c < grupoFilas.Rows[r].Cells.Count) 
                            {
                                grupoFilas.Rows[r].Cells[c].Background = brushNuevo;
                            }
                        }
                    }
                }
                _editorActual.EndChange();
                _editorActual.Focus();
            };
            dialog.Show();
        }

        // -------------------------------------------------------------
        // 1. QUITAR COLOR DE FONDO (Transparente)
        // -------------------------------------------------------------
        private void QuitarFondoSeleccion(TableCell celdaClickeada, Table tablaPadre)
        {
            if (_editorActual == null) return;
            TableRowGroup grupoFilas = (TableRowGroup)((TableRow)celdaClickeada.Parent).Parent;
            var selectionRange = _editorActual.Selection;

            _editorActual.BeginChange();
            
            if (selectionRange.IsEmpty)
            {
                celdaClickeada.Background = Brushes.Transparent;
            }
            else
            {
                TableCell? celdaInicio = null, celdaFin = null;
                TextPointer pInicio = selectionRange.Start, pFin = selectionRange.End.GetNextInsertionPosition(LogicalDirection.Backward) ?? selectionRange.End;
                if (pInicio.CompareTo(pFin) > 0) pFin = selectionRange.End;

                foreach (var r in grupoFilas.Rows) {
                    foreach (var c in r.Cells) {
                        if (pInicio.CompareTo(c.ElementStart) >= 0 && pInicio.CompareTo(c.ElementEnd) < 0) celdaInicio = c;
                        if (pFin.CompareTo(c.ElementStart) >= 0 && pFin.CompareTo(c.ElementEnd) < 0) celdaFin = c;
                    }
                }

                celdaInicio ??= celdaClickeada; celdaFin ??= celdaClickeada;
                int r1 = grupoFilas.Rows.IndexOf((TableRow)celdaInicio.Parent), c1 = ((TableRow)celdaInicio.Parent).Cells.IndexOf(celdaInicio);
                int r2 = grupoFilas.Rows.IndexOf((TableRow)celdaFin.Parent), c2 = ((TableRow)celdaFin.Parent).Cells.IndexOf(celdaFin);
                int minFila = Math.Min(r1, r2), maxFila = Math.Max(r1, r2);
                int minCol = Math.Min(c1, c2), maxCol = Math.Max(c1, c2);

                for (int r = minFila; r <= maxFila; r++) for (int c = minCol; c <= maxCol; c++) 
                    if (c < grupoFilas.Rows[r].Cells.Count) 
                        grupoFilas.Rows[r].Cells[c].Background = Brushes.Transparent;
            }
            _editorActual.EndChange();
            _editorActual.Focus();
        }

        // -------------------------------------------------------------
        // 2. CAMBIAR COLOR DE BORDE (Selección o Global)
        // -------------------------------------------------------------
        private void PintarBordeSeleccion(TableCell celdaClickeada, Table tablaPadre)
        {
            if (_editorActual == null) return;
            TableRowGroup grupoFilas = (TableRowGroup)((TableRow)celdaClickeada.Parent).Parent;

            var dialog = new KalebClipPro.Views.PaletaColores();
            dialog.Owner = Window.GetWindow(this);
            dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;

            dialog.AlSeleccionarColor = (colorNuevo) =>
            {
                var selectionRange = _editorActual.Selection;
                var brushNuevo = new SolidColorBrush(colorNuevo);

                _editorActual.BeginChange();

                int minFila = 0, maxFila = 0, minCol = 0, maxCol = 0;

                if (selectionRange.IsEmpty)
                {
                    celdaClickeada.BorderBrush = brushNuevo;
                    
                    // Extraemos las coordenadas aunque sea una sola celda
                    TableRow filaActual = (TableRow)celdaClickeada.Parent;
                    minFila = maxFila = grupoFilas.Rows.IndexOf(filaActual);
                    minCol = maxCol = filaActual.Cells.IndexOf(celdaClickeada);
                }
                else
                {
                    TableCell? celdaInicio = null, celdaFin = null;
                    TextPointer pInicio = selectionRange.Start, pFin = selectionRange.End.GetNextInsertionPosition(LogicalDirection.Backward) ?? selectionRange.End;
                    if (pInicio.CompareTo(pFin) > 0) pFin = selectionRange.End;

                    foreach (var r in grupoFilas.Rows) {
                        foreach (var c in r.Cells) {
                            if (pInicio.CompareTo(c.ElementStart) >= 0 && pInicio.CompareTo(c.ElementEnd) < 0) celdaInicio = c;
                            if (pFin.CompareTo(c.ElementStart) >= 0 && pFin.CompareTo(c.ElementEnd) < 0) celdaFin = c;
                        }
                    }

                    celdaInicio ??= celdaClickeada; celdaFin ??= celdaClickeada;
                    int r1 = grupoFilas.Rows.IndexOf((TableRow)celdaInicio.Parent), c1 = ((TableRow)celdaInicio.Parent).Cells.IndexOf(celdaInicio);
                    int r2 = grupoFilas.Rows.IndexOf((TableRow)celdaFin.Parent), c2 = ((TableRow)celdaFin.Parent).Cells.IndexOf(celdaFin);
                    
                    minFila = Math.Min(r1, r2); maxFila = Math.Max(r1, r2);
                    minCol = Math.Min(c1, c2); maxCol = Math.Max(c1, c2);

                    for (int r = minFila; r <= maxFila; r++) 
                    {
                        for (int c = minCol; c <= maxCol; c++) 
                        {
                            if (c < grupoFilas.Rows[r].Cells.Count) 
                                grupoFilas.Rows[r].Cells[c].BorderBrush = brushNuevo;
                        }
                    }
                }

                // 🌟 EL FIX ESTÁ AQUÍ 🌟
                // Solo actualizamos la memoria global (Tag) si el usuario seleccionó TODA la tabla.
                // Así evitamos que cambiar el borde de una sola celda contagie a las nuevas filas/columnas.
                int totalFilas = grupoFilas.Rows.Count;
                int totalCols = totalFilas > 0 ? grupoFilas.Rows[0].Cells.Count : 0;
                
                if (minFila == 0 && maxFila == totalFilas - 1 && minCol == 0 && maxCol == totalCols - 1)
                {
                    if (tablaPadre.Tag is object[] conf && conf.Length >= 10) {
                        conf[8] = true;         // usaBordePers = true
                        conf[9] = colorNuevo;   // colorBordePers
                        tablaPadre.Tag = conf;
                    }
                }

                _editorActual.EndChange();
                _editorActual.Focus();
            };
            dialog.Show();
        }

        // -------------------------------------------------------------
        // 3. CAMBIAR TEMA/PALETA GLOBAL DE LA TABLA
        // -------------------------------------------------------------
        private void CambiarTemaTabla(TableCell celdaClickeada, Table tablaPadre)
        {
            if (_editorActual == null) return;
            TableRowGroup grupoFilas = (TableRowGroup)((TableRow)celdaClickeada.Parent).Parent;

            // Leemos la configuración actual para saber cómo reconstruir las cebras y encabezados
            int estilo = 1; bool enc = false, tot = false, priCol = false, ultCol = false;
            if (tablaPadre.Tag is object[] config && config.Length >= 10) {
                estilo = Convert.ToInt32(config[0]);
                enc = (bool)config[4]; tot = (bool)config[5]; priCol = (bool)config[6]; ultCol = (bool)config[7];
            }

            var dialog = new KalebClipPro.Views.PaletaColores();
            dialog.Owner = Window.GetWindow(this);
            dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;

            dialog.AlSeleccionarColor = (colorNuevo) =>
            {
                _editorActual.BeginChange();

                // 1. Inyectamos el nuevo acento en la mochila
                if (tablaPadre.Tag is object[] conf && conf.Length >= 10) {
                    conf[1] = colorNuevo; 
                    tablaPadre.Tag = conf;
                }

                // 2. Preparamos las nuevas brochas temáticas
                var colorHeader = new SolidColorBrush(colorNuevo);
                var colorCebra1 = new SolidColorBrush(Color.FromArgb(60, colorNuevo.R, colorNuevo.G, colorNuevo.B));
                var colorCebra2 = new SolidColorBrush(Color.FromArgb(20, colorNuevo.R, colorNuevo.G, colorNuevo.B));

                // 3. Redibujamos inteligentemente toda la tabla respetando el layout
                int filas = grupoFilas.Rows.Count;
                if (filas == 0) return;
                int columnas = grupoFilas.Rows[0].Cells.Count;

                for (int f = 0; f < filas; f++)
                {
                    for (int c = 0; c < grupoFilas.Rows[f].Cells.Count; c++)
                    {
                        TableCell celda = grupoFilas.Rows[f].Cells[c];

                        // Patrón base de cebras
                        if (estilo == 1) celda.Background = (f % 2 == 0) ? colorCebra1 : colorCebra2;
                        else if (estilo == 2) celda.Background = (c % 2 == 0) ? colorCebra1 : colorCebra2;
                        else if (estilo == 3) celda.Background = ((f + c) % 2 == 0) ? colorCebra1 : colorCebra2;
                        else celda.Background = Brushes.Transparent;

                        // Superponer encabezados y columnas reales
                        bool esZonaRealzada = false;
                        if (enc && f == 0) esZonaRealzada = true;
                        if (tot && f == filas - 1) esZonaRealzada = true;
                        if (priCol && c == 0) esZonaRealzada = true;
                        if (ultCol && c == columnas - 1) esZonaRealzada = true;

                        if (esZonaRealzada)
                        {
                            celda.Background = colorHeader;
                        }
                    }
                }
                _editorActual.EndChange();
                _editorActual.Focus();
            };
            dialog.Show();
        }

        private void AplicarBordeASEleccion(TableCell celdaClickeada, string tipoBorde)
        {
            if (_editorActual == null) return;
            
            TableRow filaActual = (TableRow)celdaClickeada.Parent;
            TableRowGroup grupoFilas = (TableRowGroup)filaActual.Parent;
            Table tablaPadre = (Table)grupoFilas.Parent;

            int estilo = 1; Color acento = Color.FromRgb(0, 173, 181); double grosorBase = 0.5;
            bool memoriaViva = false;
            bool usaBordePers = false; Color colorBordePers = Color.FromRgb(58, 68, 77);
            bool tEnc = false, tTot = false, tPri = false, tUlt = false; string tipoBordeReal = "Todos";

            if (tablaPadre.Tag is object[] config)
            {
                try {
                    if (config.Length >= 1) estilo = Convert.ToInt32(config[0]);
                    if (config.Length >= 2 && config[1] is Color c) acento = c;
                    if (config.Length >= 3) grosorBase = Convert.ToDouble(config[2]); 
                    if (config.Length >= 10) { 
                        tEnc = (bool)config[4]; tTot = (bool)config[5]; tPri = (bool)config[6]; tUlt = (bool)config[7];
                        usaBordePers = (bool)config[8]; colorBordePers = (Color)config[9]; 
                    }
                    memoriaViva = true;
                } catch { } 
            }

            if (!memoriaViva && grupoFilas.Rows.Count > 0 && grupoFilas.Rows[0].Cells.Count > 0)
            {
                TableCell celda0 = grupoFilas.Rows[0].Cells[0];
                int fCount = grupoFilas.Rows.Count, cCount = grupoFilas.Rows[0].Cells.Count;
                int fMed = fCount > 2 ? 1 : 0;

                double maxTB = Math.Max(celda0.BorderThickness.Top, celda0.BorderThickness.Bottom);
                double maxLR = Math.Max(celda0.BorderThickness.Left, celda0.BorderThickness.Right);
                if (maxTB > 0 && maxLR == 0) tipoBordeReal = "Horizontales";
                else if (maxLR > 0 && maxTB == 0) tipoBordeReal = "Verticales";
                grosorBase = Math.Max(maxTB, maxLR);
                if (grosorBase == 0) grosorBase = 0.5;

                estilo = 0;
                foreach (var r in grupoFilas.Rows) {
                    foreach (var c in r.Cells) {
                        if (c.Background is SolidColorBrush bg && bg.Color.A > 0 && bg.Color.A < 255) {
                            estilo = 1; acento = Color.FromRgb(bg.Color.R, bg.Color.G, bg.Color.B); break;
                        }
                    }
                    if (estilo != 0) break;
                }

                if (celda0.BorderBrush is SolidColorBrush sb) {
                    if (sb.Color.A == 255 && (sb.Color.R != 58 || sb.Color.G != 68 || sb.Color.B != 77)) {
                        usaBordePers = true; colorBordePers = sb.Color; 
                    } else if (sb.Color.A == 120 && estilo == 0) {
                        acento = Color.FromRgb(sb.Color.R, sb.Color.G, sb.Color.B);
                    }
                }

                // 🌟 FIX INTELIGENTE: Diferenciar celda pintada
                bool CeldaEsRealce(TableCell c) {
                    return c.Background is SolidColorBrush bg && bg.Color.A == 255 && 
                           c.Blocks.FirstBlock is Paragraph p && p.FontWeight == FontWeights.Bold &&
                           p.Foreground is SolidColorBrush fg && fg.Color.R == 255 && fg.Color.G == 255 && fg.Color.B == 255;
                }

                if (fCount > 0 && CeldaEsRealce(grupoFilas.Rows[0].Cells[0])) tEnc = true;
                if (fCount > 1 && CeldaEsRealce(grupoFilas.Rows[fCount-1].Cells[0])) tTot = true;
                if (cCount > 0 && fCount > 1 && CeldaEsRealce(grupoFilas.Rows[fMed].Cells[0])) { 
                    tPri = true; estilo = 4; 
                    var bP = (SolidColorBrush)grupoFilas.Rows[fMed].Cells[0].Background;
                    acento = Color.FromRgb(bP.Color.R, bP.Color.G, bP.Color.B); 
                }
                if (cCount > 1 && fCount > 1 && CeldaEsRealce(grupoFilas.Rows[fMed].Cells[cCount-1])) tUlt = true;

                tablaPadre.Tag = new object[] { estilo, acento, grosorBase, tipoBordeReal, tEnc, tTot, tPri, tUlt, usaBordePers, colorBordePers }; 
            }
            
            Brush brushBorde = (estilo == 0) ? new SolidColorBrush(Color.FromRgb(58, 68, 77)) : new SolidColorBrush(Color.FromArgb(120, acento.R, acento.G, acento.B));
            if (usaBordePers) brushBorde = new SolidColorBrush(colorBordePers);

            _editorActual.BeginChange();
            if (tipoBorde == "Ninguno" || tipoBorde == "Todos" || tipoBorde == "Externos") tablaPadre.BorderThickness = new Thickness(0);

            var selectionRange = _editorActual.Selection;
            List<TableCell> celdasAfectadas = new List<TableCell>();
            int minFila = 0, maxFila = 0, minCol = 0, maxCol = 0;

            if (selectionRange.IsEmpty)
            {
                celdasAfectadas.Add(celdaClickeada);
                minFila = maxFila = grupoFilas.Rows.IndexOf(filaActual);
                minCol = maxCol = filaActual.Cells.IndexOf(celdaClickeada);
            }
            else
            {
                TableCell? celdaInicio = null, celdaFin = null;
                TextPointer pInicio = selectionRange.Start, pFin = selectionRange.End.GetNextInsertionPosition(LogicalDirection.Backward) ?? selectionRange.End;
                if (pInicio.CompareTo(pFin) > 0) pFin = selectionRange.End;

                foreach (var r in grupoFilas.Rows) foreach (var c in r.Cells) {
                    if (pInicio.CompareTo(c.ElementStart) >= 0 && pInicio.CompareTo(c.ElementEnd) < 0) celdaInicio = c;
                    if (pFin.CompareTo(c.ElementStart) >= 0 && pFin.CompareTo(c.ElementEnd) < 0) celdaFin = c;
                }

                celdaInicio ??= celdaClickeada; celdaFin ??= celdaClickeada;
                int r1 = grupoFilas.Rows.IndexOf((TableRow)celdaInicio.Parent), c1 = ((TableRow)celdaInicio.Parent).Cells.IndexOf(celdaInicio);
                int r2 = grupoFilas.Rows.IndexOf((TableRow)celdaFin.Parent), c2 = ((TableRow)celdaFin.Parent).Cells.IndexOf(celdaFin);

                minFila = Math.Min(r1, r2); maxFila = Math.Max(r1, r2);
                minCol = Math.Min(c1, c2); maxCol = Math.Max(c1, c2);

                for (int r = minFila; r <= maxFila; r++) for (int c = minCol; c <= maxCol; c++) if (c < grupoFilas.Rows[r].Cells.Count) celdasAfectadas.Add(grupoFilas.Rows[r].Cells[c]);
            }

            foreach (TableCell cell in celdasAfectadas)
            {
                TableRow row = (TableRow)cell.Parent;
                int r = grupoFilas.Rows.IndexOf(row), c = row.Cells.IndexOf(cell);
                Thickness current = cell.BorderThickness;
                Thickness newThick = new Thickness(current.Left, current.Top, current.Right, current.Bottom);

                if (tipoBorde == "Ninguno") newThick = new Thickness(0); 
                else if (tipoBorde == "Todos") newThick = new Thickness(grosorBase); 
                else if (tipoBorde == "Superior") newThick.Top = grosorBase;
                else if (tipoBorde == "Inferior") newThick.Bottom = grosorBase;
                else if (tipoBorde == "Izquierdo") newThick.Left = grosorBase;
                else if (tipoBorde == "Derecho") newThick.Right = grosorBase;
                else if (tipoBorde == "Externos") {
                    if (r == minFila) newThick.Top = grosorBase; if (r == maxFila) newThick.Bottom = grosorBase;
                    if (c == minCol) newThick.Left = grosorBase; if (c == maxCol) newThick.Right = grosorBase;
                }
                else if (tipoBorde == "Internos") {
                    if (r != minFila) newThick.Top = grosorBase; if (r != maxFila) newThick.Bottom = grosorBase;
                    if (c != minCol) newThick.Left = grosorBase; if (c != maxCol) newThick.Right = grosorBase;
                }

                cell.BorderBrush = brushBorde;
                cell.BorderThickness = newThick;
            }
            _editorActual.EndChange();
        }

        private void AplicarGrosorASeleccion(TableCell celdaClickeada, double nuevoGrosor)
        {
            if (_editorActual == null) return;
            
            TableRow filaActual = (TableRow)celdaClickeada.Parent;
            TableRowGroup grupoFilas = (TableRowGroup)filaActual.Parent;
            Table tablaPadre = (Table)grupoFilas.Parent;

            int estilo = 1; Color acento = Color.FromRgb(0, 173, 181); string tipoBordeOriginal = "Todos";
            bool memoriaViva = false; bool usaBordePers = false; Color colorBordePers = Color.FromRgb(58, 68, 77);
            bool tEnc = false, tTot = false, tPri = false, tUlt = false;
            
            if (tablaPadre.Tag is object[] config)
            {
                try {
                    if (config.Length >= 1) estilo = Convert.ToInt32(config[0]);
                    if (config.Length >= 2 && config[1] is Color c) acento = c;
                    if (config.Length >= 4 && config[3] is string t) tipoBordeOriginal = t;
                    if (config.Length >= 10) { 
                        tEnc = (bool)config[4]; tTot = (bool)config[5]; tPri = (bool)config[6]; tUlt = (bool)config[7];
                        usaBordePers = (bool)config[8]; colorBordePers = (Color)config[9]; 
                    }
                    memoriaViva = true;
                } catch { }
            }
            
            if (!memoriaViva && grupoFilas.Rows.Count > 0 && grupoFilas.Rows[0].Cells.Count > 0)
            {
                TableCell celda0 = grupoFilas.Rows[0].Cells[0];
                int fCount = grupoFilas.Rows.Count, cCount = grupoFilas.Rows[0].Cells.Count;
                int fMed = fCount > 2 ? 1 : 0;

                double maxTB = Math.Max(celda0.BorderThickness.Top, celda0.BorderThickness.Bottom);
                double maxLR = Math.Max(celda0.BorderThickness.Left, celda0.BorderThickness.Right);
                if (maxTB > 0 && maxLR == 0) tipoBordeOriginal = "Horizontales";
                else if (maxLR > 0 && maxTB == 0) tipoBordeOriginal = "Verticales";
                double grosorBase = Math.Max(maxTB, maxLR);
                if (grosorBase == 0) grosorBase = 0.5;

                estilo = 0;
                foreach (var r in grupoFilas.Rows) {
                    foreach (var c in r.Cells) {
                        if (c.Background is SolidColorBrush bg && bg.Color.A > 0 && bg.Color.A < 255) {
                            estilo = 1; acento = Color.FromRgb(bg.Color.R, bg.Color.G, bg.Color.B); break;
                        }
                    }
                    if (estilo != 0) break;
                }

                if (celda0.BorderBrush is SolidColorBrush sb) {
                    if (sb.Color.A == 255 && (sb.Color.R != 58 || sb.Color.G != 68 || sb.Color.B != 77)) {
                        usaBordePers = true; colorBordePers = sb.Color; 
                    } else if (sb.Color.A == 120 && estilo == 0) {
                        acento = Color.FromRgb(sb.Color.R, sb.Color.G, sb.Color.B);
                    }
                }

                // 🌟 FIX INTELIGENTE: Diferenciar celda pintada
                bool CeldaEsRealce(TableCell c) {
                    return c.Background is SolidColorBrush bg && bg.Color.A == 255 && 
                           c.Blocks.FirstBlock is Paragraph p && p.FontWeight == FontWeights.Bold &&
                           p.Foreground is SolidColorBrush fg && fg.Color.R == 255 && fg.Color.G == 255 && fg.Color.B == 255;
                }

                if (fCount > 0 && CeldaEsRealce(grupoFilas.Rows[0].Cells[0])) tEnc = true;
                if (fCount > 1 && CeldaEsRealce(grupoFilas.Rows[fCount-1].Cells[0])) tTot = true;
                if (cCount > 0 && fCount > 1 && CeldaEsRealce(grupoFilas.Rows[fMed].Cells[0])) { 
                    tPri = true; estilo = 4; 
                    var bP = (SolidColorBrush)grupoFilas.Rows[fMed].Cells[0].Background;
                    acento = Color.FromRgb(bP.Color.R, bP.Color.G, bP.Color.B); 
                }
                if (cCount > 1 && fCount > 1 && CeldaEsRealce(grupoFilas.Rows[fMed].Cells[cCount-1])) tUlt = true;

                tablaPadre.Tag = new object[] { estilo, acento, grosorBase, tipoBordeOriginal, tEnc, tTot, tPri, tUlt, usaBordePers, colorBordePers }; 
            }
            
            Brush brushBorde = (estilo == 0) ? new SolidColorBrush(Color.FromRgb(58, 68, 77)) : new SolidColorBrush(Color.FromArgb(120, acento.R, acento.G, acento.B));
            if (usaBordePers) brushBorde = new SolidColorBrush(colorBordePers);

            _editorActual.BeginChange();
            var selectionRange = _editorActual.Selection;
            List<TableCell> celdasAfectadas = new List<TableCell>();

            if (selectionRange.IsEmpty) celdasAfectadas.Add(celdaClickeada);
            else
            {
                TableCell? celdaInicio = null, celdaFin = null;
                TextPointer pInicio = selectionRange.Start, pFin = selectionRange.End.GetNextInsertionPosition(LogicalDirection.Backward) ?? selectionRange.End;
                if (pInicio.CompareTo(pFin) > 0) pFin = selectionRange.End;

                foreach (var r in grupoFilas.Rows) foreach (var c in r.Cells) {
                    if (pInicio.CompareTo(c.ElementStart) >= 0 && pInicio.CompareTo(c.ElementEnd) < 0) celdaInicio = c;
                    if (pFin.CompareTo(c.ElementStart) >= 0 && pFin.CompareTo(c.ElementEnd) < 0) celdaFin = c;
                }

                celdaInicio ??= celdaClickeada; celdaFin ??= celdaClickeada;
                int r1 = grupoFilas.Rows.IndexOf((TableRow)celdaInicio.Parent), c1 = ((TableRow)celdaInicio.Parent).Cells.IndexOf(celdaInicio);
                int r2 = grupoFilas.Rows.IndexOf((TableRow)celdaFin.Parent), c2 = ((TableRow)celdaFin.Parent).Cells.IndexOf(celdaFin);

                int minFila = Math.Min(r1, r2), maxFila = Math.Max(r1, r2);
                int minCol = Math.Min(c1, c2), maxCol = Math.Max(c1, c2);

                for (int r = minFila; r <= maxFila; r++) for (int c = minCol; c <= maxCol; c++) if (c < grupoFilas.Rows[r].Cells.Count) celdasAfectadas.Add(grupoFilas.Rows[r].Cells[c]);
            }

            foreach (TableCell cell in celdasAfectadas)
            {
                Thickness t = cell.BorderThickness;
                t.Left = t.Left > 0 ? nuevoGrosor : 0; 
                t.Top = t.Top > 0 ? nuevoGrosor : 0;
                t.Right = t.Right > 0 ? nuevoGrosor : 0; 
                t.Bottom = t.Bottom > 0 ? nuevoGrosor : 0;

                cell.BorderBrush = brushBorde;
                cell.BorderThickness = t;
            }
            _editorActual.EndChange();
        }
    }
}