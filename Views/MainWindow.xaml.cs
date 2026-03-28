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
        ClipboardActionService _actionService;
        
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
                ConfigurarServiciosYEventos();
                InicializarDatos();
                
                this.Topmost = false;
                _autoSaveTimer.Interval = TimeSpan.FromMilliseconds(350); 
                _autoSaveTimer.Tick += AutoSaveTimer_Tick; 
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error fatal en el diseño:\n{ex.Message}\n\nDetalle:\n{ex.InnerException?.Message}", 
                                "Error de WPF", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
            }
        }

        private void ConfigurarServiciosYEventos()
        {
            TextBoxContainer.PreviewKeyDown += (s, e) => { if (_editorActual != null) Helpers.TableSurgeonHelper.SmartEditor_PreviewKeyDown(_editorActual, s, e); };
            TextBoxContainer.PreviewMouseRightButtonUp += (s, e) => { if (_editorActual != null) Helpers.TableSurgeonHelper.MostrarMenuContextual(_editorActual, e, this); };

            // 🌟 CONFIGURAR EL SERVICIO DE ACCIONES 🌟
            _actionService = new ClipboardActionService(_sysManager);
            _actionService.OnTextoInyectado = (texto) => _ultimoTextoInyectado = texto;
            _actionService.OnEstadoCapturaCambiado = (estado) => _capturandoParaRecolector = estado;
            _actionService.OnNotificarCambioTab = (color) => { TabRecolector.IsChecked = true; NotificarCambioTab(color); };
            _actionService.OnActualizarContadorTab = () => { /* Asegúrate de que este método exista */ };
            _actionService.OnAvanzarSet = () => {
                var llaves = WorkflowActual.Sets.Keys.OrderBy(k => k).ToList();
                int indiceActual = llaves.IndexOf(_setActual);
                _setActual = llaves[(indiceActual + 1) % llaves.Count];
                NotificarCambioTab(Colors.Cyan);
            };
        }

        private void InicializarDatos()
        {
            WorkflowActual = _workflowService.CargarWorkflowDesdeDisco();
            db.InicializarBaseDeDatos();
            ListaClips.ItemsSource = MisClips;
            ListaRecolector.ItemsSource = ClipsRecolector; 
            
            SincronizarBotonesConDatos();
            InicializarSlotsVacios();
            CargarDatos();
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
                _actionService.ProcesarAtajoGlobal(wParam.ToInt32(), ClipsRecolector, MisClips, Dispatcher);
            }
            else if (msg == ClipboardManager.MSG_CLIPBOARDUPDATE)
            {
                ProcesarNuevoPortapapeles();
            }
            return IntPtr.Zero;
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
            _actionService.GuardarNuevoClipEnHistorial(db, MisClips, _ultimoTextoInyectado, 
                (estado) => _capturandoParaRecolector = estado, 
                _capturandoParaRecolector, 
                _filtrosActivos);
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
            if (CmbFuentes.SelectedItem is FontFamily fuente)
            {
                _fuentePreferida = fuente; // Guardamos la preferencia en la ventana
                Helpers.RichTextFormatterHelper.CambiarFuente(_editorActual!, fuente);
            }
        }

        private void CmbTamano_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbTamano.SelectedItem is ComboBoxItem item && double.TryParse(item.Content.ToString(), out double tamano))
            {
                Helpers.RichTextFormatterHelper.CambiarTamano(_editorActual!, tamano);
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
            Helpers.RichTextFormatterHelper.CambiarColorTexto(_editorActual!, sender, this);
        }

        // --- ALINEACIÓN ---
        private void BtnMdAlignLeft_Click(object sender, RoutedEventArgs e) { CerrarPopups(); EjecutarFormato(EditingCommands.AlignLeft); }
        private void BtnMdAlignCenter_Click(object sender, RoutedEventArgs e) { CerrarPopups(); EjecutarFormato(EditingCommands.AlignCenter); }
        private void BtnMdAlignRight_Click(object sender, RoutedEventArgs e) { CerrarPopups(); EjecutarFormato(EditingCommands.AlignRight); }

        // --- LISTAS Y SANGRÍAS ---
private void BtnMdList_Click(object sender, RoutedEventArgs e) { CerrarPopups(); EjecutarFormato(EditingCommands.ToggleBullets); }
        private void BtnMdListNum_Click(object sender, RoutedEventArgs e) { CerrarPopups(); Helpers.RichTextFormatterHelper.AplicarListaNumeradaInteligente(_editorActual!); }
        private void BtnMdIndentInc_Click(object sender, RoutedEventArgs e) { CerrarPopups(); Helpers.RichTextFormatterHelper.AplicarSangriaPersonalizada(_editorActual!, 1); }
        private void BtnMdIndentDec_Click(object sender, RoutedEventArgs e) { CerrarPopups(); Helpers.RichTextFormatterHelper.AplicarSangriaPersonalizada(_editorActual!, -1); }

        // --- INSERCIÓN COMPLEJA ---
        private void BtnMdQuote_Click(object sender, RoutedEventArgs e) { CerrarPopups(); Helpers.RichTextFormatterHelper.InsertarCita(_editorActual!); }
        
        private void BtnMdCode_Click(object sender, RoutedEventArgs e) { CerrarPopups(); Helpers.RichTextFormatterHelper.FormatearComoCodigo(_editorActual!); }
        
        private void BtnMdLink_Click(object sender, RoutedEventArgs e) { CerrarPopups(); Helpers.RichTextFormatterHelper.InsertarEnlace(_editorActual!, "https://google.com"); } // Después podemos cambiar este URL por un Input
        
        private void BtnMdMath_Click(object sender, RoutedEventArgs e) { CerrarPopups(); Helpers.RichTextFormatterHelper.FormatearComoMatematicas(_editorActual!); }
        
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