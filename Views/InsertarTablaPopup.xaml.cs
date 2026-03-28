using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace KalebClipPro.Views
{
    public partial class InsertarTablaPopup : Window
    {
        public int FilasSeleccionadas { get; private set; } = 3;
        public int ColumnasSeleccionadas { get; private set; } = 3;
        public int EstiloSeleccionado { get; private set; } = 1; 
        public string BordeSeleccionado { get; private set; } = "Todos"; 
        
        public Color ColorAcentoSeleccionado { get; private set; } = Color.FromRgb(0, 173, 181); 
        public double GrosorSeleccionado { get; private set; } = 0.5;

        // Propiedades de las Opciones de Estilo (CheckBoxes)
        public bool TieneEncabezado { get; private set; } = true;
        public bool TieneTotales { get; private set; } = false;
        public bool TienePrimeraColumna { get; private set; } = false;
        public bool TieneUltimaColumna { get; private set; } = false;

        // Propiedades para el Color de Borde Personalizado
        public Color ColorBordeSeleccionado { get; private set; } = Color.FromRgb(58, 68, 77); 
        public bool UsaColorBordePersonalizado { get; private set; } = false;

        public InsertarTablaPopup()
        {
            InitializeComponent();
            
            BtnCancelar.Click += (s, e) => { this.DialogResult = false; this.Close(); };
            BtnInsertar.Click += BtnInsertar_Click;
            
            ActualizarPreview(1); 
        }

        private void Color_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border borderSeleccionado && borderSeleccionado.Tag is string hexColor)
            {
                ColorAcentoSeleccionado = (Color)ColorConverter.ConvertFromString(hexColor);
                foreach (Border b in PanelColores.Children) b.BorderThickness = new Thickness(0);
                borderSeleccionado.BorderThickness = new Thickness(2);
                borderSeleccionado.BorderBrush = Brushes.White;
                ActualizarPreview(CmbEstilo.SelectedIndex);
            }
        }

        private void CmbEstilo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PreviewGrid != null && CmbEstilo != null) 
            {
                ActualizarPreview(CmbEstilo.SelectedIndex);
            }
        }

        private void OpcionesEstilo_Changed(object sender, RoutedEventArgs e)
        {
            if (CmbEstilo != null) ActualizarPreview(CmbEstilo.SelectedIndex);
        }

        private void Borde_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string tipo)
            {
                BordeSeleccionado = tipo;
                ActualizarPreview(CmbEstilo.SelectedIndex);
            }
        }

        // --------------------------------------------------------------------
        // REEMPLAZO: Selector de Color de Borde usando tu propia Paleta
        // --------------------------------------------------------------------
        private void BtnColorBorde_Click(object sender, MouseButtonEventArgs e) 
        {
            var dialog = new KalebClipPro.Views.PaletaColores();
            dialog.Owner = Window.GetWindow(this);
            dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            
            // Opcional: Si tu paleta tiene este método, se lo pasamos para que inicie en el color actual
            // dialog.CargarColorDesdeEditor(ColorBordeSeleccionado); 

            dialog.AlSeleccionarColor = (colorElegido) =>
            {
                ColorBordeSeleccionado = colorElegido;
                UsaColorBordePersonalizado = true; 
                
                // Actualizamos el fondo del cuadrito
                ColorBordePreviewUI.Background = new SolidColorBrush(colorElegido);
                
                // Refrescamos la vista previa de la tabla
                if (CmbEstilo != null) ActualizarPreview(CmbEstilo.SelectedIndex);
            };
            
            dialog.Show();
        }

        // --------------------------------------------------------------------
        // NUEVO: Selector de Color de Acento (Botón Arcoíris)
        // --------------------------------------------------------------------
        private void BtnColorAcentoPersonalizado_Click(object sender, MouseButtonEventArgs e)
        {
            var dialog = new KalebClipPro.Views.PaletaColores();
            dialog.Owner = Window.GetWindow(this);
            dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;

            // dialog.CargarColorDesdeEditor(ColorAcentoSeleccionado);

            dialog.AlSeleccionarColor = (colorNuevo) =>
            {
                ColorAcentoSeleccionado = colorNuevo;
                
                // 1. Apagamos el borde blanco de todos los demás circulitos
                foreach (Border b in PanelColores.Children) b.BorderThickness = new Thickness(0);
                
                // 2. Le encendemos el borde blanco al circulito arcoíris para mostrar que está seleccionado
                if (sender is Border borderArcoiris)
                {
                    borderArcoiris.BorderThickness = new Thickness(2);
                    borderArcoiris.BorderBrush = Brushes.White;
                }

                // 3. Actualizamos la vista previa de la tabla
                if (CmbEstilo != null) ActualizarPreview(CmbEstilo.SelectedIndex);
            };
            
            dialog.Show();
        }

        private void ActualizarPreview(int estilo)
        {
            if (PreviewGrid == null) return;

            PreviewGrid.Children.Clear();
            PreviewGrid.RowDefinitions.Clear();
            PreviewGrid.ColumnDefinitions.Clear();

            for (int i = 0; i < 3; i++) { PreviewGrid.RowDefinitions.Add(new RowDefinition()); }
            for (int i = 0; i < 3; i++) { PreviewGrid.ColumnDefinitions.Add(new ColumnDefinition()); }

            Color acento = ColorAcentoSeleccionado;
            var colorHeader = new SolidColorBrush(acento);
            var colorCebra1 = new SolidColorBrush(Color.FromArgb(60, acento.R, acento.G, acento.B));
            var colorCebra2 = new SolidColorBrush(Color.FromArgb(20, acento.R, acento.G, acento.B));
            var colorBordeTematico = new SolidColorBrush(Color.FromArgb(120, acento.R, acento.G, acento.B));
            var colorBordeGris = new SolidColorBrush(Color.FromRgb(58, 68, 77));

            Brush brushBordeAplicar = (estilo == 0) ? colorBordeGris : colorBordeTematico;
            if (UsaColorBordePersonalizado)
            {
                brushBordeAplicar = new SolidColorBrush(ColorBordeSeleccionado);
            }

            for (int r = 0; r < 3; r++)
            {
                for (int c = 0; c < 3; c++)
                {
                    TextBlock textoSimulado = new TextBlock { Text = "Aa", VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center, FontSize = 10 };
                    Border celda = new Border { Child = textoSimulado };
                    
                    textoSimulado.Foreground = Brushes.LightGray; 
                    celda.Background = Brushes.Transparent; 
                    celda.BorderBrush = brushBordeAplicar; 

                    // PASO 1: PINTAR EL PATRÓN BASE
                    if (estilo == 1) celda.Background = (r % 2 == 0) ? colorCebra1 : colorCebra2;      
                    else if (estilo == 2) celda.Background = (c % 2 == 0) ? colorCebra1 : colorCebra2; 
                    else if (estilo == 3) celda.Background = ((r + c) % 2 == 0) ? colorCebra1 : colorCebra2; 

                    // PASO 2: APLICAR LOS REALCES (Aplica también al estilo Sencillo)
                    bool esZonaRealzada = false;
                    if (ChkEncabezado?.IsChecked == true && r == 0) esZonaRealzada = true;
                    if (ChkTotales?.IsChecked == true && r == 2) esZonaRealzada = true; 
                    if (ChkPrimeraColumna?.IsChecked == true && c == 0) esZonaRealzada = true;
                    if (ChkUltimaColumna?.IsChecked == true && c == 2) esZonaRealzada = true; 

                    if (esZonaRealzada)
                    {
                        celda.Background = colorHeader; 
                        textoSimulado.Foreground = Brushes.White;
                        if (estilo == 0 && !UsaColorBordePersonalizado) celda.BorderBrush = colorBordeTematico; 
                    }

                    // GROSOR Y BORDES
                    double grosorBase = 0.5;
                    if (CmbGrosor != null)
                    {
                        if (CmbGrosor.SelectedIndex == 1) grosorBase = 1.0;
                        else if (CmbGrosor.SelectedIndex == 2) grosorBase = 2.0;
                        else if (CmbGrosor.SelectedIndex == 3) grosorBase = 3.0;
                    }

                    Thickness grosorBorde = new Thickness(0);
                    if (BordeSeleccionado == "Todos") grosorBorde = new Thickness(grosorBase);
                    else if (BordeSeleccionado == "Horizontales") grosorBorde = new Thickness(0, 0, 0, grosorBase);
                    else if (BordeSeleccionado == "Verticales") grosorBorde = new Thickness(0, 0, grosorBase, 0);
                    
                    celda.BorderThickness = grosorBorde;
                    
                    Grid.SetRow(celda, r);
                    Grid.SetColumn(celda, c);
                    PreviewGrid.Children.Add(celda);
                }
            }
        }

        private void BtnInsertar_Click(object sender, RoutedEventArgs e)
        {
            if (int.TryParse(TxtFilas.Text, out int f)) FilasSeleccionadas = f;
            if (int.TryParse(TxtColumnas.Text, out int c)) ColumnasSeleccionadas = c;

            if (FilasSeleccionadas < 1) FilasSeleccionadas = 1;
            if (ColumnasSeleccionadas < 1) ColumnasSeleccionadas = 1;
            
            EstiloSeleccionado = CmbEstilo.SelectedIndex; 
            
            TieneEncabezado = ChkEncabezado.IsChecked == true;
            TieneTotales = ChkTotales.IsChecked == true;
            TienePrimeraColumna = ChkPrimeraColumna.IsChecked == true;
            TieneUltimaColumna = ChkUltimaColumna.IsChecked == true;

            if (CmbGrosor.SelectedIndex == 0) GrosorSeleccionado = 0.5;
            else if (CmbGrosor.SelectedIndex == 1) GrosorSeleccionado = 1.0;
            else if (CmbGrosor.SelectedIndex == 2) GrosorSeleccionado = 2.0;
            else if (CmbGrosor.SelectedIndex == 3) GrosorSeleccionado = 3.0;
            
            this.DialogResult = true; 
            this.Close();
        }
    }
}