using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace KalebClipPro.Views
{
    public partial class PaletaColores : Window
    {
        public Color ColorSeleccionado { get; private set; } = Colors.Red;
        public Action<Color>? AlSeleccionarColor;
        
        private double _hue = 0; // Tono (0-360)
        private double _saturation = 1; // Saturación (0-1)
        private bool _actualizando = false;
        private bool _cerrando = false;

        private static List<Color> _coloresRecientes = new List<Color>();

        public PaletaColores()
        {
            InitializeComponent();
            ActualizarMotorDeColor();
            DibujarColoresRecientes(); 
            
            this.Loaded += (s, e) => 
            {
                this.Activate(); // Despierta la ventana
                TxtHex.Focus(); // Pone el foco en la caja Hex
                Keyboard.Focus(TxtHex); // Obliga a Windows a mandar las teclas ahí
                TxtHex.SelectAll(); // Selecciona el texto para que si pegas algo, se reemplace directo
            };
        }

        private void Window_Deactivated(object sender, EventArgs e)
        {
            if (_cerrando) return;
            _cerrando = true;

            // Esperamos a que el SO termine la transferencia de foco
            Application.Current.Dispatcher.BeginInvoke(new Action(() =>
            {
                this.Close();
            }), System.Windows.Threading.DispatcherPriority.ContextIdle);
        }

        // 🌟 Cierre al dar clic en Aplicar
        private void BtnAplicar_Click(object sender, RoutedEventArgs e)
        {
            if (_cerrando) return;
            _cerrando = true;

            AlSeleccionarColor?.Invoke(ColorSeleccionado);
            
            if (this.Owner != null)
            {
                this.Owner.Activate();
            }

            Application.Current.Dispatcher.BeginInvoke(new Action(() =>
            {
                this.Close();
            }), System.Windows.Threading.DispatcherPriority.ContextIdle);

            if (!_coloresRecientes.Contains(ColorSeleccionado))
            {
                _coloresRecientes.Insert(0, ColorSeleccionado); 
                if (_coloresRecientes.Count > 8) 
                {
                    _coloresRecientes.RemoveAt(8);
                }
            }
        }

        // --- LÓGICA DEL LIENZO (Tono y Saturación) ---
        private void ColorCanvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            ActualizarDesdeLienzo(e.GetPosition(ColorCanvas));
            ((UIElement)sender).CaptureMouse();
        }

        private void ColorCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
                ActualizarDesdeLienzo(e.GetPosition(ColorCanvas));
        }

        private void ColorCanvas_MouseUp(object sender, MouseButtonEventArgs e)
        {
            ((UIElement)sender).ReleaseMouseCapture();
        }

        private void ActualizarDesdeLienzo(Point p)
        {
            double x = Math.Max(0, Math.Min(p.X, ColorCanvas.ActualWidth));
            double y = Math.Max(0, Math.Min(p.Y, ColorCanvas.ActualHeight));

            CursorColor.Margin = new Thickness(x - 5, y - 5, 0, 0);

            _hue = (x / ColorCanvas.ActualWidth) * 360;
            _saturation = 1.0 - (y / ColorCanvas.ActualHeight);

            ColorPuroGradient.Color = HslToRgb(_hue, _saturation, 0.5);
            ActualizarMotorDeColor();
        }

        // --- LÓGICA DEL SLIDER (Luminosidad) ---
        private void SliderLuminosidad_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!_actualizando) ActualizarMotorDeColor();
        }

        private void SliderLuminosidad_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            SliderLuminosidad.CaptureMouse();
            ActualizarPosicionSlider(e.GetPosition(SliderLuminosidad));
        }

        private void SliderLuminosidad_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (SliderLuminosidad.IsMouseCaptured)
            {
                ActualizarPosicionSlider(e.GetPosition(SliderLuminosidad));
            }
        }

        private void SliderLuminosidad_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (SliderLuminosidad.IsMouseCaptured)
            {
                SliderLuminosidad.ReleaseMouseCapture();
            }
        }

        private void ActualizarPosicionSlider(Point p)
        {
            double y = p.Y;
            double altura = SliderLuminosidad.ActualHeight;
            double nuevoValor = 1.0 - (y / altura);
            SliderLuminosidad.Value = Math.Max(0, Math.Min(1, nuevoValor));
        }

        // --- MOTOR CENTRAL ---
        private void ActualizarMotorDeColor()
        {
            _actualizando = true;
            try
            {
                double lightness = SliderLuminosidad?.Value ?? 0.5;
                ColorSeleccionado = HslToRgb(_hue, _saturation, lightness);

                if (BrdVistaPrevia != null) BrdVistaPrevia.Background = new SolidColorBrush(ColorSeleccionado);
                if (TxtR != null) TxtR.Text = ColorSeleccionado.R.ToString();
                if (TxtG != null) TxtG.Text = ColorSeleccionado.G.ToString();
                if (TxtB != null) TxtB.Text = ColorSeleccionado.B.ToString();
                if (TxtHex != null) TxtHex.Text = $"#{ColorSeleccionado.R:X2}{ColorSeleccionado.G:X2}{ColorSeleccionado.B:X2}";
            }
            finally
            {
                _actualizando = false;
            }
        }

        // 🌟 Función para cargar el color externamente (Optimizada)
        public void CargarColorDesdeEditor(Color color)
        {
            SincronizarDesdeColor(color);
        }

        // 🌟 MÉTODO MAESTRO: Traduce un color a la UI con escudos de seguridad
        private void SincronizarDesdeColor(Color color, TextBox? origen = null)
        {
            if (_actualizando) return;
            
            try
            {
                _actualizando = true; 

                double r = color.R / 255.0; double g = color.G / 255.0; double b = color.B / 255.0;
                double max = Math.Max(r, Math.Max(g, b)); double min = Math.Min(r, Math.Min(g, b));
                double h = 0, s = 0, l = (max + min) / 2.0;

                if (max != min)
                {
                    double d = max - min;
                    s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);
                    if (max == r) h = (g - b) / d + (g < b ? 6 : 0);
                    else if (max == g) h = (b - r) / d + 2;
                    else if (max == b) h = (r - g) / d + 4;
                    h /= 6.0;
                }

                _hue = h * 360;
                _saturation = s;
                SliderLuminosidad.Value = l; 

                this.Dispatcher.BeginInvoke(new Action(() => {
                    double x = (h * ColorCanvas.ActualWidth);
                    double y = (1.0 - s) * ColorCanvas.ActualHeight;
                    CursorColor.Margin = new Thickness(x - 5, y - 5, 0, 0);
                    
                    ColorPuroGradient.Color = HslToRgb(_hue, _saturation, 0.5); 
                }), System.Windows.Threading.DispatcherPriority.Loaded);

                ColorSeleccionado = color;
                BrdVistaPrevia.Background = new SolidColorBrush(color);

                if (origen != TxtR && origen != TxtG && origen != TxtB)
                {
                    TxtR.Text = color.R.ToString();
                    TxtG.Text = color.G.ToString();
                    TxtB.Text = color.B.ToString();
                }
                if (origen != TxtHex)
                {
                    TxtHex.Text = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
                }
            }
            finally
            {
                _actualizando = false; 
            }
        }

        // --- EVENTOS DE INTERFAZ (Textos y Círculos) ---
        private void Swatch_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border brd && brd.Background is SolidColorBrush brush)
            {
                SincronizarDesdeColor(brush.Color);
            }
        }

        private void DibujarColoresRecientes()
        {
            PanelRecientes.Children.Clear();

            if (_coloresRecientes.Count == 0)
            {
                PanelRecientes.Children.Add(new TextBlock 
                { 
                    Text = "Ninguno aún", 
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#555555")), 
                    FontSize = 10,
                    VerticalAlignment = VerticalAlignment.Center 
                });
                return;
            }

            foreach (var color in _coloresRecientes)
            {
                Border circulo = new Border
                {
                    Width = 20, Height = 20, CornerRadius = new CornerRadius(10),
                    Background = new SolidColorBrush(color), Margin = new Thickness(0, 0, 8, 0),
                    Cursor = Cursors.Hand, ToolTip = $"#{color.R:X2}{color.G:X2}{color.B:X2}"
                };
                circulo.MouseLeftButtonDown += Swatch_Click;
                PanelRecientes.Children.Add(circulo);
            }
        }

        private void TxtRGB_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_actualizando) return;

            if (byte.TryParse(TxtR.Text, out byte r) && 
                byte.TryParse(TxtG.Text, out byte g) && 
                byte.TryParse(TxtB.Text, out byte b))
            {
                Color c = Color.FromRgb(r, g, b);
                SincronizarDesdeColor(c, TxtR);
            }
        }

        private void TxtHex_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_actualizando) return;

            // Limpiamos espacios en blanco por si el usuario pega algo como " # FF 00 00 "
            string hexLimpio = TxtHex.Text.Replace(" ", "").Trim();
            
            if (!hexLimpio.StartsWith("#")) 
            {
                hexLimpio = "#" + hexLimpio;
            }
            
            try
            {
                // Intentamos convertir lo que sea que haya escrito/pegado
                var colorConvertido = ColorConverter.ConvertFromString(hexLimpio);
                
                if (colorConvertido is Color c)
                {
                    SincronizarDesdeColor(c, TxtHex); // Sincronizamos si es válido
                }
            }
            catch 
            { 
                // Si el usuario borró todo y solo quedó un "#", o si escribió "Hola",
                // simplemente ignoramos el error en silencio y le dejamos seguir escribiendo
                // SIN bloquear la caja de texto.
            }
        }

        // --- MATEMÁTICAS HSL -> RGB ---
        private Color HslToRgb(double h, double s, double l)
        {
            byte r, g, b;
            if (s == 0) { r = g = b = (byte)(l * 255); }
            else
            {
                double v2 = (l < 0.5) ? (l * (1 + s)) : ((l + s) - (l * s));
                double v1 = 2 * l - v2;
                double hue = h / 360.0;
                r = (byte)(255 * HueToRgb(v1, v2, hue + (1.0 / 3)));
                g = (byte)(255 * HueToRgb(v1, v2, hue));
                b = (byte)(255 * HueToRgb(v1, v2, hue - (1.0 / 3)));
            }
            return Color.FromRgb(r, g, b);
        }

        private double HueToRgb(double v1, double v2, double vH)
        {
            if (vH < 0) vH += 1;
            if (vH > 1) vH -= 1;
            if ((6 * vH) < 1) return (v1 + (v2 - v1) * 6 * vH);
            if ((2 * vH) < 1) return v2;
            if ((3 * vH) < 2) return (v1 + (v2 - v1) * ((2.0 / 3) - vH) * 6);
            return v1;
        }
    }
}