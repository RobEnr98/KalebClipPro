using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using System.Linq;
using System.Globalization;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Documents; 
using GongSolutions.Wpf.DragDrop;
using KalebClipPro.Models;
using KalebClipPro.Services;
using KalebClipPro.Infrastructure;

// --- LIBRERÍAS DE COLORCODE ---
using ColorCode;
using ColorCode.Wpf;
using ColorCode.Styling;
using ColorCode.Common;

namespace KalebClipPro 
{
    public partial class MainWindow 
    {
        private void ListaClips_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.RemovedItems.Count > 0 && e.RemovedItems[0] is ClipData clipViejo && _editorActual != null)
            {
                GuardarEstadoSeguro(clipViejo, _editorActual);
            }

            if (_editorActual != null) 
            {
                if (_editorActual.Parent is Grid parentGrid)
                    parentGrid.Visibility = Visibility.Collapsed;
                else
                    _editorActual.Visibility = Visibility.Collapsed;
                    
                _editorActual = null;
            }

            if (ListaClips.SelectedItem is not ClipData c)
            {
                return;
            }

            var conversor = new ClipCategoryConverter();
            string tipo = conversor.Convert(c.Contenido_Plano, typeof(string), null!, CultureInfo.CurrentCulture)?.ToString() ?? "Text";
            
            string iconoTab = tipo == "Code" ? "{ }" : (tipo == "Url" ? "🌐" : "[T]");
            string colorTab = tipo == "Code" ? "#4ADE80" : (tipo == "Url" ? "#8E97FD" : "#A0AAB5");

            if (TabEditor != null)
            {
                TabEditor.IsChecked = true;
                TabEditor.ToolTip = $"Tipo: {tipo} | Origen: {c.Origen_App}";
            }

            if (TabActivaTitulo != null && TabActivaIcono != null)
            {
                TabActivaIcono.Text = iconoTab;
                TabActivaIcono.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colorTab));
                
                TabActivaTitulo.Text = $"{c.Fecha_Creacion:dd MMM. HH:mm}";
                TabActivaTitulo.ToolTip = $"Origen: {c.Origen_App}\nTipo: {tipo}";
            }

            if (!_editoresMemoria.ContainsKey(c.Guid_Clip))
            {
                var versionesDb = db.ObtenerTodasLasVersiones(c.Guid_Clip);
                string contenidoCrudo = c.ObtenerContenidoCrudo() ?? "";

                int indiceDondeEstoy = -1;

                for (int i = versionesDb.Count - 1; i >= 0; i--)
                {
                    if (versionesDb[i] == contenidoCrudo)
                    {
                        indiceDondeEstoy = i;
                        break; 
                    }
                }

                if (indiceDondeEstoy == -1)
                {
                    versionesDb.Add(contenidoCrudo);
                    indiceDondeEstoy = versionesDb.Count - 1;
                }

                _controladoresHistoria[c.Guid_Clip] = new HistorialSesion 
                { 
                    Versiones = versionesDb, 
                    IndiceActual = indiceDondeEstoy 
                };

                Grid editorWrapper = new Grid();
                editorWrapper.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                editorWrapper.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                Canvas lineCanvas = new Canvas { Background = Brushes.Transparent };
                Border lineBorder = new Border 
                { 
                    Width = 45, 
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#252B32")), 
                    BorderThickness = new Thickness(0, 0, 1, 0),
                    BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1D2227")),
                    Visibility = _mostrarNumerosLinea ? Visibility.Visible : Visibility.Collapsed,
                    Child = lineCanvas,
                    ClipToBounds = true // Crucial para que los números desaparezcan al hacer scroll
                };
                Grid.SetColumn(lineBorder, 0);

                bool esCodigo = tipo == "Code";

                RichTextBox nuevoEditor = new RichTextBox
                {
                    Foreground = esCodigo ? new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8F8F2")) : (Brush)FindResource("TextMain"), 
                    Background = Brushes.Transparent, 
                    BorderThickness = new Thickness(0),
                    Padding = new Thickness(15, 25, 15, 15),
                    AcceptsReturn = true,
                    // 🌟 FIX 1: PERMITE LA SANGRÍA CON TABULADOR
                    AcceptsTab = true, 
                    FontFamily = esCodigo ? new FontFamily("Consolas") : _fuentePreferida,
                    FontSize = 13, 
                    CaretBrush = (Brush)FindResource("AccentColor"),
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    SelectionBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4fabff")), 
                    SelectionOpacity = 0.4,
                    IsUndoEnabled = false 
                };

                // 🌟 FIX 2: INICIALIZACIÓN DE PÁRRAFO ESTILO WORD (Espacio abajo nulo)
                System.Windows.Style style = new System.Windows.Style(typeof(Paragraph));
                style.Setters.Add(new Setter(Paragraph.MarginProperty, new Thickness(0, 0, 0, 0)));
                nuevoEditor.Resources.Add(typeof(Paragraph), style);

                string textoPlano = c.Contenido_Plano;
                string? xamlOculto = null;

                if (contenidoCrudo.Contains("_||KALEB_RTF||_"))
                {
                    var partes = contenidoCrudo.Split(new string[] { "_||KALEB_RTF||_" }, StringSplitOptions.None);
                    xamlOculto = partes.Length > 1 ? partes[1] : null;
                }
                else if (contenidoCrudo.TrimStart().StartsWith("<Section xmlns="))
                {
                    xamlOculto = contenidoCrudo; 
                }

                TextRange range = new TextRange(nuevoEditor.Document.ContentStart, nuevoEditor.Document.ContentEnd);
                if (!string.IsNullOrWhiteSpace(xamlOculto))
                {
                    try
                    {
                        using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(xamlOculto)))
                        {
                            range.Load(ms, DataFormats.Xaml);
                        }
                    }
                    catch { range.Text = textoPlano; }
                }
                else
                {
                    if (esCodigo)
                    {
                        //AplicarColorCode(nuevoEditor, textoPlano);
                        Helpers.RichTextFormatterHelper.AplicarColorCode(nuevoEditor, textoPlano);
                        GuardarEstadoSeguro(c, nuevoEditor); 
                    }
                    else
                    {
                        range.Text = textoPlano;
                    }
                }

                nuevoEditor.IsUndoEnabled = true;

                Grid.SetColumn(nuevoEditor, 1);
                ActualizarComportamientoScroll(nuevoEditor);
                
                // 🌟 FIX RENDIMIENTO EXTREMO: Anti-spam y conteo matemático
                bool _updateQueued = false; // Variable local para controlar el spam
                
                Action actualizarVisuales = () => 
                {
                    if (!_mostrarNumerosLinea || nuevoEditor == null) return;

                    // 1. SISTEMA ANTI-SPAM (Debounce)
                    // Si ya hay un dibujo en la fila de espera, ignoramos los siguientes scrolls
                    if (_updateQueued) return;
                    _updateQueued = true;

                    // Usamos ContextIdle para que solo se ejecute cuando tu app no esté haciendo otras cosas críticas
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        _updateQueued = false; // Liberamos para el siguiente ciclo

                        double viewportHeight = nuevoEditor.ActualHeight;
                        if (viewportHeight <= 0) return;

                        // Encontrar la primera línea visible
                        TextPointer topVisible = nuevoEditor.GetPositionFromPoint(new Point(5, 5), true) ?? nuevoEditor.Document.ContentStart;
                        topVisible = topVisible.GetLineStartPosition(0) ?? topVisible;

                        // 2. 🚀 SÚPER MODO TURBO (Lectura de texto puro en 1 milisegundo)
                        // Extraemos el texto oculto y contamos los '\n' matemáticamente. O(N) extremo.
                        TextRange rangeOculto = new TextRange(nuevoEditor.Document.ContentStart, topVisible);
                        string textoOculto = rangeOculto.Text;
                        int numeroLinea = 1;
                        
                        for (int i = 0; i < textoOculto.Length; i++) {
                            if (textoOculto[i] == '\n') numeroLinea++;
                        }

                        // 3. MODO DIBUJO LIGERO (Solo calculamos las ~30 líneas de la pantalla)
                        TextPointer? curr = topVisible;
                        int indexCanvas = 0;
                        TextPointer? caretStartLine = nuevoEditor.CaretPosition.GetLineStartPosition(0);
                        TextBlock? currentNumText = null;

                        while (curr != null && curr.CompareTo(nuevoEditor.Document.ContentEnd) < 0)
                        {
                            Rect rect = curr.GetCharacterRect(LogicalDirection.Forward);
                            if (rect != Rect.Empty)
                            {
                                // Freno de mano: No calcular nada debajo de la ventana
                                if (rect.Top > viewportHeight + 50) break; 

                                // Detector de Word-Wrap solo para las líneas visibles
                                bool esLogica = false;
                                if (curr.Paragraph != null && curr.CompareTo(curr.Paragraph.ContentStart) == 0) esLogica = true;
                                else {
                                    TextPointer pAnt = curr.GetNextInsertionPosition(LogicalDirection.Backward);
                                    if (pAnt == null || pAnt.Paragraph != curr.Paragraph) esLogica = true;
                                    else {
                                        string t = new TextRange(pAnt, curr).Text;
                                        if (t.Contains("\n") || t.Contains("\r")) esLogica = true;
                                    }
                                }

                                if (esLogica) 
                                {
                                    if (rect.Bottom >= -50)
                                    {
                                        if (indexCanvas < lineCanvas.Children.Count)
                                            currentNumText = (TextBlock)lineCanvas.Children[indexCanvas];
                                        else
                                        {
                                            currentNumText = new TextBlock
                                            {
                                                FontFamily = new FontFamily("Consolas"),
                                                TextAlignment = TextAlignment.Right,
                                                Width = 35,
                                                Padding = new Thickness(0, 0, 8, 0)
                                            };
                                            lineCanvas.Children.Add(currentNumText);
                                        }

                                        currentNumText.Text = numeroLinea.ToString();
                                        currentNumText.FontSize = nuevoEditor.FontSize;
                                        currentNumText.Foreground = (Brush)FindResource("TextMuted");
                                        currentNumText.FontWeight = FontWeights.Normal;

                                        Canvas.SetTop(currentNumText, rect.Top);
                                        indexCanvas++;
                                    }
                                    else
                                    {
                                        currentNumText = null;
                                    }
                                    
                                    // Aumentamos el contador ya que la imprimimos
                                    numeroLinea++; 
                                }

                                // Iluminación suave que sigue a la línea base
                                if (caretStartLine != null && curr.CompareTo(caretStartLine) == 0 && currentNumText != null && nuevoEditor.IsFocused)
                                {
                                    currentNumText.Foreground = (Brush)FindResource("AccentColor");
                                    currentNumText.FontWeight = FontWeights.Bold;
                                }
                            }

                            curr = curr.GetLineStartPosition(1, out int count);
                            if (count == 0) break;
                        }

                        // Limpiamos los TextBlock que quedaron huérfanos
                        while (lineCanvas.Children.Count > indexCanvas)
                        {
                            lineCanvas.Children.RemoveAt(lineCanvas.Children.Count - 1);
                        }

                    }), DispatcherPriority.ContextIdle); 
                };

                // Asignamos los eventos rediseñados
                nuevoEditor.Loaded += (s, args) => actualizarVisuales();
                nuevoEditor.SizeChanged += (s, args) => actualizarVisuales();
                
                nuevoEditor.TextChanged += (s, args) => 
                {
                    if (_viajandoEnElTiempo) return; 
                    actualizarVisuales();
                    _autoSaveTimer.Stop(); _autoSaveTimer.Start();
                };

                nuevoEditor.SelectionChanged += (s, args) => 
                {
                    if (_viajandoEnElTiempo) return;
                    if (nuevoEditor.Selection.GetPropertyValue(TextElement.FontFamilyProperty) is FontFamily fuenteReal)
                    {
                        if (CmbFuentes != null) 
                        {
                            var item = CmbFuentes.Items.Cast<FontFamily>().FirstOrDefault(f => f.Source == fuenteReal.Source);
                            if (item != null) CmbFuentes.SelectedItem = item;
                        }
                    }
                    actualizarVisuales();
                    _autoSaveTimer.Stop(); _autoSaveTimer.Start();
                };
                
                // 🌟 FIX 3: El scroll ahora solo manda a re-dibujar, sin enlazar barras nativas
                nuevoEditor.AddHandler(ScrollViewer.ScrollChangedEvent, new ScrollChangedEventHandler((s, args) => 
                {
                    actualizarVisuales();
                }));

                nuevoEditor.PreviewMouseWheel += (s, args) => 
                {
                    if (Keyboard.Modifiers == ModifierKeys.Control)
                    {
                        double nuevoSize = nuevoEditor.FontSize + (args.Delta > 0 ? 1 : -1);
                        if (nuevoSize >= 8 && nuevoSize <= 100)
                        {
                            nuevoEditor.FontSize = nuevoSize;
                            actualizarVisuales(); // El zoom también dispara el redibujado
                        }
                        args.Handled = true; 
                    }
                };

                nuevoEditor.Loaded += (s, args) => 
                {
                    actualizarVisuales();
                };

                nuevoEditor.TextChanged += (s, args) => 
                {
                    if (_viajandoEnElTiempo) return; 

                    actualizarVisuales();
                    _autoSaveTimer.Stop();
                    _autoSaveTimer.Start();
                };

                nuevoEditor.SelectionChanged += (s, args) => 
                {
                    if (_viajandoEnElTiempo) return;

                    if (nuevoEditor.Selection.GetPropertyValue(TextElement.FontFamilyProperty) is FontFamily fuenteReal)
                    {
                        if (CmbFuentes != null) 
                        {
                            var item = CmbFuentes.Items.Cast<FontFamily>().FirstOrDefault(f => f.Source == fuenteReal.Source);
                            if (item != null) CmbFuentes.SelectedItem = item;
                        }
                    }

                    Dispatcher.BeginInvoke(new Action(() => 
                    {
                        actualizarVisuales();
                    }), DispatcherPriority.ContextIdle);

                    _autoSaveTimer.Stop();
                    _autoSaveTimer.Start();
                };
                
                nuevoEditor.AddHandler(ScrollViewer.ScrollChangedEvent, new ScrollChangedEventHandler((s, args) => 
                {
                    actualizarVisuales();
                }));

                // 🌟 2. EL NUEVO EVENTO DE ZOOM CON CTRL + RUEDA (ya no busca lineText)
                nuevoEditor.PreviewMouseWheel += (s, args) => 
                {
                    if (Keyboard.Modifiers == ModifierKeys.Control)
                    {
                        double nuevoSize = nuevoEditor.FontSize + (args.Delta > 0 ? 1 : -1);
                        
                        if (nuevoSize >= 8 && nuevoSize <= 100)
                        {
                            nuevoEditor.FontSize = nuevoSize;
                            // Al hacer zoom, recalculamos dónde deben ir los números
                            actualizarVisuales(); 
                        }
                        
                        args.Handled = true; 
                    }
                };

                // 🌟 FIX: Lógica de teclas avanzada integrada
                nuevoEditor.PreviewKeyDown += (s, args) =>
                {
                    // --- Atajos de Forzar Color (CTRL+SHIFT+F) ---
                    if (Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift) && args.Key == Key.F)
                    {
                        args.Handled = true; 
                        string textoActual = new TextRange(nuevoEditor.Document.ContentStart, nuevoEditor.Document.ContentEnd).Text;
                        if (textoActual.EndsWith("\r\n")) textoActual = textoActual.Substring(0, textoActual.Length - 2);
                        Helpers.RichTextFormatterHelper.AplicarColorCode(nuevoEditor, textoActual); GuardarEstadoSeguro(c, nuevoEditor); return;
                    }

                    // --- Viaje en el tiempo (CTRL+Z / CTRL+Y) ---
                    if (Keyboard.Modifiers == ModifierKeys.Control)
                    {
                        if (args.Key == Key.Z) { args.Handled = true; RestaurarHistorialBD(c.Guid_Clip, nuevoEditor, -1); }
                        else if (args.Key == Key.Y) { args.Handled = true; RestaurarHistorialBD(c.Guid_Clip, nuevoEditor, 1); }
                    }

                    // --- Lógica Avanzada de Teclas para Listas ---
                    TextPointer caretPos = nuevoEditor.CaretPosition;
                    Paragraph? currentParagraph = caretPos.Paragraph;

                    if (currentParagraph != null && currentParagraph.Parent is ListItem currentListItem && currentListItem.Parent is List parentList)
                    {

                        // 🌟 NUEVO: CONTROL DE TABULACIÓN PARA LISTAS (SANGRÍA) - CORREGIDO
                        if (args.Key == Key.Tab)
                        {
                            TextPointer caret = nuevoEditor.CaretPosition;
                            Paragraph? p = caret.Paragraph;

                            if (p != null)
                            {
                                // Verificamos si estamos al principio del texto
                                TextRange textBefore = new TextRange(p.ContentStart, caret);
                                bool isAtStart = string.IsNullOrEmpty(textBefore.Text);

                                // Si hay texto seleccionado, o estamos al inicio, aplicamos nuestra sangría
                                if (!nuevoEditor.Selection.IsEmpty || isAtStart || p.Parent is ListItem)
                                {
                                    args.Handled = true; // Bloqueamos el salto feo de WPF
                                    
                                    bool isShiftPressed = Keyboard.Modifiers == ModifierKeys.Shift;
                                    
                                    // +1 aumenta (derecha), -1 reduce (izquierda)
                                    Helpers.RichTextFormatterHelper.AplicarSangriaPersonalizada(nuevoEditor, isShiftPressed ? -1 : 1);
                                    
                                    actualizarVisuales();
                                    return; 
                                }
                                // Si estás en medio de una palabra normal, el Tab insertará un espacio normal (\t)
                            }
                        }

                        
                        // 🌟 MEJORA 2: CONTROL DE ENTER EN LISTA (Salir al presionar Enter en línea vacía)
                        if (args.Key == Key.Enter)
                        {
                            TextRange r = new TextRange(currentListItem.ContentStart, currentListItem.ContentEnd);
                            // Si la línea actual está vacía y presionas Enter
                            if (string.IsNullOrWhiteSpace(r.Text))
                            {
                                args.Handled = true; // Detenemos el salto de línea normal
                                
                                nuevoEditor.BeginChange();
                                
                                // Creamos un párrafo fuera de la lista
                                Paragraph newP = new Paragraph();
                                parentList.SiblingBlocks.InsertAfter(parentList, newP);
                                
                                // Borramos el ítem vacío de la lista
                                parentList.ListItems.Remove(currentListItem);
                                
                                // Si la lista ya no tiene ítems, la borramos para no dejar basura invisible
                                if (parentList.ListItems.Count == 0) parentList.SiblingBlocks.Remove(parentList);
                                
                                // Movemos el cursor al nuevo párrafo
                                nuevoEditor.CaretPosition = newP.ContentStart;
                                
                                nuevoEditor.EndChange();
                            }
                        }
                    }
                };

                editorWrapper.Children.Add(lineBorder);
                editorWrapper.Children.Add(nuevoEditor);

                TextBoxContainer.Children.Add(editorWrapper);
                _editoresMemoria[c.Guid_Clip] = nuevoEditor;
                
                actualizarVisuales();
            }

            _editorActual = _editoresMemoria[c.Guid_Clip];

            if (_editorActual.Parent is Grid pGrid)
                pGrid.Visibility = Visibility.Visible;
            else
                _editorActual.Visibility = Visibility.Visible;

            ActualizarComportamientoScroll(_editorActual);
            
            _editorActual.CaretPosition = _editorActual.Document.ContentEnd;
            _editorActual.Focus(); 
            ListaClips.Focus();
        }

        

        private void CmbInterlineado_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_editorActual == null || CmbInterlineado.SelectedItem is not ComboBoxItem item) return;
            
            string valorInterlineado = item.Content?.ToString() ?? "1.0";
            double altura = double.NaN; 
            double margenAbajo = 0;

            switch (valorInterlineado)
            {
                case "1.0": altura = double.NaN; margenAbajo = 0; break;
                case "1.15": altura = 18; margenAbajo = 2; break;
                case "1.5": altura = 24; margenAbajo = 5; break;
                case "2.0": altura = 32; margenAbajo = 10; break;
            }

            // Llamamos al Helper
            Helpers.RichTextFormatterHelper.AplicarInterlineado(_editorActual, altura, margenAbajo);
            
            // Guardamos el estado desde aquí, que es quien conoce la BD
            if (ListaClips.SelectedItem is ClipData currentClip) GuardarEstadoSeguro(currentClip, _editorActual);
            
            _editorActual.Focus(); 
        }

        private void ActualizarComportamientoScroll(RichTextBox editor)
        {
            if (_scrollHorizontalHabilitado)
            {
                editor.Document.PageWidth = 2500;
                editor.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
            }
            else
            {
                editor.Document.PageWidth = double.NaN;
                editor.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled;
            }
        }

        private void ToggleLineNumbers_Changed(object sender, RoutedEventArgs e)
        {
            _mostrarNumerosLinea = LineNumbersCheck?.IsChecked ?? false;

            if (TextBoxContainer == null) return;

            foreach (Grid wrapper in TextBoxContainer.Children.OfType<Grid>())
            {
                // Buscamos el borde de la izquierda y lo ocultamos entero
                var lineBorder = wrapper.Children.OfType<Border>().FirstOrDefault(b => b.Child is Canvas);
                if (lineBorder != null)
                {
                    lineBorder.Visibility = _mostrarNumerosLinea ? Visibility.Visible : Visibility.Collapsed;
                }
            }
        }

        private void ToggleScroll_Changed(object sender, RoutedEventArgs e)
        {
            _scrollHorizontalHabilitado = ScrollHorizontalCheck.IsChecked ?? false;

            if (_editorActual != null)
            {
                ActualizarComportamientoScroll(_editorActual);
            }
        }

        private void BtnCopiar_Click(object sender, RoutedEventArgs e)
        {
            string textoACopiar = "";

            if (TabEditor.IsChecked == true && _editorActual != null)
            {
                textoACopiar = !_editorActual.Selection.IsEmpty 
                    ? _editorActual.Selection.Text 
                    : new TextRange(_editorActual.Document.ContentStart, _editorActual.Document.ContentEnd).Text;
            }
            else if (TabRecolector.IsChecked == true && ClipsRecolector.Count > 0)
            {
                textoACopiar = string.Join(Environment.NewLine + Environment.NewLine, ClipsRecolector.Select(c => c.Contenido_Plano));
            }

            if (string.IsNullOrWhiteSpace(textoACopiar)) return;

            bool exito = false;
            for (int i = 0; i < 5; i++)
            {
                try { Clipboard.SetDataObject(textoACopiar, true); exito = true; break; } 
                catch { System.Threading.Thread.Sleep(15); }
            }

            if (exito)
            {
                BtnCopiar.Foreground = (Brush)FindResource("AccentColor");
                var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
                timer.Tick += (s, args) => { BtnCopiar.ClearValue(Button.ForegroundProperty); timer.Stop(); };
                timer.Start();
            }
        }

        private void GuardarEstadoSeguro(ClipData clipParaGuardar, RichTextBox editorParaGuardar)
        {
            if (clipParaGuardar == null || editorParaGuardar == null) return;

            _autoSaveTimer.Stop();

            string xamlParaGuardar;
            TextRange range = new TextRange(editorParaGuardar.Document.ContentStart, editorParaGuardar.Document.ContentEnd);
            
            using (MemoryStream ms = new MemoryStream())
            {
                range.Save(ms, DataFormats.Xaml);
                xamlParaGuardar = Encoding.UTF8.GetString(ms.ToArray());
            }

            string textoPlanoVista = range.Text.TrimEnd('\r', '\n'); 
            string paqueteBD = textoPlanoVista + "_||KALEB_RTF||_" + xamlParaGuardar;

            if (_controladoresHistoria.TryGetValue(clipParaGuardar.Guid_Clip, out var historia))
            {
                string versionDondeEstoyParado = historia.Versiones[historia.IndiceActual];
                
                if (paqueteBD != versionDondeEstoyParado)
                {
                    db.ActualizarClip(clipParaGuardar.Guid_Clip, paqueteBD);
                    clipParaGuardar.Contenido_Plano = paqueteBD;
                    db.GuardarVersion(clipParaGuardar.Guid_Clip, paqueteBD);

                    if (historia.IndiceActual < historia.Versiones.Count - 1)
                    {
                        historia.Versiones.RemoveRange(historia.IndiceActual + 1, 
                                                       historia.Versiones.Count - (historia.IndiceActual + 1));
                    }
                    
                    historia.Versiones.Add(paqueteBD);
                    historia.IndiceActual = historia.Versiones.Count - 1;
                }
            }
        }

        private void AutoSaveTimer_Tick(object? sender, EventArgs e)
        {
            if (ListaClips.SelectedItem is ClipData currentClip && _editorActual != null)
            {
                GuardarEstadoSeguro(currentClip, _editorActual);
            }
        }

        private bool _viajandoEnElTiempo = false;

        private void RestaurarHistorialBD(string guid, RichTextBox editor, int direccion)
        {
            if (_controladoresHistoria.TryGetValue(guid, out var historia))
            {
                int nuevoIndice = historia.IndiceActual + direccion;
                
                if (nuevoIndice >= 0 && nuevoIndice < historia.Versiones.Count)
                {
                    _viajandoEnElTiempo = true; 
                    _autoSaveTimer.Stop(); 
                    
                    historia.IndiceActual = nuevoIndice;
                    string versionCruda = historia.Versiones[nuevoIndice];
                    
                    string textoPlano = versionCruda;
                    string? xamlOculto = null; // 🌟 FIX Nulos

                    if (versionCruda.Contains("_||KALEB_RTF||_"))
                    {
                        var partes = versionCruda.Split(new string[] { "_||KALEB_RTF||_" }, StringSplitOptions.None);
                        textoPlano = partes[0];
                        xamlOculto = partes.Length > 1 ? partes[1] : null;
                    }
                    else if (versionCruda.TrimStart().StartsWith("<Section xmlns="))
                    {
                        xamlOculto = versionCruda;
                    }

                    TextRange range = new TextRange(editor.Document.ContentStart, editor.Document.ContentEnd);
                    if (!string.IsNullOrWhiteSpace(xamlOculto))
                    {
                        try { using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(xamlOculto))) { range.Load(ms, DataFormats.Xaml); } }
                        catch { range.Text = textoPlano; }
                    }
                    else { range.Text = textoPlano; }
                    
                    if (ListaClips.SelectedItem is ClipData currentClip && currentClip.Guid_Clip == guid)
                    {
                        currentClip.Contenido_Plano = versionCruda;
                        db.ActualizarClip(guid, versionCruda); 
                    }

                    _viajandoEnElTiempo = false;
                }
            }
        }
    }
}