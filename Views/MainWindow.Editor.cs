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
        // =========================================================================================
        // 1. EL DIRECTOR DE ORQUESTA (El evento principal ahora es súper fácil de leer)
        // =========================================================================================
        private void ListaClips_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 1. Guardar y ocultar el viejo
            if (e.RemovedItems.Count > 0 && e.RemovedItems[0] is ClipData clipViejo)
            {
                GuardarEstadoSeguro(clipViejo, _editorActual!);
                OcultarEditorActual();
            }

            // 2. Verificar si hay algo nuevo seleccionado
            if (ListaClips.SelectedItem is not ClipData clipNuevo) return;

            // 3. Determinar el tipo de contenido (Código, Texto, URL, etc.)
            var conversor = new ClipCategoryConverter();
            string tipo = conversor.Convert(clipNuevo.Contenido_Plano, typeof(string), null!, CultureInfo.CurrentCulture)?.ToString() ?? "Text";

            // 4. Actualizar la interfaz (Pestañas)
            ActualizarUI_Pestanas(clipNuevo, tipo);

            // 5. Crear el editor si es la primera vez que abrimos este clip
            if (!_editoresMemoria.ContainsKey(clipNuevo.Guid_Clip))
            {
                ConstruirYAlmacenarNuevoEditor(clipNuevo, tipo);
            }

            // 6. Mostrar el editor activo
            MostrarEditorActivo(clipNuevo);
        }

        // =========================================================================================
        // 2. LOS DEPARTAMENTOS (Métodos extraídos para mantener el orden)
        // =========================================================================================

        private void OcultarEditorActual()
        {
            if (_editorActual == null) return;

            if (_editorActual.Parent is Grid parentGrid)
                parentGrid.Visibility = Visibility.Collapsed;
            else
                _editorActual.Visibility = Visibility.Collapsed;
                
            _editorActual = null;
        }

        private void ActualizarUI_Pestanas(ClipData c, string tipo)
        {
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
        }

        private void MostrarEditorActivo(ClipData c)
        {
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

        // =========================================================================================
        // 3. LA FÁBRICA DE EDITORES (Aquí se construye todo por primera vez)
        // =========================================================================================
        private void ConstruirYAlmacenarNuevoEditor(ClipData c, string tipo)
        {
            // 1. Historial
            var versionesDb = db.ObtenerTodasLasVersiones(c.Guid_Clip);
            string contenidoCrudo = c.ObtenerContenidoCrudo() ?? "";
            int indiceDondeEstoy = versionesDb.LastIndexOf(contenidoCrudo);
            
            if (indiceDondeEstoy == -1)
            {
                versionesDb.Add(contenidoCrudo);
                indiceDondeEstoy = versionesDb.Count - 1;
            }

            _controladoresHistoria[c.Guid_Clip] = new HistorialSesion { Versiones = versionesDb, IndiceActual = indiceDondeEstoy };

            // 2. Construcción Visual
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
                ClipToBounds = true 
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
                AcceptsTab = true, 
                FontFamily = esCodigo ? new FontFamily("Consolas") : _fuentePreferida,
                FontSize = 13, 
                CaretBrush = (Brush)FindResource("AccentColor"),
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                SelectionBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4fabff")), 
                SelectionOpacity = 0.4,
                IsUndoEnabled = false 
            };

            System.Windows.Style style = new System.Windows.Style(typeof(Paragraph));
            style.Setters.Add(new Setter(Paragraph.MarginProperty, new Thickness(0, 0, 0, 0)));
            nuevoEditor.Resources.Add(typeof(Paragraph), style);

            // 3. Cargar Contenido
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
                try { using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(xamlOculto))) { range.Load(ms, DataFormats.Xaml); } }
                catch { range.Text = textoPlano; }
            }
            else
            {
                if (esCodigo)
                {
                    Helpers.RichTextFormatterHelper.AplicarColorCode(nuevoEditor, textoPlano);
                    GuardarEstadoSeguro(c, nuevoEditor); 
                }
                else { range.Text = textoPlano; }
            }

            nuevoEditor.IsUndoEnabled = true;
            Grid.SetColumn(nuevoEditor, 1);
            ActualizarComportamientoScroll(nuevoEditor);
            
            // 4. Asignar Eventos Mágicos
            AsignarEventosAEditorNuevo(nuevoEditor, c, lineCanvas);

            editorWrapper.Children.Add(lineBorder);
            editorWrapper.Children.Add(nuevoEditor);
            TextBoxContainer.Children.Add(editorWrapper);
            
            _editoresMemoria[c.Guid_Clip] = nuevoEditor;
        }

        private void AsignarEventosAEditorNuevo(RichTextBox nuevoEditor, ClipData c, Canvas lineCanvas)
        {
            bool _updateQueued = false; 
            
            Action actualizarVisuales = () => 
            {
                if (!_mostrarNumerosLinea || nuevoEditor == null) return;
                if (_updateQueued) return;
                _updateQueued = true;

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _updateQueued = false; 

                    double viewportHeight = nuevoEditor.ActualHeight;
                    if (viewportHeight <= 0) return;

                    TextPointer topVisible = nuevoEditor.GetPositionFromPoint(new Point(5, 5), true) ?? nuevoEditor.Document.ContentStart;
                    topVisible = topVisible.GetLineStartPosition(0) ?? topVisible;

                    TextRange rangeOculto = new TextRange(nuevoEditor.Document.ContentStart, topVisible);
                    string textoOculto = rangeOculto.Text;
                    int numeroLinea = 1;
                    
                    for (int i = 0; i < textoOculto.Length; i++) {
                        if (textoOculto[i] == '\n') numeroLinea++;
                    }

                    TextPointer? curr = topVisible;
                    int indexCanvas = 0;
                    TextPointer? caretStartLine = nuevoEditor.CaretPosition.GetLineStartPosition(0);
                    TextBlock? currentNumText = null;

                    while (curr != null && curr.CompareTo(nuevoEditor.Document.ContentEnd) < 0)
                    {
                        Rect rect = curr.GetCharacterRect(LogicalDirection.Forward);
                        if (rect != Rect.Empty)
                        {
                            if (rect.Top > viewportHeight + 50) break; 

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
                                        currentNumText = new TextBlock { FontFamily = new FontFamily("Consolas"), TextAlignment = TextAlignment.Right, Width = 35, Padding = new Thickness(0, 0, 8, 0) };
                                        lineCanvas.Children.Add(currentNumText);
                                    }

                                    currentNumText.Text = numeroLinea.ToString();
                                    currentNumText.FontSize = nuevoEditor.FontSize;
                                    currentNumText.Foreground = (Brush)FindResource("TextMuted");
                                    currentNumText.FontWeight = FontWeights.Normal;
                                    Canvas.SetTop(currentNumText, rect.Top);
                                    indexCanvas++;
                                }
                                else { currentNumText = null; }
                                
                                numeroLinea++; 
                            }

                            if (caretStartLine != null && curr.CompareTo(caretStartLine) == 0 && currentNumText != null && nuevoEditor.IsFocused)
                            {
                                currentNumText.Foreground = (Brush)FindResource("AccentColor");
                                currentNumText.FontWeight = FontWeights.Bold;
                            }
                        }

                        curr = curr.GetLineStartPosition(1, out int count);
                        if (count == 0) break;
                    }

                    while (lineCanvas.Children.Count > indexCanvas)
                    {
                        lineCanvas.Children.RemoveAt(lineCanvas.Children.Count - 1);
                    }

                }), DispatcherPriority.ContextIdle); 
            };

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
                
                Dispatcher.BeginInvoke(new Action(() => actualizarVisuales()), DispatcherPriority.ContextIdle);
                _autoSaveTimer.Stop(); _autoSaveTimer.Start();
            };
            
            nuevoEditor.AddHandler(ScrollViewer.ScrollChangedEvent, new ScrollChangedEventHandler((s, args) => actualizarVisuales()));

            nuevoEditor.PreviewMouseWheel += (s, args) => 
            {
                if (Keyboard.Modifiers == ModifierKeys.Control)
                {
                    double nuevoSize = nuevoEditor.FontSize + (args.Delta > 0 ? 1 : -1);
                    if (nuevoSize >= 8 && nuevoSize <= 100) { nuevoEditor.FontSize = nuevoSize; actualizarVisuales(); }
                    args.Handled = true; 
                }
            };

            nuevoEditor.PreviewKeyDown += (s, args) =>
            {
                if (Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift) && args.Key == Key.F)
                {
                    args.Handled = true; 
                    string textoActual = new TextRange(nuevoEditor.Document.ContentStart, nuevoEditor.Document.ContentEnd).Text;
                    if (textoActual.EndsWith("\r\n")) textoActual = textoActual.Substring(0, textoActual.Length - 2);
                    Helpers.RichTextFormatterHelper.AplicarColorCode(nuevoEditor, textoActual); GuardarEstadoSeguro(c, nuevoEditor); return;
                }

                if (Keyboard.Modifiers == ModifierKeys.Control)
                {
                    if (args.Key == Key.Z) { args.Handled = true; RestaurarHistorialBD(c.Guid_Clip, nuevoEditor, -1); }
                    else if (args.Key == Key.Y) { args.Handled = true; RestaurarHistorialBD(c.Guid_Clip, nuevoEditor, 1); }
                }

                TextPointer caretPos = nuevoEditor.CaretPosition;
                Paragraph? currentParagraph = caretPos.Paragraph;

                if (currentParagraph != null && currentParagraph.Parent is ListItem currentListItem && currentListItem.Parent is List parentList)
                {
                    if (args.Key == Key.Tab)
                    {
                        TextPointer caret = nuevoEditor.CaretPosition;
                        Paragraph? p = caret.Paragraph;

                        if (p != null)
                        {
                            TextRange textBefore = new TextRange(p.ContentStart, caret);
                            bool isAtStart = string.IsNullOrEmpty(textBefore.Text);

                            if (!nuevoEditor.Selection.IsEmpty || isAtStart || p.Parent is ListItem)
                            {
                                args.Handled = true; 
                                bool isShiftPressed = Keyboard.Modifiers == ModifierKeys.Shift;
                                Helpers.RichTextFormatterHelper.AplicarSangriaPersonalizada(nuevoEditor, isShiftPressed ? -1 : 1);
                                actualizarVisuales();
                                return; 
                            }
                        }
                    }

                    if (args.Key == Key.Enter)
                    {
                        TextRange r = new TextRange(currentListItem.ContentStart, currentListItem.ContentEnd);
                        if (string.IsNullOrWhiteSpace(r.Text))
                        {
                            args.Handled = true; 
                            nuevoEditor.BeginChange();
                            Paragraph newP = new Paragraph();
                            parentList.SiblingBlocks.InsertAfter(parentList, newP);
                            parentList.ListItems.Remove(currentListItem);
                            if (parentList.ListItems.Count == 0) parentList.SiblingBlocks.Remove(parentList);
                            nuevoEditor.CaretPosition = newP.ContentStart;
                            nuevoEditor.EndChange();
                        }
                    }
                }
            };
            
            // Forzamos un dibujado inicial
            actualizarVisuales();
        }

        // =========================================================================================
        // 4. MÉTODOS EXISTENTES (Que ya estaban limpios)
        // =========================================================================================

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

            Helpers.RichTextFormatterHelper.AplicarInterlineado(_editorActual, altura, margenAbajo);
            
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
                    string? xamlOculto = null; 

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