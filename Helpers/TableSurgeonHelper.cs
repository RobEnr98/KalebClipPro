using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

namespace KalebClipPro.Helpers
{
    public static class TableSurgeonHelper
    {
        // =========================================================================================
        // 1. EVENTOS DE TECLADO Y RATÓN (El puente con MainWindow)
        // =========================================================================================

        public static void SmartEditor_PreviewKeyDown(RichTextBox editorActual, object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Tab && editorActual != null)
            {
                var cursor = editorActual.CaretPosition;
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
                            InsertarFila(editorActual, true);
                        }
                    }
                }
            }
        }

        public static void MostrarMenuContextual(RichTextBox editorActual, MouseButtonEventArgs e, Window ownerWindow)
        {
            if (editorActual == null) return;

            Point posMouse = e.GetPosition(editorActual);
            TextPointer posicionClic = editorActual.GetPositionFromPoint(posMouse, true);

            if (posicionClic != null && posicionClic.Paragraph != null && posicionClic.Paragraph.Parent is TableCell celdaClickeada)
            {
                TableRow filaActual = (TableRow)celdaClickeada.Parent;
                TableRowGroup grupoFilas = (TableRowGroup)filaActual.Parent;
                Table tablaPadre = (Table)grupoFilas.Parent;

                if (editorActual.Selection.IsEmpty)
                {
                    editorActual.Focus();
                    editorActual.CaretPosition = posicionClic;
                }

                ContextMenu menuEspecial = new ContextMenu();
                if (ownerWindow.FindResource("MenuOscuroStyle") is Style estiloMenu) menuEspecial.Style = estiloMenu;

                var colorTexto = new SolidColorBrush(Color.FromRgb(232, 236, 239)); 
                var colorFondoSub = new SolidColorBrush(Color.FromRgb(37, 43, 50)); 
                var colorPeligro = new SolidColorBrush(Color.FromRgb(255, 100, 100));

                // --- 1. SUBMENÚ DE BORDES ---
                MenuItem itemBordes = new MenuItem { Header = "📏 Bordes", Foreground = colorTexto };
                if (ownerWindow.FindResource("MenuItemSubmenuOscuro") is Style subMenuOscuro) itemBordes.Style = subMenuOscuro;
                
                MenuItem optSinBorde = new MenuItem { Header = " Sin borde", Background = colorFondoSub, Foreground = colorTexto, Icon = new TextBlock { Text = "🔲", Foreground = Brushes.White } };
                optSinBorde.Click += (s, ev) => AplicarBordeASEleccion(editorActual, celdaClickeada, "Ninguno");
                
                MenuItem optTodosBordes = new MenuItem { Header = " Todos los bordes", Background = colorFondoSub, Foreground = colorTexto, Icon = new TextBlock { Text = "▦", Foreground = Brushes.White } };
                optTodosBordes.Click += (s, ev) => AplicarBordeASEleccion(editorActual, celdaClickeada, "Todos");

                MenuItem optBordesExternos = new MenuItem { Header = " Bordes externos", Background = colorFondoSub, Foreground = colorTexto, Icon = new TextBlock { Text = "⬜", Foreground = Brushes.White } };
                optBordesExternos.Click += (s, ev) => AplicarBordeASEleccion(editorActual, celdaClickeada, "Externos");

                MenuItem optBordesInternos = new MenuItem { Header = " Bordes internos", Background = colorFondoSub, Foreground = colorTexto, Icon = new TextBlock { Text = "➕", Foreground = Brushes.White } };
                optBordesInternos.Click += (s, ev) => AplicarBordeASEleccion(editorActual, celdaClickeada, "Internos");

                MenuItem optBordeSuperior = new MenuItem { Header = " Borde superior", Background = colorFondoSub, Foreground = colorTexto, Icon = new TextBlock { Text = "▔", Foreground = Brushes.White } };
                optBordeSuperior.Click += (s, ev) => AplicarBordeASEleccion(editorActual, celdaClickeada, "Superior");

                MenuItem optBordeInferior = new MenuItem { Header = " Borde inferior", Background = colorFondoSub, Foreground = colorTexto, Icon = new TextBlock { Text = " ", Foreground = Brushes.White } };
                optBordeInferior.Click += (s, ev) => AplicarBordeASEleccion(editorActual, celdaClickeada, "Inferior");

                MenuItem optBordeIzquierdo = new MenuItem { Header = " Borde izquierdo", Background = colorFondoSub, Foreground = colorTexto, Icon = new TextBlock { Text = "▏", Foreground = Brushes.White } };
                optBordeIzquierdo.Click += (s, ev) => AplicarBordeASEleccion(editorActual, celdaClickeada, "Izquierdo");

                MenuItem optBordeDerecho = new MenuItem { Header = " Borde derecho", Background = colorFondoSub, Foreground = colorTexto, Icon = new TextBlock { Text = "▕", Foreground = Brushes.White } };
                optBordeDerecho.Click += (s, ev) => AplicarBordeASEleccion(editorActual, celdaClickeada, "Derecho");

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

                // --- 2. SUBMENÚ DE GROSOR ---
                MenuItem itemGrosor = new MenuItem { Header = "➖ Grosor de línea", Foreground = colorTexto };
                if (ownerWindow.FindResource("MenuItemSubmenuOscuro") is Style subMenuGrosorOscuro) itemGrosor.Style = subMenuGrosorOscuro;

                MenuItem optGrosorFino = new MenuItem { Header = " Fino (0.5 px)", Background = colorFondoSub, Foreground = colorTexto };
                optGrosorFino.Click += (s, ev) => AplicarGrosorASeleccion(editorActual, celdaClickeada, 0.5);

                MenuItem optGrosorNormal = new MenuItem { Header = " Normal (1.0 px)", Background = colorFondoSub, Foreground = colorTexto };
                optGrosorNormal.Click += (s, ev) => AplicarGrosorASeleccion(editorActual, celdaClickeada, 1.0);

                MenuItem optGrosorGrueso = new MenuItem { Header = " Grueso (2.0 px)", Background = colorFondoSub, Foreground = colorTexto };
                optGrosorGrueso.Click += (s, ev) => AplicarGrosorASeleccion(editorActual, celdaClickeada, 2.0);

                MenuItem optGrosorExtra = new MenuItem { Header = " Extra Grueso (3.0 px)", Background = colorFondoSub, Foreground = colorTexto };
                optGrosorExtra.Click += (s, ev) => AplicarGrosorASeleccion(editorActual, celdaClickeada, 3.0);

                itemGrosor.Items.Add(optGrosorFino);
                itemGrosor.Items.Add(optGrosorNormal);
                itemGrosor.Items.Add(optGrosorGrueso);
                itemGrosor.Items.Add(optGrosorExtra);

                // --- 3. RESTO DEL MENÚ ---
                MenuItem itemColorCelda = new MenuItem { Header = "🎨 Pintar fondo de celda", Foreground = colorTexto };
                itemColorCelda.Click += (s, ev) => PintarFondoSeleccion(editorActual, celdaClickeada, tablaPadre, ownerWindow);

                MenuItem itemQuitarColor = new MenuItem { Header = "🧽 Quitar color de fondo", Foreground = colorTexto };
                itemQuitarColor.Click += (s, ev) => QuitarFondoSeleccion(editorActual, celdaClickeada, tablaPadre);

                MenuItem itemColorBorde = new MenuItem { Header = "🖌️ Cambiar color de borde", Foreground = colorTexto };
                itemColorBorde.Click += (s, ev) => PintarBordeSeleccion(editorActual, celdaClickeada, tablaPadre, ownerWindow);

                MenuItem itemTemaTabla = new MenuItem { Header = "🌈 Cambiar paleta de la tabla", Foreground = colorTexto };
                itemTemaTabla.Click += (s, ev) => CambiarTemaTabla(editorActual, celdaClickeada, tablaPadre, ownerWindow);

                MenuItem itemFilaArriba = new MenuItem { Header = "⬆️ Insertar fila arriba", Foreground = colorTexto };
                itemFilaArriba.Click += (s, ev) => InsertarFila(editorActual, false);

                MenuItem itemFilaAbajo = new MenuItem { Header = "⬇️ Insertar fila debajo", Foreground = colorTexto };
                itemFilaAbajo.Click += (s, ev) => InsertarFila(editorActual, true);

                MenuItem itemColIzq = new MenuItem { Header = "⬅️ Insertar columna a la izquierda", Foreground = colorTexto };
                itemColIzq.Click += (s, ev) => InsertarColumna(editorActual, false);

                MenuItem itemColDer = new MenuItem { Header = "➡️ Insertar columna a la derecha", Foreground = colorTexto };
                itemColDer.Click += (s, ev) => InsertarColumna(editorActual, true);

                MenuItem itemEliminarFila = new MenuItem { Header = "❌ Eliminar esta fila", Foreground = colorPeligro };
                itemEliminarFila.Click += (s, ev) => EliminarFilaSeleccionada(editorActual, celdaClickeada);

                MenuItem itemEliminarColumna = new MenuItem { Header = "❌ Eliminar esta columna", Foreground = colorPeligro };
                itemEliminarColumna.Click += (s, ev) => EliminarColumnaSeleccionada(editorActual, celdaClickeada);

                menuEspecial.Items.Add(itemBordes); 
                menuEspecial.Items.Add(itemGrosor); 
                menuEspecial.Items.Add(new Separator { Background = new SolidColorBrush(Color.FromRgb(58, 68, 77)) });
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

        // =========================================================================================
        // 2. CONSTRUCCIÓN PRINCIPAL DE TABLAS
        // =========================================================================================

        public static void ConstruirTabla(RichTextBox editorActual, KalebClipPro.Views.InsertarTablaPopup dialogoTabla)
        {
            if (editorActual == null || dialogoTabla == null) return; 

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

            var colorHeader = new SolidColorBrush(acento);
            var colorCebra1 = new SolidColorBrush(Color.FromArgb(60, acento.R, acento.G, acento.B));
            var colorCebra2 = new SolidColorBrush(Color.FromArgb(20, acento.R, acento.G, acento.B));
            var colorBordeTematico = new SolidColorBrush(Color.FromArgb(120, acento.R, acento.G, acento.B));
            var colorBordeGris = new SolidColorBrush(Color.FromRgb(58, 68, 77));

            Brush brushBordeAplicar = (estiloBase == 0) ? colorBordeGris : colorBordeTematico;
            if (usaBordePers) brushBordeAplicar = new SolidColorBrush(colorBordePers);

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

                    if (estiloBase == 1) celda.Background = (f % 2 == 0) ? colorCebra1 : colorCebra2;
                    else if (estiloBase == 2) celda.Background = (c % 2 == 0) ? colorCebra1 : colorCebra2;
                    else if (estiloBase == 3) celda.Background = ((f + c) % 2 == 0) ? colorCebra1 : colorCebra2;
                    else celda.Background = Brushes.Transparent; 

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

                    Thickness grosorFinal = new Thickness(0);
                    if (tipoBorde == "Todos") grosorFinal = new Thickness(grosorElegido);
                    else if (tipoBorde == "Horizontales") grosorFinal = new Thickness(0, 0, 0, grosorElegido);
                    else if (tipoBorde == "Verticales") grosorFinal = new Thickness(0, 0, grosorElegido, 0);
                    
                    celda.BorderThickness = grosorFinal;
                    filaVisual.Cells.Add(celda);
                }
            }

            editorActual.BeginChange();
            Block bloqueActual = editorActual.CaretPosition.Paragraph;
            if (bloqueActual != null && bloqueActual.SiblingBlocks != null) bloqueActual.SiblingBlocks.InsertAfter(bloqueActual, tablaWpf);
            else editorActual.Document.Blocks.Add(tablaWpf);

            var parrafoFinal = new Paragraph();
            if (tablaWpf.SiblingBlocks != null) tablaWpf.SiblingBlocks.InsertAfter(tablaWpf, parrafoFinal);

            editorActual.EndChange();
            editorActual.Focus();
            
            if (grupoFilas.Rows.Count > 0 && grupoFilas.Rows[0].Cells.Count > 0)
            {
                editorActual.CaretPosition = ((Paragraph)grupoFilas.Rows[0].Cells[0].Blocks.FirstBlock).ContentStart;
            }
        }

        // =========================================================================================
        // 3. OPERACIONES DE FILAS Y COLUMNAS
        // =========================================================================================

        public static void InsertarFila(RichTextBox editorActual, bool abajo)
        {
            if (editorActual == null) return;
            var cursor = editorActual.CaretPosition;
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

                    if (!memoriaViva && grupoFilas.Rows.Count > 0 && grupoFilas.Rows[0].Cells.Count > 0)
                    {
                        TableCell celda0 = grupoFilas.Rows[0].Cells[0];
                        int fCount = grupoFilas.Rows.Count, cCount = grupoFilas.Rows[0].Cells.Count;
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
                    editorActual.Focus();
                    if (nuevaFila.Cells.Count > 0 && nuevaFila.Cells[0].Blocks.FirstBlock is Paragraph primerBloque) editorActual.CaretPosition = primerBloque.ContentStart;
                }
            }
        }

        public static void InsertarColumna(RichTextBox editorActual, bool derecha)
        {
            if (editorActual == null) return;
            var cursor = editorActual.CaretPosition;
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
                        int fCount = grupoFilas.Rows.Count, cCount = grupoFilas.Rows[0].Cells.Count;
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
                    editorActual.Focus();
                    if (filaActual.Cells.Count > indiceInsercion && filaActual.Cells[indiceInsercion].Blocks.FirstBlock is Paragraph bloqueCelda) editorActual.CaretPosition = bloqueCelda.ContentStart;
                }
            }
        }

        public static void EliminarFilaSeleccionada(RichTextBox editorActual, TableCell celda)
        {
            if (editorActual == null) return;
            
            if (celda.Parent is TableRow filaActual && filaActual.Parent is TableRowGroup grupoFilas)
            {
                editorActual.BeginChange();
                if (grupoFilas.Rows.Count > 1) 
                { 
                    grupoFilas.Rows.Remove(filaActual); 
                }
                else if (grupoFilas.Parent is Table tablaDestinada) 
                { 
                    tablaDestinada.SiblingBlocks?.Remove(tablaDestinada); 
                }
                editorActual.EndChange();
            }
        }

        public static void EliminarColumnaSeleccionada(RichTextBox editorActual, TableCell celda)
        {
            if (editorActual == null) return;

            if (celda.Parent is TableRow filaActual && filaActual.Parent is TableRowGroup grupoFilas && grupoFilas.Parent is Table tabla)
            {
                editorActual.BeginChange();
                int colIndex = filaActual.Cells.IndexOf(celda);

                if (filaActual.Cells.Count <= 1)
                {
                    tabla.SiblingBlocks?.Remove(tabla);
                    editorActual.EndChange();
                    return;
                }

                if (tabla.Columns.Count > colIndex) { tabla.Columns.RemoveAt(colIndex); }

                foreach (TableRow fila in grupoFilas.Rows)
                {
                    if (fila.Cells.Count > colIndex)
                    {
                        fila.Cells.RemoveAt(colIndex);
                    }
                }
                editorActual.EndChange();
            }
        }

        // =========================================================================================
        // 4. FUNCIONES DE COLOR Y FORMATO
        // =========================================================================================

        public static void PintarFondoSeleccion(RichTextBox editorActual, TableCell celdaClickeada, Table tablaPadre, Window ownerWindow)
        {
            if (editorActual == null) return;
            TableRowGroup grupoFilas = (TableRowGroup)((TableRow)celdaClickeada.Parent).Parent;

            var dialog = new KalebClipPro.Views.PaletaColores();
            dialog.Owner = ownerWindow;
            dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            
            dialog.AlSeleccionarColor = (colorNuevo) =>
            {
                var selectionRange = editorActual.Selection;
                var brushNuevo = new SolidColorBrush(colorNuevo);

                editorActual.BeginChange();
                
                if (selectionRange.IsEmpty)
                {
                    celdaClickeada.Background = brushNuevo;
                }
                else
                {
                    TableCell? celdaInicio = null, celdaFin = null;
                    TextPointer pInicio = selectionRange.Start;
                    TextPointer pFin = selectionRange.End.GetNextInsertionPosition(LogicalDirection.Backward) ?? selectionRange.End;
                    
                    if (pInicio.CompareTo(pFin) > 0) pFin = selectionRange.End;

                    foreach (var r in grupoFilas.Rows) {
                        foreach (var c in r.Cells) {
                            if (pInicio.CompareTo(c.ElementStart) >= 0 && pInicio.CompareTo(c.ElementEnd) < 0) celdaInicio = c;
                            if (pFin.CompareTo(c.ElementStart) >= 0 && pFin.CompareTo(c.ElementEnd) < 0) celdaFin = c;
                        }
                    }

                    celdaInicio ??= celdaClickeada; 
                    celdaFin ??= celdaClickeada;

                    int r1 = grupoFilas.Rows.IndexOf((TableRow)celdaInicio.Parent);
                    int c1 = ((TableRow)celdaInicio.Parent).Cells.IndexOf(celdaInicio);
                    int r2 = grupoFilas.Rows.IndexOf((TableRow)celdaFin.Parent);
                    int c2 = ((TableRow)celdaFin.Parent).Cells.IndexOf(celdaFin);

                    int minFila = Math.Min(r1, r2), maxFila = Math.Max(r1, r2);
                    int minCol = Math.Min(c1, c2), maxCol = Math.Max(c1, c2);

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
                editorActual.EndChange();
                editorActual.Focus();
            };
            dialog.Show();
        }

        public static void QuitarFondoSeleccion(RichTextBox editorActual, TableCell celdaClickeada, Table tablaPadre)
        {
            if (editorActual == null) return;
            TableRowGroup grupoFilas = (TableRowGroup)((TableRow)celdaClickeada.Parent).Parent;
            var selectionRange = editorActual.Selection;

            editorActual.BeginChange();
            
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
            editorActual.EndChange();
            editorActual.Focus();
        }

        public static void PintarBordeSeleccion(RichTextBox editorActual, TableCell celdaClickeada, Table tablaPadre, Window ownerWindow)
        {
            if (editorActual == null) return;
            TableRowGroup grupoFilas = (TableRowGroup)((TableRow)celdaClickeada.Parent).Parent;

            var dialog = new KalebClipPro.Views.PaletaColores();
            dialog.Owner = ownerWindow;
            dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;

            dialog.AlSeleccionarColor = (colorNuevo) =>
            {
                var selectionRange = editorActual.Selection;
                var brushNuevo = new SolidColorBrush(colorNuevo);

                editorActual.BeginChange();

                int minFila = 0, maxFila = 0, minCol = 0, maxCol = 0;

                if (selectionRange.IsEmpty)
                {
                    celdaClickeada.BorderBrush = brushNuevo;
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

                int totalFilas = grupoFilas.Rows.Count;
                int totalCols = totalFilas > 0 ? grupoFilas.Rows[0].Cells.Count : 0;
                
                if (minFila == 0 && maxFila == totalFilas - 1 && minCol == 0 && maxCol == totalCols - 1)
                {
                    if (tablaPadre.Tag is object[] conf && conf.Length >= 10) {
                        conf[8] = true;         
                        conf[9] = colorNuevo;   
                        tablaPadre.Tag = conf;
                    }
                }

                editorActual.EndChange();
                editorActual.Focus();
            };
            dialog.Show();
        }

        public static void CambiarTemaTabla(RichTextBox editorActual, TableCell celdaClickeada, Table tablaPadre, Window ownerWindow)
        {
            if (editorActual == null) return;
            TableRowGroup grupoFilas = (TableRowGroup)((TableRow)celdaClickeada.Parent).Parent;

            int estilo = 1; bool enc = false, tot = false, priCol = false, ultCol = false;
            if (tablaPadre.Tag is object[] config && config.Length >= 10) {
                estilo = Convert.ToInt32(config[0]);
                enc = (bool)config[4]; tot = (bool)config[5]; priCol = (bool)config[6]; ultCol = (bool)config[7];
            }

            var dialog = new KalebClipPro.Views.PaletaColores();
            dialog.Owner = ownerWindow;
            dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;

            dialog.AlSeleccionarColor = (colorNuevo) =>
            {
                editorActual.BeginChange();

                if (tablaPadre.Tag is object[] conf && conf.Length >= 10) {
                    conf[1] = colorNuevo; 
                    tablaPadre.Tag = conf;
                }

                var colorHeader = new SolidColorBrush(colorNuevo);
                var colorCebra1 = new SolidColorBrush(Color.FromArgb(60, colorNuevo.R, colorNuevo.G, colorNuevo.B));
                var colorCebra2 = new SolidColorBrush(Color.FromArgb(20, colorNuevo.R, colorNuevo.G, colorNuevo.B));

                int filas = grupoFilas.Rows.Count;
                if (filas == 0) return;
                int columnas = grupoFilas.Rows[0].Cells.Count;

                for (int f = 0; f < filas; f++)
                {
                    for (int c = 0; c < grupoFilas.Rows[f].Cells.Count; c++)
                    {
                        TableCell celda = grupoFilas.Rows[f].Cells[c];

                        if (estilo == 1) celda.Background = (f % 2 == 0) ? colorCebra1 : colorCebra2;
                        else if (estilo == 2) celda.Background = (c % 2 == 0) ? colorCebra1 : colorCebra2;
                        else if (estilo == 3) celda.Background = ((f + c) % 2 == 0) ? colorCebra1 : colorCebra2;
                        else celda.Background = Brushes.Transparent;

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
                editorActual.EndChange();
                editorActual.Focus();
            };
            dialog.Show();
        }

        public static void AplicarBordeASEleccion(RichTextBox editorActual, TableCell celdaClickeada, string tipoBorde)
        {
            if (editorActual == null) return;
            
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

            editorActual.BeginChange();
            if (tipoBorde == "Ninguno" || tipoBorde == "Todos" || tipoBorde == "Externos") tablaPadre.BorderThickness = new Thickness(0);

            var selectionRange = editorActual.Selection;
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
            editorActual.EndChange();
        }

        public static void AplicarGrosorASeleccion(RichTextBox editorActual, TableCell celdaClickeada, double nuevoGrosor)
        {
            if (editorActual == null) return;
            
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

            editorActual.BeginChange();
            var selectionRange = editorActual.Selection;
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
            editorActual.EndChange();
        }
    }
}