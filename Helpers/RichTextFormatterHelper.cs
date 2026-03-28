using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using ColorCode;
using ColorCode.Wpf;
using ColorCode.Styling;
using ColorCode.Common;

namespace KalebClipPro.Helpers
{
    public static class RichTextFormatterHelper
    {
        public static void InsertarCita(RichTextBox editorActual)
        {
            if (editorActual == null) return;
            
            Paragraph p = new Paragraph(new Run(editorActual.Selection.Text)) {
                Margin = new Thickness(20, 10, 0, 10),
                FontStyle = FontStyles.Italic,
                Foreground = Brushes.Gray,
                BorderBrush = Brushes.Gray,
                BorderThickness = new Thickness(2, 0, 0, 0),
                Padding = new Thickness(10, 0, 0, 0)
            };
            editorActual.Document.Blocks.Add(p);
            editorActual.Focus();
        }

        public static void FormatearComoCodigo(RichTextBox editorActual)
        {
            if (editorActual == null) return;
            
            editorActual.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, new FontFamily("Consolas"));
            editorActual.Selection.ApplyPropertyValue(TextElement.BackgroundProperty, new SolidColorBrush(Color.FromRgb(40, 44, 52)));
            editorActual.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, Brushes.LightGreen);
            editorActual.Focus();
        }

        public static void InsertarEnlace(RichTextBox editorActual, string url)
        {
            if (editorActual == null || editorActual.Selection.IsEmpty) return;
            
            Hyperlink link = new Hyperlink(editorActual.Selection.Start, editorActual.Selection.End);
            link.NavigateUri = new Uri(url);
            link.Cursor = Cursors.Hand;
            editorActual.Focus();
        }

        public static void FormatearComoMatematicas(RichTextBox editorActual)
        {
            if (editorActual == null) return;

            TextRange range = new TextRange(editorActual.Selection.Start, editorActual.Selection.End);
            range.Text = $"${editorActual.Selection.Text}$"; 
            range.ApplyPropertyValue(TextElement.ForegroundProperty, Brushes.Orange);
            editorActual.Focus();
        }

        public static void CambiarColorTexto(RichTextBox editorActual, object senderBoton, Window ownerWindow)
        {
            if (editorActual == null) return;

            var dialog = new KalebClipPro.Views.PaletaColores();
            dialog.Owner = ownerWindow; 
            var colorActualObj = editorActual.Selection.GetPropertyValue(TextElement.ForegroundProperty);

            Button? boton = senderBoton as Button;
            if (boton != null)
            {
                Point screenPos = boton.PointToScreen(new Point(0, boton.ActualHeight));
                PresentationSource source = PresentationSource.FromVisual(ownerWindow);
                
                if (source != null)
                {
                    dialog.WindowStartupLocation = WindowStartupLocation.Manual;
                    dialog.Left = screenPos.X / source.CompositionTarget.TransformToDevice.M11;
                    dialog.Top = screenPos.Y / source.CompositionTarget.TransformToDevice.M22;
                }
            }

            if (colorActualObj is SolidColorBrush brush)
            {
                dialog.CargarColorDesdeEditor(brush.Color);
            }
            
            dialog.AlSeleccionarColor = (color) =>
            {
                var newBrush = new SolidColorBrush(color);
                editorActual.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, newBrush);
                editorActual.Focus();
            };

            dialog.Show();
        }

        public static void CambiarFuente(RichTextBox editorActual, FontFamily fuente)
        {
            if (editorActual == null || fuente == null) return;
            editorActual.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, fuente);
            editorActual.Focus();
        }

        public static void CambiarTamano(RichTextBox editorActual, double tamano)
        {
            if (editorActual == null) return;
            editorActual.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, tamano);
            editorActual.Focus();
        }

        public static void AplicarInterlineado(RichTextBox editor, double alturaLinea, double espacioAbajo)
        {
            if (editor == null || editor.Selection.IsEmpty) return;

            editor.BeginChange();

            var bloquesSeleccionados = editor.Document.Blocks
                .Where(b => editor.Selection.Contains(b.ContentStart) || 
                            editor.Selection.Contains(b.ContentEnd));

            foreach (var bloque in bloquesSeleccionados)
            {
                if (bloque is Paragraph p)
                {
                    p.LineHeight = alturaLinea;        
                    p.Margin = new Thickness(0, 0, 0, espacioAbajo); 
                }
            }

            editor.EndChange();
        }

        public static void AplicarColorCode(RichTextBox editor, string codigoCrudo)
        {
            if (editor == null || string.IsNullOrWhiteSpace(codigoCrudo)) return;

            editor.BeginChange();
            editor.Document.Blocks.Clear();

            var miEstiloVibrante = ColorCode.Styling.StyleDictionary.DefaultDark;

            ActualizarColor(miEstiloVibrante, ScopeName.Keyword, "#F92672");    
            ActualizarColor(miEstiloVibrante, ScopeName.String, "#E6DB74");     
            ActualizarColor(miEstiloVibrante, ScopeName.Comment, "#75715E");    
            ActualizarColor(miEstiloVibrante, ScopeName.Number, "#AE81FF");     
            ActualizarColor(miEstiloVibrante, ScopeName.PlainText, "#F8F8F2");  
            
            var formatter = new RichTextBoxFormatter(miEstiloVibrante);
            Paragraph p = new Paragraph { Margin = new Thickness(0) };

            formatter.FormatInlines(codigoCrudo, Languages.CSharp, p.Inlines);

            editor.Document.Blocks.Add(p);
            editor.EndChange();
        }

        public static void ActualizarColor(ColorCode.Styling.StyleDictionary dic, string scope, string hex)
        {
            if (dic.Contains(scope))
            {
                dic[scope].Foreground = hex;
            }
        }

        public static int GetPreviousListNumber(RichTextBox? editor)
        {
            if (editor == null) return 0;

            try 
            {
                TextPointer caret = editor.CaretPosition;
                Block? currentBlock = caret.Paragraph;

                if (currentBlock?.Parent is ListItem li && li.Parent is List parentList)
                {
                    currentBlock = parentList;
                }

                while (currentBlock != null)
                {
                    currentBlock = currentBlock.PreviousBlock;

                    if (currentBlock is List listaPrevia && listaPrevia.MarkerStyle == TextMarkerStyle.Decimal)
                    {
                        return listaPrevia.StartIndex + listaPrevia.ListItems.Count - 1;
                    }
                }
            } 
            catch { }
            
            return 0;
        }

        public static void AplicarListaNumeradaInteligente(RichTextBox? editor)
        {
            if (editor == null) return;

            Application.Current.Dispatcher.BeginInvoke(new Action(() =>
            {
                Keyboard.Focus(editor);
                editor.Focus();

                int ultimoNumero = GetPreviousListNumber(editor);

                EditingCommands.ToggleNumbering.Execute(null, editor);

                if (editor.CaretPosition.Paragraph?.Parent is ListItem listItem && listItem.Parent is List currentList)
                {
                    currentList.StartIndex = ultimoNumero > 0 ? ultimoNumero + 1 : 1;
                }

            }), DispatcherPriority.ApplicationIdle); 
        }

        public static void AplicarSangriaPersonalizada(RichTextBox? editor, int direccion)
        {
            if (editor == null) return;

            editor.BeginChange();

            TextSelection seleccion = editor.Selection;
            List<Paragraph> parrafosAfectados = new List<Paragraph>();

            if (seleccion.IsEmpty && editor.CaretPosition.Paragraph != null)
            {
                parrafosAfectados.Add(editor.CaretPosition.Paragraph);
            }
            else
            {
                foreach (Block block in editor.Document.Blocks)
                {
                    if (block is Paragraph p && (seleccion.Contains(p.ContentStart) || seleccion.Contains(p.ContentEnd)))
                    {
                        parrafosAfectados.Add(p);
                    }
                    else if (block is List lista)
                    {
                        foreach (ListItem li in lista.ListItems)
                        {
                            if (li.Blocks.FirstBlock is Paragraph lp && (seleccion.Contains(lp.ContentStart) || seleccion.Contains(lp.ContentEnd)))
                            {
                                parrafosAfectados.Add(lp);
                            }
                        }
                    }
                }
            }

            foreach (Paragraph p in parrafosAfectados)
            {
                if (p.Parent is ListItem li)
                {
                    double nuevaSangria = li.Margin.Left + (direccion * 35);
                    if (nuevaSangria < 0) nuevaSangria = 0;
                    li.Margin = new Thickness(nuevaSangria, li.Margin.Top, li.Margin.Right, li.Margin.Bottom);
                }
                else
                {
                    double nuevaSangria = p.Margin.Left + (direccion * 35);
                    if (nuevaSangria < 0) nuevaSangria = 0;
                    p.Margin = new Thickness(nuevaSangria, p.Margin.Top, p.Margin.Right, p.Margin.Bottom);
                }
            }

            editor.EndChange();
            
            if (editor.Parent is Grid wrapper && wrapper.Children.OfType<Border>().FirstOrDefault()?.Child is Canvas canvas)
            {
                editor.FontSize = editor.FontSize; 
            }
        }
    }
}