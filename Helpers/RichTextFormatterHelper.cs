using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

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

            // Cálculo anti-zoom de Windows movido aquí
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
    }
}