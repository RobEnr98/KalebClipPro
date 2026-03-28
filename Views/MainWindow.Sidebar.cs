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
using GongSolutions.Wpf.DragDrop;
using KalebClipPro.Models;
using KalebClipPro.Services;
using KalebClipPro.Infrastructure;      // <--- Arregla el error de ClipCategoryConverter

namespace KalebClipPro {
    public partial class MainWindow {

        private void TxtBusqueda_TextChanged(object sender, TextChangedEventArgs e) => ActualizarFiltros();
        private void Filtro_Changed(object sender, RoutedEventArgs e) => ActualizarFiltros();
        private void CalendarFiltro_SelectedDatesChanged(object sender, SelectionChangedEventArgs e) { Mouse.Capture(null); PopupFiltros.IsOpen = false; ActualizarFiltros(); }
        private void BtnFiltro_Click(object sender, RoutedEventArgs e) { PopupFiltros.IsOpen = !PopupFiltros.IsOpen; }
        private void LimpiarFiltros_Click(object sender, RoutedEventArgs e) { TxtBusqueda.Text = ""; CalendarFiltro.SelectedDate = null; ChkTexto.IsChecked = true; ChkCodigo.IsChecked = true; ChkEnlace.IsChecked = true; ActualizarFiltros(); }

        private void ActualizarFiltros()
        {
            if (ChkTexto == null || ChkCodigo == null || ChkEnlace == null || CalendarFiltro == null || TituloHistorial == null) return;

            bool verTexto = ChkTexto.IsChecked == true;
            bool verCodigo = ChkCodigo.IsChecked == true;
            bool verEnlace = ChkEnlace.IsChecked == true;
            DateTime? fechaFiltro = CalendarFiltro.SelectedDate;
            string textoBusqueda = TxtBusqueda.Text ?? "";

            if (fechaFiltro.HasValue) TituloHistorial.Text = $"HISTORIAL - {fechaFiltro.Value:dd/MM/yyyy}";
            else TituloHistorial.Text = "HISTORIAL";

            bool sinFiltros = verTexto && verCodigo && verEnlace && !fechaFiltro.HasValue && string.IsNullOrEmpty(textoBusqueda);
            if (sinFiltros)
            {
                if (_filtrosActivos) { _filtrosActivos = false; CargarDatos(); }
                return;
            }

            _filtrosActivos = true;

            List<ClipData> todosLosClips = new List<ClipData>();
            int pag = 0;
            while (true)
            {
                var lote = db.ObtenerHistorialPaginado(pag, 100);
                if (lote == null || lote.Count == 0) break;
                todosLosClips.AddRange(lote);
                pag++;
            }

            var filtrados = todosLosClips.Where(clip =>
            {
                bool coincideTexto = string.IsNullOrEmpty(textoBusqueda) || clip.Contenido_Plano.Contains(textoBusqueda, StringComparison.OrdinalIgnoreCase);
                bool coincideFecha = true;
                if (fechaFiltro.HasValue) coincideFecha = clip.Fecha_Creacion.Date == fechaFiltro.Value.Date;

                bool esEnlace = clip.Contenido_Plano.Contains("http");
                bool esCodigo = clip.Contenido_Plano.Contains("public ") || clip.Contenido_Plano.Contains("{");
                bool esTextoSimple = !esEnlace && !esCodigo;

                bool coincideTipo = (verTexto && esTextoSimple) || (verCodigo && esCodigo) || (verEnlace && esEnlace);
                return coincideTexto && coincideFecha && coincideTipo;
            }).ToList();

            MisClips.Clear();
            foreach (var c in filtrados) MisClips.Add(c);
        }

        private void PopupFiltros_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            e.Handled = true; 
            Point pos = Mouse.GetPosition(this);
            if (pos.X > SidebarPanel.ActualWidth && _editorActual != null)
                _editorActual.RaiseEvent(new MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta) { RoutedEvent = UIElement.MouseWheelEvent, Source = _editorActual });
            else
            {
                ScrollViewer? sv = ObtenerScrollViewer(ListaClips);
                if (sv != null)
                {
                    if (e.Delta > 0) { sv.LineUp(); sv.LineUp(); sv.LineUp(); }
                    else { sv.LineDown(); sv.LineDown(); sv.LineDown(); }
                }
            }
        }

        private ScrollViewer? ObtenerScrollViewer(DependencyObject depObj)
        {
            if (depObj is ScrollViewer scrollViewer) return scrollViewer;
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
            {
                var result = ObtenerScrollViewer(VisualTreeHelper.GetChild(depObj, i));
                if (result != null) return result;
            }
            return null;
        }

        private void ListaClips_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Down)
            {
                e.Handled = true;
                if (ListaClips.SelectedIndex < MisClips.Count - 1)
                {
                    ListaClips.SelectedIndex++;
                    ListaClips.ScrollIntoView(ListaClips.SelectedItem);
                }
            }
            else if (e.Key == Key.Up)
            {
                e.Handled = true;
                if (ListaClips.SelectedIndex > 0)
                {
                    ListaClips.SelectedIndex--;
                    ListaClips.ScrollIntoView(ListaClips.SelectedItem);
                }
            }
            else if (e.Key == Key.Home)
            {
                e.Handled = true;
                ListaClips.SelectedIndex = 0;
                ListaClips.ScrollIntoView(ListaClips.SelectedItem);
            }
            else if (e.Key == Key.End)
            {
                e.Handled = true;
                ListaClips.SelectedIndex = MisClips.Count - 1;
                ListaClips.ScrollIntoView(ListaClips.SelectedItem);
            }
        }

        private void ListaClips_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.RightButton == MouseButtonState.Pressed)
            {
                DependencyObject? depObj = e.OriginalSource as DependencyObject;
                Button? botonEncontrado = null;

                // Buscamos si tocaste el botón
                while (depObj != null && depObj is not ListBoxItem)
                {
                    if (depObj is Button btn && btn.Name == "BtnAsignarHotkey")
                    {
                        botonEncontrado = btn;
                        break;
                    }
                    depObj = VisualTreeHelper.GetParent(depObj);
                }

                // SI TOCASTE EL BOTÓN:
                if (botonEncontrado != null)
                {
                    _clipSeleccionadoParaMenu = botonEncontrado.DataContext as ClipData;
                    // Abrimos el menú manualmente por código
                    if (botonEncontrado.ContextMenu != null)
                    {
                        botonEncontrado.ContextMenu.PlacementTarget = botonEncontrado;
                        botonEncontrado.ContextMenu.IsOpen = true;
                    }
                }

                // MAGIA NEGRA: Sin importar si tocaste el botón o el fondo de la tarjeta,
                // matamos el clic aquí mismo. Así la Lista NUNCA se selecciona ni abre el Smart Editor.
                e.Handled = true;
                return;
            }
            
            if (e.OriginalSource is DependencyObject source)
            {
                var parent = source;
                bool clickEnTarjeta = false;
                bool clickEnScroll = false;

                while (parent != null && parent is not ListBox)
                {
                    if (parent is ScrollBar) clickEnScroll = true;
                    if (parent is Border border && border.Name == "CardBorder") clickEnTarjeta = true;
                    
                    parent = VisualTreeHelper.GetParent(parent);
                }

                if (!clickEnTarjeta && !clickEnScroll)
                {
                    ListaClips.SelectedItem = null;
                }
            }
        }

        private void ListaClips_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (_cargandoMas || _filtrosActivos) return; 

            if (e.VerticalOffset + e.ViewportHeight >= e.ExtentHeight - 100 && e.VerticalChange > 0)
            {
                _cargandoMas = true;
                _paginaActual++;
                
                Dispatcher.BeginInvoke(new Action(() => {
                    var nuevos = db.ObtenerHistorialPaginado(_paginaActual, _itemsPorPagina);
                    if (nuevos.Count > 0)
                    {
                        ScrollViewer? sv = ObtenerScrollViewer(ListaClips);
                        double offsetGuardado = sv != null ? sv.VerticalOffset : 0;

                        foreach (var n in nuevos) MisClips.Add(n);

                        if (sv != null)
                        {
                            ListaClips.UpdateLayout(); 
                            sv.ScrollToVerticalOffset(offsetGuardado);
                        }
                    }
                    _cargandoMas = false;
                }), DispatcherPriority.Background);
            }
        }

        private void BtnEliminarClip_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is ClipData clip)
            {
                // 1. Lo borramos de la base de datos
                db.EliminarClip(clip.Guid_Clip);

                // 2. Lo quitamos visualmente de la lista izquierda
                MisClips.Remove(clip);

                // 3. Si teníamos ese texto abierto en el editor, limpiamos la pantalla
                if (ListaClips.SelectedItem == clip)
                {
                    ListaClips.SelectedItem = null;
                    if (_editorActual != null)
                    {
                        _editorActual.Visibility = Visibility.Collapsed;
                        _editorActual = null;
                    }
                    if (TabEditor != null) TabEditor.Content = "Smart Editor";
                }
            }
        }

        private void ToggleSidebar_Click(object sender, RoutedEventArgs e)
        {
            TimeSpan tiempo = TimeSpan.FromMilliseconds(180); 
            IEasingFunction suavizado = new ExponentialEase { EasingMode = EasingMode.EaseOut, Exponent = 7 };

            if (isSidebarOpen)
            {
                isSidebarOpen = false;
                SidebarContentGrid.Width = SidebarContentGrid.ActualWidth;
                TextBoxContainer.Width = TextBoxContainer.ActualWidth;
                TextBoxContainer.HorizontalAlignment = HorizontalAlignment.Left;
                DoubleAnimation animWidth = new DoubleAnimation() { To = 0, Duration = tiempo, EasingFunction = suavizado };
                animWidth.Completed += (s, a) => { 
                    SidebarPanel.BeginAnimation(FrameworkElement.WidthProperty, null);
                    SidebarPanel.Width = 0; SidebarPanel.Visibility = Visibility.Collapsed; 
                    BtnShowSidebar.Visibility = Visibility.Visible; 
                    SidebarContentGrid.Width = double.NaN; TextBoxContainer.Width = double.NaN;
                    TextBoxContainer.HorizontalAlignment = HorizontalAlignment.Stretch;
                };
                SidebarPanel.BeginAnimation(FrameworkElement.WidthProperty, animWidth);
            }
            else
            {
                isSidebarOpen = true;
                SidebarPanel.Visibility = Visibility.Visible;
                BtnShowSidebar.Visibility = Visibility.Collapsed;
                SidebarContentGrid.Width = _anchoSidebarGuardado;
                TextBoxContainer.Width = TextBoxContainer.ActualWidth;
                TextBoxContainer.HorizontalAlignment = HorizontalAlignment.Left;
                DoubleAnimation animWidth = new DoubleAnimation() { To = _anchoSidebarGuardado, Duration = tiempo, EasingFunction = suavizado };
                animWidth.Completed += (s, a) => {
                    SidebarPanel.BeginAnimation(FrameworkElement.WidthProperty, null);
                    SidebarPanel.Width = _anchoSidebarGuardado;
                    SidebarContentGrid.Width = double.NaN; TextBoxContainer.Width = double.NaN;
                    TextBoxContainer.HorizontalAlignment = HorizontalAlignment.Stretch;
                };
                SidebarPanel.BeginAnimation(FrameworkElement.WidthProperty, animWidth);
            }
        }

        private void SidebarResizer_DragStarted(object sender, DragStartedEventArgs e)
        {
            _currentDragWidth = double.IsNaN(SidebarPanel.Width) ? SidebarPanel.ActualWidth : SidebarPanel.Width;
            SidebarContentGrid.Width = SidebarContentGrid.ActualWidth;
            TextBoxContainer.Width = TextBoxContainer.ActualWidth;
            TextBoxContainer.HorizontalAlignment = HorizontalAlignment.Left;
        }

        private void SidebarResizer_DragDelta(object sender, DragDeltaEventArgs e)
        {
            _currentDragWidth += e.HorizontalChange;
            if (_currentDragWidth >= 120 && _currentDragWidth <= 800)
            {
                SidebarPanel.Width = _currentDragWidth;
                _anchoSidebarGuardado = _currentDragWidth;
            }
        }

        private void SidebarResizer_DragCompleted(object sender, DragCompletedEventArgs e)
        {
            SidebarContentGrid.Width = double.NaN;
            TextBoxContainer.Width = double.NaN;
            TextBoxContainer.HorizontalAlignment = HorizontalAlignment.Stretch;
        }

        
    }
}