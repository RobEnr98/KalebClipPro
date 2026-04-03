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
using KalebClipPro.Infrastructure;

namespace KalebClipPro 
{
    public partial class MainWindow : IDropTarget 
    {
        // NOTA: _setActual ahora guarda el ID de la Carpeta (Guid), no una letra ("A", "B").

        private void BtnAsignarHotkey_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is ClipData clip)
            {
                if (clip.HotKeyIndex > 0)
                {
                    // DESANCLAR
                    int index = clip.HotKeyIndex - 1;
                    if (index >= 0 && index < 9)
                    {
                        ClipsRecolector[index].Guid_Clip = "";
                        ClipsRecolector[index].Contenido_Plano = "";
                        ClipsRecolector[index].Origen_App = "";
                    }
                }
                else
                {
                    // ANCLAR (Automático a primer vacío)
                    var primerVacio = ClipsRecolector.FirstOrDefault(c => c.EsVacio);
                    
                    if (primerVacio != null)
                    {
                        primerVacio.Guid_Clip = clip.Guid_Clip;
                        primerVacio.Contenido_Plano = clip.Contenido_Plano;
                        primerVacio.Origen_App = clip.Origen_App;
                    }
                    else
                    {
                        // FIFO (Cascada) si están llenos
                        for (int i = 8; i > 0; i--) 
                        {
                            ClipsRecolector[i].Guid_Clip = ClipsRecolector[i - 1].Guid_Clip;
                            ClipsRecolector[i].Contenido_Plano = ClipsRecolector[i - 1].Contenido_Plano;
                            ClipsRecolector[i].Origen_App = ClipsRecolector[i - 1].Origen_App;
                        }
                        ClipsRecolector[0].Guid_Clip = clip.Guid_Clip;
                        ClipsRecolector[0].Contenido_Plano = clip.Contenido_Plano;
                        ClipsRecolector[0].Origen_App = clip.Origen_App;
                    }
                }

                GuardarSlotsActuales();
                SincronizarPinesHistorial();
            }
        }

        private void AsignarSlotEspecifico_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem && menuItem.Tag != null && _clipSeleccionadoParaMenu != null)
            {
                string tagValue = menuItem.Tag.ToString() ?? "";
                ClipData clip = _clipSeleccionadoParaMenu;

                // 1. DESANCLAR (Botón Rojo del menú)
                if (tagValue == "0")
                {
                    if (clip.HotKeyIndex > 0)
                    {
                        int oldIndex = clip.HotKeyIndex - 1;
                        if (oldIndex >= 0 && oldIndex < 9)
                        {
                            ClipsRecolector[oldIndex].Guid_Clip = "";
                            ClipsRecolector[oldIndex].Contenido_Plano = "";
                            ClipsRecolector[oldIndex].Origen_App = "";
                        }
                    }
                }
                // 2. ASIGNAR A CARPETA Y SLOT ESPECÍFICO
                else
                {
                    var partes = tagValue.Split('|');
                    if (partes.Length == 2)
                    {
                        string targetCarpetaId = partes[0];
                        int targetSlot = int.Parse(partes[1]);
                        int newIndex = targetSlot - 1;

                        // Limpiamos el visual si estaba en nuestra carpeta actual
                        if (clip.HotKeyIndex > 0)
                        {
                            int oldIndex = clip.HotKeyIndex - 1;
                            if (oldIndex >= 0 && oldIndex < 9)
                            {
                                ClipsRecolector[oldIndex].Guid_Clip = "";
                                ClipsRecolector[oldIndex].Contenido_Plano = "";
                                ClipsRecolector[oldIndex].Origen_App = "";
                            }
                        }

                        // Buscamos la carpeta destino en el modelo y guardamos la referencia
                        var carpetaDestino = WorkflowActual.Carpetas.FirstOrDefault(c => c.Id == targetCarpetaId);
                        if (carpetaDestino != null)
                        {
                            carpetaDestino.Slots[newIndex].ClipIdAsignado = clip.Guid_Clip;

                            // Si es la carpeta que estamos viendo, actualizamos la UI
                            if (targetCarpetaId == _setActual)
                            {
                                ClipsRecolector[newIndex].Guid_Clip = clip.Guid_Clip;
                                ClipsRecolector[newIndex].Contenido_Plano = clip.Contenido_Plano;
                                ClipsRecolector[newIndex].Origen_App = clip.Origen_App;
                            }
                        }
                    }
                }

                _clipSeleccionadoParaMenu = null;
                GuardarSlotsActuales();
                SincronizarPinesHistorial();
            }
        }

        private void GuardarSlotsActuales()
        {
            try
            {
                var carpetaActiva = WorkflowActual.Carpetas.FirstOrDefault(c => c.Id == _setActual);
                if (carpetaActiva != null)
                {
                    // Solo guardamos el ID del clip en el modelo JSON
                    for (int i = 0; i < 9; i++)
                    {
                        carpetaActiva.Slots[i].ClipIdAsignado = ClipsRecolector[i].EsVacio ? "" : ClipsRecolector[i].Guid_Clip;
                    }
                    _workflowService.GuardarWorkflowEnDisco(WorkflowActual); 
                }
            }
            catch { }
        }

        private void InicializarSlotsVacios()
        {
            ClipsRecolector.Clear();

            // Asegurarnos de tener una carpeta seleccionada
            if (string.IsNullOrEmpty(_setActual) || !WorkflowActual.Carpetas.Any(c => c.Id == _setActual))
            {
                if (WorkflowActual.Carpetas.Count == 0)
                {
                    WorkflowActual.Carpetas.Add(_workflowService.CrearCarpetaVacia("General"));
                }
                _setActual = WorkflowActual.Carpetas.FirstOrDefault()?.Id ?? "";
            }

            var carpeta = WorkflowActual.Carpetas.FirstOrDefault(c => c.Id == _setActual);
            if (carpeta == null) return;

            foreach (var slot in carpeta.Slots)
            {
                ClipData clipParaMostrar = new ClipData { HotKeyIndex = slot.HotKeyIndex, AlturaVisual = 38 };

                if (!string.IsNullOrEmpty(slot.ClipIdAsignado))
                {
                    // Magia V2: Buscamos el texto real en la Base de Datos usando su ID
                    var clipDesdeBD = db.ObtenerClipPorId(slot.ClipIdAsignado); 
                    
                    if (clipDesdeBD != null)
                    {
                        clipParaMostrar.Guid_Clip = clipDesdeBD.Guid_Clip;
                        clipParaMostrar.Contenido_Plano = clipDesdeBD.Contenido_Plano;
                        clipParaMostrar.Origen_App = clipDesdeBD.Origen_App;
                    }
                }
                
                ClipsRecolector.Add(clipParaMostrar);
            }
            
            ActualizarContadorTab();
        }

        private void SincronizarBotonesConDatos()
        {
            if (ContenedorSets == null) return;
            ContenedorSets.Children.Clear();

            foreach (var carpeta in WorkflowActual.Carpetas)
            {
                RadioButton nuevoSet = new RadioButton
                {
                    Content = carpeta.Nombre,
                    Tag = carpeta.Id, // El Tag ahora es el Guid de la carpeta
                    Style = (Style)FindResource("SegmentedButtonStyle"),
                    IsChecked = (carpeta.Id == _setActual)
                };
                
                // --- AÑADIMOS EL MENÚ CONTEXTUAL PARA RENOMBRAR ---
                ContextMenu menuOpciones = new ContextMenu();
                MenuItem renombrarItem = new MenuItem 
                { 
                    Header = "Renombrar Set", 
                    Tag = carpeta.Id 
                };
                renombrarItem.Click += RenombrarSet_Click;
                menuOpciones.Items.Add(renombrarItem);
                
                nuevoSet.ContextMenu = menuOpciones;
                // ---------------------------------------------------

                nuevoSet.Click += BtnCambiarSet_Click;
                ContenedorSets.Children.Add(nuevoSet);
            }
        }

        private void RenombrarSet_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem item && item.Tag is string idCarpeta)
            {
                var carpeta = WorkflowActual.Carpetas.FirstOrDefault(c => c.Id == idCarpeta);
                if (carpeta == null) return;

                Window inputWindow = new Window
                {
                    Title = "Renombrar Set",
                    Width = 380,
                    Height = 210,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = this,
                    WindowStyle = WindowStyle.None, 
                    ResizeMode = ResizeMode.NoResize,
                    AllowsTransparency = true,
                    Background = Brushes.Transparent
                };

                // Permitimos mover la ventana arrastrando el fondo
                inputWindow.MouseLeftButtonDown += (s, ev) => inputWindow.DragMove();

                // Contenedor principal (Color idéntico al fondo de tu popup de tabla)
                Border mainBorder = new Border
                {
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#222831")), 
                    BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#394859")),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(8),
                    Padding = new Thickness(24)
                };

                StackPanel stack = new StackPanel();
                
                // Título Centrado (Como "Insertar Tabla Avanzada")
                stack.Children.Add(new TextBlock 
                { 
                    Text = "Renombrar Set", 
                    FontSize = 16,
                    FontWeight = FontWeights.Bold,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 0, 0, 20),
                    Foreground = Brushes.White
                });

                // Label del Input
                stack.Children.Add(new TextBlock 
                { 
                    Text = "Nombre del set:", 
                    FontSize = 12,
                    Margin = new Thickness(0, 0, 0, 5),
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A0AAB5")) 
                });

                // Envolvemos el TextBox en un Border para darle fondo oscuro y bordes redondeados
                Border txtBorder = new Border
                {
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1A1F26")), // Fondo input oscuro
                    BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#394859")), // Borde gris sutil
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(4),
                    Margin = new Thickness(0, 0, 0, 25)
                };

                TextBox txtNuevoNombre = new TextBox 
                { 
                    Text = carpeta.Nombre, 
                    Padding = new Thickness(10, 8, 10, 8),
                    Background = Brushes.Transparent, // El fondo lo da el Border
                    Foreground = Brushes.White,
                    CaretBrush = Brushes.White,
                    BorderThickness = new Thickness(0), // Quitamos el borde nativo feo
                    FontSize = 14,
                    VerticalContentAlignment = VerticalAlignment.Center
                };
                txtBorder.Child = txtNuevoNombre;
                stack.Children.Add(txtBorder);

                // Botones
                StackPanel buttonStack = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
                
                // Botón Cancelar (Texto gris, sin fondo)
                Button btnCancelar = new Button 
                { 
                    Content = "Cancelar", 
                    Margin = new Thickness(0,0,15,0), 
                    Background = Brushes.Transparent, 
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A0AAB5")), 
                    BorderThickness = new Thickness(0), 
                    Cursor = Cursors.Hand,
                    FontSize = 13
                };
                btnCancelar.Click += (s, ev) => inputWindow.Close();

                // Botón Guardar (Color Cyan sólido con bordes redondeados creados a mano)
                Border btnGuardarBorder = new Border
                {
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#00BFA5")), // Cyan exacto de tu diseño
                    CornerRadius = new CornerRadius(4),
                    Cursor = Cursors.Hand
                };
                
                TextBlock textGuardar = new TextBlock
                {
                    Text = "Guardar",
                    Foreground = Brushes.White,
                    FontWeight = FontWeights.Medium,
                    FontSize = 13,
                    Padding = new Thickness(20, 8, 20, 8),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
                btnGuardarBorder.Child = textGuardar;
                
                // Efecto hover (cambia de color al pasar el mouse)
                btnGuardarBorder.MouseEnter += (s, ev) => btnGuardarBorder.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#14CBAE"));
                btnGuardarBorder.MouseLeave += (s, ev) => btnGuardarBorder.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#00BFA5"));
                
                // Acción de guardar
                btnGuardarBorder.MouseLeftButtonUp += (s, ev) => { inputWindow.DialogResult = true; inputWindow.Close(); };
                
                buttonStack.Children.Add(btnCancelar);
                buttonStack.Children.Add(btnGuardarBorder);
                stack.Children.Add(buttonStack);

                mainBorder.Child = stack;
                inputWindow.Content = mainBorder;

                // Foco automático en el texto
                inputWindow.Loaded += (s, ev) => { txtNuevoNombre.Focus(); txtNuevoNombre.SelectAll(); };
                
                // Permitir guardar presionando la tecla Enter
                txtNuevoNombre.KeyDown += (s, ev) => 
                { 
                    if (ev.Key == Key.Enter) { inputWindow.DialogResult = true; inputWindow.Close(); } 
                };

                if (inputWindow.ShowDialog() == true && !string.IsNullOrWhiteSpace(txtNuevoNombre.Text))
                {
                    carpeta.Nombre = txtNuevoNombre.Text.Trim();
                    _workflowService.GuardarWorkflowEnDisco(WorkflowActual);
                    SincronizarBotonesConDatos(); 
                    ActualizarContadorTab();
                }
            }
        }
        // ====================================================================

        private void BtnCambiarSet_Click(object sender, RoutedEventArgs e)
        {
            if (sender is RadioButton btn && btn.Tag is string idCarpeta)
            {
                if (_setActual == idCarpeta) return;

                GuardarSlotsActuales();
                _setActual = idCarpeta;
                InicializarSlotsVacios();
                SincronizarPinesHistorial(); 
            }
        }

        private void BtnQuitarSet_Click(object sender, RoutedEventArgs e)
        {
            // 1. Regla de seguridad: No podemos quedarnos sin carpetas
            if (WorkflowActual.Carpetas.Count <= 1)
            {
                MessageBox.Show("Debes tener al menos una carpeta activa.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Information);
                return; 
            }

            // Buscamos la carpeta que el usuario quiere borrar (la actual)
            var carpetaABorrar = WorkflowActual.Carpetas.FirstOrDefault(c => c.Id == _setActual);
            
            if (carpetaABorrar != null)
            {
                // 2. Verificación de contenido para no borrar por error algo importante
                bool tieneDatos = carpetaABorrar.Slots.Any(s => !string.IsNullOrEmpty(s.ClipIdAsignado));
                if (tieneDatos)
                {
                    var respuesta = MessageBox.Show($"El set '{carpetaABorrar.Nombre}' contiene clips anclados. ¿Seguro que quieres eliminarlo?", "Confirmar", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                    if (respuesta == MessageBoxResult.No) return;
                }

                // --- MAGIA DE NAVEGACIÓN ---
                // Obtenemos la posición actual antes de borrarla
                int indiceABorrar = WorkflowActual.Carpetas.IndexOf(carpetaABorrar);

                // Borramos la carpeta del modelo y guardamos en disco
                WorkflowActual.Carpetas.Remove(carpetaABorrar);
                _workflowService.GuardarWorkflowEnDisco(WorkflowActual);

                // 3. Calculamos el nuevo índice:
                int nuevoIndice = Math.Max(0, indiceABorrar - 1);
                
                // Asignamos el ID de la carpeta que quedó en esa posición
                _setActual = WorkflowActual.Carpetas[nuevoIndice].Id;

                // 4. Refrescamos la interfaz
                SincronizarBotonesConDatos();
                InicializarSlotsVacios();
                SincronizarPinesHistorial();
            }
        }

        private void MenuAsignarSlot_Opened(object sender, RoutedEventArgs e)
        {
            if (sender is ContextMenu menu && WorkflowActual != null && WorkflowActual.Carpetas != null)
            {
                menu.Items.Clear();

                Style itemStyle = (Style)FindResource("MenuOscuroItemStyle");
                Style separatorStyle = (Style)FindResource("MenuOscuroSeparatorStyle");

                foreach (var carpeta in WorkflowActual.Carpetas)
                {
                    MenuItem setMenuItem = new MenuItem { Header = carpeta.Nombre, Style = itemStyle };

                    for (int i = 1; i <= 9; i++)
                    {
                        MenuItem slotMenuItem = new MenuItem 
                        { 
                            Header = $"Slot {i}",
                            Tag = $"{carpeta.Id}|{i}", // Separador | para no confundir guiones de los GUID
                            Style = itemStyle 
                        };
                        slotMenuItem.Click += AsignarSlotEspecifico_Click;
                        setMenuItem.Items.Add(slotMenuItem);
                    }

                    menu.Items.Add(setMenuItem);
                }

                Separator sep = new Separator { Style = separatorStyle };
                menu.Items.Add(sep);
                
                MenuItem desanclarItem = new MenuItem 
                { 
                    Header = "Desanclar", 
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EF4444")),
                    Tag = "0",
                    Style = itemStyle
                };
                desanclarItem.Click += AsignarSlotEspecifico_Click;
                menu.Items.Add(desanclarItem);
            }
        }

        private void ActualizarContadorTab()
        {
            int llenos = ClipsRecolector.Count(c => !c.EsVacio);
            var carpeta = WorkflowActual.Carpetas.FirstOrDefault(c => c.Id == _setActual);
            string nombreCarpeta = carpeta != null ? carpeta.Nombre : "Recolector";
            
            if (TabRecolector != null)
                TabRecolector.ToolTip = $"{nombreCarpeta} ({llenos} slots ocupados)";
        }

        private void CargarDatos()
        {
            try 
            {
                _paginaActual = 0;
                MisClips.Clear();
                var datos = db.ObtenerHistorialPaginado(_paginaActual, _itemsPorPagina);
                foreach (var c in datos) MisClips.Add(c);
            } 
            catch { }
        }

        private void BtnAgregarSet_Click(object sender, RoutedEventArgs e)
        {
            // Ahora la numeración es correlativa
            int numero = WorkflowActual.Carpetas.Count + 1;
            var nuevaCarpeta = _workflowService.CrearCarpetaVacia($"Set {numero}");
            
            WorkflowActual.Carpetas.Add(nuevaCarpeta);
            _workflowService.GuardarWorkflowEnDisco(WorkflowActual);

            SincronizarBotonesConDatos();

            foreach (var child in ContenedorSets.Children)
            {
                if (child is RadioButton rb && rb.Tag.ToString() == nuevaCarpeta.Id)
                {
                    rb.IsChecked = true;
                    GuardarSlotsActuales();
                    _setActual = nuevaCarpeta.Id;
                    InicializarSlotsVacios();
                    SincronizarPinesHistorial();
                    break;
                }
            }
        }

        private void BtnLimpiarRecolector_Click(object sender, RoutedEventArgs e)
        {
            foreach (var slot in ClipsRecolector)
            {
                slot.Guid_Clip = "";
                slot.Contenido_Plano = "";
                slot.Origen_App = "";
            }
            GuardarSlotsActuales();
            SincronizarPinesHistorial();
        }

        private void BtnCopiarSlot_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is ClipData clip)
            {
                Clipboard.SetText(clip.Contenido_Plano);
                string nuevoOrigen = "Exportado de " + clip.Origen_App;
                string nuevoIdGenerado = db.GuardarClipConOrigen(clip.Contenido_Plano, nuevoOrigen);
                var nuevoClip = new ClipData {
                    Guid_Clip = nuevoIdGenerado, Contenido_Plano = clip.Contenido_Plano,
                    Fecha_Creacion = DateTime.Now, Origen_App = nuevoOrigen
                };
                if (MisClips != null) MisClips.Insert(0, nuevoClip); 
            }
        }

        private void BtnEliminarSlot_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is ClipData slot)
            {
                slot.Guid_Clip = "";
                slot.Contenido_Plano = "";
                slot.Origen_App = "";
                GuardarSlotsActuales();
                SincronizarPinesHistorial();
            }
        }

        private void SincronizarPinesHistorial()
        {
            foreach (var clip in MisClips)
            {
                clip.HotKeyIndex = 0;
            }

            for (int i = 0; i < 9; i++)
            {
                if (!ClipsRecolector[i].EsVacio && !string.IsNullOrEmpty(ClipsRecolector[i].Guid_Clip))
                {
                    // MAGIA V2: Buscamos coincidencias con el ID exacto
                    var match = MisClips.FirstOrDefault(c => c.Guid_Clip == ClipsRecolector[i].Guid_Clip);
                    if (match != null) 
                    {
                        match.HotKeyIndex = i + 1; 
                    }
                }
            }

            ActualizarContadorTab(); 
        }
        
        #pragma warning disable CS0108 
        #pragma warning restore CS0108

        public new void DragOver(IDropInfo dropInfo)
        {
            if (dropInfo.Data is ClipData && dropInfo.TargetCollection == dropInfo.DragInfo.SourceCollection)
            {
                dropInfo.DropTargetAdorner = DropTargetAdorners.Insert;
                dropInfo.Effects = DragDropEffects.Move;
            }
        }

        public new void Drop(IDropInfo dropInfo)
        {
            GongSolutions.Wpf.DragDrop.DragDrop.DefaultDropHandler.Drop(dropInfo);

            for (int i = 0; i < ClipsRecolector.Count; i++)
            {
                ClipsRecolector[i].HotKeyIndex = i + 1;
            }

            GuardarSlotsActuales();
            ActualizarContadorTab();
            SincronizarPinesHistorial();
        }

        private void BlockResizer_DragDelta(object sender, System.Windows.Controls.Primitives.DragDeltaEventArgs e)
        {
            if (sender is Thumb thumb && thumb.DataContext is ClipData clip)
            {
                double nuevaAltura = clip.AlturaVisual + e.VerticalChange;
                
                if (nuevaAltura >= 38 && nuevaAltura < 800) 
                {
                    clip.AlturaVisual = nuevaAltura;
                }
            }
        }
    }
}