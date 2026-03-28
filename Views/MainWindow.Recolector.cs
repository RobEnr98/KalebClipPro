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
        private void BtnAsignarHotkey_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is ClipData clip)
            {
                if (clip.HotKeyIndex > 0)
                {
                    // DESANCLAR: Si ya tiene un pin, lo quitamos y vaciamos su slot correspondiente
                    int index = clip.HotKeyIndex - 1;
                    if (index >= 0 && index < 9)
                    {
                        ClipsRecolector[index].Contenido_Plano = "";
                        ClipsRecolector[index].Origen_App = "";
                    }
                }
                else
                {
                    // ANCLAR (Automático): Buscamos el primer slot que esté vacío
                    var primerVacio = ClipsRecolector.FirstOrDefault(c => c.EsVacio);
                    
                    if (primerVacio != null)
                    {
                        primerVacio.Contenido_Plano = clip.Contenido_Plano;
                        primerVacio.Origen_App = clip.Origen_App;
                    }
                    else
                    {
                        // Si los 9 están llenos, aplicamos FIFO (Cascada) metiéndolo en el Slot 1
                        for (int i = 8; i > 0; i--) 
                        {
                            ClipsRecolector[i].Contenido_Plano = ClipsRecolector[i - 1].Contenido_Plano;
                            ClipsRecolector[i].Origen_App = ClipsRecolector[i - 1].Origen_App;
                        }
                        ClipsRecolector[0].Contenido_Plano = clip.Contenido_Plano;
                        ClipsRecolector[0].Origen_App = clip.Origen_App;
                    }
                }

                SincronizarPinesHistorial();
            }
        }

        private void AsignarSlotEspecifico_Click(object sender, RoutedEventArgs e)
        {
            // Usamos directamente el clip que guardamos al dar clic derecho
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
                            WorkflowActual.Sets[_setActual][oldIndex].Contenido_Plano = "";
                            WorkflowActual.Sets[_setActual][oldIndex].Origen_App = "";
                            ClipsRecolector[oldIndex].Contenido_Plano = "";
                            ClipsRecolector[oldIndex].Origen_App = "";
                        }
                    }
                }
                // 2. ASIGNAR A UN SET Y SLOT
                else
                {
                    var partes = tagValue.Split('-');
                    if (partes.Length == 2)
                    {
                        string targetSet = partes[0];
                        int targetSlot = int.Parse(partes[1]);
                        int newIndex = targetSlot - 1;

                        // Limpiamos el slot viejo si ya estaba anclado en otro lado
                        if (clip.HotKeyIndex > 0 && clip.HotKeyIndex != targetSlot)
                        {
                            int oldIndex = clip.HotKeyIndex - 1;
                            if (oldIndex >= 0 && oldIndex < 9)
                            {
                                WorkflowActual.Sets[_setActual][oldIndex].Contenido_Plano = "";
                                WorkflowActual.Sets[_setActual][oldIndex].Origen_App = "";
                                ClipsRecolector[oldIndex].Contenido_Plano = "";
                                ClipsRecolector[oldIndex].Origen_App = "";
                            }
                        }

                        // Asignamos el clip al destino
                        WorkflowActual.Sets[targetSet][newIndex].Contenido_Plano = clip.Contenido_Plano;
                        WorkflowActual.Sets[targetSet][newIndex].Origen_App = clip.Origen_App;

                        // Si el destino es el que estamos viendo en pantalla, actualizamos visualmente
                        if (targetSet == _setActual)
                        {
                            ClipsRecolector[newIndex].Contenido_Plano = clip.Contenido_Plano;
                            ClipsRecolector[newIndex].Origen_App = clip.Origen_App;
                        }
                    }
                }

                // Limpiamos la variable, guardamos y sincronizamos los pines visuales de la lista
                _clipSeleccionadoParaMenu = null;
                _workflowService.GuardarWorkflowEnDisco(WorkflowActual);
                SincronizarPinesHistorial();
            }
        }

        private void GuardarSlotsActuales()
        {
            try
            {
                if (WorkflowActual.Sets.ContainsKey(_setActual))
                {
                    // Sincronizamos la lista visual con el modelo de datos
                    WorkflowActual.Sets[_setActual] = ClipsRecolector.ToList();
                    // Guardamos usando el servicio
                    _workflowService.GuardarWorkflowEnDisco(WorkflowActual); 
                }
            }
            catch { }
        }

        private void InicializarSlotsVacios()
        {
            ClipsRecolector.Clear();

            if (!WorkflowActual.Sets.ContainsKey(_setActual))
            {
                WorkflowActual.Sets.Add(_setActual, _workflowService.CrearListaVaciaDe9Slots());
            }

            var slotsGuardados = WorkflowActual.Sets[_setActual];

            foreach (var slot in slotsGuardados)
            {
                slot.AlturaVisual = 38; 
                ClipsRecolector.Add(slot);
            }
            
            ActualizarContadorTab();
        }

        private void SincronizarBotonesConDatos()
        {
            if (ContenedorSets == null) return;
            ContenedorSets.Children.Clear();

            foreach (var nombreSet in WorkflowActual.Sets.Keys.OrderBy(k => k))
            {
                RadioButton nuevoSet = new RadioButton
                {
                    Content = $"Set {nombreSet}",
                    Tag = nombreSet,
                    Style = (Style)FindResource("SegmentedButtonStyle"),
                    IsChecked = (nombreSet == _setActual)
                };
                
                nuevoSet.Click += BtnCambiarSet_Click;
                ContenedorSets.Children.Add(nuevoSet);
            }
        }

        private void BtnCambiarSet_Click(object sender, RoutedEventArgs e)
        {
            if (sender is RadioButton btn && btn.Tag is string nuevoSet)
            {
                if (_setActual == nuevoSet) return;

                GuardarSlotsActuales();
                _setActual = nuevoSet;
                InicializarSlotsVacios();
                SincronizarPinesHistorial(); 
            }
        }

        private void BtnQuitarSet_Click(object sender, RoutedEventArgs e)
        {
            int cantidadActual = WorkflowActual.Sets.Count;

            if (cantidadActual <= 2) return; 

            string idSetABorrar = ((char)('A' + cantidadActual - 1)).ToString();

            if (WorkflowActual.Sets.ContainsKey(idSetABorrar))
            {
                bool tieneDatos = WorkflowActual.Sets[idSetABorrar].Any(c => !c.EsVacio);
                if (tieneDatos)
                {
                    var respuesta = MessageBox.Show($"El Set {idSetABorrar} contiene clips guardados. ¿Estás seguro de que quieres eliminarlo?", "Confirmar", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                    if (respuesta == MessageBoxResult.No) return;
                }

                WorkflowActual.Sets.Remove(idSetABorrar);
                _workflowService.GuardarWorkflowEnDisco(WorkflowActual);

                if (_setActual == idSetABorrar)
                {
                    _setActual = ((char)(idSetABorrar[0] - 1)).ToString();
                }

                SincronizarBotonesConDatos();
                InicializarSlotsVacios();
            }
        }

        private void ActualizarCheckVisualSets()
        {
            foreach (var child in ContenedorSets.Children)
            {
                if (child is RadioButton btn && btn.Tag?.ToString() == _setActual)
                {
                    btn.IsChecked = true;
                    break; 
                }
            }
        }

        private void MenuAsignarSlot_Opened(object sender, RoutedEventArgs e)
        {
            if (sender is ContextMenu menu && WorkflowActual != null && WorkflowActual.Sets != null)
            {
                menu.Items.Clear();

                Style itemStyle = (Style)FindResource("MenuOscuroItemStyle");
                Style separatorStyle = (Style)FindResource("MenuOscuroSeparatorStyle");

                foreach (var set in WorkflowActual.Sets)
                {
                    string letraSet = set.Key;
                    MenuItem setMenuItem = new MenuItem { Header = $"Set {letraSet}", Style = itemStyle };

                    for (int i = 1; i <= 9; i++)
                    {
                        MenuItem slotMenuItem = new MenuItem 
                        { 
                            Header = $"Slot {i}",
                            Tag = $"{letraSet}-{i}",
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
            if (TabRecolector != null)
                TabRecolector.ToolTip = $"Recolector {_setActual} ({llenos} slots ocupados)"; // <--- Adaptado para el nuevo Sidebar
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
            int cantidadActual = WorkflowActual.Sets.Count;

            if (cantidadActual >= 26)
            {
                MessageBox.Show("¡Límite máximo de Sets (Z) alcanzado!", "Aviso", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            char siguienteLetra = (char)('A' + cantidadActual);
            string idNuevoSet = siguienteLetra.ToString();

            if (!WorkflowActual.Sets.ContainsKey(idNuevoSet))
            {
                WorkflowActual.Sets.Add(idNuevoSet, _workflowService.CrearListaVaciaDe9Slots());
                _workflowService.GuardarWorkflowEnDisco(WorkflowActual);
            }

            SincronizarBotonesConDatos();

            foreach (var child in ContenedorSets.Children)
            {
                if (child is RadioButton rb && rb.Tag.ToString() == idNuevoSet)
                {
                    rb.IsChecked = true;
                    GuardarSlotsActuales();
                    _setActual = idNuevoSet;
                    InicializarSlotsVacios();
                    break;
                }
            }
        }

        private void BtnLimpiarRecolector_Click(object sender, RoutedEventArgs e)
        {
            // Ahora no eliminamos la lista, solo apagamos las celdas borrando su contenido
            foreach (var slot in ClipsRecolector)
            {
                slot.Contenido_Plano = "";
                slot.Origen_App = "";
            }
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
            // Este es el botón de la papelera DENTRO de la pestaña del recolector
            if (sender is Button btn && btn.DataContext is ClipData slot)
            {
                slot.Contenido_Plano = "";
                slot.Origen_App = "";
                SincronizarPinesHistorial();
            }
        }

        private void SincronizarPinesHistorial()
        {
            // 1. Quitamos todos los pines amarillos visuales del historial
            foreach (var clip in MisClips)
            {
                clip.HotKeyIndex = 0;
            }

            // 2. Leemos los 9 slots del recolector y buscamos si esos textos están en el historial
            for (int i = 0; i < 9; i++)
            {
                if (!ClipsRecolector[i].EsVacio)
                {
                    var match = MisClips.FirstOrDefault(c => c.Contenido_Plano == ClipsRecolector[i].Contenido_Plano);
                    if (match != null) 
                    {
                        match.HotKeyIndex = i + 1; // Encendemos el pin dorado
                    }
                }
            }

            ActualizarContadorTab(); // Actualizamos el título "Recolector (X)"
        }
        
        #pragma warning disable CS0108 // Esto apaga la advertencia amarilla mágicamente
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

            ActualizarContadorTab();
        }

        private void BlockResizer_DragDelta(object sender, System.Windows.Controls.Primitives.DragDeltaEventArgs e)
        {
            // Ahora redimensionamos el DATO, no el elemento visual
            if (sender is Thumb thumb && thumb.DataContext is ClipData clip)
            {
                double nuevaAltura = clip.AlturaVisual + e.VerticalChange;
                
                // Mínimo 38px (compacto), máximo 800px (pantalla completa)
                if (nuevaAltura >= 38 && nuevaAltura < 800) 
                {
                    clip.AlturaVisual = nuevaAltura;
                }
            }
        }
    }
}