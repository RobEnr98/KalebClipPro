using System;
using System.ComponentModel;
using KalebClipPro.Models;
using System.Runtime.CompilerServices;

namespace KalebClipPro.Models
{
    public class ClipData : INotifyPropertyChanged
    {
        // Separador interno para dividir Texto Plano de Formato Rico (XAML)
        private const string RTF_SEPARATOR = "_||KALEB_RTF||_";

        private string _contenidoPlano = string.Empty;
        private int _hotKeyIndex = 0; 
        private int _workspaceId = 1; 
        private string _origenApp = string.Empty;
        private double _alturaVisual = 38; 

        public string Guid_Clip { get; set; } = string.Empty; 

        public double AlturaVisual
        {
            get => _alturaVisual;
            set
            {
                if (_alturaVisual != value)
                {
                    _alturaVisual = value;
                    OnPropertyChanged();
                }
            }
        }

        public string Contenido_Plano
        {
            get 
            {
                if (string.IsNullOrEmpty(_contenidoPlano)) return "";

                // 1. Si es formato NUEVO (paquete dual), sacamos la parte limpia
                if (_contenidoPlano.Contains("_||KALEB_RTF||_"))
                {
                    return _contenidoPlano.Split(new string[] { "_||KALEB_RTF||_" }, StringSplitOptions.None)[0];
                }
                
                // 2. Si es un registro VIEJO (XAML), borramos las etiquetas XML para que devuelva solo el texto real
                if (_contenidoPlano.TrimStart().StartsWith("<Section xmlns="))
                {
                    return System.Text.RegularExpressions.Regex.Replace(_contenidoPlano, "<.*?>", string.Empty);
                }

                // 3. Texto plano normal
                return _contenidoPlano;
            }
            set
            {
                if (_contenidoPlano != value)
                {
                    _contenidoPlano = value;
                    OnPropertyChanged(); 
                    OnPropertyChanged(nameof(EsVacio));
                }
            }
        }

        // --- 📦 MÉTODO PARA EL EDITOR ---
        // Este método devuelve el string completo (Texto + XAML) para que el editor
        // pueda reconstruir las negritas, colores, etc.
        public string ObtenerContenidoCrudo()
        {
            return _contenidoPlano;
        }

        public int HotKeyIndex
        {
            get => _hotKeyIndex;
            set
            {
                if (_hotKeyIndex != value)
                {
                    _hotKeyIndex = value;
                    OnPropertyChanged(); 
                }
            }
        }

        public int WorkspaceID
        {
            get => _workspaceId;
            set
            {
                if (_workspaceId != value)
                {
                    _workspaceId = value;
                    OnPropertyChanged();
                }
            }
        }

        public string Origen_App 
        { 
            get => _origenApp;
            set 
            {
                if (_origenApp != value)
                {
                    _origenApp = value;
                    OnPropertyChanged();
                }
            }
        }

        public DateTime Fecha_Creacion { get; set; }

        public bool EsVacio => string.IsNullOrWhiteSpace(Contenido_Plano);

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}