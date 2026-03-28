using System.Collections.Generic;

namespace KalebClipPro.Models
{
    public class HistorialSesion
    {
        public List<string> Versiones { get; set; } = new List<string>();
        public int IndiceActual { get; set; } = 0;
    }
}