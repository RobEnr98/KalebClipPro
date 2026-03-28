using System;


namespace KalebClipPro.Models
{
    public class Clip
    {
        public string ID { get; set; } = Guid.NewGuid().ToString();
        // Inicializamos con string.Empty para evitar el warning CS8618
        public string Texto { get; set; } = string.Empty; 
        public int Tipo { get; set; } // 0: Texto
        public DateTime Fecha { get; set; } = DateTime.Now;
    }
}