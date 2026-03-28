using System.Collections.Generic;

namespace KalebClipPro.Models
{
    public class WorkflowData
    {
        public string NombreWorkflow { get; set; } = "Mi Workspace";
        // Diccionario central. Clave: Letra del Set ("A", "B"). Valor: Lista de 9 Clips
        public Dictionary<string, List<ClipData>> Sets { get; set; } = new Dictionary<string, List<ClipData>>();
    }
}