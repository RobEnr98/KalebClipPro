using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace KalebClipPro.Models
{
    public class WorkflowData
    {
        public string NombreWorkflow { get; set; } = "Mi Workspace";
        public List<WorkflowFolder> Carpetas { get; set; } = new List<WorkflowFolder>();
    }

    public class WorkflowFolder
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Nombre { get; set; } = "Nueva Carpeta";
        public List<SlotData> Slots { get; set; } = new List<SlotData>();
    }

    public class SlotData
    {
        public int HotKeyIndex { get; set; }
        
        public string ClipIdAsignado { get; set; } = ""; 

        [JsonIgnore]
        public ClipData? ClipCargado { get; set; } 
    }
}