using KalebClipPro.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace KalebClipPro.Services
{
    public class WorkflowService
    {
        // Usamos un nuevo archivo v2 para evitar choques con el formato viejo
        private readonly string rutaArchivoGuardado = "KalebWorkflows_v2.json"; 

        public WorkflowData CargarWorkflowDesdeDisco()
        {
            try
            {
                if (File.Exists(rutaArchivoGuardado))
                {
                    string json = File.ReadAllText(rutaArchivoGuardado);
                    var data = JsonSerializer.Deserialize<WorkflowData>(json);
                    
                    if (data != null && data.Carpetas.Count > 0)
                        return data;
                }
            }
            catch { /* Si falla o el JSON está corrupto, creará uno nuevo abajo */ }

            // Si es la primera vez, creamos una carpeta "General" por defecto
            var workflowPorDefecto = new WorkflowData();
            workflowPorDefecto.Carpetas.Add(CrearCarpetaVacia("General"));
            GuardarWorkflowEnDisco(workflowPorDefecto);
            
            return workflowPorDefecto;
        }

        public void GuardarWorkflowEnDisco(WorkflowData workflow)
        {
            try
            {
                var opciones = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(workflow, opciones);
                File.WriteAllText(rutaArchivoGuardado, json);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error al guardar Workflows: {ex.Message}");
            }
        }

        public WorkflowFolder CrearCarpetaVacia(string nombre)
        {
            var carpeta = new WorkflowFolder { Nombre = nombre };
            
            for (int i = 1; i <= 9; i++)
            {
                // Ahora los slots nacen vacíos y listos para recibir un Guid_Clip
                carpeta.Slots.Add(new SlotData { HotKeyIndex = i, ClipIdAsignado = "" });
            }
            
            return carpeta;
        }
    }
}