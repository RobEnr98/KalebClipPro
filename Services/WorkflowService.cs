using KalebClipPro.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace KalebClipPro.Services
{
    public class WorkflowService
    {
        private readonly string rutaArchivoGuardado = "KalebWorkflows.json";

        public WorkflowData CargarWorkflowDesdeDisco()
        {
            try
            {
                if (File.Exists(rutaArchivoGuardado))
                {
                    string json = File.ReadAllText(rutaArchivoGuardado);
                    var data = JsonSerializer.Deserialize<WorkflowData>(json);
                    if (data != null && data.Sets.Count > 0)
                        return data;
                }
            }
            catch { /* Si falla, creará uno nuevo abajo */ }

            // Si está vacío o es la primera vez, creamos los sets A, B y C por defecto
            var workflowVacio = new WorkflowData();
            workflowVacio.Sets.Add("A", CrearListaVaciaDe9Slots());
            workflowVacio.Sets.Add("B", CrearListaVaciaDe9Slots());
            workflowVacio.Sets.Add("C", CrearListaVaciaDe9Slots());
            GuardarWorkflowEnDisco(workflowVacio);
            
            return workflowVacio;
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

        public List<ClipData> CrearListaVaciaDe9Slots()
        {
            var lista = new List<ClipData>();
            for (int i = 1; i <= 9; i++)
            {
                lista.Add(new ClipData { HotKeyIndex = i, Contenido_Plano = "", Origen_App = "", AlturaVisual = 38 });
            }
            return lista;
        }
    }
}