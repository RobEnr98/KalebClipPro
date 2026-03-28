using System.Collections.Generic;

namespace KalebClipPro.Services
{
    public class WorkspaceModel
    {
        public int ID_Workspace { get; set; }
        public string Nombre_Workspace { get; set; } = string.Empty;
        public string Icono { get; set; } = string.Empty;
    }

    public class WorkspaceService
    {
        // Estado actual de la aplicación
        public int WorkspaceActualID { get; private set; } = 1;
        public bool ModoLlenadoEnSerieActivo { get; set; } = false;
        
        // Puntero para saber qué número toca asignar (del 1 al 9)
        private int _siguienteSlotLibre = 1;

        public void CambiarWorkspace(int nuevoWorkspaceId)
        {
            WorkspaceActualID = nuevoWorkspaceId;
            // Al cambiar de contexto, reiniciamos el contador de slots
            _siguienteSlotLibre = 1; 
        }

        public void AlternarModoLlenadoEnSerie()
        {
            ModoLlenadoEnSerieActivo = !ModoLlenadoEnSerieActivo;
            if (!ModoLlenadoEnSerieActivo)
            {
                _siguienteSlotLibre = 1; // Reseteamos al apagarlo
            }
        }

        // Esta es la magia: decide qué atajo asignar automáticamente cuando haces Ctrl+C
        public int ObtenerSiguienteSlotYAvanzar()
        {
            if (!ModoLlenadoEnSerieActivo) 
                return 0; // 0 significa que no se asigna atajo (copiado normal)

            int slotAsignado = _siguienteSlotLibre;
            
            _siguienteSlotLibre++;
            if (_siguienteSlotLibre > 9) 
            {
                _siguienteSlotLibre = 1; // Si llega a 9, vuelve al 1 (ciclo continuo)
            }

            return slotAsignado;
        }
    }
}