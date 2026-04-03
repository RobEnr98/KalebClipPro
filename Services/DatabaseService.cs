using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using KalebClipPro.Models;

namespace KalebClipPro.Services
{
    public class DatabaseService
    {
        private string _conn = "Data Source=clips.db";

        public void InicializarBaseDeDatos()
        {
            using (var c = new SqliteConnection(_conn))
            {
                c.Open();
                var cmd = c.CreateCommand();
                
                // Tabla de Versiones (Time Travel)
                cmd.CommandText = @"
                    CREATE TABLE IF NOT EXISTS Historial_Versiones (
                        Id_Version INTEGER PRIMARY KEY AUTOINCREMENT,
                        Id_Clip TEXT, 
                        Contenido_Version TEXT,
                        Fecha_Version DATETIME
                    )";
                cmd.ExecuteNonQuery();

                // Tabla principal de Clips
                cmd.CommandText = @"
                    CREATE TABLE IF NOT EXISTS Clips_Historial (
                        ID_Clip TEXT PRIMARY KEY,
                        Contenido_Plano TEXT,
                        Tipo_Clip INTEGER,
                        Origen_App TEXT,
                        Fecha_Creacion DATETIME
                    )";
                cmd.ExecuteNonQuery();

                // --- NUEVO: Tabla de Marcos de Trabajo (Workspaces) ---
                cmd.CommandText = @"
                    CREATE TABLE IF NOT EXISTS Workspaces (
                        ID_Workspace INTEGER PRIMARY KEY AUTOINCREMENT,
                        Nombre_Workspace TEXT,
                        Icono TEXT
                    )";
                cmd.ExecuteNonQuery();

                // Insertar el marco 'General' por defecto si no existe
                cmd.CommandText = "INSERT OR IGNORE INTO Workspaces (ID_Workspace, Nombre_Workspace, Icono) VALUES (1, 'General', '📁')";
                cmd.ExecuteNonQuery();

                // Inyectar la columna de Hotkeys
                try {
                    cmd.CommandText = "ALTER TABLE Clips_Historial ADD COLUMN HotKeyIndex INTEGER DEFAULT 0";
                    cmd.ExecuteNonQuery();
                } catch { }

                // --- NUEVO: Inyectar la columna de WorkspaceID ---
                try {
                    cmd.CommandText = "ALTER TABLE Clips_Historial ADD COLUMN WorkspaceID INTEGER DEFAULT 1";
                    cmd.ExecuteNonQuery();
                } catch { }
            }
        }

        // --- MÉTODO PARA CARGA INFINITA NORMAL (Ahora filtra por Workspace) ---
        public List<ClipData> ObtenerHistorialPaginado(int pagina, int cantidad, int workspaceId = 1)
        {
            var lista = new List<ClipData>();
            int saltar = pagina * cantidad;

            using (var c = new SqliteConnection(_conn))
            {
                c.Open();
                var cmd = c.CreateCommand();
                
                cmd.CommandText = @"
                    SELECT Contenido_Plano, Origen_App, Fecha_Creacion, ID_Clip, HotKeyIndex, WorkspaceID
                    FROM Clips_Historial 
                    WHERE WorkspaceID = $wsId
                    ORDER BY Fecha_Creacion DESC 
                    LIMIT $limit OFFSET $offset";

                cmd.Parameters.AddWithValue("$wsId", workspaceId);
                cmd.Parameters.AddWithValue("$limit", cantidad);
                cmd.Parameters.AddWithValue("$offset", saltar);

                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        lista.Add(new ClipData {
                            Contenido_Plano = r.GetString(0),
                            Origen_App = r.IsDBNull(1) ? "App" : r.GetString(1),
                            Fecha_Creacion = r.GetDateTime(2),
                            Guid_Clip = r.GetString(3),
                            HotKeyIndex = r.IsDBNull(4) ? 0 : r.GetInt32(4),
                            WorkspaceID = r.IsDBNull(5) ? 1 : r.GetInt32(5)
                        });
                    }
                }
            }
            return lista;
        }

        // --- MÉTODO PARA BÚSQUEDA PROFUNDA EN DB (Ahora filtra por Workspace) ---
        public List<ClipData> BuscarHistorialPaginado(string textoBusqueda, int pagina, int cantidad, int workspaceId = 1)
        {
            var lista = new List<ClipData>();
            int saltar = pagina * cantidad;

            using (var c = new SqliteConnection(_conn))
            {
                c.Open();
                var cmd = c.CreateCommand();
                
                cmd.CommandText = @"
                    SELECT Contenido_Plano, Origen_App, Fecha_Creacion, ID_Clip, HotKeyIndex, WorkspaceID
                    FROM Clips_Historial 
                    WHERE WorkspaceID = $wsId AND Contenido_Plano LIKE $filtro 
                    ORDER BY Fecha_Creacion DESC 
                    LIMIT $limit OFFSET $offset";

                cmd.Parameters.AddWithValue("$wsId", workspaceId);
                cmd.Parameters.AddWithValue("$filtro", "%" + textoBusqueda + "%");
                cmd.Parameters.AddWithValue("$limit", cantidad);
                cmd.Parameters.AddWithValue("$offset", saltar);

                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        lista.Add(new ClipData {
                            Contenido_Plano = r.GetString(0),
                            Origen_App = r.IsDBNull(1) ? "App" : r.GetString(1),
                            Fecha_Creacion = r.GetDateTime(2),
                            Guid_Clip = r.GetString(3),
                            HotKeyIndex = r.IsDBNull(4) ? 0 : r.GetInt32(4),
                            WorkspaceID = r.IsDBNull(5) ? 1 : r.GetInt32(5)
                        });
                    }
                }
            }
            return lista;
        }

        // --- MÉTODO PARA GUARDAR (Ahora acepta Workspace y Hotkey directos) ---
        public string GuardarClipConOrigen(string texto, string origen, int hotkeyIndex = 0, int workspaceId = 1)
        {
            string nuevoId = Guid.NewGuid().ToString();
            using (var c = new SqliteConnection(_conn))
            {
                c.Open();
                var cmd = c.CreateCommand();
                
                cmd.CommandText = @"INSERT INTO Clips_Historial (ID_Clip, Contenido_Plano, Tipo_Clip, Origen_App, Fecha_Creacion, HotKeyIndex, WorkspaceID) 
                                    VALUES ($id, $txt, $tipo, $ori, $fec, $hk, $ws)";
                cmd.Parameters.AddWithValue("$id", nuevoId);
                cmd.Parameters.AddWithValue("$txt", texto);
                cmd.Parameters.AddWithValue("$tipo", 0);
                cmd.Parameters.AddWithValue("$ori", origen);
                cmd.Parameters.AddWithValue("$fec", DateTime.Now);
                cmd.Parameters.AddWithValue("$hk", hotkeyIndex);
                cmd.Parameters.AddWithValue("$ws", workspaceId);
                cmd.ExecuteNonQuery();
            }
            return nuevoId;
        }

        public List<ClipData> ObtenerHistorial()
        {
            return ObtenerHistorialPaginado(0, 30); // Por defecto lee el Workspace 1
        }

        public void ActualizarClip(string guidClip, string nuevoTexto)
        {
            using (var conexion = new SqliteConnection(_conn))
            {
                conexion.Open();
                var comando = conexion.CreateCommand();
                comando.CommandText = "UPDATE Clips_Historial SET Contenido_Plano = $txt WHERE ID_Clip = $id";
                comando.Parameters.AddWithValue("$txt", nuevoTexto);
                comando.Parameters.AddWithValue("$id", guidClip);
                comando.ExecuteNonQuery();
            }
        }

        public void GuardarVersion(string guidClip, string texto)
        {
            using (var conexion = new SqliteConnection(_conn))
            {
                conexion.Open();
                var comando = conexion.CreateCommand();
                comando.CommandText = "INSERT INTO Historial_Versiones (Id_Clip, Contenido_Version, Fecha_Version) VALUES ($id, $txt, $fec)";
                comando.Parameters.AddWithValue("$id", guidClip);
                comando.Parameters.AddWithValue("$txt", texto);
                comando.Parameters.AddWithValue("$fec", DateTime.Now);
                comando.ExecuteNonQuery();
            }
        }

        public string? ObtenerUltimaVersion(string guidClip)
        {
            using (var conexion = new SqliteConnection(_conn))
            {
                conexion.Open();
                var comando = conexion.CreateCommand();
                comando.CommandText = "SELECT Contenido_Version FROM Historial_Versiones WHERE Id_Clip = $id ORDER BY Fecha_Version DESC LIMIT 1";
                comando.Parameters.AddWithValue("$id", guidClip);
                
                var resultado = comando.ExecuteScalar();
                return resultado?.ToString();
            }
        }

        public List<string> ObtenerTodasLasVersiones(string guidClip)
        {
            var lista = new List<string>();
            using (var conexion = new SqliteConnection(_conn))
            {
                conexion.Open();
                var comando = conexion.CreateCommand();
                comando.CommandText = "SELECT Contenido_Version FROM Historial_Versiones WHERE Id_Clip = $id ORDER BY Fecha_Version ASC";
                comando.Parameters.AddWithValue("$id", guidClip);
                
                using (var lector = comando.ExecuteReader())
                {
                    while (lector.Read())
                    {
                        lista.Add(lector.GetString(0));
                    }
                }
            }
            return lista;
        }

        // =========================================================================
        // SISTEMA DE HOTKEYS (Aislado por Marco de Trabajo)
        // =========================================================================
        public void AsignarHotkey(string guidClip, int hotkeyIndex, int workspaceId = 1)
        {
            using (var c = new SqliteConnection(_conn))
            {
                c.Open();
                
                // 1. Nos aseguramos de limpiar el atajo, pero SOLO dentro del mismo marco de trabajo.
                if (hotkeyIndex > 0)
                {
                    var cmdClear = c.CreateCommand();
                    cmdClear.CommandText = "UPDATE Clips_Historial SET HotKeyIndex = 0 WHERE HotKeyIndex = $index AND WorkspaceID = $wsId";
                    cmdClear.Parameters.AddWithValue("$index", hotkeyIndex);
                    cmdClear.Parameters.AddWithValue("$wsId", workspaceId);
                    cmdClear.ExecuteNonQuery();
                }
                
                // 2. Le asignamos el atajo al clip deseado.
                var cmdSet = c.CreateCommand();
                cmdSet.CommandText = "UPDATE Clips_Historial SET HotKeyIndex = $index WHERE ID_Clip = $id";
                cmdSet.Parameters.AddWithValue("$index", hotkeyIndex);
                cmdSet.Parameters.AddWithValue("$id", guidClip);
                cmdSet.ExecuteNonQuery();
            }
        }

        public void EliminarClip(string guidClip)
        {
            try
            {
                using (var connection = new SqliteConnection(_conn)) 
                {
                    connection.Open();
                    
                    var cmdVersiones = connection.CreateCommand();
                    cmdVersiones.CommandText = "DELETE FROM Historial_Versiones WHERE Id_Clip = $id";
                    cmdVersiones.Parameters.AddWithValue("$id", guidClip);
                    cmdVersiones.ExecuteNonQuery();

                    var cmdClip = connection.CreateCommand();
                    cmdClip.CommandText = "DELETE FROM Clips_Historial WHERE ID_Clip = $id";
                    cmdClip.Parameters.AddWithValue("$id", guidClip);
                    cmdClip.ExecuteNonQuery();
                }
            }
            catch { /* Manejo de errores silencioso */ }
        }

        // =========================================================================
        // NUEVO: OBTENER CLIP POR ID (Para los slots del Workflow V2)
        // =========================================================================
        public ClipData? ObtenerClipPorId(string idClip)
        {
            using (var c = new SqliteConnection(_conn))
            {
                c.Open();
                var cmd = c.CreateCommand();
                
                cmd.CommandText = @"
                    SELECT Contenido_Plano, Origen_App, Fecha_Creacion, ID_Clip, HotKeyIndex, WorkspaceID
                    FROM Clips_Historial 
                    WHERE ID_Clip = $id LIMIT 1";

                cmd.Parameters.AddWithValue("$id", idClip);

                using (var r = cmd.ExecuteReader())
                {
                    if (r.Read())
                    {
                        return new ClipData {
                            Contenido_Plano = r.GetString(0),
                            Origen_App = r.IsDBNull(1) ? "App" : r.GetString(1),
                            Fecha_Creacion = r.GetDateTime(2),
                            Guid_Clip = r.GetString(3),
                            HotKeyIndex = r.IsDBNull(4) ? 0 : r.GetInt32(4),
                            WorkspaceID = r.IsDBNull(5) ? 1 : r.GetInt32(5)
                        };
                    }
                }
            }
            return null; // Retorna null si ese ID ya no existe en la BD
        }
    }
}