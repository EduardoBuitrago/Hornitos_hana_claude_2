using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Lotes_Hornitos
{
    internal static class Program
    {
        // Se guarda en una variable para que el recolector de basura no lo elimine.
        private static System.Threading.Timer latido;
        private static System.Threading.Timer vigilante;
        private static int vigilanteFallos = 0;

        /// <summary>
        /// Cuenta los procesos de SAP Business One (cliente) que corren en ESTA sesion de Windows.
        /// Devuelve -1 si no se pudo consultar.
        /// </summary>
        private static int ContarProcesosSapEnMiSesion()
        {
            try
            {
                int miSesion = Process.GetCurrentProcess().SessionId;
                int n = 0;
                foreach (var p in Process.GetProcesses())
                {
                    try
                    {
                        if (p.SessionId == miSesion && p.ProcessName.StartsWith("SAP Business One", StringComparison.OrdinalIgnoreCase)) n++;
                    }
                    catch { }
                    finally { p.Dispose(); }
                }
                return n;
            }
            catch { return -1; }
        }

        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // DIAGNOSTICO: si el addon se cae por un error no controlado, antes se
            // cerraba sin dejar ninguna linea en el log. Ahora queda registrado.
            try
            {
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                Application.ThreadException += (s, e) =>
                {
                    Classes.Ini.RegistrarError("ERROR NO CONTROLADO (hilo): " + e.Exception.Message + " - " + e.Exception.StackTrace);
                };
                AppDomain.CurrentDomain.UnhandledException += (s, e) =>
                {
                    var ex = e.ExceptionObject as Exception;
                    Classes.Ini.RegistrarError("ERROR NO CONTROLADO (el addon se cierra: " + e.IsTerminating + "): " + (ex != null ? ex.Message + " - " + ex.StackTrace : e.ExceptionObject));
                };
                AppDomain.CurrentDomain.ProcessExit += (s, e) =>
                {
                    Classes.Ini.RegistrarError("El proceso del addon termino.");
                };
            }
            catch { }

            if (Classes.Ini.Run())
            {
                // Una linea cada 10 minutos: si en el log se corta, ahi se cayo el addon.
                try
                {
                    var asm = System.Reflection.Assembly.GetExecutingAssembly();
                    Classes.Ini.RegistrarError("Addon iniciado. Version " + asm.GetName().Version + ", archivo compilado el " + System.IO.File.GetLastWriteTime(asm.Location).ToString("yyyy-MM-dd HH:mm:ss") + ", en " + asm.Location);
                }
                catch { Classes.Ini.RegistrarError("Addon iniciado."); }
                latido = new System.Threading.Timer(_ => Classes.Ini.RegistrarError("[latido] addon activo"), null, TimeSpan.FromMinutes(10), TimeSpan.FromMinutes(10));

                // VIGILANTE (procesos que se quedaban vivos): si SAP se cierra a la fuerza (por ejemplo,
                // cerrando la sesion de escritorio remoto) SAP no le avisa al addon y el proceso queda
                // corriendo, acumulandose y bloqueando la instalacion de versiones nuevas.
                // Cada minuto se revisa si sigue habiendo un cliente de SAP en ESTA sesion de Windows
                // (no se miran las de otros usuarios). Si no hay ninguno 3 veces seguidas, el addon se cierra.
                // Solo se activa si al iniciar SI se encontro el proceso de SAP (asi, si el nombre del
                // proceso fuera distinto en otra version, el vigilante no cierra el addon por error).
                int sapAlIniciar = ContarProcesosSapEnMiSesion();
                if (sapAlIniciar > 0)
                {
                    Classes.Ini.RegistrarError("Vigilante activo: procesos de SAP en esta sesion = " + sapAlIniciar);
                    vigilante = new System.Threading.Timer(_ =>
                    {
                        int n = ContarProcesosSapEnMiSesion();
                        if (n < 0) return; // no se pudo consultar: no se hace nada
                        if (n > 0) { vigilanteFallos = 0; return; }
                        vigilanteFallos++;
                        if (vigilanteFallos >= 3)
                        {
                            Classes.Ini.RegistrarError("Vigilante: ya no hay un cliente de SAP en esta sesion; el addon se cierra.");
                            Environment.Exit(0);
                        }
                    }, null, TimeSpan.FromMinutes(1), TimeSpan.FromMinutes(1));
                }
                else
                {
                    Classes.Ini.RegistrarError("Vigilante NO activo: no se encontro el proceso de SAP en esta sesion (resultado " + sapAlIniciar + ").");
                }
                Application.Run();
            }
            else
            {
                Application.Exit();
            }
        }
    }
}
