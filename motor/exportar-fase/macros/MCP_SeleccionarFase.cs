using System; using System.Collections; using System.Collections.Generic; using System.IO;
using Tekla.Structures.Model;

// Selecciona en el modelo TODO lo de una fase: partes y bulones (Skill exportar-fase, Ferma).
// Sirve antes de sacar reportes ("Partes seleccionadas"): sin los bulones, el reporte de bulones sale vacío.
// No modifica el modelo: solo cambia la selección.
// Parámetros en %USERPROFILE%\TeklaMCP_logs\export_params.txt: MODELO=..., FASE=...
// Log: %USERPROFILE%\TeklaMCP_logs\seleccionar_fase_log.txt

//Keep: Tekla.Technology.Akit.UserScript
namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        static readonly string DIR_LOG = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "TeklaMCP_logs");

        public static void Run(IScript akit)
        {
            var log = new List<string>();
            log.Add("INICIO=" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            try
            {
                Dictionary<string, string> p = LeerParametros(Path.Combine(DIR_LOG, "export_params.txt"));
                Model model = new Model();
                if (!model.GetConnectionStatus()) { log.Add("ERROR: sin conexion con Tekla"); return; }
                string nombre = Path.GetFileNameWithoutExtension(model.GetInfo().ModelName);
                log.Add("MODELO=" + nombre);
                if (!p.ContainsKey("MODELO") || !string.Equals(p["MODELO"].Trim(), nombre, StringComparison.OrdinalIgnoreCase))
                {
                    log.Add("ERROR: el modelo abierto no es el esperado. No se hace nada.");
                    return;
                }
                if (!ParametrosRecientes(p, log)) return;
                int fase;
                if (!p.ContainsKey("FASE") || !int.TryParse(p["FASE"].Trim(), out fase)) { log.Add("ERROR: FASE invalida"); return; }

                var sel = new ArrayList();
                int nPartes = 0, nBulones = 0;
                ModelObjectEnumerator en = model.GetModelObjectSelector().GetAllObjects();
                en.SelectInstances = true;
                while (en.MoveNext())
                {
                    ModelObject o = en.Current;
                    bool esParte = o is Part, esBulon = o is BoltGroup;
                    if (!esParte && !esBulon) continue;
                    Phase ph;
                    if (!o.GetPhase(out ph) || ph.PhaseNumber != fase) continue;
                    sel.Add(o);
                    if (esParte) nPartes++; else nBulones++;
                }
                log.Add("FASE=" + fase);
                log.Add("PARTES=" + nPartes);
                log.Add("GRUPOS_BULONES=" + nBulones);
                if (sel.Count == 0) { log.Add("ERROR: nada en la fase"); return; }
                new Tekla.Structures.Model.UI.ModelObjectSelector().Select(sel);
                log.Add("SELECCIONADOS=" + sel.Count);
                log.Add("FIN=OK");
            }
            catch (Exception ex) { log.Add("EXCEPCION: " + ex.Message + " " + ex.StackTrace); }
            finally
            {
                Directory.CreateDirectory(DIR_LOG);
                File.WriteAllLines(Path.Combine(DIR_LOG, "seleccionar_fase_log.txt"), log.ToArray());
            }
        }

        // Los parámetros los escribe "exportar_fase.py preparar --crear" con CREADO=fecha. Si son viejos, no se usan:
        // así nunca se exporta la fase anterior por accidente (2026-11-10).
        static bool ParametrosRecientes(Dictionary<string, string> p, List<string> log)
        {
            DateTime c;
            if (!p.ContainsKey("CREADO") || !DateTime.TryParseExact(p["CREADO"].Trim(), "yyyy-MM-dd HH:mm:ss",
                    System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out c))
            { log.Add("ERROR: export_params.txt sin CREADO. Correr de nuevo 'preparar --crear'. No se hace nada."); return false; }
            if ((DateTime.Now - c).TotalMinutes > 120)
            { log.Add("ERROR: export_params.txt es viejo (" + p["CREADO"].Trim() + "). Correr de nuevo 'preparar --crear'. No se hace nada."); return false; }
            return true;
        }

        static Dictionary<string, string> LeerParametros(string ruta)
        {
            var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (!File.Exists(ruta)) return d;
            foreach (string linea in File.ReadAllLines(ruta))
            {
                int i = linea.IndexOf('=');
                if (i > 0) d[linea.Substring(0, i).Trim()] = linea.Substring(i + 1);
            }
            return d;
        }
    }
}
