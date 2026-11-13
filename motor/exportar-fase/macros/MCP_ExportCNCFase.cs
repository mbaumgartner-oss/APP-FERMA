using System; using System.Collections; using System.Collections.Generic; using System.IO;
using Tekla.Structures.Model; using Tekla.Structures.Model.Operations;

// Exporta los CNC (.nc1) de TODAS las partes de una fase con la config "CNC FERMA" (Skill exportar-fase, Ferma).
// Equivale a: seleccionar la fase -> Archivos CN -> CNC FERMA -> Partes seleccionadas -> Crear, con otra carpeta.
// No modifica el modelo: solo cambia la selección y escribe archivos en DESTINO.
// Parámetros en %USERPROFILE%\TeklaMCP_logs\export_params.txt (CLAVE=valor, uno por línea):
//   MODELO=<nombre exacto del modelo abierto, sin .db1>   -> si no coincide, NO hace nada
//   FASE=102
//   DESTINO=<carpeta existente y sin .nc1 adentro>        -> nunca pisa archivos
//   CONFIG_CNC=Todo   (opcional) = la LINEA .ncfs marcada dentro de "CNC FERMA.ncf"; con "CNC FERMA" falla
// Log: %USERPROFILE%\TeklaMCP_logs\export_cnc_log.txt

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
                ModelInfo mi = model.GetInfo();
                string nombre = Path.GetFileNameWithoutExtension(mi.ModelName);
                log.Add("MODELO=" + nombre);
                log.Add("RUTA=" + mi.ModelPath);
                if (!p.ContainsKey("MODELO") || !string.Equals(p["MODELO"].Trim(), nombre, StringComparison.OrdinalIgnoreCase))
                {
                    log.Add("ERROR: el modelo abierto no es el esperado (" + (p.ContainsKey("MODELO") ? p["MODELO"] : "sin MODELO") + "). No se hace nada.");
                    return;
                }
                if (!ParametrosRecientes(p, log)) return;
                int fase;
                if (!p.ContainsKey("FASE") || !int.TryParse(p["FASE"].Trim(), out fase)) { log.Add("ERROR: FASE invalida"); return; }
                if (!p.ContainsKey("DESTINO")) { log.Add("ERROR: falta DESTINO"); return; }
                string destino = p["DESTINO"].Trim();
                if (!destino.EndsWith("\\")) destino += "\\";
                string config = p.ContainsKey("CONFIG_CNC") ? p["CONFIG_CNC"].Trim() : "Todo";
                if (!Directory.Exists(destino)) { log.Add("ERROR: no existe DESTINO " + destino); return; }
                if (Directory.GetFiles(destino, "*.nc1").Length > 0) { log.Add("ERROR: DESTINO ya tiene archivos .nc1; no se pisa nada"); return; }
                log.Add("FASE=" + fase); log.Add("DESTINO=" + destino); log.Add("CONFIG_CNC=" + config);

                // Todas las partes de la fase
                var partes = new ArrayList();
                var marcas = new List<string>();
                ModelObjectEnumerator en = model.GetModelObjectSelector().GetAllObjects();
                en.SelectInstances = true;
                while (en.MoveNext())
                {
                    Part pa = en.Current as Part;
                    if (pa == null) continue;
                    Phase ph;
                    if (!pa.GetPhase(out ph) || ph.PhaseNumber != fase) continue;
                    partes.Add(pa);
                    string pos = "";
                    pa.GetReportProperty("PART_POS", ref pos);
                    if (!marcas.Contains(pos)) marcas.Add(pos);
                }
                log.Add("PARTES_EN_FASE=" + partes.Count);
                log.Add("MARCAS_DISTINTAS=" + marcas.Count);
                if (partes.Count == 0) { log.Add("ERROR: la fase no tiene partes. No se hace nada."); return; }

                new Tekla.Structures.Model.UI.ModelObjectSelector().Select(partes);
                bool ok = Operation.CreateNCFilesFromSelected(config, destino);
                log.Add("CreateNCFilesFromSelected=" + ok);

                string[] creados = Directory.GetFiles(destino, "*.nc1");
                log.Add("NC1_CREADOS=" + creados.Length);
                marcas.Sort();
                foreach (string m in marcas)
                {
                    bool hay = File.Exists(Path.Combine(destino, m + ".nc1"));
                    log.Add((hay ? "  OK   " : "  SIN  ") + m);
                }
                log.Add("FIN=OK");
            }
            catch (Exception ex) { log.Add("EXCEPCION: " + ex.Message + " " + ex.StackTrace); }
            finally
            {
                Directory.CreateDirectory(DIR_LOG);
                File.WriteAllLines(Path.Combine(DIR_LOG, "export_cnc_log.txt"), log.ToArray());
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
