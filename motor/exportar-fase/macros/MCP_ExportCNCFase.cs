using System; using System.Collections; using System.Collections.Generic; using System.IO;
using Tekla.Structures; using Tekla.Structures.Model; using Tekla.Structures.Model.Operations;

// Exporta los CNC (.nc1) de las fases pedidas con la config "CNC FERMA" (Skill exportar-fase, Ferma).
// (2026-11-14) Por TÍPICOS: todas las piezas de los conjuntos cuya pieza principal está en la fase (las mismas que
// seleccionó MCP_SeleccionarFase, leídas de partes_fase_<n>.txt). Varias fases en una sola corrida.
// No modifica el modelo: solo cambia la selección y escribe archivos en el destino de cada fase.
// Parámetros en %USERPROFILE%\TeklaMCP_logs\export_params.txt (CLAVE=valor, uno por línea):
//   MODELO=<nombre exacto del modelo abierto, sin .db1>   -> si no coincide, NO hace nada
//   FASES=46,47 (o FASE=46) · DESTINO_CNC_46=<carpeta> (o DESTINO=<carpeta> con una sola fase)
//   CONFIG_CNC=Todo   = la LINEA .ncfs marcada dentro de "CNC FERMA.ncf"; con "CNC FERMA" falla
// El destino tiene que existir y no tener .nc1: nunca pisa archivos.
// Log: export_cnc_log_<fase>.txt por fase y export_cnc_log.txt (resumen).

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
                if (!p.ContainsKey("MODELO") || !string.Equals(p["MODELO"].Trim(), nombre, StringComparison.OrdinalIgnoreCase))
                { log.Add("ERROR: el modelo abierto no es el esperado. No se hace nada."); return; }
                if (!ParametrosRecientes(p, log)) return;
                List<int> fases = Fases(p);
                if (fases.Count == 0) { log.Add("ERROR: FASES invalida"); return; }
                string config = p.ContainsKey("CONFIG_CNC") ? p["CONFIG_CNC"].Trim() : "Todo";
                log.Add("FASES=" + string.Join(",", fases.ConvertAll(x => x.ToString()).ToArray()));
                foreach (int fase in fases)
                {
                    var lf = new List<string>();
                    lf.Add("INICIO=" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    lf.Add("MODELO=" + nombre);
                    lf.Add("FASE=" + fase);
                    try { ExportarFase(model, p, fases, fase, config, lf); }
                    catch (Exception ex) { lf.Add("EXCEPCION: " + ex.Message + " " + ex.StackTrace); }
                    File.WriteAllLines(Path.Combine(DIR_LOG, "export_cnc_log_" + fase + ".txt"), lf.ToArray());
                    log.Add("FASE " + fase + ": " + (lf.Contains("FIN=OK") ? "OK" : "ERROR (ver export_cnc_log_" + fase + ".txt)"));
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

        static void ExportarFase(Model model, Dictionary<string, string> p, List<int> fases, int fase, string config, List<string> log)
        {
            string k = "DESTINO_CNC_" + fase;
            string destino = p.ContainsKey(k) ? p[k].Trim() : (fases.Count == 1 && p.ContainsKey("DESTINO") ? p["DESTINO"].Trim() : "");
            if (destino == "") { log.Add("ERROR: falta " + k); return; }
            if (!destino.EndsWith("\\")) destino += "\\";
            if (!Directory.Exists(destino)) { log.Add("ERROR: no existe DESTINO " + destino); return; }
            if (Directory.GetFiles(destino, "*.nc1").Length > 0) { log.Add("ERROR: DESTINO ya tiene archivos .nc1; no se pisa nada"); return; }
            log.Add("DESTINO=" + destino); log.Add("CONFIG_CNC=" + config);

            List<int> ids = PartesDeFase(model, p, fases, fase, log);
            if (ids == null) return;
            var partes = new ArrayList();
            var marcas = new List<string>();
            foreach (int id in ids)
            {
                Part pa = model.SelectModelObject(new Identifier(id)) as Part;
                if (pa == null) continue;
                partes.Add(pa);
                string pos = Marca(pa);
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

        // ---------------------------------------------------------------- piezas de la fase por TÍPICOS (igual en las 3 macros)
        // Devuelve los IDs de las piezas de los conjuntos (de primer nivel) cuya pieza principal está en la fase.
        // Usa partes_fase_<n>.txt si es de esta misma exportación (mismo CREADO); si no, calcula TODAS las fases de FASES.
        static List<int> PartesDeFase(Model model, Dictionary<string, string> p, List<int> fases, int fase, List<string> log)
        {
            string creado = p["CREADO"].Trim();
            List<int> ya = LeerIds(fase, creado);
            if (ya != null) { log.Add("PIEZAS=desde partes_fase_" + fase + ".txt"); return ya; }

            var subs = new HashSet<int>();
            var conjuntos = new List<KeyValuePair<int, Assembly>>();   // (fase de la pieza principal, conjunto)
            ModelObjectEnumerator en = model.GetModelObjectSelector().GetAllObjectsWithType(ModelObject.ModelObjectEnum.ASSEMBLY);
            while (en.MoveNext())
            {
                Assembly a = en.Current as Assembly;
                if (a == null) continue;
                foreach (object s in a.GetSubAssemblies()) { Assembly sa = s as Assembly; if (sa != null) subs.Add(sa.Identifier.ID); }
                Part main = a.GetMainPart() as Part;
                Phase ph;
                int fm = (main != null && main.GetPhase(out ph)) ? ph.PhaseNumber : -1;
                conjuntos.Add(new KeyValuePair<int, Assembly>(fm, a));
            }
            var porFase = new Dictionary<int, List<int>>();
            var otraFase = new Dictionary<int, List<string>>();    // pieza de un típico de la fase con otra fase
            var enTipicos = new HashSet<int>();
            foreach (int f in fases) { porFase[f] = new List<int>(); otraFase[f] = new List<string>(); }
            foreach (var kv in conjuntos)
            {
                Assembly a = kv.Value;
                if (subs.Contains(a.Identifier.ID) || !porFase.ContainsKey(kv.Key)) continue;   // sub-conjunto u otra fase
                var partes = new List<Part>();
                Juntar(a, partes, 0);
                foreach (Part pa in partes)
                {
                    if (enTipicos.Contains(pa.Identifier.ID)) continue;
                    enTipicos.Add(pa.Identifier.ID);
                    porFase[kv.Key].Add(pa.Identifier.ID);
                    Phase ph;
                    if (pa.GetPhase(out ph) && ph.PhaseNumber != kv.Key) otraFase[kv.Key].Add(Marca(pa) + " (fase " + ph.PhaseNumber + ")");
                }
            }
            // piezas con la fase puesta pero de un típico de OTRA fase: no van (solo se informan)
            var fuera = new Dictionary<int, List<string>>();
            foreach (int f in fases) fuera[f] = new List<string>();
            en = model.GetModelObjectSelector().GetAllObjects();
            Fuera(en, fuera, enTipicos);
            foreach (int f in fases)
                EscribirIds(f, creado, porFase[f], otraFase[f], fuera[f]);
            log.Add("PIEZAS=calculadas por tipicos para las fases " + string.Join(",", fases.ConvertAll(x => x.ToString()).ToArray()));
            return porFase[fase];
        }

        static void Fuera(ModelObjectEnumerator en, Dictionary<int, List<string>> fuera, HashSet<int> enTipicos)
        {
            while (en.MoveNext())
            {
                Part pa = en.Current as Part;
                if (pa == null || enTipicos.Contains(pa.Identifier.ID)) continue;
                Phase ph;
                if (!pa.GetPhase(out ph) || !fuera.ContainsKey(ph.PhaseNumber)) continue;
                string m = Marca(pa);
                if (!fuera[ph.PhaseNumber].Contains(m)) fuera[ph.PhaseNumber].Add(m);
            }
        }

        static void Juntar(Assembly a, List<Part> partes, int nivel)
        {
            if (nivel > 10) return;
            Part main = a.GetMainPart() as Part;
            if (main != null) partes.Add(main);
            foreach (object o in a.GetSecondaries()) { Part pa = o as Part; if (pa != null) partes.Add(pa); }
            foreach (object s in a.GetSubAssemblies()) { Assembly sa = s as Assembly; if (sa != null) Juntar(sa, partes, nivel + 1); }
        }

        static string Marca(Part pa) { string m = ""; pa.GetReportProperty("PART_POS", ref m); return m; }

        static string ArchivoIds(int fase) { return Path.Combine(DIR_LOG, "partes_fase_" + fase + ".txt"); }

        static List<int> LeerIds(int fase, string creado)
        {
            string ruta = ArchivoIds(fase);
            if (!File.Exists(ruta)) return null;
            string[] l = File.ReadAllLines(ruta);
            if (l.Length == 0 || l[0].Trim() != "CREADO=" + creado) return null;
            var ids = new List<int>();
            foreach (string x in l) { int id; if (x.StartsWith("ID=") && int.TryParse(x.Substring(3), out id)) ids.Add(id); }
            return ids;
        }

        static void EscribirIds(int fase, string creado, List<int> ids, List<string> otraFase, List<string> fuera)
        {
            var l = new List<string>();
            l.Add("CREADO=" + creado);
            l.Add("FASE=" + fase);
            foreach (string m in otraFase) l.Add("OTRA_FASE=" + m);
            foreach (string m in fuera) l.Add("FUERA=" + m);
            foreach (int id in ids) l.Add("ID=" + id);
            Directory.CreateDirectory(DIR_LOG);
            File.WriteAllLines(ArchivoIds(fase), l.ToArray());
        }

        static List<int> Fases(Dictionary<string, string> p)
        {
            var r = new List<int>();
            string t = p.ContainsKey("FASES") ? p["FASES"] : (p.ContainsKey("FASE") ? p["FASE"] : "");
            foreach (string x in t.Split(',')) { int f; if (int.TryParse(x.Trim(), out f) && !r.Contains(f)) r.Add(f); }
            return r;
        }

        // Los parámetros los escribe "exportar_fase.py preparar --crear" con CREADO=fecha. Si son viejos, no se usan:
        // así nunca se exporta la fase anterior por accidente (2026-11-10).
        static bool ParametrosRecientes(Dictionary<string, string> p, List<string> log)
        {
            DateTime c;
            if (!p.ContainsKey("CREADO") || !DateTime.TryParseExact(p["CREADO"].Trim(), "yyyy-MM-dd HH:mm:ss",
                    System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out c))
            { log.Add("ERROR: export_params.txt sin CREADO. Correr de nuevo 'preparar --crear'. No se hace nada."); return false; }
            if ((DateTime.Now - c).TotalMinutes > 240)
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
