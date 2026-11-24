using System; using System.Collections; using System.Collections.Generic; using System.IO;
using Tekla.Structures.Filtering; using Tekla.Structures.Filtering.Categories;
using Tekla.Structures; using Tekla.Structures.Model; using Tekla.Structures.Drawing;
using Part = Tekla.Structures.Model.Part; using Assembly = Tekla.Structures.Model.Assembly;
using Phase = Tekla.Structures.Model.Phase; using ModelObject = Tekla.Structures.Model.ModelObject;
using ModelObjectEnumerator = Tekla.Structures.Model.ModelObjectEnumerator;
using BoltGroup = Tekla.Structures.Model.BoltGroup;

// Imprime a PDF los planos de las fases pedidas (Skill exportar-fase, Ferma), con el motor PDF propio de Tekla.
// (2026-11-14) Por TÍPICOS:
//   TIP = dibujos de CONJUNTO cuya pieza principal está en la fase          -> DESTINO_TIP_<fase>\<marca>.pdf
//   SUB = dibujos de PARTE de las piezas de esos conjuntos (partes_fase_<n>.txt, aunque la pieza tenga otra fase)
//         -> DESTINO_SUB_<fase>\<marca>.pdf
// Varias fases en una sola pasada por los dibujos. No modifica el modelo ni los planos. Nunca pisa un PDF que ya exista.
// Parámetros en %USERPROFILE%\TeklaMCP_logs\export_params.txt:
//   MODELO, CREADO, FASES=46,47 (o FASE), DESTINO_SUB_<f>/DESTINO_TIP_<f> (o DESTINO_SUB/DESTINO_TIP con una fase),
//   [SOLO=SUB|TIP], [PRUEBA=<marca>]
// Log: imprimir_pdf_log_<fase>.txt por fase y imprimir_pdf_log.txt (resumen).

//Keep: Tekla.Technology.Akit.UserScript
namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        static readonly string DIR_LOG = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "TeklaMCP_logs");

        class Destino { public string Sub, Tip; public List<string> Log = new List<string>(); public int nSub, nTip, nErr, nExiste, nDesact; public HashSet<int> Partes; }

        public static void Run(IScript akit)
        {
            var log = new List<string>();
            log.Add("INICIO=" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            var dest = new Dictionary<int, Destino>();
            try
            {
                Dictionary<string, string> p = LeerParametros(Path.Combine(DIR_LOG, "export_params.txt"));
                Model model = new Model();
                if (!model.GetConnectionStatus()) { log.Add("ERROR: sin conexion con Tekla"); return; }
                string nombre = Path.GetFileNameWithoutExtension(model.GetInfo().ModelName);
                log.Add("MODELO=" + nombre);
                if (!p.ContainsKey("MODELO") || !string.Equals(p["MODELO"].Trim(), nombre, StringComparison.OrdinalIgnoreCase))
                { log.Add("ERROR: el modelo abierto no es el esperado. No se hace nada."); return; }
                if (!ParametrosRecientes(p, log)) return;
                List<int> fases = Fases(p);
                if (fases.Count == 0) { log.Add("ERROR: FASES invalida"); return; }
                string solo = p.ContainsKey("SOLO") ? p["SOLO"].Trim().ToUpper() : "";
                string prueba = p.ContainsKey("PRUEBA") ? p["PRUEBA"].Trim() : "";
                log.Add("FASES=" + string.Join(",", fases.ConvertAll(x => x.ToString()).ToArray()));

                // pieza -> fase (por típicos)
                var faseDePieza = new Dictionary<int, int>();
                foreach (int f in fases)
                {
                    var d = new Destino();
                    d.Log.Add("INICIO=" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    d.Log.Add("MODELO=" + nombre);
                    d.Log.Add("FASE=" + f);
                    d.Sub = p.ContainsKey("DESTINO_SUB_" + f) ? p["DESTINO_SUB_" + f].Trim() : (fases.Count == 1 && p.ContainsKey("DESTINO_SUB") ? p["DESTINO_SUB"].Trim() : "");
                    d.Tip = p.ContainsKey("DESTINO_TIP_" + f) ? p["DESTINO_TIP_" + f].Trim() : (fases.Count == 1 && p.ContainsKey("DESTINO_TIP") ? p["DESTINO_TIP"].Trim() : "");
                    dest[f] = d;
                    if (!Directory.Exists(d.Sub) || !Directory.Exists(d.Tip)) { d.Log.Add("ERROR: no existe DESTINO_SUB o DESTINO_TIP"); d.Partes = null; continue; }
                    d.Log.Add("DESTINO_SUB=" + d.Sub); d.Log.Add("DESTINO_TIP=" + d.Tip);
                    List<int> ids = PartesDeFase(model, p, fases, f, d.Log);
                    if (ids == null) { d.Partes = null; continue; }
                    d.Partes = new HashSet<int>(ids);
                    foreach (int id in ids) faseDePieza[id] = f;
                }

                DrawingHandler dh = new DrawingHandler();
                if (!dh.GetConnectionStatus()) { log.Add("ERROR: sin conexion con dibujos"); return; }

                var attr = new DPMPrinterAttributes();
                attr.OutputType = DotPrintOutputType.PDF;
                attr.PaperSize = DotPrintPaperSize.Auto;
                attr.ScalingMethod = DotPrintScalingType.Auto;
                attr.ScaleFactor = 1.0;
                attr.Orientation = DotPrintOrientationType.Auto;
                attr.ColorMode = DotPrintColor.BlackAndWhite;
                attr.NumberOfCopies = 1;
                attr.OpenFileWhenFinished = false;
                attr.PrintToMultipleSheet = DotPrintToMultipleSheet.Off;

                DrawingEnumerator de = dh.GetDrawings();
                while (de.MoveNext())
                {
                    Drawing dw = de.Current;
                    string tipo = null; int fase = -1;
                    if (dw is SinglePartDrawing)
                    {
                        tipo = "SUB";
                        int id = ((SinglePartDrawing)dw).PartIdentifier.ID;
                        if (!faseDePieza.TryGetValue(id, out fase)) continue;
                    }
                    else if (dw is AssemblyDrawing)
                    {
                        tipo = "TIP";
                        Tekla.Structures.Model.Assembly a = model.SelectModelObject(((AssemblyDrawing)dw).AssemblyIdentifier) as Tekla.Structures.Model.Assembly;
                        Tekla.Structures.Model.ModelObject mo = a != null ? a.GetMainPart() : null;
                        Tekla.Structures.Model.Phase ph;
                        if (mo == null || !mo.GetPhase(out ph)) continue;
                        fase = ph.PhaseNumber;
                    }
                    if (tipo == null || !dest.ContainsKey(fase) || dest[fase].Partes == null) continue;
                    if (solo != "" && solo != tipo) continue;
                    Destino d = dest[fase];
                    string marca = dw.Mark.Trim().TrimStart('[').TrimEnd(']').Trim();
                    if (prueba != "" && !string.Equals(prueba, marca, StringComparison.OrdinalIgnoreCase)) continue;
                    string destino = Path.Combine(tipo == "SUB" ? d.Sub : d.Tip, marca + ".pdf");
                    string estado = dw.UpToDateStatus.ToString();
                    if (File.Exists(destino)) { d.nExiste++; d.Log.Add("  YA EXISTE (no se pisa) " + tipo + " " + marca); continue; }
                    bool ok = dh.PrintDrawing(dw, attr, destino);
                    bool hay = File.Exists(destino);
                    if (ok && hay) { if (tipo == "SUB") d.nSub++; else d.nTip++; }
                    else d.nErr++;
                    if (estado != "DrawingIsUpToDate") d.nDesact++;
                    d.Log.Add("  " + (ok && hay ? "OK  " : "ERR ") + tipo + " " + marca + "  estado=" + estado + "  print=" + ok + " archivo=" + hay);
                }
                foreach (var kv in dest)
                {
                    Destino d = kv.Value;
                    if (d.Partes == null) continue;
                    d.Log.Add("SUB_IMPRESOS=" + d.nSub);
                    d.Log.Add("TIP_IMPRESOS=" + d.nTip);
                    d.Log.Add("ERRORES=" + d.nErr);
                    d.Log.Add("YA_EXISTIAN=" + d.nExiste);
                    d.Log.Add("NO_ACTUALIZADOS=" + d.nDesact);
                    d.Log.Add("FIN=OK");
                    log.Add("FASE " + kv.Key + ": SUB " + d.nSub + " · TIP " + d.nTip + " · errores " + d.nErr);
                }
                log.Add("FIN=OK");
            }
            catch (Exception ex) { log.Add("EXCEPCION: " + ex.Message + " " + ex.StackTrace); }
            finally
            {
                Directory.CreateDirectory(DIR_LOG);
                foreach (var kv in dest)
                    File.WriteAllLines(Path.Combine(DIR_LOG, "imprimir_pdf_log_" + kv.Key + ".txt"), kv.Value.Log.ToArray());
                File.WriteAllLines(Path.Combine(DIR_LOG, "imprimir_pdf_log.txt"), log.ToArray());
            }
        }

        // ---------------------------------------------------------------- piezas de la fase por TÍPICOS (igual en las 3 macros)
        // Devuelve los IDs de las piezas de los conjuntos (de primer nivel) cuya pieza principal está en la fase.
        // Usa partes_fase_<n>.txt si es de esta misma exportación (mismo CREADO); si no, calcula TODAS las fases de FASES
        // y escribe un archivo por fase con: ID= (piezas), BULON= (bulones que van), OTRA_FASE=, FUERA=, BULON_FUERA= (avisos).
        static List<int> PartesDeFase(Model model, Dictionary<string, string> p, List<int> fases, int fase, List<string> log)
        {
            string creado = p["CREADO"].Trim();
            List<int> ya = LeerIds(fase, creado);
            if (ya != null) { log.Add("PIEZAS=desde partes_fase_" + fase + ".txt"); return ya; }

            var porFase = new Dictionary<int, List<int>>();
            var otraFase = new Dictionary<int, List<string>>();
            var fuera = new Dictionary<int, List<string>>();
            var bulones = new Dictionary<int, List<int>>();
            var bulFuera = new Dictionary<int, List<string>>();
            foreach (int f in fases)
            {
                porFase[f] = new List<int>(); otraFase[f] = new List<string>(); fuera[f] = new List<string>();
                bulones[f] = new List<int>(); bulFuera[f] = new List<string>();
            }
            var vistos = new HashSet<int>();          // conjuntos de primer nivel ya revisados
            var tipicoDe = new Dictionary<int, int>(); // pieza -> fase del típico (solo piezas de típicos de las fases)

            // 1) Conjuntos de las fases: Tekla devuelve las piezas cuyo conjunto tiene esa fase (filtro, sin recorrer todo)
            bool porFiltro = true;
            ModelObjectEnumerator en;
            try { en = model.GetModelObjectSelector().GetObjectsByFilter(Filtro(fases, true)); }
            catch { en = null; porFiltro = false; }
            if (en != null) JuntarTipicos(en, fases, vistos, porFase, otraFase, tipicoDe);
            int total = 0;
            foreach (int f in fases) total += porFase[f].Count;
            if (total == 0)
            {
                // el filtro no trajo nada (o no anduvo en esta versión): se recorre el modelo entero, como antes
                porFiltro = false;
                JuntarTipicos(model.GetModelObjectSelector().GetAllObjectsWithType(ModelObject.ModelObjectEnum.ASSEMBLY),
                              fases, vistos, porFase, otraFase, tipicoDe);
            }
            log.Add("METODO=" + (porFiltro ? "filtro por fase del conjunto" : "recorrido completo del modelo"));

            // 2) (2026-11-24) BULONES COMO LOS SACA EL EQUIPO (video del usuario): filtro de vista "Conjunto · Fase = N",
            //    se selecciona todo lo que muestra y se saca el reporte. Entra todo bulón cuyo CONJUNTO es de la fase,
            //    tenga la fase que tenga el bulón. Reemplaza la regla del 2026-11-14 (fase del bulón + toca un típico).
            var yaBulon = new HashSet<int>();
            int bulPorFiltro = 0;
            foreach (int f in fases)
            {
                ModelObjectEnumerator eb = null;
                try { eb = model.GetModelObjectSelector().GetObjectsByFilter(Filtro(new List<int> { f }, true)); }
                catch { eb = null; }
                if (eb == null) continue;
                while (eb.MoveNext())
                {
                    BoltGroup bg = eb.Current as BoltGroup;
                    if (bg == null || yaBulon.Contains(bg.Identifier.ID)) continue;
                    yaBulon.Add(bg.Identifier.ID);
                    bulones[f].Add(bg.Identifier.ID);
                    bulPorFiltro++;
                }
            }
            if (bulPorFiltro == 0)
            {
                // El filtro no devolvió bulones (otra versión de Tekla): se toman los bulones de las piezas de los típicos.
                // Cada bulón va con el conjunto de la pieza que se abulona (PartToBeBolted), que es su conjunto en Tekla.
                foreach (int f in fases)
                    foreach (int id in porFase[f])
                    {
                        Part pa = model.SelectModelObject(new Identifier(id)) as Part;
                        if (pa == null) continue;
                        ModelObjectEnumerator eb = pa.GetBolts();
                        while (eb.MoveNext())
                        {
                            BoltGroup bg = eb.Current as BoltGroup;
                            if (bg == null || bg.PartToBeBolted == null || yaBulon.Contains(bg.Identifier.ID)) continue;
                            int ft;
                            if (!tipicoDe.TryGetValue(bg.PartToBeBolted.Identifier.ID, out ft)) continue;
                            yaBulon.Add(bg.Identifier.ID);
                            bulones[ft].Add(bg.Identifier.ID);
                        }
                    }
            }
            log.Add("METODO_BULONES=" + (bulPorFiltro > 0 ? "filtro Conjunto Fase (como en la vista)" : "bulones de las piezas de los tipicos"));

            // 3) Objetos con la fase puesta (filtro): piezas que quedan afuera y bulones con la fase que son de un
            //    conjunto de otra fase (no van; se avisan)
            try { en = model.GetModelObjectSelector().GetObjectsByFilter(Filtro(fases, false)); }
            catch { en = model.GetModelObjectSelector().GetAllObjects(); }
            while (en.MoveNext())
            {
                ModelObject o = en.Current;
                Phase ph;
                if (o == null || !o.GetPhase(out ph) || !porFase.ContainsKey(ph.PhaseNumber)) continue;
                int f = ph.PhaseNumber;
                Part pa = o as Part;
                if (pa != null)
                {
                    if (!tipicoDe.ContainsKey(pa.Identifier.ID)) { string m = Marca(pa); if (!fuera[f].Contains(m)) fuera[f].Add(m); }
                    continue;
                }
                BoltGroup bg = o as BoltGroup;
                if (bg == null || yaBulon.Contains(bg.Identifier.ID)) continue;
                var une = new List<string>();
                foreach (Part q in PiezasDe(bg)) une.Add(Marca(q));
                bulFuera[f].Add(string.Join("+", une.ToArray()));
            }
            foreach (int f in fases)
                EscribirIds(f, creado, porFase[f], otraFase[f], fuera[f], bulones[f], bulFuera[f]);
            log.Add("PIEZAS=calculadas por tipicos para las fases " + string.Join(",", fases.ConvertAll(x => x.ToString()).ToArray()));
            return porFase[fase];
        }

        static FilterExpression Filtro(List<int> fases, bool porConjunto)
        {
            var col = new BinaryFilterExpressionCollection();
            foreach (int f in fases)
            {
                BinaryFilterExpression e = porConjunto
                    ? new BinaryFilterExpression(new AssemblyFilterExpressions.Phase(), StringOperatorType.IS_EQUAL, new StringConstantFilterExpression(f.ToString()))
                    : new BinaryFilterExpression(new ObjectFilterExpressions.Phase(), NumericOperatorType.IS_EQUAL, new NumericConstantFilterExpression(f));
                col.Add(new BinaryFilterExpressionItem(e, BinaryFilterOperatorType.BOOLEAN_OR));
            }
            return col;
        }

        // Recorre piezas (o conjuntos) y junta las piezas de cada conjunto de PRIMER NIVEL cuya pieza principal está en una fase.
        static void JuntarTipicos(ModelObjectEnumerator en, List<int> fases, HashSet<int> vistos, Dictionary<int, List<int>> porFase,
                                  Dictionary<int, List<string>> otraFase, Dictionary<int, int> tipicoDe)
        {
            while (en.MoveNext())
            {
                ModelObject o = en.Current;
                Assembly a = o as Assembly;
                if (a == null) { Part pa = o as Part; if (pa != null) a = pa.GetAssembly(); }
                a = PrimerNivel(a);
                if (a == null || vistos.Contains(a.Identifier.ID)) continue;
                vistos.Add(a.Identifier.ID);
                Part main = a.GetMainPart() as Part;
                Phase ph;
                int fm = (main != null && main.GetPhase(out ph)) ? ph.PhaseNumber : -1;
                if (!porFase.ContainsKey(fm)) continue;
                var partes = new List<Part>();
                Juntar(a, partes, 0);
                foreach (Part pa in partes)
                {
                    if (tipicoDe.ContainsKey(pa.Identifier.ID)) continue;
                    tipicoDe[pa.Identifier.ID] = fm;
                    porFase[fm].Add(pa.Identifier.ID);
                    if (pa.GetPhase(out ph) && ph.PhaseNumber != fm) otraFase[fm].Add(Marca(pa) + " (fase " + ph.PhaseNumber + ")");
                }
            }
        }

        static Assembly PrimerNivel(Assembly a)
        {
            for (int i = 0; a != null && i < 10; i++)
            {
                Assembly padre = a.GetAssembly() as Assembly;
                if (padre == null || padre.Identifier.ID == 0 || padre.Identifier.ID == a.Identifier.ID) break;
                a = padre;
            }
            return a;
        }

        static List<Part> PiezasDe(BoltGroup bg)
        {
            var r = new List<Part>();
            if (bg.PartToBeBolted != null) r.Add(bg.PartToBeBolted);
            if (bg.PartToBoltTo != null) r.Add(bg.PartToBoltTo);
            foreach (object o in bg.GetOtherPartsToBolt()) { Part q = o as Part; if (q != null) r.Add(q); }
            return r;
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
            return LeerLista(fase, "ID=");
        }

        static List<int> LeerLista(int fase, string clave)
        {
            var ids = new List<int>();
            string ruta = ArchivoIds(fase);
            if (!File.Exists(ruta)) return ids;
            foreach (string x in File.ReadAllLines(ruta)) { int id; if (x.StartsWith(clave) && int.TryParse(x.Substring(clave.Length), out id)) ids.Add(id); }
            return ids;
        }

        static void EscribirIds(int fase, string creado, List<int> ids, List<string> otraFase, List<string> fuera, List<int> bulones, List<string> bulFuera)
        {
            var l = new List<string>();
            l.Add("CREADO=" + creado);
            l.Add("FASE=" + fase);
            foreach (string m in otraFase) l.Add("OTRA_FASE=" + m);
            foreach (string m in fuera) l.Add("FUERA=" + m);
            foreach (string m in bulFuera) l.Add("BULON_FUERA=" + m);
            foreach (int id in ids) l.Add("ID=" + id);
            foreach (int id in bulones) l.Add("BULON=" + id);
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
