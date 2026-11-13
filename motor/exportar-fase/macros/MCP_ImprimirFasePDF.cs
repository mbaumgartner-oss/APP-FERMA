using System; using System.Collections; using System.Collections.Generic; using System.IO;
using Tekla.Structures.Model; using Tekla.Structures.Drawing;

// Imprime a PDF los planos de una fase (Skill exportar-fase, Ferma), con el motor PDF propio de Tekla (sin impresora).
//   SUB = dibujos de PARTE cuya pieza está en la fase          -> DESTINO_SUB\<marca>.pdf
//   TIP = dibujos de CONJUNTO cuya pieza principal está en la fase -> DESTINO_TIP\<marca>.pdf
// Se buscan por la FASE de la pieza, no por la marca (las marcas de típicos no llevan la fase).
// No modifica el modelo ni los planos. Nunca pisa un PDF que ya exista.
// Parámetros en %USERPROFILE%\TeklaMCP_logs\export_params.txt:
//   MODELO=..., FASE=..., DESTINO_SUB=..., DESTINO_TIP=..., [SOLO=SUB|TIP], [PRUEBA=<marca>] (imprime solo esa)
// Log: %USERPROFILE%\TeklaMCP_logs\imprimir_pdf_log.txt

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
                { log.Add("ERROR: el modelo abierto no es el esperado. No se hace nada."); return; }
                if (!ParametrosRecientes(p, log)) return;
                int fase;
                if (!p.ContainsKey("FASE") || !int.TryParse(p["FASE"].Trim(), out fase)) { log.Add("ERROR: FASE invalida"); return; }
                string dSub = p.ContainsKey("DESTINO_SUB") ? p["DESTINO_SUB"].Trim() : "";
                string dTip = p.ContainsKey("DESTINO_TIP") ? p["DESTINO_TIP"].Trim() : "";
                string solo = p.ContainsKey("SOLO") ? p["SOLO"].Trim().ToUpper() : "";
                string prueba = p.ContainsKey("PRUEBA") ? p["PRUEBA"].Trim() : "";
                if (!Directory.Exists(dSub) || !Directory.Exists(dTip)) { log.Add("ERROR: no existe DESTINO_SUB o DESTINO_TIP"); return; }
                log.Add("FASE=" + fase); log.Add("DESTINO_SUB=" + dSub); log.Add("DESTINO_TIP=" + dTip);

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

                int nSub = 0, nTip = 0, nErr = 0, nExiste = 0, nDesact = 0;
                DrawingEnumerator de = dh.GetDrawings();
                while (de.MoveNext())
                {
                    Drawing dw = de.Current;
                    string tipo = null; Tekla.Structures.Model.ModelObject mo = null;
                    if (dw is SinglePartDrawing)
                    { tipo = "SUB"; mo = model.SelectModelObject(((SinglePartDrawing)dw).PartIdentifier); }
                    else if (dw is AssemblyDrawing)
                    {
                        tipo = "TIP";
                        Tekla.Structures.Model.Assembly a = model.SelectModelObject(((AssemblyDrawing)dw).AssemblyIdentifier) as Tekla.Structures.Model.Assembly;
                        if (a != null) mo = a.GetMainPart();
                    }
                    if (tipo == null || mo == null) continue;
                    if (solo != "" && solo != tipo) continue;
                    Tekla.Structures.Model.Phase ph;
                    if (!mo.GetPhase(out ph) || ph.PhaseNumber != fase) continue;
                    string marca = dw.Mark.Trim().TrimStart('[').TrimEnd(']').Trim();
                    if (prueba != "" && !string.Equals(prueba, marca, StringComparison.OrdinalIgnoreCase)) continue;
                    string destino = Path.Combine(tipo == "SUB" ? dSub : dTip, marca + ".pdf");
                    string estado = dw.UpToDateStatus.ToString();
                    if (File.Exists(destino)) { nExiste++; log.Add("  YA EXISTE (no se pisa) " + tipo + " " + marca); continue; }
                    bool ok = dh.PrintDrawing(dw, attr, destino);
                    bool hay = File.Exists(destino);
                    if (ok && hay) { if (tipo == "SUB") nSub++; else nTip++; }
                    else nErr++;
                    if (estado != "DrawingIsUpToDate") nDesact++;
                    log.Add("  " + (ok && hay ? "OK  " : "ERR ") + tipo + " " + marca + "  estado=" + estado + "  print=" + ok + " archivo=" + hay);
                }
                log.Add("SUB_IMPRESOS=" + nSub);
                log.Add("TIP_IMPRESOS=" + nTip);
                log.Add("ERRORES=" + nErr);
                log.Add("YA_EXISTIAN=" + nExiste);
                log.Add("NO_ACTUALIZADOS=" + nDesact);
                log.Add("FIN=OK");
            }
            catch (Exception ex) { log.Add("EXCEPCION: " + ex.Message + " " + ex.StackTrace); }
            finally
            {
                Directory.CreateDirectory(DIR_LOG);
                File.WriteAllLines(Path.Combine(DIR_LOG, "imprimir_pdf_log.txt"), log.ToArray());
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
