using System; using System.Collections.Generic; using System.IO;
using Tekla.Structures.Model;

// Lee el modelo abierto (nombre y ruta) para la Skill exportar-fase (Ferma). No modifica nada.
// Log: %USERPROFILE%\TeklaMCP_logs\probe_modelo_log.txt  (PROBE_FECHA sirve para saber que el log es de recién)

//Keep: Tekla.Technology.Akit.UserScript
namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        static readonly string DIR_LOG = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "TeklaMCP_logs");

        public static void Run(IScript akit)
        {
            var log = new List<string>();
            log.Add("PROBE_FECHA=" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            try
            {
                Model model = new Model();
                if (!model.GetConnectionStatus()) { log.Add("ERROR: sin conexion con Tekla"); return; }
                ModelInfo mi = model.GetInfo();
                log.Add("MODEL_NAME=" + mi.ModelName);
                log.Add("MODEL_PATH=" + mi.ModelPath);
                string xs = "";
                Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_FIRM", ref xs); log.Add("XS_FIRM=" + xs);
                xs = ""; Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_PROJECT", ref xs); log.Add("XS_PROJECT=" + xs);
            }
            catch (Exception ex) { log.Add("EXCEPCION: " + ex.Message); }
            finally
            {
                Directory.CreateDirectory(DIR_LOG);
                File.WriteAllLines(Path.Combine(DIR_LOG, "probe_modelo_log.txt"), log.ToArray());
            }
        }
    }
}
