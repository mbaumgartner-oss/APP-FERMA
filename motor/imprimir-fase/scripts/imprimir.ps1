# Skill imprimir-fase (Ferma) - trabajador de Windows. Lo llama imprimir_fase.py; no se usa a mano.
#   -modo xlsx2pdf : cada .xlsx -> PDF temporal (respeta A4/orientación del archivo; el .xlsx NO se guarda)
#   -modo info     : páginas y tamaño (mm) de cada PDF
#   -modo imprimir : manda cada PDF a su impresora: tamaño de hoja y orientación por página, B/N, simple faz, 1 copia.
#                    Con "salida" en el trabajo imprime a archivo (prueba sin papel con "Microsoft Print to PDF").
# Entrada: -entrada <json>  ·  Salida: JSON por stdout.
# PDF: se dibuja la página a escala 1:1 (renderizada con el motor PDF de Windows, hasta 600 dpi).
param([Parameter(Mandatory)][ValidateSet('xlsx2pdf','info','imprimir')][string]$modo,
      [Parameter(Mandatory)][string]$entrada)
$ErrorActionPreference = 'Stop'
$trabajos = Get-Content -Raw -Encoding UTF8 $entrada | ConvertFrom-Json

function Salir($obj) { [Console]::OutputEncoding = [Text.Encoding]::UTF8; $obj | ConvertTo-Json -Depth 6 -Compress; exit 0 }

if ($modo -eq 'xlsx2pdf') {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    # OJO: el PDF de Excel toma el tamaño de la impresora predeterminada de Windows: con "Microsoft Print to PDF"
    # una hoja A4 sale Carta (2026-11-09) y Excel no deja cambiar ActivePrinter por COM. Por eso se devuelve la
    # hoja y orientación configuradas en el Excel, y al imprimir la página se ajusta a esa hoja.
    $papel = @{ 9 = 'A4'; 8 = 'A3'; 66 = 'A2'; 67 = 'A1' }
    $res = @()
    try {
        foreach ($t in $trabajos) {
            try {
                $wb = $xl.Workbooks.Open($t.archivo, 0, $true)   # solo lectura
                $su = $wb.Worksheets.Item(1).PageSetup
                $hoja = $papel[[int]$su.PaperSize]; $horiz = ([int]$su.Orientation -eq 2)
                $wb.ExportAsFixedFormat(0, $t.pdf)               # 0 = PDF
                $wb.Close($false)
                $res += @{ archivo = $t.archivo; pdf = $t.pdf; ok = $true; hoja = $hoja; horizontal = $horiz; papel = [int]$su.PaperSize }
            } catch { $res += @{ archivo = $t.archivo; ok = $false; error = "$_" } }
        }
    } finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    Salir $res
}

# --- PDF con el motor de Windows (WinRT) ---
Add-Type -AssemblyName System.Runtime.WindowsRuntime
Add-Type -AssemblyName System.Drawing
[void][Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
[void][Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
$asTaskOp = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
$asTaskAc = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' } | Select-Object -First 1
function Esperar($op, [Type]$tipo) { $t = $asTaskOp.MakeGenericMethod($tipo).Invoke($null, @($op)); [void]$t.Wait(-1); $t.Result }
function EsperarAccion($ac) { $t = $asTaskAc.Invoke($null, @($ac)); [void]$t.Wait(-1) }

function AbrirPdf($ruta) {
    $ms = New-Object IO.MemoryStream(, [IO.File]::ReadAllBytes($ruta))   # por bytes: anda igual en el Drive
    $ras = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($ms)
    Esperar ([Windows.Data.Pdf.PdfDocument]::LoadFromStreamAsync($ras)) ([Windows.Data.Pdf.PdfDocument])
}
$DIP_MM = 25.4 / 96
$script:DPI_MAX = 400
$script:BN = $true          # blanco y negro puro (1 bit). Se puede apagar con FERMA_IMPRIMIR_COLOR=1
if ($env:FERMA_IMPRIMIR_COLOR -eq '1') { $script:BN = $false; $script:DPI_MAX = 600 }

if ($modo -eq 'info') {
    $res = @()
    foreach ($t in $trabajos) {
        try {
            $doc = AbrirPdf $t.archivo
            $pags = @()
            for ($i = 0; $i -lt $doc.PageCount; $i++) {
                $p = $doc.GetPage($i)
                $pags += @{ ancho_mm = [math]::Round($p.Size.Width * $DIP_MM, 1); alto_mm = [math]::Round($p.Size.Height * $DIP_MM, 1); giro = [int]$p.Rotation }
                $p.Dispose()
            }
            $res += @{ archivo = $t.archivo; ok = $true; paginas = $pags }
        } catch { $res += @{ archivo = $t.archivo; ok = $false; error = "$_" } }
    }
    Salir $res
}

function Renderizar($pagina, [double]$dpi, [bool]$tramar = $false) {
    $op = New-Object Windows.Data.Pdf.PdfPageRenderOptions
    $op.DestinationWidth = [uint32][math]::Round($pagina.Size.Width * $dpi / 96)
    $op.DestinationHeight = [uint32][math]::Round($pagina.Size.Height * $dpi / 96)
    $op.BackgroundColor = [Windows.UI.Color]::FromArgb(255, 255, 255, 255)
    $st = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
    EsperarAccion ($pagina.RenderToStreamAsync($st, $op))
    $net = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($st.GetInputStreamAt(0))
    $ms = New-Object IO.MemoryStream; $net.CopyTo($ms); $st.Dispose()
    $ms.Position = 0
    $img = [System.Drawing.Bitmap]::FromStream($ms)
    if (-not $script:BN) { return $img }
    # (2026-11-14) Blanco y negro puro (1 bit): la hoja viaja ~25 veces más liviana que en color y la impresora la
    # procesa mucho más rápido. Lo que no es casi blanco queda negro, así las líneas finas y grises no se pierden.
    # (2026-11-17) Los informes de Excel tienen títulos con fondo gris: con umbral salían negros y el texto no se
    # veía. En las páginas de Excel los grises se TRAMAN (punteado) y el texto negro queda encima, legible.
    if ($tramar) { $bn = [Bn1]::Tramar($img, 100, 235) } else { $bn = Convertir1Bit $img }
    $img.Dispose(); $ms.Dispose()
    return $bn
}

Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @'
using System; using System.Drawing; using System.Drawing.Imaging; using System.Runtime.InteropServices;
public static class Bn1 {
    // Bitmap -> 1 bit por pixel. Umbral alto (casi blanco = blanco, el resto negro): no se pierden líneas grises finas.
    public static Bitmap Convertir(Bitmap src, int umbral) {
        int w = src.Width, h = src.Height;
        Bitmap c32 = src.PixelFormat == PixelFormat.Format32bppArgb ? src : src.Clone(new Rectangle(0, 0, w, h), PixelFormat.Format32bppArgb);
        Bitmap dst = new Bitmap(w, h, PixelFormat.Format1bppIndexed);
        dst.SetResolution(src.HorizontalResolution, src.VerticalResolution);
        BitmapData a = c32.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        BitmapData b = dst.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.WriteOnly, PixelFormat.Format1bppIndexed);
        byte[] fila = new byte[a.Stride]; byte[] sal = new byte[b.Stride];
        for (int y = 0; y < h; y++) {
            Marshal.Copy(new IntPtr(a.Scan0.ToInt64() + (long)y * a.Stride), fila, 0, a.Stride);
            Array.Clear(sal, 0, sal.Length);
            for (int x = 0; x < w; x++) {
                int i = x * 4; int lum = (fila[i] * 114 + fila[i + 1] * 587 + fila[i + 2] * 299) / 1000;
                if (lum >= umbral) sal[x >> 3] |= (byte)(0x80 >> (x & 7));   // bit en 1 = blanco (paleta 0 negro, 1 blanco)
            }
            Marshal.Copy(sal, 0, new IntPtr(b.Scan0.ToInt64() + (long)y * b.Stride), b.Stride);
        }
        c32.UnlockBits(a); dst.UnlockBits(b);
        if (!object.ReferenceEquals(c32, src)) c32.Dispose();
        return dst;
    }

    // Para informes (Excel): lum <= negro -> negro (texto, bordes); lum >= blanco -> blanco; en el medio, trama
    // ordenada (Bayer 8x8) con la densidad del gris. Un fondo gris queda punteado y el texto negro se lee encima.
    static readonly int[] BAYER = {
         0,32, 8,40, 2,34,10,42, 48,16,56,24,50,18,58,26, 12,44, 4,36,14,46, 6,38, 60,28,52,20,62,30,54,22,
         3,35,11,43, 1,33, 9,41, 51,19,59,27,49,17,57,25, 15,47, 7,39,13,45, 5,37, 63,31,55,23,61,29,53,21 };
    public static Bitmap Tramar(Bitmap src, int negro, int blanco) {
        int w = src.Width, h = src.Height;
        Bitmap c32 = src.PixelFormat == PixelFormat.Format32bppArgb ? src : src.Clone(new Rectangle(0, 0, w, h), PixelFormat.Format32bppArgb);
        Bitmap dst = new Bitmap(w, h, PixelFormat.Format1bppIndexed);
        dst.SetResolution(src.HorizontalResolution, src.VerticalResolution);
        BitmapData a = c32.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        BitmapData b = dst.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.WriteOnly, PixelFormat.Format1bppIndexed);
        byte[] fila = new byte[a.Stride]; byte[] sal = new byte[b.Stride];
        int rango = blanco - negro;
        for (int y = 0; y < h; y++) {
            Marshal.Copy(new IntPtr(a.Scan0.ToInt64() + (long)y * a.Stride), fila, 0, a.Stride);
            Array.Clear(sal, 0, sal.Length);
            int fb = (y & 7) * 8;
            for (int x = 0; x < w; x++) {
                int i = x * 4; int lum = (fila[i] * 114 + fila[i + 1] * 587 + fila[i + 2] * 299) / 1000;
                bool bl;
                if (lum >= blanco) bl = true;
                else if (lum <= negro) bl = false;
                else bl = (lum - negro) * 64 > BAYER[fb + (x & 7)] * rango;   // más claro = más puntos blancos
                if (bl) sal[x >> 3] |= (byte)(0x80 >> (x & 7));
            }
            Marshal.Copy(sal, 0, new IntPtr(b.Scan0.ToInt64() + (long)y * b.Stride), b.Stride);
        }
        c32.UnlockBits(a); dst.UnlockBits(b);
        if (!object.ReferenceEquals(c32, src)) c32.Dispose();
        return dst;
    }
}
'@
function Convertir1Bit($img) { [Bn1]::Convertir($img, 230) }

# modo imprimir: cada trabajo es UN solo envío a la impresora con muchas páginas (de varios PDF).
# Mandar de a un archivo es muchísimo más lento en la impresora (usuario, 2026-11-09): se agrupa por impresora y hoja.
# Trabajo: { nombre, impresora, salida?, paginas: [ { archivo, pagina, hoja, horizontal } ] }
$res = @()
foreach ($t in $trabajos) {
    $docs = @{}
    try {
        foreach ($p in $t.paginas) { if (-not $docs.ContainsKey($p.archivo)) { $docs[$p.archivo] = AbrirPdf $p.archivo } }
        $ps = New-Object System.Drawing.Printing.PrinterSettings
        $ps.PrinterName = $t.impresora
        if (-not $ps.IsValid) { throw "impresora no encontrada: $($t.impresora)" }
        $ps.Copies = 1
        $ps.Duplex = [System.Drawing.Printing.Duplex]::Simplex
        if ($t.salida) { $ps.PrintToFile = $true; $ps.PrintFileName = $t.salida }
        $papeles = @{}
        foreach ($pz in $ps.PaperSizes) { if (-not $papeles.ContainsKey($pz.PaperName)) { $papeles[$pz.PaperName] = $pz } }
        # (2026-11-20) Con otro driver la hoja puede llamarse distinto ("ISO A3", "A3 (297 x 420 mm)"): se busca por medida.
        $MM = @{ A4 = 210, 297; A3 = 297, 420; A2 = 420, 594; A1 = 594, 841; A0 = 841, 1189 }
        foreach ($h in $t.paginas) {
            if ($papeles.ContainsKey($h.hoja)) { continue }
            $m = $MM[$h.hoja]
            $hallada = $null
            if ($m) {
                $a = $m[0] / 0.254; $b = $m[1] / 0.254     # mm -> 1/100 de pulgada
                foreach ($pz in $ps.PaperSizes) {
                    $c = [math]::Min($pz.Width, $pz.Height); $l = [math]::Max($pz.Width, $pz.Height)
                    if ([math]::Abs($c - $a) -le 8 -and [math]::Abs($l - $b) -le 8) { $hallada = $pz; break }
                }
            }
            if (-not $hallada) { throw "la impresora $($t.impresora) no tiene hoja $($h.hoja)" }
            $papeles[$h.hoja] = $hallada
        }

        $pd = New-Object System.Drawing.Printing.PrintDocument
        $pd.PrinterSettings = $ps
        $pd.DocumentName = $t.nombre   # nombre en la cola
        $pd.PrintController = New-Object System.Drawing.Printing.StandardPrintController   # sin ventana
        $pd.OriginAtMargins = $false
        $estado = @{ i = 0; docs = $docs; t = $t; papeles = $papeles; n = @($t.paginas).Count }
        $pd.add_QueryPageSettings({ param($s, $e)
            $h = $estado.t.paginas[$estado.i]
            $e.PageSettings.PaperSize = $estado.papeles[$h.hoja]
            $e.PageSettings.Landscape = [bool]$h.horizontal
            $e.PageSettings.Color = $false
        })
        $pd.add_PrintPage({ param($s, $e)
            $q = $estado.t.paginas[$estado.i]
            $pag = $estado.docs[$q.archivo].GetPage([uint32]$q.pagina)
            $mm2 = ($pag.Size.Width * $DIP_MM) * ($pag.Size.Height * $DIP_MM)
            # (2026-11-14) 400 dpi máximo (antes 600) y tope de ~40 Mpx por hoja: lo mismo de nítido en papel,
            # mucho menos para procesar en la PC y en la impresora.
            $dpi = [math]::Min($script:DPI_MAX, [math]::Sqrt(40e6 / ($mm2 / 645.16)))
            $img = Renderizar $pag $dpi ([bool]$q.excel)
            $b = $e.PageBounds   # hoja completa, en 1/100 de pulgada
            # Misma hoja -> escala 1:1. Si difiere (PDF de Excel en Carta) se ajusta a la hoja, centrado.
            $iw = $pag.Size.Width / 0.96; $ih = $pag.Size.Height / 0.96
            $k = [math]::Min($b.Width / $iw, $b.Height / $ih)
            $w = $iw * $k; $h = $ih * $k
            $x = ($b.Width - $w) / 2 - $e.PageSettings.HardMarginX; $y = ($b.Height - $h) / 2 - $e.PageSettings.HardMarginY
            $e.Graphics.DrawImage($img, [float]$x, [float]$y, [float]$w, [float]$h)
            $img.Dispose(); $pag.Dispose()
            $estado.i++
            $e.HasMorePages = ($estado.i -lt $estado.n)
        })
        $pd.Print()
        $pd.Dispose()
        $res += @{ nombre = $t.nombre; ok = $true; paginas = $estado.i }
    } catch { $res += @{ nombre = $t.nombre; ok = $false; error = "$_" } }
    finally { $docs = $null; [GC]::Collect() }
}
Salir $res
