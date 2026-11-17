<#
  Lanzador de retos y examen final (Windows + Word 365 + Edge o Chrome).
  Abre el documento base en Word (75% de la pantalla) y el panel de indicaciones (25% restante).

  Uso (normalmente mediante los archivos .bat):
    Iniciar_Reto1.bat | Iniciar_Reto2.bat | Iniciar_Reto3.bat | Iniciar_Examen_Final.bat
    Opciones:  -Modo Derecha    Word 75% de ancho / panel 25% de ancho (por defecto: Abajo, 75% de alto / 25% de alto)
               -Nuevo           descarta el avance anterior y copia un documento base limpio
               -Url https://... usa el panel publicado en tu sitio web
    Reacomodar_Ventanas.bat     vuelve a colocar Word y el panel ya abiertos (75% / 25%)
#>
param(
  [ValidateSet('final','reto1','reto2','reto3')][string]$Examen = 'final',
  [ValidateSet('Abajo','Derecha')][string]$Modo = 'Abajo',
  [ValidateRange(50,90)][int]$PorcentajeWord = 75,
  [string]$Url = '',
  [switch]$Nuevo,
  [switch]$SoloAcomodar
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinTools {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
}
"@
[void][WinTools]::SetProcessDPIAware()   # coordenadas en pixeles reales

$raiz = $PSScriptRoot
$mapa = @{
  final = @{ html='examen_final.html';   carpeta='Proyecto_Final'; doc='Proyecto_Final.docx'; base='assets\docs\final\Proyecto_Final_Base.docx';
             extras=@('assets\docs\final\Texto_Complementario.docx','assets\docs\final\Datos_Portafolio.txt','assets\docs\final\Grafico_Financiero.jpg');
             panel='*Examen TBox - Proyecto Final*'; word='*Proyecto_Final*' }
  reto1 = @{ html='reto_formato.html';   carpeta='Reto_1'; doc='Reto_1.docx'; base='assets\docs\reto1\Reto1_Base.docx';
             extras=@('assets\images\logo_Corporativo.jpg'); panel='*Examen TBox - Reto 1*'; word='*Reto_1*' }
  reto2 = @{ html='reto_formato2.html';  carpeta='Reto_2'; doc='Reto_2.docx'; base='assets\docs\reto2\Reto2_Base.docx';
             extras=@('assets\images\bono.jpg','assets\images\dinero.png'); panel='*Examen TBox - Reto 2*'; word='*Reto_2*' }
  reto3 = @{ html='reto_formato3.html';  carpeta='Reto_3'; doc='Reto_3.docx'; base='assets\docs\reto3\Reto3_Base.docx';
             extras=@(); panel='*Examen TBox - Reto 3*'; word='*Reto_3*' }
}
$cfg = $mapa[$Examen]

# Distribucion de la pantalla (area de trabajo, sin barra de tareas)
$wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$p = $PorcentajeWord / 100.0
if ($Modo -eq 'Abajo') {
  $wx=$wa.X; $wy=$wa.Y; $ww=$wa.Width; $wh=[int]($wa.Height*$p)
  $px=$wa.X; $py=$wa.Y+$wh; $pw=$wa.Width; $ph=$wa.Height-$wh
} else {
  $wx=$wa.X; $wy=$wa.Y; $ww=[int]($wa.Width*$p); $wh=$wa.Height
  $px=$wa.X+$ww; $py=$wa.Y; $pw=$wa.Width-$ww; $ph=$wa.Height
}
function Mover($h,$x,$y,$w,$hh) { [void][WinTools]::ShowWindow($h,9); [void][WinTools]::SetWindowPos($h,[IntPtr]::Zero,$x,$y,$w,$hh,0x0040) }
function BuscarWord($patron) { Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -like $patron } | Select-Object -First 1 }
function BuscarPanel($patron) { Get-Process | Where-Object { $_.ProcessName -match 'msedge|chrome' -and $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -like $patron } | Select-Object -First 1 }

# Solo reacomodar: vuelve a colocar las ventanas ya abiertas (si no se indica el reto, busca todos)
if ($SoloAcomodar) {
  $lista = if ($PSBoundParameters.ContainsKey('Examen')) { @($Examen) } else { @($mapa.Keys) }
  $movidas = 0
  foreach ($k in $lista) {
    $w = BuscarWord $mapa[$k].word;  if ($w) { Mover $w.MainWindowHandle $wx $wy $ww $wh; $movidas++ }
    $b = BuscarPanel $mapa[$k].panel; if ($b) { Mover $b.MainWindowHandle $px $py $pw $ph; $movidas++ }
  }
  if ($movidas -eq 0) { Write-Warning 'No se encontraron ventanas de Word ni del panel para reacomodar.' } else { Write-Host 'Ventanas reacomodadas.' }
  exit
}

# 1) Carpeta de trabajo (Documentos\TBox_Word\<reto>) con el documento base y los archivos del reto
$trabajo = Join-Path (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'TBox_Word') $cfg.carpeta
$doc = Join-Path $trabajo $cfg.doc
if ($Nuevo -and (Test-Path $trabajo)) { Remove-Item $trabajo -Recurse -Force }
if (-not (Test-Path $trabajo)) { New-Item -ItemType Directory -Path $trabajo | Out-Null }
if (-not (Test-Path $doc)) {
  Copy-Item (Join-Path $raiz $cfg.base) $doc
  foreach ($e in $cfg.extras) { Copy-Item (Join-Path $raiz $e) $trabajo -Force }
}

# 2) Abrir Word con el documento y acomodar la ventana
Write-Host 'Abriendo Word...'
Start-Process -FilePath $doc
$hWord = [IntPtr]::Zero
for ($i=0; $i -lt 80 -and $hWord -eq [IntPtr]::Zero; $i++) {
  Start-Sleep -Milliseconds 500
  $w = BuscarWord $cfg.word
  if ($w) { $hWord = $w.MainWindowHandle }
}
if ($hWord -ne [IntPtr]::Zero) { Mover $hWord $wx $wy $ww $wh } else { Write-Warning 'No se pudo ubicar la ventana de Word; acomodela manualmente o use Reacomodar_Ventanas.bat.' }

# 3) Abrir el panel de indicaciones (Edge o Chrome en modo aplicacion)
if ($Url -eq '') { $Url = ([System.Uri](Join-Path $raiz $cfg.html)).AbsoluteUri }
$sep = if ($Url -match '\?') { '&' } else { '?' }
$Url = $Url + $sep + 'modo=' + $Modo.ToLower()
$candidatos = @(
  "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
  "$env:ProgramFiles\Google\Chrome\Application\chrome.exe", "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe")
$nav = $candidatos | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $nav) { Write-Warning 'No se encontro Edge ni Chrome. Se abre el panel en el navegador predeterminado.'; Start-Process $Url; exit }
$perfil = Join-Path $env:TEMP ('tbox_panel_' + $Examen)
Write-Host 'Abriendo panel de indicaciones...'
Start-Process -FilePath $nav -ArgumentList "--app=$Url", "--user-data-dir=`"$perfil`"", '--no-first-run', '--no-default-browser-check'
$hPanel = [IntPtr]::Zero
for ($i=0; $i -lt 40 -and $hPanel -eq [IntPtr]::Zero; $i++) {
  Start-Sleep -Milliseconds 500
  $b = BuscarPanel $cfg.panel
  if ($b) { $hPanel = $b.MainWindowHandle }
}
if ($hPanel -ne [IntPtr]::Zero) { Mover $hPanel $px $py $pw $ph } else { Write-Warning 'No se pudo ubicar la ventana del panel; use el boton Acoplar 25 % o Reacomodar_Ventanas.bat.' }
Write-Host "Listo. Trabaja en Word y guarda tu archivo en: $trabajo"
