<#ip#>
$nombreAdaptador = $null
foreach ($adapter in Get-NetAdapter)
{
	if ($adapter.status -eq "Up" -And $adapter.InterfaceDescription.Substring(0, 7) -ne "Hyper-V")
	{
		$nombreAdaptador = $adapter.Name
		break
	}
}
$ip = (Get-NetIPConfiguration -InterfaceAlias $nombreAdaptador).ipv4address.ipaddress


<#nombre so#>
$osaux1 = (Get-CimInstance Win32_OperatingSystem -Property *).name.Split("|")[0]
$osaux2 = (Get-CimInstance Win32_OperatingSystem -Property *).osarchitecture
$sonombre = $osaux1 + " " + $osaux2


<#placa-base#>
$pbmarca = (Get-WmiObject win32_baseboard).manufacturer
$pbmodelo = (Get-WmiObject win32_baseboard).product
$pbserie = (Get-WmiObject win32_baseboard).serialnumber
$pbcapacidad = ""
$pbtipo = ""


<#procesador#>
$cpuaux = (Get-WmiObject Win32_Processor).Name.Split(" ")
$cpumarca = $cpuaux[0]
$cpumodelo = $cpuaux[1] + " " + $cpuaux[2]
$cpuserie = (Get-CimInstance win32_processor -Property *).serialnumber
$cpucapacidad = $cpuaux[$cpuaux.Length - 1]
$cputipo = ""


<#ram#>
$rammarca = ""
$rammodelo = ""
$ramserie = "" 
$ramcapacidad = ""
$ramcapacidadT = 0
$ramtipo = ""
$ramaux = Get-WMIObject Win32_PhysicalMemory
foreach ($ram in $ramaux)
{
	$rammarca += "`n" + $ram.manufacturer;
	$ramserie += "`n" + $ram.serialnumber;
	$ramcapacidad += "`n" + $($ram.capacity/1024/1024/1024)+" GB";
	$ramcapacidadT += $ram.capacity;
	if ($ram.SMBIOSMemoryType -eq 25)
	{
		$ramtipo += "`n" + "DDR3"
	}
	if ($ram.SMBIOSMemoryType -eq 26)
	{
		$ramtipo += "`n" + "DDR4"
	}
}
$ramcapacidadT = $ramcapacidadT/1024/1024/1024


<#video#>
$videomarca = ""
$videomodelo = ""
$videoserie = ""
$videocapacidad = "" 
$videotipo = "" 
$videoaux = Get-WmiObject win32_VideoController
foreach ($video in $videoaux)
{
	$videomarca += "`n" + $video.caption.Split("")[0];
	$videomodelo += "`n" + $video.caption.Split("")[0] + $video.caption.Split("")[1] + $video.caption.Split("")[2];
	$videoserie += "`n" + $video.PNPDeviceID;
	$videocapacidad += "`n" + $($video.AdapterRAM/1024/1024/1024) + " GB"; 
}


<#disco#>
$discomarca = ""
$discomodel = ""
$discoserie = ""
$discocapacidad = ""
$discotipo = "" 
$discoaux = Get-PhysicalDisk | Select-Object *;
foreach ($disco in $discoaux)
{
	$discomarca += "`n" + $disco.friendlyname.Split("")[0];
	$discomodel += "`n" + $disco.model;
	$discoserie += "`n" + $disco.serialnumber;
	$discocapacidad += "`n ~" + [math]::Round($($disco.size/1024/1024/1024),2) + " GB";
	$discotipo += "`n" + $disco.bustype; 
}


<#lectora#>


<#excel#>
$rutaexcel = "$(Get-Location)\inventario.xlsx"
if ([System.IO.File]::Exists($rutaexcel))
{
	$excel = New-Object -ComObject Excel.Application
	$libro = $excel.Workbooks.Open($rutaexcel)
}
else
{
	[system.Windows.Forms.MessageBox]::Show("No existe el archivo-formato inventario.xlsx") 
	exit;
} 

<#mostrar excel#>
#$excel.visible = $true

<#escojer hoja por nombre#>
$hoja = $libro.Sheets.Item('SSC')

<#poner valores#>
$hoja.range("g14") = $sonombre 

$hoja.Range("h10") = $ip

$hoja.range("b18") = $pbmarca 
$hoja.range("c18") = $pbmodelo
$hoja.range("d18") = $pbserie
$hoja.range("e18") = $pbcapacidad
$hoja.range("f18") = $pbtipo

$hoja.range("b19") = $cpumarca
$hoja.range("c19") = $cpumodelo 
$hoja.range("d19") = $cpuserie
$hoja.range("e19") = $cpucapacidad
$hoja.range("f19") = $cputipo

$hoja.range("b20") = $rammarca
$hoja.range("c20") = $rammodelo
$hoja.range("d20") = $ramserie
$hoja.range("e20") = $ramcapacidad + "`n--------`n" + $ramcapacidadT + " GB"  
$hoja.range("f20") = $ramtipo

$hoja.range("b21") = $videomarca
$hoja.range("c21") = $videomodelo
$hoja.range("d21") = $videoserie
$hoja.range("e21") = $videocapacidad
$hoja.range("f21") = $videotipo

$hoja.range("b22") = $discomarca
$hoja.range("c22") = $discomodelo
$hoja.range("d22") = $discoserie
$hoja.range("e22") = $discocapacidad
$hoja.range("f22") = $discotipo


<#guardar excel#> 
try
{
	$libro.SaveAs($rutaexcel);
	$libro.close($true);
	[System.Windows.Forms.MessageBox]::Show("Grabado correctamente"); 
}
catch [System.Exception]
{
	[System.windows.Forms.MessageBox]::Show("Error al Grabar,, vuelva a intentar"); 
}
