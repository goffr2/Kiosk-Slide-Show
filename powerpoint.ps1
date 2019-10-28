New-Item -Force C:\Users\kuser\Slideshow -Type Directory
New-Item -Force C:\Users\kuser\Slideshow\logs -Type Directory

copy-item \\itnas\netshare\KioskSlideShow\Slideshow.ppsx C:\Users\kuser\Slideshow\Slideshow.ppsx -Recurse -Force
Start-Process C:\Users\kuser\Slideshow\Slideshow.ppsx
$A = 1
While ($A = 2)
{
$LastWritten = Get-Item \\itnas\netshare\KioskSlideShow\Slideshow.ppsx |Select-Object -ExpandProperty  LastWriteTime
start-sleep -Seconds 10
$CurrentWritten = Get-Item \\itnas\netshare\KioskSlideShow\Slideshow.ppsx |Select-Object -ExpandProperty  LastWriteTime
if ($CurrentWritten -ne $LastWritten)
{
	$Process = Get-Process POWERPNT |select-object -ExpandProperty Id
	Stop-Process $Process
	copy-item \\itnas\netshare\KioskSlideShow\Slideshow.ppsx C:\Users\kuser\Slideshow\Slideshow.ppsx -Recurse -Force
	Start-Process C:\Users\kuser\Slideshow\Slideshow.ppsx
}
else{

}

}
