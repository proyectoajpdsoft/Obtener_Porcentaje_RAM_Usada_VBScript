' Obtener el porcentaje de memoria RAM usada en el equipo

dim memoriaUsada, porcentajeMemoriaUsada
dim memoriaTotal

' Obtenemos la memoria RAM usada del equipo actual mediante WMI
equipo = "."
set oWMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & equipo & "\root\cimv2")
set cTotalRAM = oWMI.ExecQuery("Select * from Win32_ComputerSystem")

' Obtenemos la memoria RAM total del equipo
memoriaTotal = 0
for each slotRAM in cTotalRAM
	memoriaTotal = slotRAM.totalPhysicalMemory
next

' Obtenemos la memoria RAM libre en el equipo
set cRAMUsada = oWMI.ExecQuery("Select freePhysicalMemory from Win32_OperatingSystem")
for each slotRAM in cRAMUsada
	memoriaUsada = slotRAM.freePhysicalMemory * 1024
next

' Comprobamos que no se haya producido error en la obtención del total de RAM
on error resume next
siError = cTotalRAM.Count
if (err.number <> 0) then
  siError = true
else
  siError = false
end if
on error goto 0

' Comprobamos que no se haya producido error en la obtención del total de RAM usada
on error resume next
siError = cRAMUsada.Count
if (err.number <> 0) then
  siError = true
else
  siError = false
end if
on error goto 0

if (not siError and memoriaTotal <> 0) then
	' Obtenemos el porcentaje de RAM usada en base al total de memoria RAM y el total usada
	porcentajeMemoriaUsada = round (100 - (memoriaUsada / memoriaTotal) * 100, 2)
	' Mostramos el resultado por consola
	Wscript.StdOut.WriteLine "Memoria RAM usada: " & porcentajeMemoriaUsada & "%"
else
    Wscript.StdOut.WriteLine "No se ha podido obtener la Memoria RAM usada"
end if
Wscript.StdOut.flush
WScript.Quit