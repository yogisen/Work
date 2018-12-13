'==========================================================================
' NAME: Script Audit PC
' DATE  : 090505
' COMMENT: V1.0
'==========================================================================

On Error Resume Next

Dim objNetwork, objDrive, intDrive, intNetLetter
Dim FSO, ObjFile, WshNetwork, SrtNewText, strText
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshNetwork = WScript.CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")

hostname = WshNetwork.computername
username = WshNetwork.username
domainname = WshNetwork.UserDomain


'==========================================================================
'mettez ici l'emplacement du dossier dans lequel sera enregistrer le rapport
'si ce dossier n'existe pas, il sera crée
path = ""

'sharename = "logs_"&username&""

If  objFSO.FolderExists (""&path&"") Then
  Else
Set objFolder = objFSO.CreateFolder (""&path&"")
End If

'save softwares configuration In a txt file
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile(""&path&""& hostname & " - "&username&".txt", True)
filepath = ""&path&""& hostname & " - "&username&".txt"

'==========================================================================
'INFOS GENERALES (Nom du poste, nom de l'utilisateur actuel)
'==========================================================================
objTextFile.WriteLine "== GENERAL =========================================================="&VbCrLf&""
objTextFile.WriteLine "Nom du poste : "& hostname &""
Set objWMISvc = GetObject( "winmgmts:\\.\root\cimv2" )
Set colItems = objWMISvc.ExecQuery( "Select * from Win32_ComputerSystem", , 48 )
For Each objItem in colItems
    strComputerDomain = objItem.Domain
    If objItem.PartOfDomain Then
        objTextFile.WriteLine "Domaine : " & strComputerDomain
    Else
        objTextFile.WriteLine "Workgroup : " & strComputerDomain
    End If
objTextFile.WriteLine "Marque : " & objItem.Manufacturer
objTextFile.WriteLine "Modele : " & objItem.Model
Next
Set SNSet = GetObject("winmgmts:").InstancesOf ("Win32_BIOS")
for each SN in SNSet
objTextFile.WriteLine "Serial : " & SN.SerialNumber
Next
objTextFile.WriteLine

'==========================================================================
'INFOS SYSTEME
'==========================================================================
objTextFile.WriteLine "== SYSTEME EXPLOITATION ============================================="&VbCrLf&""

Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")

For Each objOperatingSystem in colOperatingSystems

    objTextFile.WriteLine "OS : " & objOperatingSystem.Caption
    objTextFile.WriteLine "Version: " & objOperatingSystem.Version
    objTextFile.WriteLine "Service Pack : " & objOperatingSystem.ServicePackMajorVersion
    objTextFile.WriteLine "Numero Serie : " & objOperatingSystem.SerialNumber
    objTextFile.WriteLine "Type Version : " & objOperatingSystem.BuildType
    dtmConvertedDate.Value = objOperatingSystem.InstallDate
    dtmInstallDate = dtmConvertedDate.GetVarDate
    objTextFile.WriteLine "Date d'Installation : " & dtmInstallDate
    objTextFile.WriteLine
Next

'==========================================================================
'INFORMATIONS MATERIEL
'==========================================================================


objTextFile.WriteLine "== PROCESSEUR & RAM ================================================="&VbCrLf&""

'-- Processeur : nom
    Set colComputer = objWMIService.ExecQuery("Select * from Win32_Processor")
    For Each objComputer In colComputer
        objTextFile.WriteLine "Processeur :"& vbTab &"" & objComputer.Name
    Next

'-- Processeur : type
Set colSettings = objWMIService.ExecQuery _
 ("SELECT * FROM Win32_Processor")
For Each objProcessor In colSettings
    objTextFile.WriteLine "Type Processeur : "& vbTab &"" & objProcessor.Description
Next
objTextFile.WriteLine
'-- RAM
Set colComputer = objWMIService.ExecQuery _
("Select * from Win32_ComputerSystem")
For Each objComputer In colComputer
    objTextFile.WriteLine "RAM Total : "&FormatNumber (objComputer.TotalPhysicalMemory /1048576+1, 0) & "Mo"
    objTextFile.WriteLine
Next

For Each objMem In GetObject("winmgmts:{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2").InstancesOf("Win32_PhysicalMemory")
    objTextFile.WriteLine ""& vbTab &"+"&  objMem.BankLabel &" : "& objMem.Capacity/1024/1024 & " Mo"
Next

' Disques et partition
objTextFile.WriteLine ""&VbCrLf&"== DISQUE DUR & PARTITIONS =========================================="&VbCrLf&""

objTextFile.WriteLine "Disques Physiques :"
Set colItems = objWMIService.ExecQuery(_
"Select * from Win32_DiskDrive")

            maxsize = 0
            drive = 0

            For Each objItem in colItems
If objItem.Partitions > 0 Then
objTextFile.WriteLine ""&vbTab &"Nom : "& vbTab &"" & objItem.Name
objTextFile.WriteLine ""&vbTab &"Taille : " & Int(objItem.Size /(1073741824)) & " GB"
objTextFile.WriteLine ""&vbTab &"Modèle : " & objItem.Model
objTextFile.WriteLine ""&vbTab &"Nbr de partition(s) : "& vbTab &"" & objItem.Partitions &VbCrLf&""

maxsize = maxsize + Int(objItem.Size /(1073741824))
drive = drive + 1
Else
End If
Next

objTextFile.WriteLine"Espace Total : "&maxsize&" Go dans "&drive&" disques durs"

objTextFile.WriteLine ""&VbCrLf&"Partitions"

    '-- Partitions et Tx Occupation
        Set colDisks = objWMIService.ExecQuery _
        ("Select * from Win32_LogicalDisk Where DriveType = 3")
        For Each objDisk In colDisks
            intFreeSpace = objDisk.FreeSpace /1000000000             'espace disque libre
            intTotalSpace = objDisk.Size /1000000000
            pctFreeSpace = intFreeSpace / intTotalSpace
            intOccupedSpace = intTotalSpace - intFreeSpace
            pctOccupedSpace = intOccupedSpace / intTotalSpace
            Disk = objDisk.DeviceID
    objTextFile.WriteLine ""& vbTab &" "&Disk&" Occupé à "&FormatPercent(pctOccupedSpace,0)&" - Taille Totale : " &FormatNumber (intTotalSpace,0)&" Go ("&FormatNumber(intFreeSpace,0)&" Go utilisés)"
Next

'==========================================================================
'RESEAU
'==========================================================================
objTextFile.WriteLine ""&VbCrLf&"== RESEAU ==========================================================="&VbCrLf&""
Set colNicConfigs = objWMIService.ExecQuery _
 ("Select * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")


Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colAdapters = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

n = 1
objTextFile.WriteLine

For Each objAdapter in colAdapters
   objTextFile.WriteLine "Carte Réseau - " & objAdapter.Description
  objTextFile.WriteLine "  Adresse MAC : " & vbTab & objAdapter.MACAddress

   If Not IsNull(objAdapter.IPAddress) Then
      For i = 0 To UBound(objAdapter.IPAddress)
         objTextFile.WriteLine "  Adresse IP : " &vbTab & vbTab & objAdapter.IPAddress(i)
      Next
   End If

   If Not IsNull(objAdapter.IPSubnet) Then
      For i = 0 To UBound(objAdapter.IPSubnet)
         objTextFile.WriteLine "  Sous Réseau: " &vbTab & vbTab & objAdapter.IPSubnet(i)
      Next
   End If

   If Not IsNull(objAdapter.DefaultIPGateway) Then
      For i = 0 To UBound(objAdapter.DefaultIPGateway)
         objTextFile.WriteLine "  Passerelle : " &vbTab & vbTab & objAdapter.DefaultIPGateway(i)
      Next
   End If

    If IsNull(objAdapter.DefaultIPGateway) Then
    objTextFile.WriteLine
            objTextFile.WriteLine "  Pas de DNS Specifié"
 Else

   If Not IsNull(objAdapter.DNSServerSearchOrder) Then
      For i = 0 To UBound(objAdapter.DNSServerSearchOrder)
         objTextFile.WriteLine "  DNS :" & objAdapter.DNSServerSearchOrder(i)
      Next
   End If

   objTextFile.WriteLine "  Domaine DNS : " & objAdapter.DNSDomain

   If Not IsNull(objAdapter.DNSDomainSuffixSearchOrder) Then
      For i = 0 To UBound(objAdapter.DNSDomainSuffixSearchOrder)
         objTextFile.WriteLine "    Liste de recherche suffixe DNS : " & _
             objAdapter.DNSDomainSuffixSearchOrder(i)
      Next
   End If
End If

' Configuration WINS
 If IsNull(objAdapter.WINSPrimaryServer) And IsNull(objAdapter.WINSSecondaryServer) Then
  objTextFile.WriteLine
   objTextFile.WriteLine "Pas de WINS specifié"
    objTextFile.WriteLine
   Else
   objTextFile.WriteLine "  WINS"
   objTextFile.WriteLine "  ----"
   objTextFile.WriteLine "    Serveur WINS Primaire :   " & objAdapter.WINSPrimaryServer
   objTextFile.WriteLine "    Serveur WINS Secondaire : " & objAdapter.WINSSecondaryServer
 End If
  n = n + 1
Next

Function WMIDateStringToDate(utcDate)
   WMIDateStringToDate = CDate(Mid(utcDate, 5, 2)  & "/" & _
       Mid(utcDate, 7, 2)  & "/" & _
           Left(utcDate, 4)    & " " & _
               Mid (utcDate, 9, 2) & ":" & _
                   Mid(utcDate, 11, 2) & ":" & _
                      Mid(utcDate, 13, 2))
End Function

objTextFile.WriteLine ""&VbCrLf&"=========================================================================="
objTextFile.WriteLine "==                       File saved @ "&Now&"               =="
objTextFile.WriteLine "=========================================================================="&VbCrLf&""

objTextFile.close
WScript.Quit(1)