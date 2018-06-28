'Adhoc_starting'
'VBScript: José Gerlan Pereira de Paula, 23/06/2015'

'Nome do computador'
Dim WshNetwork
Set WshNetwork = CreateObject("WScript.Network")
ComputerName = WshNetwork.ComputerName
Rem WScript.Echo ComputerName
'------------------'
'Nome Admin'
'Inutilizado no momento...'
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkLoginProfile")
Dim indiceColItems
Dim nameAdmin
indiceColItems = 0
nameAdmin = ""
For Each objItem in colItems
	if indiceColItems = 1 Then
		nameAdmin = "" & objItem.Name
	End If
	indiceColItems = indiceColItems +1
Next
userAdmin = """"& nameAdmin &""""
'------------------'
Dim inputRede
Rem InputBox(prompt[, title][, default][, xpos][, ypos][, helpfile, context])
inputRede = Inputbox("Defina o nome da rede:","Nome da rede",ComputerName & "_redeauxiliar")

If inputRede = "" Then
	MsgBox "Script encerrado! Click OK para sair"
Else
	Dim inputSenha
	inputSenha = Inputbox("Defina a senha:","Senha da rede","suasenhaaqui")
	If inputSenha = "" Then
		MsgBox "Script encerrado! Click OK para sair"
	Else
		WScript.Echo "Obs: Não se esqueça de configurar manualmente o compartilhamento de internet (e ip) com a nova rede criada na central de rede e compartilhamento! Para cada 'nome' de rede só precisa uma vez."
		Dim oShell
		Dim comandoAux
		Dim comandoFull
		
		Set oShell = WScript.CreateObject("WScript.shell")
		comandoAux = "echo Iniciando modo ADHOC! & echo Nome da rede: " & inputRede & " & "
		comandoFull = comandoFull & comandoAux
		comandoAux = "netsh wlan set hostednetwork mode=allow ssid=" & inputRede & " key=" & inputSenha
		comandoFull = comandoFull & comandoAux
		comandoAux = " & netsh wlan start hostednetwork"
		comandoFull = comandoFull & comandoAux
		comandoFull = comandoFull & " & echo Deseja cancelar conexao? Pressione enter para sair!"
		comandoAux = "netsh wlan stop hostednetwork"
		oShell.Run "cmd  /K title= Adhoc Starting & @echo off & " & comandoFull & " & pause & " & comandoAux & " & exit"
	End If
End If

Set oShell = Nothing