Dim s_data, d, nd,objStr1,objStr2,objStr3,strIP


Sub StartUp()
Dim x,y,xw,yh,arrFileLines()
ForReading = 1
strIP = "config\ipadress.ini"
objStr1 = "<TABLE border='1'>" & vbCrLf & "<TR>" & vbCrLf & "<TH><input type='CheckBox' name='CheckboxOption' id='pin0' onClick='ForEach'></TD>" & vbCrLf & "<TH><img src='config/5.png'>" & vbCrLf & "<TH><FONT SIZE='' COLOR='#0000FF'>Узел:</FONT>" & vbCrLf & "<TH><FONT SIZE='' COLOR='#0000FF'>IP Adress:</FONT>" & vbCrLf & "<TH><FONT SIZE='' COLOR='#0000FF'>Ping ms</FONT>" & vbCrLf & "<TH><img src='config/1.png'>" & vbCrLf & "<TH><img src='config/2.png'>" & vbCrLf & "<TH><FONT SIZE='2' COLOR='#FF0000'>Потери %</FONT>" & vbCrLf & "</TR>" & vbCrLf
objStr3 = "</TABLE>"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objIP = objFSO.OpenTextFile(strIP, ForReading)
	n = 0
	Do Until objIP.AtEndOfStream
	Redim Preserve arrFileLines(n)
	arrFileLines(n) = objIP.ReadLine
	arr = Split(arrFileLines(n), "|")
	text0 = arr(0)
	text1 = arr(1)
	text2 = arr(2)
	iparr = Split(text2, ".")
	ip0 = iparr(0)
	ip1 = iparr(1)
	ip2 = iparr(2)
	ip3 = iparr(3)
	'msgbox ip0 & "." & ip1 & "." & ip2 & "." & ip3
	if text0 = "True" then flag = "checked" else flag = ""
	objStr2 = "<TR>"& vbCrLf &"<TD><input type='CheckBox' name='CheckboxOption' id='pin" & n+1 & "' " & flag & "></TD>"& vbCrLf &"<TH id='Dostup"& n+1 &"'><img src='config/4.png' alt='"& text1 & vbCrLf & text2 & vbCrLf & "ON-OFF'></TH>"& vbCrLf &"<TH id='t"& n+1 &"' style='text-align: left'>" & text1 &"</TH>"& vbCrLf &"<TH id='CompName"& n+1 &"' style='text-align: left'>" & text2 & "</TH>" & vbCrLf &"<TH id='Time"& n+1 &"' style='text-align: left' value='-1'>-1</TH>"& vbCrLf &"<TH id='d" & n+1 & "' style='text-align: left'>0</TH>"& vbCrLf &"<TH id='nd"& n+1 &"' style='text-align: left'>0</TH>"& vbCrLf &"<TH id='pr"& n+1 &"' style='text-align: left'>0</TH>"& vbCrLf &"</TR> "& vbCrLf 
	objStr_new = objStr_new  & objStr2
	n = n + 1
	Loop
	objIP.Close
	Set objIP = Nothing

	xw = 560
	yh = 180 + 25*arrLength(arrFileLines)
	x = (window.screen.width - xw) / 2
	y = (window.screen.height - yh) / 2
	If x < 0 Then x = 0
	If y < 0 Then y = 0
	window.resizeTo xw,yh
	window.moveTo x,y
	
	objWin.InnerHTML = objStr1 & objStr_new & objStr3
	Version_Div.InnerHTML = "Версия: " & oHTA.Version
	document.title="Ping IP v" & oHTA.Version
	ProcessList.InnerHTML = time
	sTime.innerHTML = time
	iTimerID = window.setInterval("RefreshList", 1000)
	if pin0.Checked then ForEachTrue 
	'else ForEachFalse
	'end if
End Sub

Function arrLength(vArray)
	ItemCount = 0
		For ItemIndex = 0 To UBound(vArray)
			If Not(vArray(ItemIndex)) = Empty Then
				ItemCount = ItemCount + 1
			End If
		Next
	arrLength = ItemCount
End Function
Function countEmptySlots(arr)
    Dim x, c
    c = 0
    For x = 0 To ubound(arr)
    	If arr(x) = vbUndefined Then c = c + 1
    Next
    countEmptySlots = c
End Function
Sub ForEach
If pin0.Checked then ForEachTrue '
If Not pin0.Checked then ForEachFalse
 End Sub
Sub ForEachTrue
 For Each checkbox In CheckboxOption
  checkbox.Checked = True 
 Next
 End Sub
Sub ForEachFalse
	For Each checkbox In CheckboxOption
		checkbox.Checked = False
	Next
 End Sub
Function PING_S(t, CompName, Dostup, time0, d, nd)
	if CompName <> "" then
		Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & CompName & "'")
		For Each objStatus In objPing
			If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then
				Dostup.innerHTML = "<IMAGE id='img' SRC='config/2.png' alt='"& t & vbCrLf & CompName & vbCrLf & "OFF'>"
				time0.innerHTML = "-1"
				nd.innerHTML = nd.innerHTML + 1
				Logpin = document.getElementByID("setLog").checked
				If Logpin Then
				Set objLogs = CreateObject("Scripting.FileSystemObject")
				If Not objLogs.FileExists("log\" & CompName & "_" & Date & ".log") Then
				Set writelogs = objLogs.OpenTextFile("log\" & CompName & "_" & Date & ".log", 8, True)
				writelogs.WriteLine(Now & " создание лога!!")
				writelogs.close
				End If
				Set f = objLogs.OpenTextFile("log\" & CompName & "_" & Date & ".log", 1)
				buffer = f.ReadAll()
				f.Close()
				Set f = objLogs.OpenTextFile("log\" & CompName & "_" & Date & ".log", 2, True)
				s = Now & " " & CompName & " '" & time0.innerHTML & "' не доступен " & nd.innerHTML
				f.WriteLine(s)
				f.WriteLine(buffer)
				f.Close()
				end if
				
			Else
				nTIME = (time0.innerHTML * (nd.innerHTML + d.innerHTML) + objStatus.ResponseTime) \ (nd.innerHTML + d.innerHTML + 1)
				if nTIME < 100 Then Dostup.innerHTML = "<IMAGE id='img' SRC='config/1.gif'alt='"& t & vbCrLf & CompName & vbCrLf & "ON'>"
				if nTIME >= 100 Then Dostup.innerHTML = "<IMAGE id='img' SRC='config/1.gif'alt='"& t & vbCrLf & CompName & vbCrLf & "ON'>"
				time0.innerHTML = nTIME
				d.innerHTML = d.innerHTML + 1
				Logpin = document.getElementByID("setLog").checked
				If Logpin Then
				Set objLogs = CreateObject("Scripting.FileSystemObject")
				If Not objLogs.FileExists("log\" & CompName & "_" & Date & ".log") Then
				Set writelogs = objLogs.OpenTextFile("log\" & CompName & "_" & Date & ".log", 8, True)
				writelogs.WriteLine(Now & " создание лога!!")
				writelogs.close
				End If
				Set f = objLogs.OpenTextFile("log\" & CompName & "_" & Date & ".log", 1)
				buffer = f.ReadAll()
				f.Close()
				Set f = objLogs.OpenTextFile("log\" & CompName & "_" & Date & ".log", 2, True)
				s = Now & " " & CompName & " '" & time0.innerHTML & "' доступен " & d.innerHTML
				f.WriteLine(s)
				f.WriteLine(buffer)
				f.Close()
				end if
			End if
		Next
		
	end if

End Function
Function Procent(d0, nd0, pr0)
np = (d0.innerHTML * 1 + nd0.innerHTML * 1)
	if np <> 0 then
		pr0.innerHTML = (nd0.innerHTML * 100)\np
	end if
end Function
sub proc()
	Dim arrFileLines()
	sTime.innerHTML = time
	ForReading = 1
	strIP = "config\ipadress.ini"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objIP = objFSO.OpenTextFile(strIP, ForReading)
	n = 0
	Do Until objIP.AtEndOfStream
	Redim Preserve arrFileLines(n)
	arrFileLines(n) = objIP.ReadLine
	arr = Split(arrFileLines(n), "|")
	text1 = arr(1)
	text2 = arr(2)
	call Procent(document.getElementByID("d" & n+1), document.getElementByID("nd" & n+1), document.getElementByID("pr" & n+1))
		n = n + 1
	Loop
	objIP.Close
	Set objIP = Nothing
end sub
Sub TreeWalk(tag)
    For Each t In tag.childNodes
        msgbox t.nodeName
        If t.hasChildNodes() Then TreeWalk(t)
    Next
End Sub
sub pingfun()
	Dim arrFileLines()
	sTime.innerHTML = time
	ForReading = 1
	strIP = "config\ipadress.ini"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objIP = objFSO.OpenTextFile(strIP, ForReading)
	n = 0
	Do Until objIP.AtEndOfStream
	Redim Preserve arrFileLines(n)
	arrFileLines(n) = objIP.ReadLine
	arr = Split(arrFileLines(n), "|")
	
	text1 = arr(1)
	text2 = arr(2)
	boolpin = document.getElementByID("pin" & n+1).checked
	If boolpin Then
	
		'call PING_S(text2, document.getElementByID("Dostup" & n+1), document.getElementByID("Time" & n+1), document.getElementByID("d" & n+1), document.getElementByID("nd" & n+1))
		call PING_S(text1, document.getElementByID("CompName" & n+1).innerHTML, document.getElementByID("Dostup" & n+1), document.getElementByID("Time" & n+1), document.getElementByID("d" & n+1), document.getElementByID("nd" & n+1))
	'msgbox document.getElementByID("CompName" & n+1).innerHTML & " " & document.getElementByID("Time" & n+1).innerHTML
	End If
	n = n + 1
	Loop
	objIP.Close
	Set objIP = Nothing

	call proc()
end sub
sub start_onclick()

	Dim boolSetTimeProcessing
	pingfun
	boolSetTimeProcessing = document.getElementByID("SetTimeProcessing").checked
	If boolSetTimeProcessing Then
	start_timer
	End If
end sub

sub objXML()
	Dim xmlDocument, compValue, sysValue, xmlNewNode
	SET xmlDoc=CreateObject("Msxml2.DOMDocument.3.0") 
		xmlDoc.async="false"
		xmlDoc.load("config\APS.XML") 
	Set AValue=xmlDoc.documentElement.selectSingleNode("COMPNAME") 
		
		boolpin1 = AValue.getAttribute("pin1")
		if boolpin1 = "True" then
		With document.getElementByID("pin1")
		.checked = Not .checked
		End With
		end if
		CompName1.innerHTML = AValue.getAttribute("CompName1")
		t1.innerHTML = AValue.getAttribute("t1")
		

end sub

sub XMLsave_onclick()
	Dim arrFileLines()
	ForReading = 1
	strIP = "config\ipadress.ini"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objIP = objFSO.OpenTextFile(strIP, ForReading)
	n = 0
	Do Until objIP.AtEndOfStream
	Redim Preserve arrFileLines(n)
	arrFileLines(n) = objIP.ReadLine
	arr = Split(arrFileLines(n), "|")
	text1 = arr(1)
	text2 = arr(2)
	n = n + 1
	Loop
	'msgbox arrLength(arrFileLines)
	objIP.Close
	Set objIP = Nothing
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	set c2f=objFSO.OpenTextFile("config\ipadress.ini",2,-1)
	'msgbox arrLength(arrFileLines)
	for i = 0 to arrLength(arrFileLines)*1 'objSIP.value
	boolpin = document.getElementByID("pin" & i+1).checked 
	If boolpin Then objTrue = "True|" else objTrue = "False|"
	objText = objTrue & document.getElementByID("t" & i+1).innerHTML & "|" & document.getElementByID("CompName" & i+1).innerHTML
	'msgbox objText
	c2f.WriteLine objText
	next
	c2f.close
end sub

Sub start_timer()
 s_data = CLng(document.getElementById("SetTime").Value * sTs.Value)
 Call MyTimer()
End Sub
Sub MyTimer()
 Dim TimerID
 s_data = s_data - 1
 If s_data < 0 Then
 Call Alert()
 Exit Sub
 End If
 TimerID=SetTimeout("MyTimer()",1000)
End Sub
Sub Alert()
	Call start_onclick()
End Sub
Sub RefreshList
    strHTML = time
    ProcessList.InnerHTML = strHTML
End Sub
Sub WindowOnLoad
	'XMLsave_onclick()
	StartUp
	'XMLsave_onclick()
	'objXML
End Sub
sub CloseButton_onclick()
'XMLsave_onclick()
end sub
sub options_onclick()
end sub
sub optione_onclick()
end sub
Sub OnClickButtonbtnSET
Dim x,y,xw,yh,arrFileLines()
ForReading = 1
strIP = "config\ipadress.ini"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objIP = objFSO.OpenTextFile(strIP, ForReading)
	n = 0
	Do Until objIP.AtEndOfStream
	Redim Preserve arrFileLines(n)
	arrFileLines(n) = objIP.ReadLine
	n = n + 1
	Loop
	objIP.Close
	Set objIP = Nothing
	m = 150 + (25*arrLength(arrFileLines))
	varReturn = window.ShowModalDialog("Setting.hta", Fill_Dict_TextBox, "dialogHeight:" & m & "px;dialogWidth:300px") 
End Sub