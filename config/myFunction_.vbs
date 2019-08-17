Dim s_data, d, nd
i = 0
'd = 0
'nd = 0
 Sub StartUp()
 Dim x,y,xw,yh
 xw = 550
 yh = 410
 x = (window.screen.width - xw) / 2
 y = (window.screen.height - yh) / 2
 If x < 0 Then x = 0
 If y < 0 Then y = 0
 window.resizeTo xw,yh
 window.moveTo x,y
 'msgbox "x = " & x & " y = " & y
 End Sub
Function PING_S(CompName, Dostup, time0, d, nd)
	if CompName <> "" then
		Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & CompName & "'")
		For Each objStatus In objPing
			If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then
				Dostup.innerHTML = "<IMAGE id='img' SRC='config/2.png' alt='OFF'>"
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
				if nTIME < 100 Then Dostup.innerHTML = "<IMAGE id='img' SRC='config/1.png'alt='ON'>"
				if nTIME >= 100 Then Dostup.innerHTML = "<IMAGE id='img' SRC='config/3.png'alt='ON'>"
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
		'if pr0.innerHTML < 70 then
		'pr0.innerHTML = "<TH COLOR='#000000'>" & pr0.innerHTML & "</TH>"
		'end if
	end if
end Function
sub proc()
	call Procent(d1, nd1, pr1)
	call Procent(d2, nd2, pr2)
	call Procent(d3, nd3, pr3)
	call Procent(d4, nd4, pr4)
	call Procent(d5, nd5, pr5)
	call Procent(d6, nd6, pr6)
	call Procent(d7, nd7, pr7)
	call Procent(d8, nd8, pr8)
	call Procent(d9, nd9, pr9)
	call Procent(d10, nd10, pr10)
	sTime.innerHTML = time
end sub
sub pingfun()
	boolpin1 = document.getElementByID("pin1").checked
	If boolpin1 Then
		call PING_S(CompName1.innerHTML, Dostup1, Time1, d1, nd1)
	End If
	boolpin2 = document.getElementByID("pin2").checked
	If boolpin2 Then
		call PING_S(CompName2.innerHTML, Dostup2, Time2, d2, nd2)
	End If
	boolpin3 = document.getElementByID("pin3").checked
	If boolpin3 Then
		call PING_S(CompName3.innerHTML, Dostup3, Time3, d3, nd3)
	End If
	boolpin4 = document.getElementByID("pin4").checked
	If boolpin4 Then
		call PING_S(CompName4.innerHTML, Dostup4, Time4, d4, nd4)
	End If
	boolpin5 = document.getElementByID("pin5").checked
	If boolpin5 Then
		call PING_S(CompName5.innerHTML, Dostup5, Time5, d5, nd5)
	End If
	boolpin6 = document.getElementByID("pin6").checked
	If boolpin6 Then
		call PING_S(CompName6.innerHTML, Dostup6, Time6, d6, nd6)
	End If
	boolpin7 = document.getElementByID("pin7").checked
	If boolpin7 Then
		call PING_S(CompName7.innerHTML, Dostup7, Time7, d7, nd7)
	End If
	boolpin8 = document.getElementByID("pin8").checked
	If boolpin8 Then
		call PING_S(CompName8.innerHTML, Dostup8, Time8, d8, nd8)
	End If
	boolpin9 = document.getElementByID("pin9").checked
	If boolpin9 Then
		call PING_S(CompName9.innerHTML, Dostup9, Time9, d9, nd9)
	End If
	boolpin10 = document.getElementByID("pin10").checked
	If boolpin10 Then
		call PING_S(CompName10.innerHTML, Dostup10, Time10, d10, nd10)
	End If
	call proc()
end sub
sub start_onclick()
'document.getElementById("option").style.visibility = "hidden"
'document.getElementById("options").style.visibility = "visible"
'document.getElementById("optione").style.visibility = "hidden"
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
		value1 = AValue.getAttribute("CompName1")
		CompName1.innerHTML = value1
		text1 = AValue.getAttribute("t1")
		t1.innerHTML = text1
		
		boolpin2 = AValue.getAttribute("pin2")
		if boolpin2 = "True" then
		With document.getElementByID("pin2")
		.checked = Not .checked
		End With
		end if
		value2 = AValue.getAttribute("CompName2")
		CompName2.innerHTML = value2
		text2 = AValue.getAttribute("t2")
		t2.innerHTML = text2
		
		boolpin3 = AValue.getAttribute("pin3")
		if boolpin3 = "True" then
		With document.getElementByID("pin3")
		.checked = Not .checked
		End With
		end if
		value3 = AValue.getAttribute("CompName3")
		CompName3.innerHTML = value3
		text3 = AValue.getAttribute("t3")
		t3.innerHTML = text3
		
		boolpin4 = AValue.getAttribute("pin4")
		if boolpin4 = "True" then
		With document.getElementByID("pin4")
		.checked = Not .checked
		End With
		end if
		value4 = AValue.getAttribute("CompName4")
		CompName4.innerHTML = value4
		text4 = AValue.getAttribute("t4")
		t4.innerHTML = text4
		
		boolpin5 = AValue.getAttribute("pin5")
		if boolpin5 = "True" then
		With document.getElementByID("pin5")
		.checked = Not .checked
		End With
		end if
		value5 = AValue.getAttribute("CompName5")
		CompName5.innerHTML = value5
		text5 = AValue.getAttribute("t5")
		t5.innerHTML = text5
		
		boolpin6 = AValue.getAttribute("pin6")
		if boolpin6 = "True" then
		With document.getElementByID("pin6")
		.checked = Not .checked
		End With
		end if
		value6 = AValue.getAttribute("CompName6")
		CompName6.innerHTML = value6
		text6 = AValue.getAttribute("t6")
		t6.innerHTML = text6
		
		boolpin7 = AValue.getAttribute("pin7")
		if boolpin7 = "True" then
		With document.getElementByID("pin7")
		.checked = Not .checked
		End With
		end if
		value7 = AValue.getAttribute("CompName7")
		CompName7.innerHTML = value7
		text7 = AValue.getAttribute("t7")
		t7.innerHTML = text7
		
		boolpin8 = AValue.getAttribute("pin8")
		if boolpin8 = "True" then
		With document.getElementByID("pin8")
		.checked = Not .checked
		End With
		end if
		value8 = AValue.getAttribute("CompName8")
		CompName8.innerHTML = value8
		text8 = AValue.getAttribute("t8")
		t8.innerHTML = text8
		
		boolpin9 = AValue.getAttribute("pin9")
		if boolpin9 = "True" then
		With document.getElementByID("pin9")
		.checked = Not .checked
		End With
		end if
		value9 = AValue.getAttribute("CompName9")
		CompName9.innerHTML = value9
		text9 = AValue.getAttribute("t9")
		t9.innerHTML = text9
		
		boolpin10 = AValue.getAttribute("pin10")
		if boolpin10 = "True" then
		With document.getElementByID("pin10")
		.checked = Not .checked
		End With
		end if
		value10 = AValue.getAttribute("CompName10")
		CompName10.innerHTML = value10
		text10 = AValue.getAttribute("t10")
		t10.innerHTML = text10
end sub

sub XMLsave_onclick()
	Set FSO = CreateObject("Scripting.FileSystemObject")
	set c2f=fso.OpenTextFile("config\APS.XML",2,-1)
		boolpin1 = document.getElementByID("pin1").checked
		if boolpin1 = false then
		boolpin1 = "False"
		else boolpin1 = "True"
		end if
		boolpin2 = document.getElementByID("pin2").checked
		if boolpin2 = false then
		boolpin2 = "False"
		else boolpin2 = "True"
		end if
		boolpin3 = document.getElementByID("pin3").checked
		if boolpin3 = false then
		boolpin3 = "False"
		else boolpin3 = "True"
		end if
		boolpin4 = document.getElementByID("pin4").checked
		if boolpin4 = false then
		boolpin4 = "False"
		else boolpin4 = "True"
		end if
		boolpin5 = document.getElementByID("pin5").checked
		if boolpin5 = false then
		boolpin5 = "False"
		else boolpin5 = "True"
		end if
		boolpin6 = document.getElementByID("pin6").checked
		if boolpin6 = false then
		boolpin6 = "False"
		else boolpin6 = "True"
		end if
		boolpin7 = document.getElementByID("pin7").checked
		if boolpin7 = false then
		boolpin7 = "False"
		else boolpin7 = "True"
		end if
		boolpin8 = document.getElementByID("pin8").checked
		if boolpin8 = false then
		boolpin8 = "False"
		else boolpin8 = "True"
		end if
		boolpin9 = document.getElementByID("pin9").checked
		if boolpin9 = false then
		boolpin9 = "False"
		else boolpin9 = "True"
		end if
		boolpin10 = document.getElementByID("pin10").checked
		if boolpin10 = false then
		boolpin10 = "False"
		else boolpin10 = "True"
		end if
		c2f.WriteLine("<?xml version='1.0' encoding='windows-1251'?>")
		c2f.WriteLine("<APS version='1.50' date='22.01.2012'>" )
		c2f.WriteLine("<COMPNAME")
		c2f.WriteLine("pin1" & " ='" & boolpin1 & "' t1" & " ='" & t1.innerHTML & "' CompName1='" & CompName1.innerHTML & "'")
		c2f.WriteLine("pin2" & " ='" & boolpin2 & "' t2" & " ='" & t2.innerHTML & "' CompName2='" & CompName2.innerHTML & "'")
		c2f.WriteLine("pin3" & " ='" & boolpin3 & "' t3" & " ='" & t3.innerHTML & "' CompName3='" & CompName3.innerHTML & "'")
		c2f.WriteLine("pin4" & " ='" & boolpin4 & "' t4" & " ='" & t4.innerHTML & "' CompName4='" & CompName4.innerHTML & "'")
		c2f.WriteLine("pin5" & " ='" & boolpin5 & "' t5" & " ='" & t5.innerHTML & "' CompName5='" & CompName5.innerHTML & "'")
		c2f.WriteLine("pin6" & " ='" & boolpin6 & "' t6" & " ='" & t6.innerHTML & "' CompName6='" & CompName6.innerHTML & "'")
		c2f.WriteLine("pin7" & " ='" & boolpin7 & "' t7" & " ='" & t7.innerHTML & "' CompName7='" & CompName7.innerHTML & "'")
		c2f.WriteLine("pin8" & " ='" & boolpin8 & "' t8" & " ='" & t8.innerHTML & "' CompName8='" & CompName8.innerHTML & "'")
		c2f.WriteLine("pin9" & " ='" & boolpin9 & "' t9" & " ='" & t9.innerHTML & "' CompName9='" & CompName9.innerHTML & "'")
		c2f.WriteLine("pin10" & " ='" & boolpin10 & "' t10" & " ='" & t10.innerHTML & "' CompName10='" & CompName10.innerHTML & "'")
		c2f.WriteLine("/></APS>")
		
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
       'strHTML = ""
       strHTML = time
	   'strComputer = "."
       'Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
       'Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process")
       
       'For Each objProcess in colProcesses
       '    strHTML = strHTML & objProcess.Name & "<BR>"
       'Next
       
       ProcessList.InnerHTML = strHTML
    End Sub
Sub WindowOnLoad
	StartUp
	objXML
	Version_Div.InnerHTML = "Версия: " & oHTA.Version
	document.title="Ping IP v" & oHTA.Version
	ProcessList.InnerHTML = time
	sTime.innerHTML = time
	time1.innerHTML = "-1"
	d1.innerHTML = "0"
	nd1.innerHTML = "0"
	pr1.innerHTML = "0"
	time2.innerHTML = "-1"
	d2.innerHTML = "0"
	nd2.innerHTML = "0"
	pr2.innerHTML = "0"
	time3.innerHTML = "-1"
	d3.innerHTML = "0"
	nd3.innerHTML = "0"
	pr3.innerHTML = "0"
	time4.innerHTML = "-1"
	d4.innerHTML = "0"
	nd4.innerHTML = "0"
	pr4.innerHTML = "0"
	time5.innerHTML = "-1"
	d5.innerHTML = "0"
	nd5.innerHTML = "0"
	pr5.innerHTML = "0"
	time6.innerHTML = "-1"
	d6.innerHTML = "0"
	nd6.innerHTML = "0"
	pr6.innerHTML = "0"
	time7.innerHTML = "-1"
	d7.innerHTML = "0"
	nd7.innerHTML = "0"
	pr7.innerHTML = "0"
	time8.innerHTML = "-1"
	d8.innerHTML = "0"
	nd8.innerHTML = "0"
	pr8.innerHTML = "0"
	time9.innerHTML = "-1"
	d9.innerHTML = "0"
	nd9.innerHTML = "0"
	pr9.innerHTML = "0"
	time10.innerHTML = "-1"
	d10.innerHTML = "0"
	nd10.innerHTML = "0"
	pr10.innerHTML = "0"
	
	iTimerID = window.setInterval("RefreshList", 1000)
	'document.getElementById("option").style.visibility = "hidden"
	'document.getElementById("options").style.visibility = "visible"
	'document.getElementById("optione").style.visibility = "hidden"
End Sub

sub CloseButton_onclick()
XMLsave_onclick()
objXML
end sub
sub options_onclick()

document.getElementById("options").style.visibility = "visible"
document.getElementById("optione").style.visibility = "visible"
end sub
sub optione_onclick()
document.getElementById("optione").style.visibility = "visible"
document.getElementById("option").style.visibility = "hidden"
end sub
Sub OnClickButtonbtnSET
	varReturn = window.ShowModalDialog("Setting.hta", Fill_Dict_TextBox, "dialogHeight:350px;dialogWidth:280px") 
End Sub