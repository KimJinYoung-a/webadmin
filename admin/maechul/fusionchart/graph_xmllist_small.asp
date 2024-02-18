<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/maechul/fusionchart/maechul_class.asp" -->

<% 
dim vGubun, yyyy1, yyyy2,dateview1 , datecancle,bancancle,accountdiv,sitename,i, mm1, mm2, defaultdate1, monthday
dim ipkumdatesucc, vParam, Omaechul_list

	yyyy1 			= request("yyyy1")
	yyyy2 			= request("yyyy2")
	dateview1 		= request("dateview1")
	datecancle 		= request("datecancle")
	bancancle 		= request("bancancle") 
	accountdiv 		= request("accountdiv")			
	sitename 		= request("sitename") 
	ipkumdatesucc 	= request("ipkumdatesucc")
	mm1 			= request("mm1")
	mm2 			= request("mm2")
	monthday		= request("monthday")

	vParam = request("param")

	yyyy1 = split(vParam,"^^")(1)
	yyyy1 = split(yyyy1,"=")(1)
	
	yyyy2 = split(vParam,"^^")(2)
	yyyy2 = split(yyyy2,"=")(1)

	datecancle = split(vParam,"^^")(3)
	datecancle = split(datecancle,"=")(1)
	
	bancancle = split(vParam,"^^")(4)
	bancancle = split(bancancle,"=")(1)
	
	accountdiv = split(vParam,"^^")(5)
	accountdiv = split(accountdiv,"=")(1)
	
	sitename = split(vParam,"^^")(6)
	sitename = split(sitename,"=")(1)
	
	dateview1 = split(vParam,"^^")(7)
	dateview1 = split(dateview1,"=")(1)
	
	ipkumdatesucc = split(vParam,"^^")(8)
	ipkumdatesucc = split(ipkumdatesucc,"=")(1)
	
	mm1 = split(vParam,"^^")(9)
	mm1 = split(mm1,"=")(1)
	
	mm2 = split(vParam,"^^")(10)
	mm2 = split(mm2,"=")(1)
	
	monthday = split(vParam,"^^")(11)
	monthday = split(monthday,"=")(1)
	
	
	defaultdate1 = dateadd("d",-60,year(now) & "-" &TwoNumber(month(now)) & "-" & day(now))		'날짜값이 없을때 기본값으로 60이전까지 검색
	if yyyy2 = "" then yyyy2 = year(now)
	if yyyy1 = "" then yyyy1 = CInt(yyyy2)-2
	if mm1 = "" then mm1 = "01"
	if mm2 = "" then mm2 = month(now)
	mm2 = TwoNumber(mm2)
	if bancancle = "" then bancancle = "1"
	if dateview1 = "" then dateview1 = "yes"
	
	vGubun = split(vParam,"^^")(12)
	vGubun = split(vGubun,"=")(1)

	
	'<!-- //-->


	set Omaechul_list = new Cmaechul_list
	Omaechul_list.FRectStartdate = yyyy1
	Omaechul_list.frectdatecancle = datecancle
	Omaechul_list.frectbancancle = bancancle
	Omaechul_list.frectaccountdiv = accountdiv
	Omaechul_list.frectsitename = sitename
	Omaechul_list.frectipkumdatesucc = ipkumdatesucc		
	Omaechul_list.fmonth = mm1
	Omaechul_list.fmonthday = monthday
	Omaechul_list.fmaechul_graph_new()
	

	Dim vXML, vLoofCount, j, vLineColor, vStart, vEnd, vEndDay
	vLoofCount = CInt(yyyy2) - CInt(yyyy1)
	vXML = ""
%>

<?xml version='1.0' encoding='EUC-KR' ?>
<chart chartBottomMargin='2' formatNumberScale='0' drawAnchors='1' showLimits='0' divLineThickness='1' decimalPrecision='1' chartTopMargin='2' showShadow='1' canvasBorderColor='CBCBCB' animation='1' baseFontColor='666666' bgColor='FFFFFF' formatNumber='1' legendBorderColor='FFFFFF' canvasPadding='3' legendBgColor='FFFFFF' chartRightMargin='2' legendPadding='2' legendShadow='0' divLineIsDashed='1' showBorder='0' legendBorderThickness='0' placeValuesInside='1' chartLeftMargin='0' canvasBorderThickness='1' baseFontSize='11' divLineDashGap='3' setAdaptiveYMin='1' anchorRadius='5' plotBorderAlpha='20' >
<%
	If vGubun = "1" Then
	'############################################### 실금액 통계 ###############################################
		vXML = vXML & "<categories>" & vbCrLf
		
		If monthday = "m" Then
		'############################################### 월매출 통계 ###############################################
			For i = 1 to 12
				vXML = vXML & "	<category name='" & TwoNumber(i) & "월' showName='1' showLine='1' />" & vbCrLf
			Next
			vXML = vXML & "</categories>" & vbCrLf
			
			For j = 0 To vLoofCount
				If j = vLoofCount Then
					vLineColor = ChartLineColor("top")
				Else
					vLineColor = ChartLineColor(j)
				End If
				vStart = j*12
				If ((j+1)*12) > (Omaechul_list.ftotalcount-1) Then
					vEnd = Omaechul_list.ftotalcount - 1
				Else
					vEnd = (j+1)*12-1
				End IF
				vXML = vXML & "<dataset seriesName='" & CInt(yyyy2)-CInt(vLoofCount)+CInt(j) & "' color='" & vLineColor & "' showValues='0' >" & vbCrLf
					For i = vStart to vEnd
						vXML = vXML & "	<set value='" & Omaechul_list.flist(i).fsubtotalprice & "' />" & vbCrLf
					Next
				vXML = vXML & "</dataset>" & vbCrLf
			Next
		Else
		'############################################### 일매출 통계 ###############################################
			For i = 1 to 31
				vXML = vXML & "	<category name='" & TwoNumber(i) & "' showName='1' showLine='1' />" & vbCrLf
			Next
			vXML = vXML & "</categories>" & vbCrLf
			
			For j = 0 To vLoofCount
				vEndDay = Day(DateAdd("d", -1, (DateAdd("m",1,(yyyy1+j)&"-"&mm1))))
				
				If j = vLoofCount Then
					vLineColor = ChartLineColor("top")
				Else
					vLineColor = ChartLineColor(j)
				End If
				vStart = j*vEndDay
				If ((j+1)*vEndDay) > (Omaechul_list.ftotalcount-1) Then
					vEnd = Omaechul_list.ftotalcount - 1
				Else
					vEnd = ((j+1)*vEndDay)-1
				End IF
				vXML = vXML & "<dataset seriesName='" & CInt(yyyy1)+CInt(j) & "-" & mm1 & "' color='" & vLineColor & "' showValues='0' >" & vbCrLf
					For i = vStart to vEnd
						vXML = vXML & "	<set value='" & Omaechul_list.flist(i).fsubtotalprice & "' />" & vbCrLf
					Next
				vXML = vXML & "</dataset>" & vbCrLf
			Next
		End If
		
	ElseIf vGubun = "2" Then
	'############################################### 순수익 통계 ###############################################
		vXML = vXML & "<categories>" & vbCrLf
		
		If monthday = "m" Then
		'############################################### 월매출 통계 ###############################################
			For i = 1 to 12
				vXML = vXML & "	<category name='" & TwoNumber(i) & "월' showName='1' showLine='1' />" & vbCrLf
			Next
			vXML = vXML & "</categories>" & vbCrLf
			
			For j = 0 To vLoofCount
				If j = vLoofCount Then
					vLineColor = ChartLineColor("top")
				Else
					vLineColor = ChartLineColor(j)
				End If
				vStart = j*12
				If ((j+1)*12) > (Omaechul_list.ftotalcount-1) Then
					vEnd = Omaechul_list.ftotalcount - 1
				Else
					vEnd = (j+1)*12-1
				End IF
				vXML = vXML & "<dataset seriesName='" & CInt(yyyy2)-CInt(vLoofCount)+CInt(j) & "' color='" & vLineColor & "' showValues='0' >" & vbCrLf
					For i = vStart to vEnd
						vXML = vXML & "	<set value='" & Omaechul_list.flist(i).fsunsuik & "' />" & vbCrLf
					Next
				vXML = vXML & "</dataset>" & vbCrLf
			Next
		Else
		'############################################### 일매출 통계 ###############################################
			For i = 1 to 31
				vXML = vXML & "	<category name='" & TwoNumber(i) & "' showName='1' showLine='1' />" & vbCrLf
			Next
			vXML = vXML & "</categories>" & vbCrLf
			
			For j = 0 To vLoofCount
				vEndDay = Day(DateAdd("d", -1, (DateAdd("m",1,(yyyy1+j)&"-"&mm1))))
				
				If j = vLoofCount Then
					vLineColor = ChartLineColor("top")
				Else
					vLineColor = ChartLineColor(j)
				End If
				vStart = j*vEndDay
				If ((j+1)*vEndDay) > (Omaechul_list.ftotalcount-1) Then
					vEnd = Omaechul_list.ftotalcount - 1
				Else
					vEnd = ((j+1)*vEndDay)-1
				End IF
				vXML = vXML & "<dataset seriesName='" & CInt(yyyy1)+CInt(j) & "-" & mm1 & "' color='" & vLineColor & "' showValues='0' >" & vbCrLf
					For i = vStart to vEnd
						vXML = vXML & "	<set value='" & Omaechul_list.flist(i).fsunsuik & "' />" & vbCrLf
					Next
				vXML = vXML & "</dataset>" & vbCrLf
			Next
		End If
		
	ElseIf vGubun = "3" Then
	'############################################### 총건수 통계 ###############################################
		vXML = vXML & "<categories>" & vbCrLf
		
		If monthday = "m" Then
		'############################################### 월매출 통계 ###############################################
			For i = 1 to 12
				vXML = vXML & "	<category name='" & TwoNumber(i) & "월' showName='1' showLine='1' />" & vbCrLf
			Next
			vXML = vXML & "</categories>" & vbCrLf
			
			For j = 0 To vLoofCount
				If j = vLoofCount Then
					vLineColor = ChartLineColor("top")
				Else
					vLineColor = ChartLineColor(j)
				End If
				vStart = j*12
				If ((j+1)*12) > (Omaechul_list.ftotalcount-1) Then
					vEnd = Omaechul_list.ftotalcount - 1
				Else
					vEnd = (j+1)*12-1
				End IF
				vXML = vXML & "<dataset seriesName='" & CInt(yyyy2)-CInt(vLoofCount)+CInt(j) & "' color='" & vLineColor & "' showValues='0' >" & vbCrLf
					For i = vStart to vEnd
						vXML = vXML & "	<set value='" & Omaechul_list.flist(i).ftotalcount & "' />" & vbCrLf
					Next
				vXML = vXML & "</dataset>" & vbCrLf
			Next
		Else
		'############################################### 일매출 통계 ###############################################
			For i = 1 to 31
				vXML = vXML & "	<category name='" & TwoNumber(i) & "' showName='1' showLine='1' />" & vbCrLf
			Next
			vXML = vXML & "</categories>" & vbCrLf
			
			For j = 0 To vLoofCount
				vEndDay = Day(DateAdd("d", -1, (DateAdd("m",1,(yyyy1+j)&"-"&mm1))))
				
				If j = vLoofCount Then
					vLineColor = ChartLineColor("top")
				Else
					vLineColor = ChartLineColor(j)
				End If
				vStart = j*vEndDay
				If ((j+1)*vEndDay) > (Omaechul_list.ftotalcount-1) Then
					vEnd = Omaechul_list.ftotalcount - 1
				Else
					vEnd = ((j+1)*vEndDay)-1
				End IF
				vXML = vXML & "<dataset seriesName='" & CInt(yyyy1)+CInt(j) & "-" & mm1 & "' color='" & vLineColor & "' showValues='0' >" & vbCrLf
					For i = vStart to vEnd
						vXML = vXML & "	<set value='" & Omaechul_list.flist(i).ftotalcount & "' />" & vbCrLf
					Next
				vXML = vXML & "</dataset>" & vbCrLf
			Next
		End If
		
	End If

	Set Omaechul_list = nothing
	
	Response.Write vXML
%>
</chart>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->

<%
Function ChartLineColor(vGubun)
	Select Case vGubun
		Case "0"
			ChartLineColor = "8BBA00"
		Case "1"
			ChartLineColor = "A66EDD"
		Case "2"
			ChartLineColor = "F6BD0F"
		Case "3"
			ChartLineColor = "A8A8FF"
		Case "4"
			ChartLineColor = "FD9696"
		Case "5"
			ChartLineColor = "F3B7F3"
		Case "6"
			ChartLineColor = "AC5353"
		Case "7"
			ChartLineColor = "AFB1FF"
		Case "8"
			ChartLineColor = "A2A2A2"
		Case "9"
			ChartLineColor = "E4E4E4"
		Case Else
			ChartLineColor = "0000FF"
	End Select
End Function
%>