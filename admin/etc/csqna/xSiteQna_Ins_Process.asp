<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 600 ''초단위
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/admin/etc/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/admin/etc/csqna/lib/xSiteQnALib.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim IS_TEST_MODE : IS_TEST_MODE = False
Dim sellsite, selldate, csGubun, arrRows, isSuccess
dim nowdate, fromdate, todate, currdate
Dim i
sellsite	= requestCheckVar(html2db(request("sellsite")),32)
selldate	= requestCheckVar(html2db(request("selldate")),32)
csGubun		= requestCheckVar(html2db(request("mode")),32)

If (sellsite = "") or (csGubun = "") then
	response.write "잘못된 접근입니다."
	dbget.close : response.end
End If

dim IS_SELLDATE_FIXED : IS_SELLDATE_FIXED = False
if (selldate = "") then
	'// 오늘까지 일괄로 가져오기
	Call GetCSCheckStatus(sellsite, csGubun, selldate, isSuccess)
	fromdate = selldate
	todate = Left(Now, 10)
else
	fromdate = selldate
	todate = selldate
	IS_SELLDATE_FIXED = True
end if

Select Case sellsite
	Case "wconcept1010"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if
		currdate = fromdate

		If (csGubun = "reqQnA") Then
			do while (currdate <= todate)
				response.write "<br />" & sellsite & " : " & currdate & "<br />"
				'//W컨셉
				Call GetCSQnA_wconcept(currdate)
				'http://localhost:11117/admin/etc/csqna/xSiteQna_Ins_Process.asp?sellsite=wconcept1010&mode=reqQnA

				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
			Loop
		ElseIf (csGubun = "resQnA") Then
			arrRows = getCSAnswerComplete(sellsite)
			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					Call resCSQnA_wconcept(arrRows(0, i), arrRows(1, i), arrRows(2, i))
				Next
			Else
				rw "None"
			End If
			'http://localhost:11117/admin/etc/csqna/xSiteQna_Ins_Process.asp?sellsite=wconcept1010&mode=resQnA
		Else
			response.write "잘못된 접근입니다."
			dbget.close : response.end
		End If
	Case "auction1010", "gmarket1010"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if
		currdate = fromdate

		If (csGubun = "reqQnA") Then
			do while (currdate <= todate)
				response.write "<br />" & sellsite & " : " & currdate & "<br />"
				'//ebay
				If sellsite = "auction1010" Then
					Call GetCSQnA_ebay(sellsite, currdate, "1")
					Call GetCSQnA_ebay(sellsite, currdate, "2")
				Else
					Call GetCSQnA_ebay(sellsite, currdate, "3")
				End If
				'http://localhost:11117/admin/etc/csqna/xSiteQna_Ins_Process.asp?sellsite=auction1010&mode=reqQnA
				'http://localhost:11117/admin/etc/csqna/xSiteQna_Ins_Process.asp?sellsite=gmarket1010&mode=reqQnA

				selldate = currdate      
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
			Loop
		ElseIf (csGubun = "resQnA") Then
			arrRows = getCSAnswerCompleteMallId(sellsite, "1")
			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					Call resCSQnA_ebay(arrRows(0, i), arrRows(1, i), arrRows(3, i), sellsite)
				Next
			Else
				rw "None"
			End If
			'http://localhost:11117/admin/etc/csqna/xSiteQna_Ins_Process.asp?sellsite=benepia1010&mode=resQnA
		ElseIf (csGubun = "reqCS") Then
			do while (currdate <= todate)
				response.write "<br />" & sellsite & " : " & currdate & "<br />"
				Call GetCSQnA_ebay_refer(sellsite, currdate)
				'http://localhost:11117/admin/etc/csqna/xSiteQna_Ins_Process.asp?sellsite=gmarket1010&mode=reqCS
				selldate = currdate      
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
			Loop
		ElseIf (csGubun = "resCS") Then
			arrRows = getCSAnswerCompleteMallId(sellsite, "2")
			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					Call resCSQnA_ebay_refer(arrRows(0, i), arrRows(1, i), sellsite)
				Next
			Else
				rw "None"
			End If
			'http://localhost:11117/admin/etc/csqna/xSiteQna_Ins_Process.asp?sellsite=benepia1010&mode=resQnA
		Else
			response.write "잘못된 접근입니다."
			dbget.close : response.end
		End If
	Case "benepia1010"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if
		currdate = fromdate

		If (csGubun = "reqQnA") Then
			do while (currdate <= todate)
				response.write "<br />" & sellsite & " : " & currdate & "<br />"
				'//W컨셉
				Call GetCSQnA_benepia(currdate)
				'http://localhost:11117/admin/etc/csqna/xSiteQna_Ins_Process.asp?sellsite=benepia1010&mode=reqQnA

				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
			Loop
		ElseIf (csGubun = "resQnA") Then
			arrRows = getCSAnswerComplete(sellsite)
			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					Call resCSQnA_benepia(arrRows(0, i), arrRows(1, i))
				Next
			Else
				rw "None"
			End If
			'http://localhost:11117/admin/etc/csqna/xSiteQna_Ins_Process.asp?sellsite=benepia1010&mode=resQnA
		Else
			response.write "잘못된 접근입니다."
			dbget.close : response.end
		End If
	Case "kakaostore"
		'//카카오톡스토어
		If (csGubun = "reqQnA") Then
			rw sellsite
			Call GetCSQnA_kakaostore()
			'http://localhost:11117/admin/etc/csqna/xSiteQna_Ins_Process.asp?sellsite=kakaostore&mode=reqQnA
		ElseIf (csGubun = "resQnA") Then
			arrRows = getCSAnswerComplete(sellsite)
			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					Call resCSQnA_kakaostore(arrRows(0, i), arrRows(1, i))
				Next
			Else
				rw "None"
			End If
			'http://localhost:11117/admin/etc/csqna/xSiteQna_Ins_Process.asp?sellsite=kakaostore&mode=resQnA
		Else
			response.write "잘못된 접근입니다."
			dbget.close : response.end
		End If
	Case "boribori1010"
		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if
		currdate = fromdate
		If (csGubun = "reqQnA") Then
			do while (currdate <= todate)
				response.write "<br />" & sellsite & " : " & currdate & "<br />"
				'//보리보리
				Call GetCSQnA_boribori(currdate)
				'http://localhost:11117/admin/etc/csqna/xSiteQna_Ins_Process.asp?sellsite=boribori1010&mode=reqQnA

				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
			Loop
		ElseIf (csGubun = "resQnA") Then
			arrRows = getCSAnswerCompleteMallId(sellsite, "1")
			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					Call resCSQnA_boribori(arrRows(0, i), arrRows(1, i))
				Next
			Else
				rw "None"
			End If
			'http://localhost:11117/admin/etc/csqna/xSiteQna_Ins_Process.asp?sellsite=boribori1010&mode=resQnA
		ElseIf (csGubun = "reqCS") Then
			do while (currdate <= todate)
				response.write "<br />" & sellsite & " : " & currdate & "<br />"
				'//보리보리
				Call GetCSQnA_boribori_Refer(currdate)
				'http://localhost:11117/admin/etc/csqna/xSiteQna_Ins_Process.asp?sellsite=boribori1010&mode=reqCS

				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
			Loop
		ElseIf (csGubun = "resCS") Then
			arrRows = getCSAnswerCompleteMallId(sellsite, "2")
			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					Call resCSQnA_boribori_Refer(arrRows(0, i), arrRows(1, i))
				Next
			Else
				rw "None"
			End If
			'http://localhost:11117/admin/etc/csqna/xSiteQna_Ins_Process.asp?sellsite=boribori1010&mode=resCS
		End If
	Case Else
		response.write "잘못된 접근입니다."
		dbget.close : response.end
End Select

If (IS_TEST_MODE = False) and (IS_SELLDATE_FIXED = False) Then
	If (selldate < Left(Now(), 10)) Then
		Call SetCSCheckStatus(sellsite, csGubun, Left(DateAdd("d", 1, CDate(selldate)), 10), "N")
	ElseIf (selldate = Left(Now(), 10)) Then
		Call SetCSCheckStatus(sellsite, csGubun, selldate, "Y")
	End If
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
