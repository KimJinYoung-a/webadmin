<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/datamart/qna/qna_summaryCls.asp"-->
<%
Dim sSdate, sEdate, sType, dayMode, summaryName, makerid, i
Dim arrList, oQna, TotalQnaAllCnt, TotalQnasecretYCnt, TotalQnasecretNCnt, TotalreplyYCnt, TotalreplyNCnt, TotalSumReplyDayCnt, Totalsnssend1Cnt, Totalsnssend2Cnt, Totalsnssend3Cnt, Totalsnssend4Cnt, Totalsnssend5Cnt
Dim TermQnaAllCnt, TermQnasecretYCnt, TermQnasecretNCnt, TermreplyYCnt, TermreplyNCnt, TermSumReplyDayCnt, Termsnssend1Cnt, Termsnssend2Cnt, Termsnssend3Cnt, Termsnssend4Cnt, Termsnssend5Cnt
sSdate			= requestCheckVar(request("iSD2"),10)
sEdate			= requestCheckVar(request("iED2"),10)
sType			= requestCheckVar(request("sType2"),10)
makerid			= requestCheckVar(request("makerid2"),32)

If sSdate = "" Then sSdate = DateSerial(Year(Now()), Month(Now()), 1)
If sEdate = "" Then sEdate = Date()
If sType = "" Then sType = "category"
If dayMode = "" Then dayMode = "D"

SET oQna = new cQnaSummary
	oQna.FRectSdate			= sSdate
	oQna.FRectEdate			= sEdate
	oQna.FRectSType			= sType
	oQna.FRectTopCnt		= "Y"
	arrList = oQna.fnQnaSummayReport

	TotalQnaAllCnt		= oQna.FTotalQnaAllCnt
	TotalQnasecretYCnt	= oQna.FTotalQnasecretYCnt
	TotalQnasecretNCnt	= oQna.FTotalQnasecretNCnt
	TotalreplyYCnt		= oQna.FTotalreplyYCnt
	TotalreplyNCnt		= oQna.FTotalreplyNCnt
	TotalSumReplyDayCnt	= oQna.FTotalSumReplyDayCnt
	Totalsnssend1Cnt	= oQna.FTotalsnssend1Cnt
	Totalsnssend2Cnt	= oQna.FTotalsnssend2Cnt
	Totalsnssend3Cnt	= oQna.FTotalsnssend3Cnt
	Totalsnssend4Cnt	= oQna.FTotalsnssend4Cnt
	Totalsnssend5Cnt	= oQna.FTotalsnssend5Cnt

	TermQnaAllCnt		= oQna.FTermQnaAllCnt		
	TermQnasecretYCnt	= oQna.FTermQnasecretYCnt	
	TermQnasecretNCnt	= oQna.FTermQnasecretNCnt	
	TermreplyYCnt		= oQna.FTermreplyYCnt		
	TermreplyNCnt		= oQna.FTermreplyNCnt		
	TermSumReplyDayCnt	= oQna.FTermSumReplyDayCnt	
	Termsnssend1Cnt		= oQna.FTermsnssend1Cnt		
	Termsnssend2Cnt		= oQna.FTermsnssend2Cnt		
	Termsnssend3Cnt		= oQna.FTermsnssend3Cnt		
	Termsnssend4Cnt		= oQna.FTermsnssend4Cnt		
	Termsnssend5Cnt		= oQna.FTermsnssend5Cnt		

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=qnaSummary_" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="application/vnd.ms-excel;charset=euc-kr">
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="7.1%"></td>
	<td width="7.1%">Q&A 전체(건)</td>
	<td width="7.1%">Q&A 공개(건)</td>
	<td width="7.1%">Q&A 비공개(건)</td>
	<td width="7.1%">답변(건)</td>
	<td width="7.1%">미답변(건)</td>
	<td width="7.1%">지연<br />미답변(건)</td>
	<td width="7.1%">답변율(%)</td>
	<td width="7.1%">평균<br />답변일(일)</td>
	<td width="7.1%">알림 문자<br />1차 발송(건)</td>
	<td width="7.1%">알림 문자<br />2차 발송(건)</td>
	<td width="7.1%">알림 문자<br />3차 발송(건)</td>
	<td width="7.1%">알림 문자<br />4차 발송(건)</td>
	<td width="7.1%">알림 문자<br />5차 발송(건)</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td>누적 전체</td>
	<td><%= TotalQnaAllCnt %></td>
	<td><%= TotalQnasecretYCnt %></td>
	<td><%= TotalQnasecretNCnt %></td>
	<td><%= TotalreplyYCnt %></td>
	<td><%= TotalreplyNCnt %></td>
	<td><%= Totalsnssend1Cnt %></td>
	<td>
	<%
		If TotalreplyYCnt <> 0 Then
			response.write Round(TotalreplyYCnt / TotalQnaAllCnt * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td>
	<%
		If TotalSumReplyDayCnt <> 0 Then
			response.write Round((TotalSumReplyDayCnt * 1.0 / TotalQnaAllCnt * 1.0), 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td><%= Totalsnssend1Cnt %></td>
	<td><%= Totalsnssend2Cnt %></td>
	<td><%= Totalsnssend3Cnt %></td>
	<td><%= Totalsnssend4Cnt %></td>
	<td><%= Totalsnssend5Cnt %></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td>기간 전체</td>
	<td><%= TermQnaAllCnt %></td>
	<td><%= TermQnasecretYCnt %></td>
	<td><%= TermQnasecretNCnt %></td>
	<td><%= TermreplyYCnt %></td>
	<td><%= TermreplyNCnt %></td>
	<td><%= Termsnssend1Cnt %></td>
	<td>
	<%
		If TermreplyYCnt <> 0 Then
			response.write Round(TermreplyYCnt / TermQnaAllCnt * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td>
	<%
		If TermSumReplyDayCnt <> 0 Then
			response.write Round((TermSumReplyDayCnt * 1.0 / TermQnaAllCnt * 1.0), 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td><%= Termsnssend1Cnt %></td>
	<td><%= Termsnssend2Cnt %></td>
	<td><%= Termsnssend3Cnt %></td>
	<td><%= Termsnssend4Cnt %></td>
	<td><%= Termsnssend5Cnt %></td>
</tr>
<tr align="center" bgcolor="#D2D2D2">
	<td colspan="15"></td>
</tr>
<% If IsArray(arrList) Then %>
<% For i=0 To Ubound(arrList, 2) %>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td><%= arrList(0, i) %></td>
	<td><%= arrList(1, i) %></td>
	<td><%= arrList(2, i) %></td>
	<td><%= arrList(3, i) %></td>
	<td><%= arrList(4, i) %></td>
	<td><%= arrList(5, i) %></td>
	<td><%= arrList(7, i) %></td>
	<td>
	<%
		If arrList(4, i) <> 0 Then
			response.write Round(arrList(4, i) / arrList(1, i) * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td>
	<%
		If arrList(6, i) <> 0 Then
			response.write Round(arrList(6, i) / arrList(1, i), 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td><%= arrList(7, i) %></td>
	<td><%= arrList(8, i) %></td>
	<td><%= arrList(9, i) %></td>
	<td><%= arrList(10, i) %></td>
	<td><%= arrList(11, i) %></td>
</tr>
<% Next %>
<% End If %>
</table>
</body>
</html>
<% set oQna = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->