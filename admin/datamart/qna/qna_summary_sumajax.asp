<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/datamart/qna/qna_summaryCls.asp"-->
<%
Response.CharSet = "euc-kr"
Dim oQna, sSdate, sEdate, sType, dayMode, makerid, arrList, i
Dim TermQnaAllCnt, TermQnasecretYCnt, TermQnasecretNCnt, TermreplyYCnt, TermreplyNCnt, TermSumReplyDayCnt, Termsnssend1Cnt, Termsnssend2Cnt, Termsnssend3Cnt, Termsnssend4Cnt, Termsnssend5Cnt
sSdate			= requestCheckVar(request("iSD"),10)
sEdate			= requestCheckVar(request("iED"),10)
sType			= requestCheckVar(request("sType"),10)
dayMode			= requestCheckVar(request("dayMode"),1)
makerid			= requestCheckVar(request("makerid"),32)

SET oQna = new cQnaSummary
	oQna.FRectSdate			= sSdate
	oQna.FRectEdate			= sEdate
	oQna.FRectSType			= sType
	oQna.FRectDayMode		= dayMode
	oQna.FRectMakerid		= makerid
	arrList = oQna.fnQnaSummayReportByTerm

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
%>
<p></p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="30" bgcolor="#FFFFFF">
	<td colspan="15">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				<strong>기간별 통계</strong>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="8.3%"></td>
	<td width="8.3%">Q&A 전체(건)</td>
	<td width="8.3%">Q&A 공개(건)</td>
	<td width="8.3%">Q&A 비공개(건)</td>
	<td width="8.3%">답변(건)</td>
	<td width="8.3%">미답변(건)</td>
	<td width="8.3%">답변율(%)</td>
	<td width="8.3%">알림 문자<br />1차 발송(건)</td>
	<td width="8.3%">알림 문자<br />2차 발송(건)</td>
	<td width="8.3%">알림 문자<br />3차 발송(건)</td>
	<td width="8.3%">알림 문자<br />4차 발송(건)</td>
	<td width="8.3%">알림 문자<br />5차 발송(건)</td>
</tr>
<tr align="center" bgcolor="GOLD" height="30">
	<td>기간 전체</td>
	<td><%= TermQnaAllCnt %></td>
	<td><%= TermQnasecretYCnt %></td>
	<td><%= TermQnasecretNCnt %></td>
	<td><%= TermreplyYCnt %></td>
	<td><%= TermreplyNCnt %></td>
	<td>
	<%
		If TermreplyYCnt <> 0 Then
			response.write Round(TermreplyYCnt / TermQnaAllCnt * 100, 1)
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
<% If IsArray(arrList) Then %>
<% For i=0 To Ubound(arrList, 2) %>
<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td><%= Replace(arrList(0, i), "-", ".") %></td>
	<td><%= arrList(1, i) %></td>
	<td><%= arrList(2, i) %></td>
	<td><%= arrList(3, i) %></td>
	<td><%= arrList(4, i) %></td>
	<td><%= arrList(5, i) %></td>
	<td>
	<%
		If arrList(4, i) <> 0 Then
			response.write Round(arrList(4, i) / arrList(1, i) * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td><%= arrList(6, i) %></td>
	<td><%= arrList(7, i) %></td>
	<td><%= arrList(8, i) %></td>
	<td><%= arrList(9, i) %></td>
	<td><%= arrList(10, i) %></td>
</tr>
<% Next %>
<% End If %>
</table>
<% SET oQna = nothing %>
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->