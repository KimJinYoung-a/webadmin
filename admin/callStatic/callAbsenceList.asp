<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim yyyymmdd
dim i

yyyymmdd	= req("yyyymmdd", "")

Dim strSql
strSql = " db_datamart.dbo.sp_Ten_Call_Absence_List ('" & yyyymmdd & "')"

db3_rsget.CursorLocation = adUseClient
db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly, adLockReadOnly, adCmdStoredProc	

Dim rs 
If Not db3_rsget.EOF Then
	rs = db3_rsget.getRows()
End If 
db3_rsget.close


%>

<script language='javascript'>

window.onload = function () 
{
	document.getElementById("divShow").innerHTML = document.getElementById("divHide").innerHTML;
}

</script>


<div id="divShow"></div>
<p>
 
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">

        <td></td>
        <td>수신번호</td>
        <td>발신번호</td>
        <td>전화온시간</td>
        <td>비고</td>

	</tr>
<%

'' 요일명 리턴
Function getWeekDay(ByVal val)
	Select Case Weekday(val)
	Case "1" getWeekDay = "<span style='color:red;'>일요일</span>"
	Case "2" getWeekDay = "월요일"
	Case "3" getWeekDay = "화요일"
	Case "4" getWeekDay = "수요일"
	Case "5" getWeekDay = "목요일"
	Case "6" getWeekDay = "금요일"
	Case "7" getWeekDay = "<span style='color:red;'>토요일</span>"
	End Select 
End Function


Dim rowCnt
Dim sRs(20)

Dim dststr, dst, bigo, dcontext, lastapp, lastdata

If IsArray(rs) Then 
%>

	<%For i=0 To UBound(rs,2)%>
    <tr align="center" bgcolor="#FFFFFF">
	<%
		dst = rs(1,i)
		dcontext	= rs(3,i)
		lastapp		= rs(4,i)
		lastdata	= rs(5,i)
        
        dststr = ""
		If dst <> "" Then
			If InStr(dst,"07075490429") > 0 Then 
				dststr = "콜센터_헌트"
				sRs(1) = sRs(1) + 1
			ElseIf InStr(dst,"07075490556") > 0 Then 
				dststr = "사무실_헌트"
				sRs(2) = sRs(2) + 1
			ElseIf InStr(dst,"07075490449") > 0 Then 
				dststr = "대표번호"
				sRs(3) = sRs(3) + 1
			ElseIf InStr(dst,"07075490448") > 0 Then 
				dststr = "대표번호"
				sRs(4) = sRs(4) + 1
			End If 
		End If 
		dststr = dststr & "(" & dst & ")"
		sRs(5) = sRs(5) + 1


		bigo = "멘트청취후 끊음"
		If dcontext = "tr_context" Then 
			bigo = "돌려준전화"
		ElseIf lastapp = "Busy" Or lastapp = "BackGround" Then 
			Select Case Replace(lastdata,"tenbyten/","")
				Case "tenbyten_call_main"
				bigo = "대표전화안내멘트중끊음"
				Case "tenbyten_call_recall"
				bigo = "모든상담원통화중멘트중끊음"
				Case "tenbyten_main"
				bigo = "대표전화안내멘트중끊음"
				Case "tenbyten_call_lunch"
				bigo = "점심시간안내멘트중끊음"
				Case "tenbyten_call_workafter"
				bigo = "업무후안내멘트중끊음"
				Case "tenbyten_call_workbefore"
				bigo = "업무전안내멘트중끊음"
				Case "tenbyten_forword"
				bigo = "포워딩안내멘트중끊음"
				Case Else 
				bigo = ""
			End Select 
		End If 

	%>
		<td><%=i+1%></td>
		<td><%=dststr%></td>
		<td><%=rs(2,i)%></td>
		<td><%=rs(0,i)%></td>
		<td><%=bigo%></td>
	</tr>
	<%Next%>
<%
End If 
%>
</table>


<!-- 서머리 -->
<div id="divHide" style="display:none;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>일별 부재중전화 내역</td>
		<td>수신번호</td>
		<td>콜수</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td rowspan="5"><%=yyyymmdd%><br><%=getWeekDay(yyyymmdd)%></td>
		<td>콜센터_헌트(07075490429)</td>
		<td align="right"><%=FormatNumber(sRs(1),0)%></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>사무실_헌트(07075490556)</td>
		<td align="right"><%=FormatNumber(sRs(2),0)%></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>대표번호(07075490449)</td>
		<td align="right"><%=FormatNumber(sRs(3),0)%></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>대표번호(07075490448)</td>
		<td align="right"><%=FormatNumber(sRs(4),0)%></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>합계</td>
		<td align="right"><%=FormatNumber(sRs(5),0)%></td>
	</tr>
</table>
</div>
<!-- 서머리 -->

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
