<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim yyyy, mm
dim i

yyyy	= req("yyyy1", Left(Date,4))
mm		= req("mm1", Mid(Date,6,2))

Dim strSql
strSql = " db_datamart.dbo.sp_Ten_Call_Absence_Month ('" & yyyy & "-" & mm & "')"

db3_rsget.CursorLocation = adUseClient
db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly, adLockReadOnly, adCmdStoredProc	

Dim rs 
If Not db3_rsget.EOF Then
	rs = db3_rsget.getRows()
End If 
db3_rsget.close


%>

<script language='javascript'>


</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
	       	년월: &nbsp;<% DrawYMBox yyyy,mm %>
			&nbsp;
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->
 
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">

        <td>날짜</td>
        <td>요일</td>
        <td width="100">콜센터_헌트(07075490429)</td>
        <td width="100">사무실_헌트(07075490556)</td>
        <td width="100">대표번호(07075490449)</td>
        <td width="100">대표번호(07075490448)</td>
        <td width="100">건수</td>

</tr>
<%
Dim rowCnt
Dim sRs(20)

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

If IsArray(rs) Then 
	rowCnt = UBound(rs,2) + 1
%>

	<%For i=0 To UBound(rs,2)%>
    <tr align="center" bgcolor="#FFFFFF">
	<%
		' Row 합산
		sRs(1) = sRs(1) + CDbl(rs(1,i))
		sRs(2) = sRs(2) + CDbl(rs(2,i))
		sRs(3) = sRs(3) + CDbl(rs(3,i))
		sRs(4) = sRs(4) + CDbl(rs(4,i))
		sRs(5) = sRs(5) + CDbl(rs(5,i))

	%>
		<td><%=rs(0,i)%></td>
		<td><%=getWeekDay(rs(0,i))%></td>
		<td><%=FormatNumber(rs(2,i),0)%></td>
		<td><%=FormatNumber(rs(3,i),0)%></td>
		<td><%=FormatNumber(rs(4,i),0)%></td>
		<td><%=FormatNumber(rs(5,i),0)%></td>
		<td><a href="callAbsenceList.asp?yyyymmdd=<%=rs(0,i)%>"><%=FormatNumber(rs(1,i),0)%></a></td>
	</tr>
	<%Next%>
    <tr align="center" bgcolor="#FFFFFF">
	<%
	%>
    	<td><b>합계 or 평균</b></td>
    	<td></td>
		<td><b><%=FormatNumber(sRs(2),0)%></b></td>
		<td><b><%=FormatNumber(sRs(3),0)%></b></td>
		<td><b><%=FormatNumber(sRs(4),0)%></b></td>
		<td><b><%=FormatNumber(sRs(5),0)%></b></td>
		<td><b><%=FormatNumber(sRs(1),0)%></b></td>
    </tr>
<%
End If 
%>
</table>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
