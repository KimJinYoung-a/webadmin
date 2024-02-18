<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/board/holiday/holidayCls.asp"-->
<%
Dim oHoliday, num, mode, strSQL
Dim holiday, logics_holiday, upche_holiday, holiday_name
num = request("num")
mode = request("mode")

If mode = "U" Then
	holiday			= request("holiday")
	logics_holiday	= request("logics_holiday")
	upche_holiday	= request("upche_holiday")
	holiday_name	= request("holiday_name")

	strSQL = ""
	strSQL = strSQL & " UPDATE [db_sitemaster].[dbo].[LunarToSolar] "
	strSQL = strSQL & " SET holiday_name= '"& holiday_name &"' "
	strSQL = strSQL & " , holiday = '"& Chkiif(holiday="2", "2", "0") &"' "
	strSQL = strSQL & " , logics_holiday = '"& Chkiif(logics_holiday="2", "2", "0") &"' "
	strSQL = strSQL & " , upche_holiday = '"& Chkiif(upche_holiday="2", "2", "0") &"' "
	strSQL = strSQL & " WHERE num = '"&num&"' "
	dbget.Execute strSQL
	response.write "<script>opener.location.reload();window.close();</script>"
	response.end
End If

Dim dayColor

Set oHoliday = new CHoliday
	oHoliday.FRectNum = num
	oHoliday.getHolidayOneItem

	If oHoliday.FOneItem.FWeek ="일" Then
		dayColor = "RED"
	ElseIf oHoliday.FOneItem.FWeek ="토" Then
		dayColor = "BLUE"
	Else
		dayColor = "BLACK"
	End If

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function memoUpdate(){
	var frm;
	frm = document.frm;

	if (confirm('저장 하시겠습니까?')){
		frm.submit();
	}
}
</script>
<form name="frm" method="post" action="popholiday.asp">
<input type="hidden" name="mode" value="U">
<input type="hidden" name="num" value="<%= num %>">
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="50" width="30%">
    <td align="center" bgcolor="#E8E8FF">날짜</td>
    <td bgcolor="#FFFFFF">
		<font color="<%= dayColor %>"><%= oHoliday.FOneItem.FSolar_date %>&nbsp;(<%= oHoliday.FOneItem.FWeek %>)</font>
	</td>
</tr>
<tr height="50">
    <td align="center" bgcolor="#E8E8FF">휴일명</td>
    <td bgcolor="#FFFFFF">
		<input type="text" name="holiday_name" value="<%= oHoliday.FOneItem.FHoliday_name %>">
	</td>
</tr>
<tr height="50">
    <td align="center" bgcolor="#E8E8FF">휴무 설정</td>
    <td bgcolor="#FFFFFF">
		<label><input type="checkbox" <% If oHoliday.FOneItem.FHoliday ="1" or oHoliday.FOneItem.FHoliday ="2" Then response.write "checked" %> name="holiday" class="checkbox" value="2">텐바이텐</label>&nbsp;
		<label><input type="checkbox" <%= Chkiif(oHoliday.FOneItem.FLogics_holiday ="2", "checked", "") %> name="logics_holiday" class="checkbox" value="2">물류</label>&nbsp;
		<label><input type="checkbox" <%= Chkiif(oHoliday.FOneItem.FUpche_holiday ="2", "checked", "") %> name="upche_holiday" class="checkbox" value="2">업체</label>
	</td>
</tr>
<tr align="center" height="25" bgcolor="#FFFFFF">
    <td colspan="2" align="center">
    	<input type="button" value="변경" class="button" onClick="memoUpdate();">
    </td>
</tr>
</table>
</form>
<% Set oHoliday = nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->