<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/offshop/staff/offshop_employee_managementCls.asp"-->

<%
	Dim i, cWorkCode, vWorkCode, vStartWork, vEndWork
	vWorkCode = Request("wc")
	
	If vWorkCode <> "" Then
		Set cWorkCode = New cEmployeeManagementClass_list
		cWorkCode.FRectWorkCode = vWorkCode
		cWorkCode.fWorkCodeView()
		
		vStartWork = fnChangeTimeType(cWorkCode.FOneItem.FStartWork)
		vEndWork = fnChangeTimeType(cWorkCode.FOneItem.FEndWork)
		Set cWorkCode = Nothing
	End IF
	
	Set cWorkCode = New cEmployeeManagementClass_list
	cWorkCode.fWorkCodeList()
%>

<script type="text/javascript">
<!--
function jsEditWorkCode(wc)
{
	location.href = "<%=CurrURL()%>?wc="+wc+"";
}

function goSaveWorkCode()
{
	if(frm1.workcode.value == "")
	{
		alert("근무코드를 입력하세요.");
		frm1.workcode.focus();
		return false;
	}
	if(frm1.startwork.value == "")
	{
		alert("출근시간을 입력하세요.");
		frm1.startwork.focus();
		return false;
	}
	return true;
}
//-->
</script>

※ <b>출근시간 입력시</b> 시간을 입력할때는 <b>09:00 형식</b>으로 입력하고<br><b>일반글자로 입력</b>할때는 <b>퇴근시간은 반드시 비워</b>두세요.
<form name="frm1" action="offshop_employee_workcode_proc.asp" method="post" style="margin:0px;" onSubmit="return goSaveWorkCode();">
<input type="hidden" name="action" value="<%=CHKIIF(vWorkCode<>"","update","insert")%>">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">근무코드</td>
    <td align="center" width="80">출근시간</td>
    <td align="center" width="80">퇴근시간</td>
    <td align="center" width="70">근무시간</td>
    <td align="center" width="60"></td>
</tr>
<tr bgcolor="#B7F0B1" height="50">
	<td align="center"><input type="text" size="5" name="workcode" value="<%= vWorkCode %>" maxlength="2" style="text-align:center;" <%=CHKIIF(vWorkCode<>"","readonly","")%>><br>대문자, 변경불가</td>
	<td align="center"><input type="text" size="7" name="startwork" value="<%= vStartWork %>" style="text-align:center;"></td>
	<td align="center"><input type="text" size="7" name="endwork" value="<%= vEndWork %>" maxlength="5" style="text-align:center;"></td>
	<td align="center">자동계산</td>
	<td align="center"><input type="submit" value="저장" class="button"></td>
</tr>
<%
	For i = 0 To cWorkCode.FTotalCount -1
%>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= cWorkCode.flist(i).FWorkCode %></td>
	<td align="center"><%= fnChangeTimeType(cWorkCode.FList(i).FStartWork) %></td>
	<td align="center"><%= fnChangeTimeType(cWorkCode.FList(i).FEndWork) %></td>
	<td align="center"><%= fnWorkTimeCalc(cWorkCode.FList(i).FStartWork, cWorkCode.FList(i).FEndWork) %></td>
	<td align="center"><input type="button" value="수정" class="button" onClick="jsEditWorkCode('<%= cWorkCode.flist(i).FWorkCode %>');"></td>
</tr>
<%
	Next
%>
</table>
</form>

<% Set cWorkCode = Nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->