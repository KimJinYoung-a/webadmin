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
<!-- #include virtual="/lib/classes/member/fingerprints/fingerprints_cls.asp" -->
<!-- #include virtual="/lib/classes/offshop/staff/offshop_employee_managementCls.asp"-->

<%
	Dim vEmpNo, vWorkDate, vWorkType, oposcodeList, i
	vEmpNo = Request("empno")
	vWorkDate = Request("wdate")
	vWorkType = Request("type")
	
	set oposcodeList = new cfingerprints_list
		oposcodeList.FPageSize=100
		oposcodeList.FCurrPage= 1
		oposcodeList.fposcode_list
%>

<script type="text/javascript">
<!--
function goSaveWorkTime()
{
	if(frm1.placeid.value == "")
	{
		alert("근무장소를 선택하세요.");
		frm1.placeid.focus();
		return false;
	}
	if(frm1.inoutdate.value == "")
	{
		alert("시간을 입력하세요.");
		frm1.inoutdate.focus();
		return false;
	}
	if(frm1.inouttime.value == "")
	{
		alert("시간을 입력하세요.");
		frm1.inouttime.focus();
		return false;
	}
	return true;
}
//-->
</script>

<form name="frm1" action="inoutTime_input_proc.asp" method="post" style="margin:0px;" onSubmit="return goSaveWorkTime();">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">사번</td>
	<td align="center" bgcolor="#FFFFFF"><%=vEmpNo%><input type="hidden" name="empno" value="<%=vEmpNo%>"></td>
</tr>
<tr height="30">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">근무일</td>
	<td align="center" bgcolor="#FFFFFF"><%=vWorkDate%><input type="hidden" name="yyyymmdd" value="<%=vWorkDate%>"></td>
</tr>
<tr height="30">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">출근/퇴근</td>
	<td align="center" bgcolor="#FFFFFF"><%=CHKIIF(vWorkType="0","출근","퇴근")%><input type="hidden" name="inouttype" value="<%=vWorkType%>"></td>
</tr>
<tr height="30">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">근무장소</td>
	<td align="center" bgcolor="#FFFFFF">
	<select name="placeid" class="select">
	<option value="">-선택-</option>
	<% for i=0 to oposcodeList.FResultCount-1 %>
		<option value="<%= oposcodeList.FItemList(i).fplaceid %>"><%= oposcodeList.FItemList(i).fplaceiname %></option>
	<% next %>
	</select>
	</td>
</tr>
<tr height="30">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">시간</td>
	<td align="center" bgcolor="#FFFFFF"><input type="text" name="inoutdate" size="10" value="<%=vWorkDate%>"> <input type="text" size="5" name="inouttime" value="00:00" maxlength="5">
	<br>※ 입력형식 예) <%=date()%> 00:00
	</td>
</tr>
<tr height="50">
	<td align="right" colspan="2" bgcolor="#FFFFFF"><input type="submit" value="저장" class="button"></td>
</tr>
</table>
</form>

<% set oposcodeList = Nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->