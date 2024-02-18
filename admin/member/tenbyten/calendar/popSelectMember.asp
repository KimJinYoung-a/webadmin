<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%

dim department_id
department_id = requestCheckVar(Request("department_id"),32)


dim arrList
dim Memberlist
set Memberlist = new CCooperate
Memberlist.FRectDepartmentID = department_id

if (department_id <> "") then
	arrList = Memberlist.fnGetMemberList
end if

dim intLoop

%>
<link rel="stylesheet" href="/bct.css" type="text/css">
<script language="javascript">

function putItem(valName, valId) {
	var frm = document.frm;

	if (frm.department_id.value == "") {
		alert("부서를 선택하세요.");
		frm.department_id.focus();
		return;
	}

	var result = opener.addMemberItem(valName, valId);
	if (result != "") {
		alert(result);
	}
}

</script>
<form name="frm" method="GET" action="">
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr height="10" valign="bottom" bgcolor="F4F4F4">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="bottom" bgcolor="F4F4F4">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top" bgcolor="F4F4F4"><b>부서 선택</b></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center" width="80">부 서</td>
	<td>
		<%= drawSelectBoxDepartment("department_id", department_id) %>
		&nbsp;
		<input type="button" class="button" value=" 검 색 " onClick="frm.submit();">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2">
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
		<tr bgcolor="#EFEFEF" height="30">
			<td align="center">팀</td>
			<td width="80" align="center">직급</td>
			<td width="100" align="center">이름</td>
			<td width="100" align="center">선택</td>
		</tr>
<%
	IF isArray(arrList) THEN
		For intLoop = 0 To UBound(arrList,2)
%>
	    	<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
				<td align="left"><%=arrList(1,intLoop)%><%=chkIIF(Not(arrList(7,intLoop)="" or isNull(arrList(7,intLoop))),"<br /><font color=darkgray>" & arrList(7,intLoop) & "</font>","")%></td>
				<td align="center"><%=arrList(2,intLoop)%></td>
				<td align="center"><%=arrList(3,intLoop)%>
				<%
					If Trim(arrList(6,intLoop)) <> "" Then
						If arrList(6,intLoop) = "no" Then
							Response.Write "<br>[" & "<font color=green>휴가중</font>" & "]"
						Else
							Response.Write "<br>[" & "<font color=green>반차 "&arrList(6,intLoop)&"</font>" & "]"
						End IF
					End If
				%>
				</td>
				<td align="center">
					<input type="button" class="button" value=" 추 가 " onClick="putItem('<%=arrList(1,intLoop)%> - <%=arrList(3,intLoop)%>&nbsp;<%=arrList(2,intLoop)%>', '<%= arrList(8,intLoop) %>')">
				</td>
	    	</tr>
<%
		Next
	Else
%>
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="4" align="center" class="page_link">[데이터가 없습니다.]</td>
		</tr>
<%
	End If
%>
</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center" colspan="2" height="45">
		<input type="button" class="button" value=" 닫 기 " onClick="self.close()">
	</td>
</tr>
</table>
</form>
<%

set Memberlist = Nothing

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
