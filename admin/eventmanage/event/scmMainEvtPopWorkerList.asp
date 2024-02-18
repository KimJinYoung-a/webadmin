<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<!-- #include virtual="/lib/classes/member_board/boardCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
	Dim iTotCnt, arrList,intLoop
	Dim cPartList, arrPartList, Memberlist, i, sTeam
	Dim eCode, oneWorkName, mode, userid
	Dim department_id
	
	mode =requestCheckVar(Request("mode"),1)
	userid =requestCheckVar(Request("userid"),34)
	eCode = NullFillWith(Request("eCode"),"")
	 
	department_id = requestCheckVar(Request("department_id"),10)
	if (department_id = "") then
		department_id = GetUserDepartmentID("", session("ssBctId"))
	end if
	oneWorkName = getEvtoneWorkname(eCode)
 
	
	set Memberlist = new CCooperate
	Memberlist.FRectDepartmentID = department_id
	arrList = Memberlist.fnGetMemberList

	If mode = "U" Then
		Dim strSQL
		strSQL = ""
		strSQL = strSQL & " UPDATE db_event.dbo.tbl_event_display SET " & VBCRLF
		strSQL = strSQL & " partMDid = '"&userid&"' " & VBCRLF
		strSQL = strSQL & " WHERE evt_code = '"&eCode&"' " & VBCRLF
		dbget.execute(strSQL)
		Response.Write	"<script language='javascript'>" &_
						"alert('저장하였습니다.');" &_
						"opener.location.reload();"&_
						"window.close();"&_
						"</script>"&_
		dbget.close()	:	response.End
	End If
%>

<script language="javascript">
function goPartSelect(part)
{
	document.location.href = "scmMainEvtPopWorkerList.asp?eCode=<%=eCode%>&department_id=" + part + "";
}
function evtworkSEL(wkID){
	if(confirm("선택하신 담당자로 지정하시겠습니까?")){
		document.efrm.userid.value = wkID;
		document.efrm.submit();
	}
}
</script>

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left" style="padding-bottom:10;" colspan="2">
		<%= drawChSelectBoxDepartment("department_id", department_id,"onChange='goPartSelect(this.value)'")%> 
	</td>
</tr>
</table>
<%
	If oneWorkName <> "" Then
		response.write "※ 현재 담당자 : "&oneWorkName&" "
	End If
%>
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
				<td align="left"><%=arrList(1,intLoop)%><%=chkIIF(arrList(0,intLoop)="eoslove","-포토","")%></td>
				<td align="center"><%=arrList(2,intLoop)%></td>
				<td align="center"><%=arrList(3,intLoop)%>
				<%
					If Trim(arrList(6,intLoop)) <> "" Then
						If arrList(6,intLoop) = "no" Then
							Response.Write "[" & "<font color=green>휴</font>" & "]"
						Else
							Response.Write "[" & "<font color=green>반"&arrList(6,intLoop)&"</font>" & "]"
						End IF
					End If
				%>
				</td>
				<td align="center">
					<input type="button" value="지정" class="button" onClick="evtworkSEL('<%=arrList(0,intLoop)%>')">
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
<form name="efrm" action="<%= CurrURL %>" method="POST" style="margin:0px;">
<input type="hidden" name="mode" value="U">
<input type="hidden" name="eCode" value="<%=eCode%>">
<input type="hidden" name="userid" value="">
</form>
<% Set Memberlist = nothing %>