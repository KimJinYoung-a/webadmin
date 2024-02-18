<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 부서 , 당담자 검색
' History : 2018.01.26 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
	Dim iTotCnt, arrList,intLoop
	Dim cPartList, arrPartList, Memberlist, i, sWorker, vChecked, iDoc_Idx, sTeam, sRefer
	dim department_id, srchWorker
	
	iDoc_Idx= NullFillWith(requestCheckVar(Request("didx"),10),"0")
	sWorker = NullFillWith(Request("worker"),"")
	sRefer	= NullFillWith(Request("refer"),"")
	srchWorker = NullFillWith(requestCheckVar(Request("srchWorker"),10),"")
	'sTeam	= NullFillWith(Request("team"),g_MyTeam)
	
	department_id = requestCheckVar(Request("department_id"),10)
	if (department_id = "") then
		department_id = GetUserDepartmentID("", session("ssBctId"))
	end if
	
	If InStr(sTeam,",") > 0 Then
		sTeam = session("ssAdminPsn")
	End IF
	
	If sWorker <> "" Then
		sWorker = sWorker & ","
	End If
	
	If sRefer <> "" Then
		sRefer = sRefer & ","
	End If

	If srchWorker<>"" then
		department_id = ""
	end if
	
	set Memberlist = new CCooperate
	Memberlist.FDoc_Idx = iDoc_Idx
	Memberlist.FRectDepartmentID = department_id
	Memberlist.FRectWorker = srchWorker			'이름으로 검색(2015.09.01; 허진원)
	arrList = Memberlist.fnGetMemberList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function workerselect(userid,username,department,departmentname) {
	var o_MDid = opener.document.itemreg.MDid;
	var o_req_department = opener.document.itemreg.req_department;

	opener.$("#divMDidname").html(username)
	o_MDid.value = userid;
	opener.$("#divdepartmentname").html(departmentname)
	o_req_department.value = department;

	temp_workerlist_js()
	return;
	//window.close();
}

function temp_workerlist_js(){
	document.getElementById("temp_workerlist").value = opener.$("#divMDidname").html();
}

function goPartSelect(part)
{
	document.location.href = "popdepartmentselect.asp?didx=<%=iDoc_Idx%>&worker=<%=sWorker%>&department_id=" + part + "";
}
</script>

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left" style="padding-bottom:3;" colspan="2">
		<input type="radio" id="selSrchTm" name="selSrch" <%=chkIIF(srchWorker="","checked","")%>>
		소속부서 : 
		<%= drawChSelectBoxDepartment("department_id", department_id,"onfocus=""document.getElementById('selSrchTm').checked=true"" onChange=""goPartSelect(this.value)""") %> 
	</td>
</tr>
<tr>
	<td align="left" style="padding-bottom:10;" colspan="2">
		<input type="radio" id="selSrchNm" name="selSrch" <%=chkIIF(srchWorker="","","checked")%>>
		<form name="srcFrm" method="GET">
		<input type="hidden" name="didx" value="<%=iDoc_Idx%>" />
		<!--<input type="hidden" name="worker" value="<%=sWorker%>" />-->
		작업자명 :
		<input type="text" name="srchWorker" value="<%=srchWorker%>" class="text" style="width:100px" onclick="document.getElementById('selSrchNm').checked=true;" />
		<input type="submit" value="검색" class="button" />
		</form>
	</td>
</tr>
<input type="hidden" name="temp_workerlist" id="temp_workerlist" value="" size="50" readonly>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#EFEFEF" height="30">			
	<td align="center">팀</td>
	<!--<td width="80" align="center">직급</td>-->
	<td width="100" align="center">이름</td>
	<td width="100" align="center">선택</td>
</tr>
<%
	IF isArray(arrList) THEN
		For intLoop = 0 To UBound(arrList,2)
			vChecked = ""
			If Instr(1, sWorker, arrList(0,intLoop)) <> 0 Then
				vChecked = "checked"
			End If
%>
	    	<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
				<td align="left"><%=arrList(1,intLoop)%><%=chkIIF(Not(arrList(7,intLoop)="" or isNull(arrList(7,intLoop))),"<br /><font color=darkgray>" & arrList(7,intLoop) & "</font>","")%></td>
				<!--<td align="center"><%=arrList(2,intLoop)%></td>-->
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
					<input type="button" value="선택" class="button" onClick="workerselect('<%=arrList(0,intLoop)%>','<%=arrList(3,intLoop)%>','<%=arrList(4,intLoop)%>','<%=arrList(1,intLoop)%>')">
					<input type="hidden" name="workername" value="<%=arrList(3,intLoop)%>">
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

<script>
temp_workerlist_js()
</script>

<%
set Memberlist = nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->