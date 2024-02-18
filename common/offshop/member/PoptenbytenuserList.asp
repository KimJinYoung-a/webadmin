<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  직원리스트
' History : 2012.06.11 강준구 최초생성
'			2012.07.17 한용민 이동생성(공용 팝업으로 수정)
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
Dim iTotCnt, arrList,intLoop ,cPartList, arrPartList, Memberlist, i, sWorker, vChecked, iDoc_Idx, sTeam, sRefer
dim username ,boxname,department_id
	iDoc_Idx= NullFillWith(requestCheckVar(Request("didx"),10),"0")
	sWorker = NullFillWith(Request("worker"),"")
	sRefer	= NullFillWith(Request("refer"),"")
	sTeam	= NullFillWith(Request("team"),g_MyTeam)
	username= NullFillWith(requestCheckVar(Request("username"),32),"")
	boxname= NullFillWith(requestCheckVar(Request("boxname"),32),"")
	department_id = requestCheckVar(Request("department_id"),10)

	if (department_id = "") then
		department_id = GetUserDepartmentID("", session("ssBctId"))
	end if

If sWorker <> "" Then
	sWorker = sWorker & ","
End If

If sRefer <> "" Then
	sRefer = sRefer & ","
End If

Set cPartList = new CCooperate
	arrPartList = cPartList.fnPartList
Set cPartList = Nothing

set Memberlist = new CCooperate
	Memberlist.FDoc_Idx = iDoc_Idx
	Memberlist.FTeam = department_id
	Memberlist.FRectDepartmentID = department_id
	arrList = Memberlist.fnGetMemberList
%>

<script language="javascript">

function fFindText(strText,writeText)
{
	var arrText = strText.split(",");
	var trueorfalse = false;

	for(var i=0; i<arrText.length; i++)
	{
		if(writeText == arrText[i])
		{
			trueorfalse = true;
		}
	}

	return trueorfalse;
}

function workerselect(wid,wname)
{
	var o_wname = opener.document.getElementsByName("username")[0];
	var o_wid = opener.document.getElementsByName("<%= boxname %>")[0];

	o_wname.value =  wname;
	o_wid.value =  wid;

	temp_workerlist_js()
	window.close();
}

function temp_workerlist_js()
{
	document.getElementById("temp_workerlist").value = '<%=sWorker%>';
}

function goPartSelect(part)
{
	document.location.href = "/common/offshop/member/PoptenbytenuserList.asp?didx=<%=iDoc_Idx%>&worker=<%=sWorker%>&boxname=<%=boxname%>&username=<%=username%>&department_id=" + part + "";
}

</script>

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left" style="padding-bottom:10;" colspan="2">
		<%= drawChSelectBoxDepartment("department_id", department_id,"onChange='goPartSelect(this.value)'") %>
	</td>
</tr>
<tr>
	<td align="left" style="padding-bottom:3;">담당자 : <input type="text" name="temp_workerlist" id="temp_workerlist" value="" size="10" readonly></td>
	<td align="right"><input type="button" value="닫 기" class="button" onClick="window.close()"></td>
</tr>
<tr>
	<td colspan="2"><font color="blue">※ 선택된 담당자를 삭제 하시려면 해당 담당자를 한번 더 클릭하시면 삭제가 됩니다.<br>&nbsp;&nbsp;&nbsp;&nbsp;담당자 입력란을 비워두지 마시고 채워둔 후 삭제 하세요.</font></td>
</tr>
</table>
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
			vChecked = ""
			If Instr(1, sWorker, arrList(0,intLoop)) <> 0 Then
				vChecked = "checked"
			End If
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
					<input type="button" value="지정" class="button" onClick="workerselect('<%=arrList(0,intLoop)%>','<%=arrList(3,intLoop)%>')">
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
Set Memberlist = nothing
%>
