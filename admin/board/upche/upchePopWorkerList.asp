<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<!-- #include virtual="/lib/classes/board/companyrequestcls.asp" -->
<%
	Dim iTotCnt, arrList,intLoop
	Dim cPartList, arrPartList, Memberlist, i, sTeam
	Dim id, oneWorkName, mode, userid

	mode =requestCheckVar(Request("mode"),1)
	userid =requestCheckVar(Request("userid"),34)
	id = NullFillWith(Request("id"),"")
	sTeam	= NullFillWith(Request("team"),g_MyTeam)
	oneWorkName = getUpcheoneWorkname(id)

	Set cPartList = new CCooperate
	arrPartList = cPartList.fnPartList
	Set cPartList = Nothing
	
	set Memberlist = new CCooperate
	Memberlist.FTeam = sTeam
	arrList = Memberlist.fnGetMemberList

	If mode = "U" Then
		Dim strSQL
		strSQL = ""
		strSQL = strSQL & " UPDATE [db_cs].[dbo].tbl_company_request SET " & VBCRLF
		strSQL = strSQL & " workid = '"&userid&"' " & VBCRLF
		strSQL = strSQL & " WHERE id = '"&id&"' " & VBCRLF
		dbget.execute(strSQL)
		Response.Write	"<script language='javascript'>" &_
						"alert('�����Ͽ����ϴ�.');" &_
						"opener.location.reload();"&_
						"window.close();"&_
						"</script>"&_
		dbget.close()	:	response.End
	End If
%>

<script language="javascript">
function goPartSelect(part)
{
	document.location.href = "upchePopWorkerList.asp?id=<%=id%>&team=" + part + "";
}
function evtworkSEL(wkID){
	if(confirm("�����Ͻ� ����ڷ� �����Ͻðڽ��ϱ�?")){
		document.efrm.userid.value = wkID;
		document.efrm.submit();
	}
}
</script>

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left" style="padding-bottom:10;" colspan="2">
		<select name="team" class="select" onChange="goPartSelect(this.value)">
		<%
			
			IF isArray(arrPartList) THEN
				For intLoop = 0 To UBound(arrPartList,2)
					If arrPartList(0,intLoop) <> "1" Then
						Response.Write "<option value=""" & arrPartList(0,intLoop) & """ "
						If CStr(arrPartList(0,intLoop)) = CStr(sTeam) Then
							Response.Write "selected"
						End If
						Response.Write ">" & arrPartList(1,intLoop) & "</option>"
					End If
				Next
			End If
		%>
		</select>
	</td>
</tr>
</table>
<%
	If oneWorkName <> "" Then
		response.write "�� ���� ����� : "&oneWorkName&" "
	End If
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#EFEFEF" height="30">			
	<td align="center">��</td>
	<td width="80" align="center">����</td>
	<td width="100" align="center">�̸�</td>
	<td width="100" align="center">����</td>
</tr>
<%
	IF isArray(arrList) THEN
		For intLoop = 0 To UBound(arrList,2)
%>
	    	<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
				<td align="left"><%=arrList(1,intLoop)%><%=chkIIF(arrList(0,intLoop)="eoslove","-����","")%></td>
				<td align="center"><%=arrList(2,intLoop)%></td>
				<td align="center"><%=arrList(3,intLoop)%>
				<%
					If Trim(arrList(6,intLoop)) <> "" Then
						If arrList(6,intLoop) = "no" Then
							Response.Write "[" & "<font color=green>��</font>" & "]"
						Else
							Response.Write "[" & "<font color=green>��"&arrList(6,intLoop)&"</font>" & "]"
						End IF
					End If
				%>
				</td>
				<td align="center">
					<input type="button" value="����" class="button" onClick="evtworkSEL('<%=arrList(0,intLoop)%>')">
				</td>
	    	</tr>
<%
		Next
	Else
%>
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="4" align="center" class="page_link">[�����Ͱ� �����ϴ�.]</td>
		</tr>
<%
	End If
%>
</table>
<form name="efrm" action="<%= CurrURL %>" method="POST" style="margin:0px;">
<input type="hidden" name="mode" value="U">
<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="userid" value="">
</form>
<% Set Memberlist = nothing %>