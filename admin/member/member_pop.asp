<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  파트관리자 담당자리스트 폼
' History : 2011.01.25 김진영 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/admin/partpersonCls.asp"-->
<%
	Dim iTotCnt, arrList,intLoop, arrPartList
	Dim Memberlist, i, sWorker, vChecked, idx, sTeam
	idx = requestCheckVar(request("idx"),10)
	sWorker = NullFillWith(Request("worker"),"")
	sTeam	= NullFillWith(Request("team"),g_MyTeam)
	If sWorker <> "" Then
		sWorker = sWorker & ","
	End If

	Set Memberlist = new Partlist
	Memberlist.FTeam = sTeam
	arrList = Memberlist.fnMemberList
	arrPartList = Memberlist.fnPartList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function checkworker(tmp)
{
	var Element = document.getElementsByName("workerid")[tmp];
	if (Element.checked == true)
	{
		Element.checked = false;
	}
	else
	{
		Element.checked = true;
	}
}
function workerselect(wid,wname)
{
	var o_wname = opener.document.iform.doc_workername;
	var o_wid = opener.document.iform.doc_worker;

	if(!(o_wid.value.match(wid)))
	{
		if(o_wid.value == "")
		{
			o_wname.value = o_wname.value + "" + wname;
			o_wid.value = o_wid.value + "" + wid;
		}
		else
		{
			o_wname.value = o_wname.value + "," + wname;
			o_wid.value = o_wid.value + "," + wid;
		}
	}
	else
	{
		o_wname.value = o_wname.value.replace(wname,"");
		o_wid.value = o_wid.value.replace(wid,"");

		o_wname.value = o_wname.value.replace(",,",",");
		o_wid.value = o_wid.value.replace(",,",",");

		if(o_wid.value.substring(0,1) == ",")
		{
			o_wname.value = o_wname.value.substring(1,o_wname.value.length);
			o_wid.value = o_wid.value.substring(1,o_wid.value.length);
		}
		if(o_wid.value.substring(o_wid.value.length-1,o_wid.value.length) == ",")
		{
			o_wname.value = o_wname.value.substring(0,o_wname.value.length-1);
			o_wid.value = o_wid.value.substring(0,o_wid.value.length-1);
		}
	}

	temp_workerlist_js()
return;
}
function temp_workerlist_js()
{
	document.getElementById("temp_workerlist").value = opener.document.iform.doc_workername.value;
}
function goPartSelect(part)
{
	document.location.href = "member_pop.asp?worker=<%=sWorker%>&team=" + part + "";
}
function clearList(){
	$('#temp_workerlist').val('');
	opener.document.iform.doc_workername.value = '';
	opener.document.iform.doc_worker.value = '';
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
<!--
		<a href="member_pop.asp?idx=<%=idx%>&worker=<%=sWorker%>&team=9,10">운영관리팀</a>
		|<a href="member_pop.asp?idx=<%=idx%>&worker=<%=sWorker%>&team=11,12,14,16">텐바이텐온라인사업팀</a>
		|<a href="member_pop.asp?idx=<%=idx%>&worker=<%=sWorker%>&team=13,18">오프라인팀</a>
		|<a href="member_pop.asp?idx=<%=idx%>&worker=<%=sWorker%>&team=7">시스템팀</a>
		|<a href="member_pop.asp?idx=<%=idx%>&worker=<%=sWorker%>&team=8">경영관리팀</a>
		|<a href="member_pop.asp?idx=<%=idx%>&worker=<%=sWorker%>&team=15">아이띵소팀</a>
		|<a href="member_pop.asp?idx=<%=idx%>&worker=<%=sWorker%>&team=17">패션사업팀</a>
-->
	</td>
</tr>

<tr>
	<td align="left" style="padding-bottom:3;">
		<input type="text" name="temp_workerlist" id="temp_workerlist" value="" size="50" readonly>
		<input type="button" value="Clear" class="button" onClick="clearList();">
	</td>
	<td align="right"><input type="button" value="닫 기" class="button" onClick="window.close()"></td>
</tr>
<tr>
	<td colspan="2">※ 색깔이 바뀐 라인을 클릭하셔도 체크가 됩니다.</td>
</tr>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#EFEFEF" height="30">
	<td align="center">팀</td>
	<!--<td width="80" align="center">직급</td>-->
	<td width="100" align="center">이름</td>
	<td width="50" align="center">선택</td>
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
				<td align="left" style="cursor:pointer" onClick="workerselect('<%=arrList(5,intLoop)%>','<%=arrList(3,intLoop)%>')"><%=arrList(1,intLoop)%></td>
				<!--<td align="center" style="cursor:pointer" onClick="workerselect('<%=arrList(5,intLoop)%>','<%=arrList(3,intLoop)%>')"><%=arrList(2,intLoop)%></td>-->
				<td align="center" style="cursor:pointer" onClick="workerselect('<%=arrList(5,intLoop)%>','<%=arrList(3,intLoop)%>')"><%=arrList(3,intLoop)%></td>
				<td align="center">
					<input type="button" value="선택" class="button" onClick="workerselect('<%=arrList(5,intLoop)%>','<%=arrList(3,intLoop)%>')">
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

<script>temp_workerlist_js()</script>

<%
	Set Memberlist = nothing
%>