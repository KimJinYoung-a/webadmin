<%@ language=vbscript %>
<% option explicit %>
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
	Memberlist.FRectWorker = srchWorker			'�̸����� �˻�(2015.09.01; ������)
	arrList = Memberlist.fnGetMemberList
%>

<script type="text/javascript">
function checkworker(tmp) {
	var Element = document.getElementsByName("workerid")[tmp];
	if (Element.checked == true) {
		Element.checked = false;
	} else {
		Element.checked = true;
	}
}

function fFindText(strText,writeText) {
	var arrText = strText.split(",");
	var trueorfalse = false;

	for(var i=0; i<arrText.length; i++) {
		if(writeText == arrText[i]) {
			trueorfalse = true;
		}
	}

	return trueorfalse;
}

function workerselect(wid,wname) {
	var o_wname = opener.document.frm.doc_workername;
	var o_wid = opener.document.frm.doc_worker;
	var chktemp = opener.document.forms["frm"].elements["doc_worker"];


	<% ''####### �۾��� ������ �Ѹ� �ǰ� �ش޶�� ����� ����(20120608). ���� ������� �׼� PopWorkerList_20120608bakup.asp �� ����. %>
	if(!(fFindText(chktemp.value,wid)))
	{
		if(o_wid.value != "" && o_wid.value.split(",").length == 1)
		{
			alert("�۾��ڴ� 1�� �����մϴ�.\n\n�� �۾��ڸ���Ʈâ�� ������ ������� �о��ּ���.");
			return;
		}
	}

	if(!(fFindText(chktemp.value,wid)))
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
	//window.close();
}

function referselect(rid,rname)
{
	var o_rname = opener.document.frm.doc_refername;
	var o_rid = opener.document.frm.doc_refer;
	var chktempp = opener.document.forms["frm"].elements["doc_refer"];


	//if(!(chktempp.createTextRange().findText(rid,rid.length,0)))
	if(!(fFindText(chktempp.value,rid)))
	{
		if(o_rid.value == "")
		{
			o_rname.value = o_rname.value + "" + rname;
			o_rid.value = o_rid.value + "" + rid;
		}
		else
		{
			o_rname.value = o_rname.value + "," + rname;
			o_rid.value = o_rid.value + "," + rid;
		}
	}
	else
	{
		o_rname.value = o_rname.value.replace(rname,"");
		o_rid.value = o_rid.value.replace(rid,"");

		o_rname.value = o_rname.value.replace(",,",",");
		o_rid.value = o_rid.value.replace(",,",",");

		if(o_rid.value.substring(0,1) == ",")
		{
			o_rname.value = o_rname.value.substring(1,o_rname.value.length);
			o_rid.value = o_rid.value.substring(1,o_rid.value.length);
		}


		if(o_rid.value.substring(o_rid.value.length-1,o_rid.value.length) == ",")
		{
			o_rname.value = o_rname.value.substring(0,o_rname.value.length-1);
			o_rid.value = o_rid.value.substring(0,o_rid.value.length-1);
		}
	}

	temp_referlist_js()
return;
	//window.close();
}

function allcheck(g)
{
	if(g == "o")
	{
		document.getElementById("allchk").style.display = "none";
		document.getElementById("nonechk").style.display = "";
		for(i = 0; i < document.getElementsByName("workerid").length; i++) 
		{ 
			document.getElementsByName("workerid").item(i).checked = true; 
		}
	}
	else
	{
		document.getElementById("allchk").style.display = "";
		document.getElementById("nonechk").style.display = "none";
		for(i = 0; i < document.getElementsByName("workerid").length; i++) 
		{ 
			document.getElementsByName("workerid").item(i).checked = false; 
		}
	}
}


function temp_workerlist_js()
{
	document.getElementById("temp_workerlist").value = opener.document.frm.doc_workername.value;
}

function temp_referlist_js()
{
	document.getElementById("temp_referlist").value = opener.document.frm.doc_refername.value;
}

function goPartSelect(part)
{
	document.location.href = "PopWorkerList.asp?didx=<%=iDoc_Idx%>&worker=<%=sWorker%>&department_id=" + part + "";
}
</script>

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left" style="padding-bottom:3;" colspan="2">
		<input type="radio" id="selSrchTm" name="selSrch" <%=chkIIF(srchWorker="","checked","")%>>
		�ҼӺμ� : 
		<%= drawChSelectBoxDepartment("department_id", department_id,"onfocus=""document.getElementById('selSrchTm').checked=true"" onChange=""goPartSelect(this.value)""") %> 
		<!--
		����Կ� ���Ͽ� ���� ��!! 2012-03-06
		<a href="PopWorkerList.asp?didx=<%=iDoc_Idx%>&worker=<%=sWorker%>&team=9,10">�������</a>
		|<a href="PopWorkerList.asp?didx=<%=iDoc_Idx%>&worker=<%=sWorker%>&team=11,12,14,16">�ٹ����ٿ¶��λ����</a>
		|<a href="PopWorkerList.asp?didx=<%=iDoc_Idx%>&worker=<%=sWorker%>&team=13,18">����������</a>
		|<a href="PopWorkerList.asp?didx=<%=iDoc_Idx%>&worker=<%=sWorker%>&team=7">�ý�����</a>
		|<a href="PopWorkerList.asp?didx=<%=iDoc_Idx%>&worker=<%=sWorker%>&team=8">�濵������</a>
		|<a href="PopWorkerList.asp?didx=<%=iDoc_Idx%>&worker=<%=sWorker%>&team=15,19">���̶����</a>
		|<a href="PopWorkerList.asp?didx=<%=iDoc_Idx%>&worker=<%=sWorker%>&team=17">�мǻ����</a>
		//-->
	</td>
</tr>
<tr>
	<td align="left" style="padding-bottom:10;" colspan="2">
		<input type="radio" id="selSrchNm" name="selSrch" <%=chkIIF(srchWorker="","","checked")%>>
		<form name="srcFrm" method="GET">
		<input type="hidden" name="didx" value="<%=iDoc_Idx%>" />
		<!--<input type="hidden" name="worker" value="<%=sWorker%>" />-->
		�۾��ڸ� :
		<input type="text" name="srchWorker" value="<%=srchWorker%>" class="text" style="width:100px" onclick="document.getElementById('selSrchNm').checked=true;" />
		<input type="submit" value="�˻�" class="button" />
		</form>
	</td>
</tr>
<!--
<tr>
	<td align="left" style="padding-bottom:3;">
		<div id="allchk" style="display:'';">
		<input type="button" value="��ü����" class="button" onClick="allcheck('o')">
		</div>
		<div id="nonechk" style="display:'none';">
		<input type="button" value="��ü����" class="button" onClick="allcheck('n')">
		</div>
	</td>
	<td align="right" style="padding-bottom:3;"><input type="button" value="��������" class="button" onClick="workerselect()"></td>
</tr>
//-->
<tr>
	<td colspan="2"><font color="red">�� �۾��ڰ� �������� ��� �۾��ڸ� �����(����)���� �Ͻð� ���� �۾��ڸ� �����ڷ� �����Ͻø�
		<br>&nbsp;&nbsp;&nbsp;&nbsp;�˴ϴ�. �۾��� ��� �Ϸᰡ �Ǹ� ����� ������ �ϷḦ ��Ű�� �˴ϴ�.</font>
	</td>
</tr>
<tr>
	<td align="left" style="padding-bottom:3;">�۾� : <input type="text" name="temp_workerlist" id="temp_workerlist" value="" size="50" readonly></td>
	<td align="right"><input type="button" value="�� ��" class="button" onClick="window.close()"></td>
</tr>
<tr>
	<td align="left" style="padding-bottom:3;">���� : <input type="text" name="temp_referlist" id="temp_referlist" value="" size="50" readonly></td>
	<td align="right"></td>
</tr>
<tr>
	<td colspan="2"><font color="blue">�� ���õ� �۾��ڸ� ���� �Ͻ÷��� �ش� �۾��ڸ� �ѹ� �� Ŭ���Ͻø� ������ �˴ϴ�.<br>&nbsp;&nbsp;&nbsp;&nbsp;�۾��� �Է¶��� ������� ���ð� ä���� �� ���� �ϼ���.</font></td>
</tr>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#EFEFEF" height="30">			
	<td align="center">��</td>
	<!--<td width="80" align="center">����</td>-->
	<td width="100" align="center">�̸�</td>
	<td width="100" align="center">����</td>
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
							Response.Write "<br>[" & "<font color=green>�ް���</font>" & "]"
						Else
							Response.Write "<br>[" & "<font color=green>���� "&arrList(6,intLoop)&"</font>" & "]"
						End IF
					End If
				%>
				</td>
				<td align="center">
					<input type="button" value="�۾�" class="button" onClick="workerselect('<%=arrList(0,intLoop)%>|<%=arrList(4,intLoop)%>','<%=arrList(3,intLoop)%>')">
					<input type="hidden" name="workername" value="<%=arrList(3,intLoop)%>">
					&nbsp;
					<input type="button" value="����" class="button" onClick="referselect('<%=arrList(0,intLoop)%>|<%=arrList(4,intLoop)%>','<%=arrList(3,intLoop)%>')">
					<input type="hidden" name="refername" value="<%=arrList(3,intLoop)%>">
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
<!--
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td align="right" style="padding-top:3;"><input type="button" value="��������" class="button" onClick="workerselect()"></td>
</tr>
</table>
//-->

<script>
temp_workerlist_js()
temp_referlist_js()
</script>

<%
	set Memberlist = nothing
%>