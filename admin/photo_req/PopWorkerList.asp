<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->

<%
	Dim iTotCnt, arrList,intLoop
	Dim Memberlist, i, sWorker, vChecked, iDoc_Idx, sTeam, sRefer
	iDoc_Idx= NullFillWith(requestCheckVar(Request("didx"),10),"0")
	sWorker = NullFillWith(Request("worker"),"")
	sRefer	= NullFillWith(Request("refer"),"")
	sTeam	= NullFillWith(Request("team"),g_MyTeam)
	
	If sWorker <> "" Then
		sWorker = sWorker & ","
	End If
	
	If sRefer <> "" Then
		sRefer = sRefer & ","
	End If
	
	set Memberlist = new CCooperate
	Memberlist.FDoc_Idx = iDoc_Idx
	Memberlist.FTeam = sTeam
	arrList = Memberlist.fnGetMemberList
%>

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
	var o_wname = opener.document.frm.doc_workername;
	var o_wid = opener.document.frm.doc_worker;
	var chktemp = opener.document.forms["frm"].elements["doc_worker"];


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
		<% If iDoc_Idx <> "0" Then %>
			if(o_wid.value.split(",").length == 1)
			{
				alert('작업자를 1명이상 선택하셔야 합니다.');
				return;
			}
		<% End If %>

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
</script>

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left" style="padding-bottom:10;" colspan="2">
		<a href="PopWorkerList.asp?didx=<%=iDoc_Idx%>&worker=<%=sWorker%>&team=9,10">운영관리팀</a>
		|<a href="PopWorkerList.asp?didx=<%=iDoc_Idx%>&worker=<%=sWorker%>&team=11,12,14,16">텐바이텐온라인사업팀</a>
		|<a href="PopWorkerList.asp?didx=<%=iDoc_Idx%>&worker=<%=sWorker%>&team=13,18">오프라인팀</a>
		|<a href="PopWorkerList.asp?didx=<%=iDoc_Idx%>&worker=<%=sWorker%>&team=7">시스템팀</a>
		|<a href="PopWorkerList.asp?didx=<%=iDoc_Idx%>&worker=<%=sWorker%>&team=8">경영관리팀</a>
		|<a href="PopWorkerList.asp?didx=<%=iDoc_Idx%>&worker=<%=sWorker%>&team=15,19">아이띵소팀</a>
		|<a href="PopWorkerList.asp?didx=<%=iDoc_Idx%>&worker=<%=sWorker%>&team=17">패션사업팀</a>
	</td>
</tr>
<!--
<tr>
	<td align="left" style="padding-bottom:3;">
		<div id="allchk" style="display:'';">
		<input type="button" value="전체선택" class="button" onClick="allcheck('o')">
		</div>
		<div id="nonechk" style="display:'none';">
		<input type="button" value="전체해제" class="button" onClick="allcheck('n')">
		</div>
	</td>
	<td align="right" style="padding-bottom:3;"><input type="button" value="선택적용" class="button" onClick="workerselect()"></td>
</tr>
//-->
<tr>
	<td align="left" style="padding-bottom:3;">작업 : <input type="text" name="temp_workerlist" id="temp_workerlist" value="" size="50" readonly></td>
	<td align="right"><input type="button" value="닫 기" class="button" onClick="window.close()"></td>
</tr>
<tr>
	<td align="left" style="padding-bottom:3;">참조 : <input type="text" name="temp_referlist" id="temp_referlist" value="" size="50" readonly></td>
	<td align="right"></td>
</tr>
<tr>
	<td colspan="2">
		<font color="blue">※ 선택된 작업자를 삭제 하시려면 해당 작업자를 한번 더 클릭하시면 삭제가 됩니다.<br>&nbsp;&nbsp;&nbsp;&nbsp;작업자 입력란을 비워두지 마시고 채워둔 후 삭제 하세요.</font>
	</td>
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
					<input type="button" value="작업" class="button" onClick="workerselect('<%=arrList(0,intLoop)%>|<%=arrList(4,intLoop)%>','<%=arrList(3,intLoop)%>')">
					<input type="hidden" name="workername" value="<%=arrList(3,intLoop)%>">
					&nbsp;
					<input type="button" value="참조" class="button" onClick="referselect('<%=arrList(0,intLoop)%>|<%=arrList(4,intLoop)%>','<%=arrList(3,intLoop)%>')">
					<input type="hidden" name="refername" value="<%=arrList(3,intLoop)%>">
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
<!--
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td align="right" style="padding-top:3;"><input type="button" value="선택적용" class="button" onClick="workerselect()"></td>
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