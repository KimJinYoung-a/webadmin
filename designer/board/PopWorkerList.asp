<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/board/upche_qnacls.asp" -->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
	Dim iTotCnt, arrList,intLoop
	Dim cPartList, arrPartList, Memberlist, i, sWorker, vChecked, iDoc_Idx 
	Dim department_id
	iDoc_Idx= NullFillWith(requestCheckVar(Request("didx"),10),"0")
	sWorker = NullFillWith(Request("workerid"),"") 
	department_id = requestCheckVar(Request("department_id"),10)
	if (department_id = "") then department_id = "31"
	If sWorker <> "" Then
		sWorker = sWorker & ","
	End If
 
	set Memberlist = new CUpcheQnADetail
	Memberlist.FTeam = department_id
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
	
	workerselect();
}

function workerselect()
{
	var w_id = document.getElementsByName("workerid").length;
	var k_id = new Array();
	var k_name = new Array();
	var k_date = new Array();
	var m = 0;
	
	for(var i=0; i < w_id ; i++){
	    if (document.getElementsByName("workerid")[i].checked == true)
	    {
			k_id[m] = document.getElementsByName("workerid")[i].value;
	        k_name[m] = document.getElementsByName("workername")[i].value;
	        m = m+1;
	    }
	}
	opener.document.frm.workername.value = k_name;
	opener.document.frm.workerid.value = k_id;
	window.close();
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
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="2" height="25"><!--부서수정 2014-11-->
			<%= drawChSelectBoxDepartment("department_id", department_id,"onChange=""self.location='PopWorkerList.asp?worker="&sWorker&"&department_id='+this.value""") %>
		<!--	<select name="team" class="select" onChange="self.location='PopWorkerList.asp?worker=<%=sWorker%>&team='+this.value">
			<option value="">::부서선택::</option>
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
			</select>-->
			<!-- 사장님에 의하여 수정 됨!! 2012-03-16
			<a href="PopWorkerList.asp?worker=<%=sWorker%>&team=9,10">운영관리팀</a>
			|<a href="PopWorkerList.asp?worker=<%=sWorker%>&team=11,12,14,16">텐바이텐온라인사업팀</a>
			|<a href="PopWorkerList.asp?worker=<%=sWorker%>&team=13">오프라인팀</a>
			|<a href="PopWorkerList.asp?worker=<%=sWorker%>&team=7">시스템팀</a>
			|<a href="PopWorkerList.asp?worker=<%=sWorker%>&team=8">경영관리팀</a>
			|<a href="PopWorkerList.asp?worker=<%=sWorker%>&team=15">아이띵소팀</a>
			|<a href="PopWorkerList.asp?worker=<%=sWorker%>&team=17">패션사업팀</a>
			-->
		</td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left" style="padding-bottom:3;">※ 색깔이 바뀐 라인을 클릭하셔도 체크가 됩니다.</td>
		<td align="right" style="padding-bottom:3;"><input type="button" value="선택적용" class="button" onClick="workerselect()"></td>
	</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#EFEFEF" height="30">			
	<td align="center">팀/파트</td>
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
				<td align="left" style="cursor:pointer" onClick="checkworker('<%=intLoop%>')"><%=arrList(1,intLoop)%>
				<% if Not(arrList(5,intLoop)="" or isNull(arrList(5,intLoop))) then Response.Write " " & arrList(5,intLoop) %>
				</td>
				<td align="center" style="cursor:pointer" onClick="checkworker('<%=intLoop%>')"><%=arrList(3,intLoop)%></td>
				<td align="center">
					<input type="radio" name="workerid" value="<%=arrList(0,intLoop)%>" <%=vChecked%> onClick="workerselect()">
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
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td align="right" style="padding-top:3;"><input type="button" value="선택적용" class="button" onClick="workerselect()"></td>
</tr>
</table>