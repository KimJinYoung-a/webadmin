<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<%
	Dim iTotCnt, arrList,intLoop, vParam
	Dim iPageSize, iCurrentpage ,iDelCnt, sSearchTeam, sDoc_Status, sDoc_Status1, sDoc_AnsOX, sSearchMine, sUserName, sSearching, sContent
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt, vTempArr11, vOnlyNewList
	Dim sDoc_Type

	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	sSearchTeam		= NullFillWith(requestCheckVar(Request("search_team"),20),"")
	sDoc_Status		= NullFillWith(requestCheckVar(Request("doc_status"),10),"0")
	sDoc_Status1	= NullFillWith(requestCheckVar(Request("doc_status1"),10),"x")
	sDoc_Type		= NullFillWith(requestCheckVar(Request("doc_type"),10),"")
	sDoc_AnsOX		= NullFillWith(requestCheckVar(Request("ans_ox"),1),"")
	sSearchMine		= NullFillWith(requestCheckVar(Request("onlymine"),1),"o")
	sUserName		= NullFillWith(requestCheckVar(Request("username"),10),"")
	sSearching		= NullFillWith(requestCheckVar(Request("searching"),10),"")
	sContent		= NullFillWith(requestCheckVar(Request("content"),100),"")
	vOnlyNewList	= NullFillWith(requestCheckVar(Request("onlynewlist"),1),"")
	iPageSize 		= CHKIIF(g_VertiHoriz="h",7,15)
	iPerCnt 		= 10
	
	If sSearching = "doc_idx" AND IsNumeric(sContent) = False Then
		Response.Write "<script language='javascript'>alert('����No �̻��� ���ڷθ� �Է��ϼž� �մϴ�.');history.back();</script>"
	End IF
	
	vParam = "&iC="&iCurrentpage&"&s_search_team="&sSearchTeam&"&s_status="&sDoc_Status&"&s_type="&sDoc_Type&"&s_ans_ox="&sDoc_AnsOX&"&s_onlymine="&sSearchMine&"&onlynewlist="&vOnlyNewList&"&username="&sUserName&"searching="&sSearching&"content="&sContent&""
	'<!-- �д� �������� ���� �Ķ���͸����� �Ǿ� �ִ°� �־ Ȥ�ó� �� �Ͽ� ����Ͽ� �Ķ���͸��� �ٲ㼭 �ְ� �޾ҽ�. //-->

	Dim cooperatelist , i
	
		set cooperatelist = new CCooperate
	 	cooperatelist.FCPage = iCurrentpage
	 	cooperatelist.FPSize = iPageSize
	 	cooperatelist.FTeam = sSearchTeam
	 	
	 	cooperatelist.FDoc_IsRefer = "o"
	 	If sDoc_Status = "6" Then
	 		cooperatelist.FDoc_Status = "x"
	 	Else
	 		cooperatelist.FDoc_Status = sDoc_Status
		End If
		cooperatelist.FDoc_Status1 = sDoc_Status1
	 	
	 	cooperatelist.FOnlyNewList = vOnlyNewList
	 	cooperatelist.FDoc_Type = sDoc_Type
	 	cooperatelist.FDoc_AnsOX = sDoc_AnsOX
	 	cooperatelist.FDoc_MineOX = sSearchMine
	 	cooperatelist.FDoc_UserName = sUserName
	 	cooperatelist.FDoc_Searching = sSearching
	 	cooperatelist.FDoc_Content = sContent
		arrList = cooperatelist.fnGetCooperateList
		iTotCnt = cooperatelist.FTotCnt
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">

<script language="javascript">
function code_manage()
{
	window.open('PopManageCode.asp','coopcode','width=410,height=570');
}
function goWrite(didx)
{
	top.coopDetail.location.href = "cooperate_read.asp?<%=CHKIIF(sDoc_Status="6","ischamjo=o&","")%>didx="+didx+"";
}
function jsGoPage(iP){
	document.frmpage.iC.value = iP;
	document.frmpage.submit();
}
function mine()
{
	if(!(document.frm.onlyminechk.checked))
	{
		document.frm.onlymine.value = "x";
	}
	else
	{
		document.frm.onlymine.value = "o";
	}
}
function issystem(value)
{
}

function popCooperate(){
	 var winCooperate = window.open("popIndex.asp","popCooperate","width="+screen.availWidth+", height="+ screen.availHeight +",resizable=yes, scrollbars=yes"); 
	 winCooperate.focus();
}

function jsonlynewlist()
{
	frm.onlynewlist.value = "o";
	frm.submit();
}

function goPopDetail(didx)
{
	 window.open("cooperate_read.asp?<%=CHKIIF(sDoc_Status="6","ischamjo=o&","")%>ispop=pop&didx="+didx+"","","width=900, height=1000,resizable=yes, scrollbars=yes");
}

//���ڰ��� ǰ�Ǽ� ��� - ��������������ȣ(scmidx) 
function jsRegEapp(scmidx){ 
	var winEapp = window.open("/admin/approval/eapp/regeapp.asp","popE","width=1000,height=600,scrollbars=yes");
	document.frmEapp.iSL.value = scmidx;   
	document.frmEapp.target = "popE";
	document.frmEapp.submit();
	winEapp.focus();
}
//���ڰ��� ǰ�Ǽ� ���뺸��
function jsViewEapp(reportidx,reportstate){
	var winEapp = window.open("/admin/approval/eapp/modeapp.asp?iRM=M01"+reportstate+"&iridx="+reportidx,"popE","");
	winEapp.focus();
}
</script>
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
<div style="height:100%;overflow-y:auto;">
<table width="100%" cellpadding="0" cellspacing="1" class="a" border="0"> 
<tr> 
	<td height="25"><font color="#4E9FC6"><b>������������ > <%=NaviTitle(sDoc_Status)%></font></b></font></td>
</tr>
</table>
<form name="frmEapp" method="post" action="/admin/approval/eapp/regeapp.asp">
<input type="hidden" name="tC" value="">
<input type="hidden" name="ieidx" value="38">  
<input type="hidden" name="iSL" value="">
</form>
<form name="frm" action="index.asp" method="get" style="margin:0px;">
<input type="hidden" name="doc_status" value="<%=sDoc_Status%>">
<input type="hidden" name="onlynewlist" value="<%=vOnlyNewList%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<table class="a">
		<tr>
			<td>
		     	<%=CommonCode("w","doc_type",sDoc_Type)%>&nbsp;
				������� : <input type="text" name="username" value="<%=sUserName%>" size="10">
				<% If sDoc_Status = "6" Then %>
				&nbsp;
				<select name='doc_status1' class='select'>
					<option value='x'>-ó������-</option>
					<option value='0' <%=CHKIIF(sDoc_Status1="0","selected","")%>>��ó�� ��ü</option>
					<option value='1' <%=CHKIIF(sDoc_Status1="1","selected","")%>>���</option>
					<option value='2' <%=CHKIIF(sDoc_Status1="2","selected","")%>>�۾���</option>
					<option value='3' <%=CHKIIF(sDoc_Status1="3","selected","")%>>�۾��Ϸ�</option>
					<option value='4' <%=CHKIIF(sDoc_Status1="4","selected","")%>>�ݷ�</option>
					<option value='5' <%=CHKIIF(sDoc_Status1="5","selected","")%>>�ݷ� �� �����Ϸ�</option>
				</select>
				<% End If %>
			</td>
			<td rowspan="2" style="padding:0 0 0 <%=CHKIIF(sDoc_Status="6","30","80")%>px;" align="right" valign="top"><input type="submit" value=" ��  �� " class="button" style="width:70px;height:50px;" onfocus="this.blur();"></td>
		</tr>
		<tr>
			<td>
				<select name="searching" class="select">
					<option value="">-����-</option>
					<option value="doc_idx" <%=CHKIIF(sSearching="doc_idx","selected","")%>>����No</option>
					<option value="title" <%=CHKIIF(sSearching="title","selected","")%>>����</option>
					<option value="content" <%=CHKIIF(sSearching="content","selected","")%>>����</option>
				</select>
				<input type="text" name="content" value="<%=sContent%>" size="41">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="CHKIIF(g_VertiHoriz="h","left","right")">
		<input type="hidden" name="onlymine" value="<%=sSearchMine%>">
		<!--<label id="onlymine"><input type="checkbox" name="onlyminechk" onClick="mine()" value="o" <% If sSearchMine = "o" Then %>checked<% End If %>>���� �۾��� ����</label>//-->
		<input type="button" value="��Ȯ�αۺ���" class="button" onClick="jsonlynewlist()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� ����, ���� �˻��� ����� ���� �� �ֽ��ϴ�.
	</td>
</tr>
</table>
</form>

<p><% If CInt(session("ssAdminLsn")) = 1 AND CInt(session("ssAdminPsn")) = 7 Then %><input type="button" class="button" value="�ڵ����" onClick="code_manage()"><% End If %><p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#FFFFFF">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			�˻���� : <b><%= iTotCnt %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td width="40" align="center">NO</td>
		<td width="60" align="center">�������</td>
		<td>����</td>
		<td width="80" align="center">�߿䵵</td>
		<td width="70" align="center">�����</td>
		<td width="40" align="center">�亯</td>
		<td  align="center">���系��</td> 
	</tr>
<%
	IF isArray(arrList) THEN
		For intLoop =0 To UBound(arrList,2)
		
			If sDoc_Status = "6" Then
				vTempArr11 = arrList(12,intLoop)
			Else
				vTempArr11 = arrList(11,intLoop)
			End If
%>
		<tr align="center" bgcolor="<%=CHKIIF(isNull(vTempArr11),"#D4FFFF","#FFFFFF")%>" height="30" onmouseout="this.style.backgroundColor='<%=CHKIIF(isNull(vTempArr11),"#D4FFFF","#FFFFFF")%>'" onmouseover="this.style.backgroundColor='#F1F1F1'">
			<td style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=arrList(0,intLoop)%></td>
			<td style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=arrList(7,intLoop)%></td>
			<td align="left">
				<font color="silver"><%=CommonCode("v","doc_type",arrList(2,intLoop))%></font>
				<br><span style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=db2html(arrList(1,intLoop))%></span>
				&nbsp;&nbsp;&nbsp;<img src="http://fiximage.10x10.co.kr/web2012/category/product_btn_view.png" border="0" style="cursor:pointer" onClick="goPopDetail('<%=arrList(0,intLoop)%>');">
			</td>
			<td style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=CommonCode("v","doc_important",arrList(3,intLoop))%></td>
			<td style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=FormatDatetime(arrList(6,intLoop),2)%></td>
			<td style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=arrList(8,intLoop)%></td>
			<td nowrap>  <!--'�ý��۰��� �� �����϶��� ���縮��Ʈ �����ش� 2014.03.06 ������ �߰�-->
			<%IF (arrList(2,intLoop)="3" )  THEN %>
				<div>
				<% if  isNull(arrList(12,intLoop)) then %>
			  <font color="Gray">ǰ�Ǽ� ���ۼ�</font>
				<% else %>
				<%=fnGetReportState(arrList(13,intLoop))%>&nbsp; 
				<input type="button" class="button"  value="ǰ�Ǽ� ����" onClick="jsViewEapp('<%=arrList(12,intLoop)%>','<%= arrList(13,intLoop)%>');">
				<% end if%> 
			</div>
			<%IF arrList(13,intLoop) = 7 THEN%>
				<div style="padding:3px">
				<% if isNull(arrList(14,intLoop)) then %>
				<input type="button" class="button"  value="�� ���߰�ȹ�� ǰ��" onClick="jsRegEapp('<%=arrList(0,intLoop)%>');" >
				<% else %>
				<%=fnGetReportState(arrList(15,intLoop))%><br>
				<input type="button" class="button"  value="���߰�ȹ�� ����" onClick="jsViewEapp('<%=arrList(14,intLoop)%>','<%= arrList(15,intLoop)%>');">
				<% end if%> 
			 </div>
			 <%END IF%>
		<%END IF%>
			</td>
		</tr>
<%
		Next
	Else
%>
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="20" align="center" class="page_link">[�����Ͱ� �����ϴ�.]</td>
	</tr>
<%
	End If
%>


	<!-- ����¡ó�� -->
	
	
	<%
	iStartPage = (Int((iCurrentpage-1)/iPerCnt)*iPerCnt) + 1
	
	If (iCurrentpage mod iPerCnt) = 0 Then
		iEndPage = iCurrentpage
	Else
		iEndPage = iStartPage + (iPerCnt-1)
	End If
	%>
	
	<form name="frmpage" method="get" action="index.asp">
	<input type="hidden" name="iC" value="">
	<input type="hidden" name="search_team" value="<%=sSearchTeam%>">
	<input type="hidden" name="doc_status" value="<%=sDoc_Status%>">
	<input type="hidden" name="doc_status1" value="<%=sDoc_Status1%>">
	<input type="hidden" name="doc_type" value="<%=sDoc_Type%>">
	<input type="hidden" name="ans_ox" value="<%=sDoc_AnsOX%>">
	<input type="hidden" name="onlymine" value="<%=sSearchMine%>">
	<input type="hidden" name="username" value="<%=sUserName%>">
	<input type="hidden" name="searching" value="<%=sSearching%>">
	<input type="hidden" name="content" value="<%=sContent%>">
	<input type="hidden" name="onlynewlist" value="<%=vOnlyNewList%>">
	<tr height="50" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
			<% else %>[pre]<% end if %>
	        <%
				for ix = iStartPage  to iEndPage
					if (ix > iTotalPage) then Exit for
					if Cint(ix) = Cint(iCurrentpage) then
			%>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="red">[<%=ix%>]</font></a>
			<%		else %>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
			<%
					end if
				next
			%>
	    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
			<% else %>[next]<% end if %>
		</td>
	</tr>
    </form>
</table>

<%
	set cooperatelist = nothing
%>
<script type="text/javascript">
document.bgColor = "white";
</script>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->