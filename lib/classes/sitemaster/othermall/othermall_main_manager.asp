<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2007.11.09 �ѿ�� ����
'###########################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/othermall/othermall_main_contents_managecls.asp" -->

<%
dim research,isusing, fixtype, linktype, poscode, validdate
dim page

isusing = request("isusing")
research= request("research")
poscode = request("poscode")
fixtype = request("fixtype")
page    = request("page")
validdate= request("validdate")

if ((research="") and (isusing="")) then 
    isusing = "Y"
    validdate = "on"
end if

if page="" then page=1

dim oposcode
set oposcode = new CMainContentsCode
oposcode.FRectPosCode = poscode

if (poscode<>"") then
    oposcode.GetOneContentsCode
end if

dim oMainContents
set oMainContents = new CMainContents
oMainContents.FPageSize = 10
oMainContents.FCurrPage = page
oMainContents.FRectIsusing = isusing
oMainContents.FRectfixtype = fixtype
oMainContents.FRectPosCode = poscode
oMainContents.FRectvaliddate = validdate
oMainContents.GetMainContentsList

dim i
%>
<script language='javascript'>
function NextPage(page){
    frm.page.value = page;
    frm.submit();
}

function popPosCodeManage(){
    var popwin = window.open('/admin/sitemaster/othermall/othermall_popmainposcodeedit.asp','mainposcodeedit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function AddNewMainContents(idx){
    var popwin = window.open('/admin/sitemaster/othermall/othermall_popmaincontentsedit.asp?idx=' + idx,'mainposcodeedit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function AssignTest(){
    if (document.frm.poscode.value == ""){
		alert("������ġ�� �������ּ���");
		document.frm.poscode.focus();
	}
	else{
		 var popwin = window.open('','refreshFrm_Main_Test','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm_Main_Test";
		 refreshFrm.action = "<%=uploadUrl%>/chtml/othermall_contents_Test_JS.asp?poscode=" + document.frm.poscode.value;
		 refreshFrm.submit();
	}
}

function AssignReal(){
    if (document.frm.poscode.value == ""){
		alert("������ġ�� �������ּ���");
		document.frm.poscode.focus();
	}
	else{
		 var popwin = window.open('','refreshFrm_Main','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm_Main";
		 refreshFrm.action = "<%=uploadUrl%>/chtml/othermall_make_main_contents_JS.asp?poscode=" + document.frm.poscode.value;
		 refreshFrm.submit();
	}
}


function AssignDailyTest(idx){
	 var popwin = window.open('','refreshFrm_Main_Test','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main_Test";
	 refreshFrm.action = "<%=othermall%>/chtml/othermall_make_main_contents_byidx_Test_JS.asp?idx=" + idx;
	 refreshFrm.submit();
}

function AssignDailyReal(idx){
	 var popwin = window.open('','refreshFrm_Main','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main";
	 refreshFrm.action = "<%=othermall%>/chtml/othermall_make_main_contents_byidx_JS.asp?idx=" + idx;
	 refreshFrm.submit();
}


function AssignFlashReal(){
    if (document.frm.poscode.value == ""){
		alert("������ġ�� �������ּ���");
		document.frm.poscode.focus();
	}
	else{
		 var popwin = window.open('','refreshFrm_Main','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm_Main";
		 refreshFrm.action = "<%=othermall%>/chtml/othermall_make_main_flash_Text.asp?poscode=" + document.frm.poscode.value;
		 refreshFrm.submit();
	}
}
</script>

<!--ǥ ������-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>�ܺθ� ���������� ����</strong></font>
			</td>		
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
		    <input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >��������
		    ��뱸��
			<select name="isusing" >
			<option value="">��ü
			<option value="Y" <% if isusing="Y" then response.write "selected" %> >�����
			<option value="N" <% if isusing="N" then response.write "selected" %> >������
			</select>
			���뱸��
			<% call DrawFixTypeCombo ("fixtype", fixtype, "") %>
			������ġ
			<% call DrawMainPosCodeCombo("poscode",poscode, "") %>
			<% if C_ADMIN_AUTH then %>
			<input type="button" value="�ڵ����" onClick="popPosCodeManage();">
			<% end if %>			
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>		
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10" valign="top">
		<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_06.gif"></td>
		<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
	</tr>
	</form>
</table>
<!--ǥ ��峡-->

<table width="100%" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
    <td><a href="http://gseshop.10x10.co.kr/index_preview.asp?yyyymmdd=<%= Left(CStr(now()),10) %>" target="refreshFrm_Main">�������</a></td>
    <td colspan="2">
    <% if (poscode<>"") then %>
    <% if oposcode.FOneItem.Flinktype="F" then %>
    <a href="javascript:AssignFlashReal('<%= poscode %>');"><img src="/images/refreshcpage.gif" border="0"> Flash Real ����</a>
    <% elseif (oposcode.FOneItem.Ffixtype <> "D") and (oposcode.FOneItem.Ffixtype <> "R") then %>
    <a href="javascript:AssignTest('<%= poscode %>');"><img src="/images/icon_search.jpg" border="0"> �̸�����</a> 
    &nbsp;&nbsp;
    <a href="javascript:AssignReal('<%= poscode %>');"><img src="/images/refreshcpage.gif" border="0"> Real ����</a>
    <% end if %>
    <% end if %>
    </td>
    <td colspan="10" align="right"><a href="javascript:AddNewMainContents('0');"><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
    <td width="60">idx</td>
    <td width="100">���и�</td>
    <td width="150">�̹���</td>
    <td width="50">��ũ<br>����</td>
    <td width="80">�ݿ�<br>�ֱ�</td>
    <td width="76">������</td>
    <td width="76">������</td>
    <td width="40">��뿩��</td>
    <td width="80">�����</td>
    <td></td>
</tr>
<% for i=0 to oMainContents.FResultCount - 1 %>
<% if (oMainContents.FItemList(i).IsEndDateExpired) or (oMainContents.FItemList(i).FIsusing="N") then %>
<tr bgcolor="#DDDDDD">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
    <td align="center"><%= oMainContents.FItemList(i).Fidx %></td>
    <td align="center"><a href="?poscode=<%= oMainContents.FItemList(i).Fposcode %>"><%= oMainContents.FItemList(i).Fposname %></a></td>
    <td ><a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img height="66" src="<%= oMainContents.FItemList(i).getImageUrl %>" border="0"></a></td>
    <td align="center"><%= oMainContents.FItemList(i).getlinktypeName %></td>
    <td align="center"><%= oMainContents.FItemList(i).getfixtypeName %></td>
    <td align="center"><%= oMainContents.FItemList(i).FStartdate %></td>
    <td align="center">
    <% if (oMainContents.FItemList(i).IsEndDateExpired) then %>
    <font color="#777777"><%= Left(oMainContents.FItemList(i).FEnddate,10) %></font>
    <% else %>
    <%= Left(oMainContents.FItemList(i).FEnddate,10) %>
    <% end if %>
    </td>
    <td align="center"><%= oMainContents.FItemList(i).FIsusing %></td>
    <td align="center"><%= oMainContents.FItemList(i).Freguserid %></td>
    <td>
    <% if Not(oMainContents.FItemList(i).IsEndDateExpired or oMainContents.FItemList(i).FIsusing="N" or oMainContents.FItemList(i).Flinktype="F") then %>
    <a href="javascript:AssignDailyTest('<%= oMainContents.FItemList(i).Fidx %>');"><img src="/images/icon_search.jpg" border="0"> �̸�����</a> 
    &nbsp;
    <a href="javascript:AssignDailyReal('<%= oMainContents.FItemList(i).Fidx %>');"><img src="/images/refreshcpage.gif" border="0"> Real ����</a>
    <% else %>
    &nbsp;
    <% end if %> 
    </td>
</tr>
<% next %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
		    <% if oMainContents.HasPreScroll then %>
				<a href="javascript:NextPage('<%= oMainContents.StarScrollPage-1 %>');">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
			<% for i=0 + oMainContents.StarScrollPage to oMainContents.FScrollCount + oMainContents.StarScrollPage - 1 %>
				<% if i>oMainContents.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
				<% end if %>
			<% next %>
			<% if oMainContents.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>');">[next]</a>
			<% else %>
				[next]
			<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<%
set oposcode = Nothing
set oMainContents = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->