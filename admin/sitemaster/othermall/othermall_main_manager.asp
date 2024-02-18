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
oMainContents.FPageSize = 20
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

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">		
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
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    <input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >��������				
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<% if C_ADMIN_AUTH then %>
			<input type="button" value="�ڵ����" class="button" onClick="popPosCodeManage();">
			<% end if %>
		</td>
		<td align="right">
			<a href="javascript:AddNewMainContents('0');"><img src="/images/icon_new_registration.gif" border="0"></a>
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if oMainContents.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= oMainContents.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= oMainContents.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
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
	    <td colspan="10" align="right"></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td>idx</td>
	    <td>���и�</td>
	    <td>�̹���</td>
	    <td>��ũ<br>����</td>
	    <td>�ݿ�<br>�ֱ�</td>
	    <td>������</td>
	    <td>������</td>
	    <td>��뿩��</td>
	    <td>�����</td>
	    <td></td>
    </tr>
	<% for i=0 to oMainContents.FResultCount - 1 %>
	<% if (oMainContents.FItemList(i).IsEndDateExpired) or (oMainContents.FItemList(i).FIsusing="N") then %>
	<tr align="center" bgcolor="#DDDDDD">
	<% else %>
    <tr align="center" bgcolor="#FFFFFF">
	<% end if %>
	    <td align="center"><%= oMainContents.FItemList(i).Fidx %></td>
	    <td align="center"><a href="?poscode=<%= oMainContents.FItemList(i).Fposcode %>"><%= oMainContents.FItemList(i).Fposname %></a></td>
	    <td ><a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img width=60 height=60 src="<%= oMainContents.FItemList(i).getImageUrl %>" border="0"></a></td>
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
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
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
	</tr>
</table>
<%
set oposcode = Nothing
set oMainContents = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->