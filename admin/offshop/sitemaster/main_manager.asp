<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###############################################
' PageName : main_manager.asp
' Discription : ����Ʈ ���� ����
' History : 2008.04.11 ������ : �Ǽ������� ����
'			2009.04.19 �ѿ�� 2009�� �°� ����
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/sitemaster/offshopmain_ContentsManageCls.asp" -->
<%
dim research,isusing, fixtype, linktype, poscode, validdate ,page ,i, oposcode, oMainContents
	isusing = requestCheckVar(request("isusing"),1)
	research= requestCheckVar(request("research"),2)
	poscode = requestCheckVar(request("poscode"),10)
	fixtype = requestCheckVar(request("fixtype"),10)
	page    = requestCheckVar(request("page"),10)
	validdate= requestCheckVar(request("validdate"),2)

	if ((research="") and (isusing="")) then 
	    isusing = "Y"
	    validdate = "on"
	end if
	
	if page="" then page=1

set oposcode = new CMainContentsCode
	oposcode.FRectPosCode = poscode
	
	if (poscode<>"") then
	    oposcode.GetOneContentsCode
	end if

set oMainContents = new CMainContents
	oMainContents.FPageSize = 10
	oMainContents.FCurrPage = page
	oMainContents.FRectIsusing = isusing
	oMainContents.FRectfixtype = fixtype
	oMainContents.FRectPosCode = poscode
	oMainContents.FRectvaliddate = validdate
	oMainContents.GetPoint1010ContentsList
%>

<script type="text/javascript">

function NextPage(page){
    frm.page.value = page;
    frm.submit();
}

function popPosCodeManage(){
    var popwin = window.open('/admin/offshop/sitemaster/popmainposcodeedit.asp','mainposcodeedit','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function AddNewMainContents(idx){
    var popwin = window.open('/admin/offshop/sitemaster/popmaincontentsedit.asp?idx=' + idx,'mainposcodeedit','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}


function AssignFlashReal(pc,lt){
    if (document.frm.poscode.value == ""){
		alert("������ġ�� �������ּ���");
		document.frm.poscode.focus();
	}
	else{
		 var popwin = window.open('','refreshFrm_Main','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm_Main";

		 if(lt=="F") {
			 refreshFrm.action = "<%=wwwUrl%>/offshop/flash/make_main_flash_Text.asp?poscode=" + document.frm.poscode.value;
		 }
			 refreshFrm.submit();
	}
}

</script>

<!-- ��� �˻��� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
	    <input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >��������
	    &nbsp;
	    ��뱸��
		<select name="isusing" class="select">
		<option value="">��ü
		<option value="Y" <% if isusing="Y" then response.write "selected" %> >�����
		<option value="N" <% if isusing="N" then response.write "selected" %> >������
		</select>
		&nbsp;&nbsp;
		���뱸��
		<% call DrawFixTypeCombo ("fixtype", fixtype, "") %>
		
		&nbsp;&nbsp;
		������ġ
		<% call DrawPoint1010PosCodeCombo("poscode",poscode, "") %>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�">
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<br>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
    <td></td>
    <td colspan="2">
	    <% if (poscode<>"") then %>
	         <% if (oposcode.FOneItem.Ffixtype="R") then %>   	    

	    	 <%	elseif oposcode.FOneItem.Flinktype="F" or oposcode.FOneItem.Flinktype="B" then %>
		        <a href="javascript:AssignFlashReal('<%= poscode %>','<%=oposcode.FOneItem.Flinktype%>');"><img src="/images/refreshcpage.gif" border="0"> Flash Real ����</a>
		          <% elseif (oposcode.FOneItem.Ffixtype <> "D") and (oposcode.FOneItem.Ffixtype <> "R") then %>
	    	    <a href="javascript:AssignTest('<%= poscode %>');"><img src="/images/icon_search.jpg" border="0"> �̸�����</a> 
	    	    &nbsp;&nbsp;
	    	    <a href="javascript:AssignReal('<%= poscode %>');"><img src="/images/refreshcpage.gif" border="0"> Real ����</a>
	    	<% end if %>
		<% end if %>
    </td>
    <td colspan="10" align="right">
    	<% if C_ADMIN_AUTH then %>
			<input type="button" class="button" value="�ڵ����" onClick="popPosCodeManage();">&nbsp;
		<% end if %>
    	<a href="javascript:AddNewMainContents('0');"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%=oMainContents.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oMainContents.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width=60>idx</td>
    <td>���и�</td>
    <td>�̹���</td>
    <td width=80>��ũ����</td>
    <td width=80>�ݿ��ֱ�</td>
    <td>������</td>
    <td>������</td>
    <td width=60>��뿩��</td>
    <td width=60>�켱����</td>
    <td width=90>�����</td>
    <td width=90>���</td>
</tr>
<%
if oMainContents.FResultCount > 0 then
	
for i=0 to oMainContents.FResultCount - 1
%>
<% if (oMainContents.FItemList(i).IsEndDateExpired) or (oMainContents.FItemList(i).FIsusing="N") then %>
<tr bgcolor="#DDDDDD" align="center">
<% else %>
<tr bgcolor="#FFFFFF" align="center">
<% end if %>
    <td><%= oMainContents.FItemList(i).Fidx %></td>
    <td><a href="?poscode=<%= oMainContents.FItemList(i).Fposcode %>"><%= oMainContents.FItemList(i).Fposname %></a></td>
    <td>
		<% if not(oMainContents.FItemList(i).Fimagewidth="" or isnull(oMainContents.FItemList(i).Fimagewidth)) then %>
			<%
			'�̹��� ����� ���� ǥ��(���� 300px)
			if oMainContents.FItemList(i).Fimagewidth>300 then
			%>
				<img src="<%= oMainContents.FItemList(i).getImageUrl %>" border="0" width=300>
			<% else %>
				<img src="<%= oMainContents.FItemList(i).getImageUrl %>" border="0">
			<% end if %>
		<% else %>
			<img src="<%= oMainContents.FItemList(i).getImageUrl %>" border="0" width=300>
		<% end if %>
    </td>
    <td><%= oMainContents.FItemList(i).getlinktypeName %></td>
    <td><%= oMainContents.FItemList(i).getfixtypeName %></td>
    <td><%= oMainContents.FItemList(i).FStartdate %></td>
    <td>
		<% if (oMainContents.FItemList(i).IsEndDateExpired) then %>
			<font color="#777777"><%= Left(oMainContents.FItemList(i).FEnddate,10) %></font>
		<% else %>
			<%= Left(oMainContents.FItemList(i).FEnddate,10) %>
		<% end if %>
    </td>
    <td><%= oMainContents.FItemList(i).FIsusing %></td>
    <td><%=oMainContents.FItemList(i).forderidx %></td>
    <td><%= oMainContents.FItemList(i).Freguserid %></td>
    <td>
		<input type="button" value="����" onclick="AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');" class="button">

		<% if (oMainContents.FItemList(i).Ffixtype="R") then %>   
		
		<% elseif Not(oMainContents.FItemList(i).IsEndDateExpired or oMainContents.FItemList(i).FIsusing="N" or oMainContents.FItemList(i).Flinktype="F" or oMainContents.FItemList(i).Flinktype="B" or oMainContents.FItemList(i).Ffixtype="R") then %>
			<br>
			<a href="javascript:AssignDailyTest('<%= oMainContents.FItemList(i).Fidx %>');"><img src="/images/icon_search.jpg" border="0"> �̸�����</a> 
			<br>
			<a href="javascript:AssignDailyReal('<%= oMainContents.FItemList(i).Fidx %>');"><img src="/images/refreshcpage.gif" border="0"> Real ����</a>
		<% end if %> 
    </td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="20" align="center">
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
<% else %>
<tr bgcolor="#FFFFFF">
    <td colspan="20" align="center">
		�˻� ����� �����ϴ�
    </td>
</tr>
<% end if %>
</table>

<%
set oposcode = Nothing
set oMainContents = Nothing
%>

<form name="refreshFrm" method="post">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->