<%@ language=vbscript %>
<% option explicit %>
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
<!-- #include virtual="/lib/classes/sitemasterclass/brandStreet_ManageCls.asp" -->
<%
'###############################################
' PageName : BrandStreet_main.asp
' Discription : �귣�� ��Ʈ��Ʈ ���� ����
' History : 2008.04.11 ������ : ����Ʈ���� ���� ����/����
'###############################################

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
set oposcode = new CStreetMainCode
oposcode.FRectPosCode = poscode

if (poscode<>"") then
    oposcode.GetOneContentsCode
end if

dim oMainContents
set oMainContents = new CStreetMainCont
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
    var popwin = window.open('/admin/sitemaster/lib/popStreetPosCodeEdit.asp','StreetPosCodeEdit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function AddNewMainContents(idx){
    var popwin = window.open('/admin/sitemaster/lib/popStreetContentsEdit.asp?idx=' + idx,'StreetContentEdit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
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
		 refreshFrm.action = "<%=uploadUrl%>/chtml/make_StreetMain_Image_JS.asp?poscode=" + document.frm.poscode.value;
		 refreshFrm.submit();
	}
}

function AssignDailyFlashReal(idx){
	 var popwin = window.open('','refreshFrm_Main','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main";
	 refreshFrm.action = "<%=uploadUrl%>/chtml/make_StreetMain_Flash_byIdx_JS.asp?idx=" + idx;
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
		 refreshFrm.action = "<%=uploadUrl%>/chtml/make_StreetMain_Flash_JS.asp?poscode=" + document.frm.poscode.value;
		 refreshFrm.submit();
	}
}
</script>
<!-- ��� �˻��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
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
		<% call DrawMainPosCodeCombo("poscode",poscode, "") %>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
    <td><a href="http://www.10x10.co.kr/index_preview.asp?yyyymmdd=<%= Left(CStr(now()),10) %>" target="refreshFrm_Main">�������</a></td>
    <td colspan="2">
	    <%
	    	if (poscode<>"") then
	    		if Not(oposcode.FOneItem.Flinktype="F") then
	    %>
	    <a href="javascript:AssignReal('<%= poscode %>');"><img src="/images/refreshcpage.gif" border="0"> Real ����</a>
	    <% elseif (oposcode.FOneItem.Ffixtype <> "D") and (oposcode.FOneItem.Ffixtype <> "R") then %>
	    <a href="javascript:AssignFlashReal('<%= poscode %>');"><img src="/images/refreshcpage.gif" border="0"> Flash Real ����</a>
	    <%
	    		end if
	    	end if
	    %>
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
	<td colspan="10">
		�˻���� : <b><%=oMainContents.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oMainContents.FtotalPage%></b>
	</td>
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
<tr bgcolor="#DDDDDD">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
    <td align="center"><%= oMainContents.FItemList(i).Fidx %></td>
    <td align="center"><a href="?poscode=<%= oMainContents.FItemList(i).Fposcode %>"><%= oMainContents.FItemList(i).Fposname %></a></td>
    <td align="center">
    <%
    	if Not(oMainContents.FItemList(i).getImageUrl="" or isNull(oMainContents.FItemList(i).getImageUrl)) then
    		if right(oMainContents.FItemList(i).getImageUrl,3)="swf" then
    %>
    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><font color="darkblue">[�÷�������]</font></embed></a>
    <%		else %>
    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img height="66" src="<%= oMainContents.FItemList(i).getImageUrl %>" border="0"></a>
    <%
    		end if
    	end if
    %>
    </td>
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
    <% if Not(oMainContents.FItemList(i).IsEndDateExpired or oMainContents.FItemList(i).FIsusing="N" or oMainContents.FItemList(i).Flinktype<>"F") then %>
    <a href="javascript:AssignDailyFlashReal('<%= oMainContents.FItemList(i).Fidx %>');"><img src="/images/refreshcpage.gif" border="0"> Real ����</a>
    <% else %>
    &nbsp;
    <% end if %> 
    </td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="10" align="center">
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