	<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/categoryCdm_md_choicecls.asp"-->
<%
dim page, cdl,cdm, isusing
cdl = request("cdl")
cdm = request("cdm")
page = request("page")
isusing = request("isusing")

if cdl = "" then cdl ="110"
if page="" then page=1


dim omd
set omd = New CMDChoice
omd.FCurrPage = page
omd.FPageSize=20
omd.FRectCDL = cdl
omd.FRectCDM = cdm
omd.FRectIsUsing = isusing
omd.GetMDChoiceList

dim i
%>
<script language='javascript'>
<!--
function popItemWindow(tgf){
	if (document.Listfrm.cdl.value == ""){
		alert("��ī�װ��� ������ �ּ���!");
		document.Listfrm.cdl.focus();
		return;
	}
	
	if (document.Listfrm.cdm.value == ""){
		alert("��ī�װ��� ������ �ּ���!");
		document.Listfrm.cdm.focus();
		return;
	}
	
	var popup_item = window.open("/common/pop_CateItemList.asp?cdl=" + document.Listfrm.cdl.value + "&cdm="+ document.Listfrm.cdm.value+"&target=" + tgf+"&sd=best", "popup_item", "width=800,height=500,scrollbars=yes,status=no");
	popup_item.focus();
	
}

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];

		if (frm.name.indexOf('frmBuyPrc')!= -1) {

			pass = ((pass)||(frm.cksel.checked));
		}

	}

	if (!pass) {
		return false;
	}
	return true;
}

function delitems(upfrm){
	if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}

	if (document.Listfrm.cdl.value == ""){
		alert("��ī�װ��� ������ �ּ���!");
		document.Listfrm.cdl.focus();
		return;
	}
	
	if (document.Listfrm.cdm.value == ""){
		alert("��ī�װ��� ������ �ּ���!");
		document.Listfrm.cdm.focus();
		return;
	}
	var ret = confirm('���� �������� �����Ͻðڽ��ϱ�?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + frm.itemid.value + "," ;
				}
			}
		}
		
		upfrm.cdl.value = Listfrm.cdl.value;
		upfrm.cdm.value = Listfrm.cdm.value;
		upfrm.mode.value="del";
		upfrm.submit();

	}
}

function changeUsing(upfrm){
	if (!CheckSelected()){
		alert('��ǰ�� ������ �ּ���.');
		return;
	}

	if (document.Listfrm.cdl.value == ""){
		alert("��ī�װ��� ������ �ּ���!");
		document.Listfrm.cdl.focus();
		return;
	}
	
	if (document.Listfrm.cdm.value == ""){
		alert("��ī�װ��� ������ �ּ���!");
		document.Listfrm.cdm.focus();
		return;
	}
	
	if (upfrm.allusing.value=='Y'){
		var ret = confirm('�����Ͻ� ��ǰ�� ����� ���� �����մϴ�.');
	} else {
		var ret = confirm('�����Ͻ� ��ǰ�� ������ ���� �����մϴ�.');
	}


	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + frm.itemid.value + "," ;
				}
			}
		}
		upfrm.cdl.value = Listfrm.cdl.value;
		upfrm.cdm.value = Listfrm.cdm.value;
		upfrm.mode.value="isUsingValue";
		upfrm.submit();

	}
}

function AddIttems(){
	if (document.Listfrm.cdl.value == ""){
		alert("��ī�װ��� ������ �ּ���!");
		document.Listfrm.cdl.focus();
		return;
	}
	
	if (document.Listfrm.cdm.value == ""){
		alert("��ī�װ��� ������ �ּ���!");
		document.Listfrm.cdm.focus();
		return;
	}
	
	var ret = confirm(arrFrm.itemid.value + '�������� �߰��Ͻðڽ��ϱ�?');
	if (ret){
		arrFrm.itemid.value = arrFrm.itemid.value;
		arrFrm.cdl.value = Listfrm.cdl.value;
		arrFrm.cdm.value = Listfrm.cdm.value;
		arrFrm.mode.value="add";
		alert(arrFrm.cdl.value);
		arrFrm.submit();
	}
}

function AddIttems2(){
	if (document.Listfrm.cdl.value == ""){
		alert("��ī�װ��� ������ �ּ���!");
		document.Listfrm.cdl.focus();
		return;
	}
	
	if (document.Listfrm.cdm.value == ""){
		alert("��ī�װ��� ������ �ּ���!");
		document.Listfrm.cdm.focus();
		return;
	}
	
   if (document.arrFrm.itemidarr.value == ""){
		alert("�����۹�ȣ��  �����ּ���!");
		document.arrFrm.itemidarr.focus();
	}
	else if (confirm(arrFrm.itemidarr.value + '�������� �߰��Ͻðڽ��ϱ�?')){
		arrFrm.itemid.value = arrFrm.itemidarr.value;
		arrFrm.cdl.value = Listfrm.cdl.value;
		arrFrm.cdm.value = Listfrm.cdm.value;
		arrFrm.mode.value="add";
		arrFrm.submit();
	}
}

function RefreshCaMDChoiceRec(){
	if (document.Listfrm.cdl.value == ""){
		alert("��ī�װ��� ������ �ּ���!");
		document.Listfrm.cdl.focus();
		return;
	}
	
	if (document.Listfrm.cdm.value == ""){
		alert("��ī�װ��� ������ �ּ���!");
		document.Listfrm.cdm.focus();
		return;
	} 
	
	 var popwin = window.open('','refreshFrm','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm";
	 refreshFrm.cdl.value = document.Listfrm.cdl.value;
	 refreshFrm.cdm.value = document.Listfrm.cdm.value;
	 refreshFrm.action = "<%=wwwUrl%>/chtml/make_category_channel_text_mdchoice.asp";
	 refreshFrm.submit();
	 
}

function RefreshXMLMDChoiceRec(){
	if (document.Listfrm.cdl.value == ""){
		alert("ī�װ��� ������ �ּ���!");
		document.Listfrm.cdl.focus();
	}else{
		var popwin = window.open('','refreshFrm','');
		popwin.focus();
		refreshFrm.target = "refreshFrm";
		refreshFrm.cdl.value = document.Listfrm.cdl.value;
		refreshFrm.cdm.value = document.Listfrm.cdm.value;
		refreshFrm.action = "<%=wwwUrl%>/chtml/make_cate_XML_CDmdchoice.asp";
		refreshFrm.submit();
	 }
}

// ī�װ� ����� ���
function changecontent(){
	document.Listfrm.submit();
}

//-->
</script>
<form name="refreshFrm" method="post">
<input type="hidden" name="cdl">
<input type="hidden" name="cdm">
</form>
<!-- ��� �˻��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="Listfrm" method="get" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		ī�װ� :
		<% DrawSelectBoxCategoryLarge "cdl", cdl %>&nbsp;
		<% DrawSelectBoxCategoryMid "cdm",cdl,cdm %>&nbsp;/&nbsp;
		������� :
		<select name="isusing" onchange="this.form.submit();" class="select">
			<option value=""  <% if isusing="" then response.write "selected" %>>��ü</option>
			<option value="Y" <% if isusing="Y" then response.write "selected" %>>���</option>
			<option value="N" <% if isusing="N" then response.write "selected" %>>������</option>
		</select>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�">
	</td>
</tr>
</table>
<!-- �˻� �� -->
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td colspan="2">
				<img src="/images/icon_reload.gif" onClick="RefreshCaMDChoiceRec()" style="cursor:pointer" align="absmiddle" alt="html�����">
				����Ʈ�� ����
				<img src="/images/icon_reload.gif" onClick="RefreshXMLMDChoiceRec()" style="cursor:pointer" align="absmiddle" alt="html�����">
				2011 xml�� ����Ʈ�� ����
			</td>
		</tr>
		</form>
		<form name="arrFrm" method="post" action="doCdM_MDChoice.asp">
		<input type="hidden" name="cdl">
		<input type="hidden" name="cdm">
		<input type="hidden" name="mode">
		<input type="hidden" name="itemid">
		<tr>
			<td colspan="2" align="right">
				<input type="text" name="itemidarr" value="" size="80" class="input">
				<input type="button" value="������ ���� �߰�" onclick="AddIttems2()" class="button">
			</td>
		</tr>
		<tr>
			<td>
				<input type="button" value="���þ����� ����" onClick="delitems(arrFrm)" class="button">
				<select name="allusing"  class="select">
					<option value="Y">���� ��� -> Y </option>
					<option value="N">���� ��� -> N </option>
				</select><input type="button" value="����" class="button" onclick="changeUsing(arrFrm);">
			</td>
			<td align="right"><input type="button" value="������ �߰�" onclick="popItemWindow('arrFrm.itemid')" class="button"></td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FAFAFA" height="22">
	<td colspan="8">&nbsp;�˻��� ��ǰ�� : <%=omd.FTotalCount%> ��</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td align="center">��ī�װ���</td>
	<td align="center">��ī�װ���</td>	
	<td align="center">Image</td>
	<td align="center">ItemID</td>
	<td align="center">��ǰ��</td>
	<td align="center">�������</td>
	<td align="center">ǰ������</td>
</tr>
<% for i=0 to omd.FResultCount-1 %>
<form name="frmBuyPrc<%=i%>" method="post" action="" >
<input type="hidden" name="itemid" value="<%= omd.FItemList(i).FItemID %>">
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="cksel" cdl="<%= omd.FItemList(i).Fcdl %>" cdm="<%= omd.FItemList(i).Fcdm %>" onClick="AnCheckClick(this);"></td>
	<td align="center"><%= omd.FItemList(i).Fcode_nm %></td>
	<td align="center"><%= omd.FItemList(i).Fmcode_nm %></td>
	<td align="center"><img src="<%= omd.FItemList(i).Fimagesmall %>" width="50" height="50"></td>
	<td align="center"><%= omd.FItemList(i).FItemID %></td>
	<td align="center"><%= omd.FItemList(i).FItemname %></td>
	<td align="center"><%= omd.FItemList(i).Fisusing %></td>
	<td align="center"><% if omd.FItemList(i).IsSoldOut=true then Response.Write "<font color=red>ǰ��</font>" %></td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" align="center">
	<% if omd.HasPreScroll then %>
		<a href="?page=<%= omd.StarScrollPage-1 %>&cdl=<%=cdl%>&cdm=<%=cdm%>&menupos=<%= menupos %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + omd.StarScrollPage to omd.FScrollCount + omd.StarScrollPage - 1 %>
		<% if i>omd.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&cdl=<%=cdl%>&cdm=<%=cdm%>&menupos=<%= menupos %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if omd.HasNextScroll then %>
		<a href="?page=<%= i %>&cdl=<%=cdl%>&cdm=<%=cdm%>&menupos=<%= menupos %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
set omd = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->