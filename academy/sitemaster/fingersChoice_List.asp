<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/sitemaster/fingersChoiceCls.asp"-->
<%
dim page, MenuId, isusing
MenuId = RequestCheckvar(request("MenuId"),6)
page = RequestCheckvar(request("page"),10)
isusing = RequestCheckvar(request("isusing"),1)

if page="" then page=1


dim oFingers
set oFingers = New CFingersChoice
oFingers.FCurrPage = page
oFingers.FPageSize=21
oFingers.FRectMenuId = MenuId
oFingers.FRectIsUsing = isusing
oFingers.GetFingersChoiceList

dim i
%>
<script language='javascript'>
<!--
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
		alert('���ð��°� �����ϴ�.');
		return;
	}

	var ret = confirm('���� ���¸� �����Ͻðڽ��ϱ�?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.lec_idx.value = upfrm.lec_idx.value + frm.lec_idx.value + "," ;
				}
			}
		}
		upfrm.mode.value="del";
		upfrm.submit();

	}
}

function changeUsing(upfrm){
	if (!CheckSelected()){
		alert('���¸� ������ �ּ���.');
		return;
	}

	if (upfrm.allusing.value=='Y'){
		var ret = confirm('�����Ͻ� ���¸� ����� ���� �����մϴ�.');
	} else {
		var ret = confirm('�����Ͻ� ���¸� ������ ���� �����մϴ�.');
	}


	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.lec_idx.value = upfrm.lec_idx.value + frm.lec_idx.value + "," ;
				}
			}
		}
		upfrm.MenuId.value = Listfrm.MenuId.value;
		upfrm.mode.value="isUsingValue";
		upfrm.submit();

	}
}

// ��������
function changeSort(upfrm) {
	if (!CheckSelected()){
		alert('���¸� ������ �ּ���.');
		return;
	}
	var ret = confirm('�����Ͻ� ������ ������ �����Ͻ� ��ȣ�� �����Ͻðڽ��ϱ�?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.lec_idx.value = upfrm.lec_idx.value + frm.lec_idx.value + "," ;
					upfrm.sortNo.value = upfrm.sortNo.value + frm.sortNo.value + "," ;
				}
			}
		}
		upfrm.MenuId.value = Listfrm.MenuId.value;
		upfrm.mode.value="ChangeSort";
		upfrm.submit();

	}
}

function AddIttems(){
	var ret = confirm(arrFrm.lec_idx.value + '���¸� �߰��Ͻðڽ��ϱ�?');
	if (ret){
		arrFrm.lec_idx.value = arrFrm.lec_idx.value;
		arrFrm.MenuId.value = Listfrm.MenuId.value;
		arrFrm.mode.value="add";
		arrFrm.submit();
	}
}

function AddIttems2(){
	if (document.Listfrm.MenuId.value == ""){
		alert("�Է��� �ָ޴��� ������ �ּ���!");
		document.Listfrm.MenuId.focus();
	}
	else if (document.arrFrm.lecIdxarr.value == ""){
		alert("���¹�ȣ��  �����ּ���!");
		document.arrFrm.lecIdxarr.focus();
	}
	else if (confirm(arrFrm.lecIdxarr.value + '���¸� �߰��Ͻðڽ��ϱ�?')){
		arrFrm.lec_idx.value = arrFrm.lecIdxarr.value;
		arrFrm.MenuId.value = Listfrm.MenuId.value;
		arrFrm.mode.value="add";
		arrFrm.submit();
	}
}

function RefreshCaFingersChoiceRec(){
	if (document.Listfrm.MenuId.value == ""){
		alert("�ָ޴��� ������ �ּ���!");
		document.Listfrm.MenuId.focus();
	}
	 else{
			  var popwin = window.open('','refreshFrm','');
			  popwin.focus();
			  refreshFrm.target = "refreshFrm";
			  refreshFrm.MenuId.value = document.Listfrm.MenuId.value;
			  refreshFrm.action = "<%=wwwFingers%>/chtml/make_FingersChoice_JS.asp";
			  refreshFrm.submit();
	 }
}

// �ָ޴� ����� ���
function changecontent(){}

//-->
</script>
<form name="refreshFrm" method="post">
<input type="hidden" name="MenuId">
</form>
<!-- ��� �˻��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="Listfrm" method="get" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		�ָ޴� :
		<Select name="MenuId" Class="select">
			<option value="">����</option>
			<option value="10" <% if MenuId="10" then Response.Write "selected"%>>������ü</option>
			<option value="20" <% if MenuId="20" then Response.Write "selected"%>>������ Ŭ����</option>
			<option value="30" <% if MenuId="30" then Response.Write "selected"%>>��Ŭ�� Ŭ����</option>
			<option value="40" <% if MenuId="40" then Response.Write "selected"%>>��Ʃ�����ũ��</option>			
		</select>
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
				<img src="/images/icon_reload.gif" onClick="RefreshCaFingersChoiceRec()" style="cursor:pointer" align="absmiddle" alt="html�����">
				����Ʈ�� ����
			</td>
		</tr>
		</form>
		<form name="arrFrm" method="post" action="doFingersChoice.asp">
		<input type="hidden" name="MenuId">
		<input type="hidden" name="mode">
		<input type="hidden" name="lec_idx">
		<input type="hidden" name="sortNo">
		<tr>
			<td>
				<input type="button" value="���ð��� ����" onClick="delitems(arrFrm)" class="button"> /
				<select name="allusing"  class="select">
					<option value="Y">���� ��� -> Y </option>
					<option value="N">���� ��� -> N </option>
				</select><input type="button" value="����" class="button" onclick="changeUsing(arrFrm);"> /
				<input type="button" value="��������" class="button" onclick="changeSort(arrFrm);">
			</td>
			<td align="right">
				<input type="text" name="lecIdxarr" value="" size="50" class="input">
				<input type="button" value="���� �߰�" onclick="AddIttems2()" class="button">
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FAFAFA" height="22">
	<td colspan="8">&nbsp;�˻��� ���¼� : <%=oFingers.FTotalCount%> ��</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td align="center">�ָ޴���</td>
	<td align="center">Image</td>
	<td align="center">���¹�ȣ</td>
	<td align="center">���¸�</td>
	<td align="center">����</td>
	<td align="center">�������</td>
	<td align="center">��������</td>
</tr>
<% for i=0 to oFingers.FResultCount-1 %>
<form name="frmBuyPrc<%=i%>" method="post" action="" >
<input type="hidden" name="lec_idx" value="<%= oFingers.FItemList(i).Flec_idx %>">
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center"><%= getLecMenuName(oFingers.FItemList(i).FMenuId) %></td>
	<td align="center"><img src="<%= oFingers.FItemList(i).Fimagesmall %>" width="50" height="50"></td>
	<td align="center"><%= oFingers.FItemList(i).Flec_idx %></td>
	<td align="center"><%= oFingers.FItemList(i).Flec_title %></td>
	<td align="center"><input type="text" name="sortNo" value="<%= oFingers.FItemList(i).FsortNo %>" size="3" style="text-align:right;" /></td>
	<td align="center"><%= oFingers.FItemList(i).Fisusing %></td>
	<td align="center">
		<% if oFingers.FItemList(i).IsSoldOut then %>
		<font color="red">����</font>
		<% end if %>
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" align="center">
	<% if oFingers.HasPreScroll then %>
		<a href="?page=<%= oFingers.StarScrollPage-1 %>&MenuId=<%=MenuId%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oFingers.StarScrollPage to oFingers.FScrollCount + oFingers.StarScrollPage - 1 %>
		<% if i>oFingers.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&MenuId=<%=MenuId%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oFingers.HasNextScroll then %>
		<a href="?page=<%= i %>&MenuId=<%=MenuId%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
set oFingers = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyClose.asp" -->
