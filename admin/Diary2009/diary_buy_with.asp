<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/Diary2009/classes/DiaryCls.asp"-->
<%
dim page, isusing, cate
page = request("page")
cate = request("cated")
isusing = request("isusing")

if page="" then page=1

dim omd
set omd = New DiaryCls
omd.FCurrPage = page
omd.FPageSize=20
omd.FRectCDL = cate
omd.FRectIsUsing = isusing
omd.GetWithBuyList

dim i
%>
<script language='javascript'>
<!--
function popItemWindow(tgf){
	if (document.Listfrm.cated.value == ""){
		alert("ī�װ��� ������ �ּ���!");
		document.Listfrm.cated.focus();
	}
	else{
		var popup_item = window.open("/common/pop_CateItemList.asp?cdl=010&cdm=010&target=" + tgf, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
		popup_item.focus();
	}
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
		upfrm.mode.value="del";
		upfrm.submit();
	}
}

function changeUsing(upfrm){
	if (!CheckSelected()){
		alert('��ǰ�� ������ �ּ���.');
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
		upfrm.mode.value="isUsingValue";
		upfrm.submit();
	}
}

//ī�װ� ���� ����
function changeCate(upfrm) {
	if(document.Listfrm.cated.value == ""){
		alert('ī�װ��� ������ �ּ���.');
		document.Listfrm.cated.focus();
		return;
	}
	if (!CheckSelected()){
		alert('��ǰ�� ������ �ּ���.');
		return;
	}
	var ret = confirm('�����Ͻ� ���̾ ī�װ��� �����Ͻðڽ��ϱ�?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.idx.value = upfrm.idx.value + frm.idx.value + "," ;
					upfrm.cate.value = document.Listfrm.cated.value ;
				}
			}
		}
		upfrm.mode.value="modify";
		upfrm.submit();
	}
}

function AddIttems(){
	var ret = confirm(arrFrm.itemid.value + '�������� �߰��Ͻðڽ��ϱ�?');
	if (ret){
		arrFrm.itemid.value = arrFrm.itemid.value;
		arrFrm.cate.value = document.Listfrm.cated.value;
		arrFrm.mode.value="add";
		arrFrm.submit();
	}
}

function AddIttems2(){
	if (document.Listfrm.cated.value == ""){
		alert("ī�װ��� ������ �ּ���!");
		document.Listfrm.cated.focus();
	}
	else if (document.arrFrm.itemidarr.value == ""){
		alert("�����۹�ȣ��  �����ּ���!");
		document.arrFrm.itemidarr.focus();
	}
	else if (confirm(arrFrm.itemidarr.value + '�������� �߰��Ͻðڽ��ϱ�?')){
		arrFrm.itemid.value = arrFrm.itemidarr.value;
		arrFrm.cate.value = document.Listfrm.cated.value;
		arrFrm.mode.value="add";
		arrFrm.submit();
	}
}

function modify_catename(idx,mode){
	window.open('/admin/Diary2009/pop_Diartycate.asp?idx='+idx+'&mode='+mode);
}
//-->
</script>
<!-- ��� �˻��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="Listfrm" method="get" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		���̾ ī�װ� :
		<select name="cated">
			<option value="">-����-
			<option value="10" <% if cate="10" then response.write "selected" %>>����
			<option value="20" <% if cate="20" then response.write "selected" %>>�Ϸ���Ʈ
			<option value="30" <% if cate="30" then response.write "selected" %>>����
			<option value="40" <% if cate="40" then response.write "selected" %>>����
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
</form>
<!-- �˻� �� -->
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
		<form name="arrFrm" method="post" action="doWithBuy.asp">
		<input type="hidden" name="cate">
		<input type="hidden" name="mode">
		<input type="hidden" name="itemid">
		<input type="hidden" name="sortNo">
		<input type="hidden" name="idx">
		<tr>
			<td colspan="2" align="right">
				<input type="text" name="itemidarr" value="" size="80" class="input" onkeypress="if(event.keyCode==13){return false;}">
				<input type="button" value="������ ���� �߰�" onclick="AddIttems2()" class="button">
			</td>
		</tr>		
		<tr>
			<td>
				<input type="button" value="���þ����� ����" onClick="delitems(arrFrm)" class="button"> /
				<select name="allusing"  class="select">
					<option value="Y">���� ��� -> Y </option>
					<option value="N">���� ��� -> N </option>
				</select><input type="button" value="����" class="button" onclick="changeUsing(arrFrm);"> /
				<input type="button" value="��������" class="button" onclick="changeSort(arrFrm);"> /
				<input type="button" value="ī�װ�����" class="button" onclick="changeCate(arrFrm);"> /
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
	<td align="center">���̾ ī�װ���</td>
	<td align="center">Image</td>
	<td align="center">ItemID</td>
	<td align="center">��ǰ��</td>
	<td align="center">����</td>
	<td align="center">�������</td>
	<td align="center">ǰ������</td>
</tr>
<%
Dim catename
%>
<% for i=0 to omd.FResultCount-1 %>
<%
	If omd.FItemList(i).FCdl = 10 Then catename = "����"
	If omd.FItemList(i).FCdl = 20 Then catename = "�Ϸ���Ʈ"
	If omd.FItemList(i).FCdl = 30 Then catename = "����"
	If omd.FItemList(i).FCdl = 40 Then catename = "����"
%>
<form name="frmBuyPrc<%=i%>" method="post" action="" >
<input type="hidden" name="itemid" value="<%= omd.FItemList(i).FItemID %>">
<input type="hidden" name="idx" value="<%= omd.FItemList(i).FIdx %>">
<input type="hidden" name="cate" value="<%= omd.FItemList(i).FCdl %>">
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center"><a href="javascript:modify_catename('<%=omd.FItemList(i).Fidx%>','modify');"><%= catename %></a></td>
	<td align="center"><img src="<%= omd.FItemList(i).Fimagesmall %>" width="50" height="50"></td>
	<td align="center"><%= omd.FItemList(i).FItemID %></td>
	<td align="center"><%= omd.FItemList(i).FItemname %></td>
	<td align="center"><input type="text" name="sortNo" value="<%= omd.FItemList(i).FsortNo %>" size="3" style="text-align:right;" /></td>
	<td align="center"><%= omd.FItemList(i).Fisusing %></td>
	<td align="center">
		<% if omd.FItemList(i).IsSoldOut then %>
		<font color="red">ǰ��</font>
		<% end if %>
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" align="center">
	<% if omd.HasPreScroll then %>
		<a href="?page=<%= omd.StarScrollPage-1 %>&cdl=<%=cdl%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + omd.StartScrollPage to omd.FScrollCount + omd.StartScrollPage - 1 %>
		<% if i>omd.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&cate=<%=cate%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if omd.HasNextScroll then %>
		<a href="?page=<%= i %>&cate=<%=cate%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[next]</a>
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
