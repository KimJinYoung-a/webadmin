<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.04.18 �ѿ�� ī�װ���md�� �̵�/ �߰�/����
'	Description : ���������� �ǰ���
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_tabitem_cls.asp"-->
<%
dim page, cdl, isusing
	cdl = request("cdl")
	page = request("page")
	isusing = request("isusing")
	
	if page="" then page=1


dim oip
	set oip = New Cmain_tabitem_list
	oip.FCurrPage = page
	oip.FPageSize=20
	oip.FRectCDL = cdl
	oip.FRectIsUsing = isusing
	oip.Getmain_tabitem

dim i
%>
<script language='javascript'>

	function popItemWindow(tgf){
		if (document.Listfrm.cdl.value == ""){
			alert("ī�װ����� ������ �ּ���!");
			document.Listfrm.cdl.focus();
		}
		else{
			var popup_item = window.open("/common/pop_CateItemList.asp?cdl=" + document.refreshFrm.cdl.value + "&target=" + tgf, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
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
			upfrm.cdl.value = Listfrm.cdl.value;
			upfrm.mode.value="isUsingValue";
			upfrm.submit();
	
		}
	}
	
	// ��������
	function changeSort(upfrm) {
		if (!CheckSelected()){
			alert('��ǰ�� ������ �ּ���.');
			return;
		}
		var ret = confirm('�����Ͻ� ��ǰ�� ������ �����Ͻ� ��ȣ�� �����Ͻðڽ��ϱ�?');
	
		if (ret){
			var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.itemid.value = upfrm.itemid.value + frm.itemid.value + "," ;
						upfrm.sortNo.value = upfrm.sortNo.value + frm.sortNo.value + "," ;
					}
				}
			}
			upfrm.cdl.value = Listfrm.cdl.value;
			upfrm.mode.value="ChangeSort";
			upfrm.submit();
	
		}
	}
	
	function AddIttems(){
		var ret = confirm(arrFrm.itemid.value + '�������� �߰��Ͻðڽ��ϱ�?');
		if (ret){
			arrFrm.itemid.value = arrFrm.itemid.value;
			arrFrm.cdl.value = Listfrm.cdl.value;
			arrFrm.mode.value="add";
			arrFrm.submit();
		}
	}
	
	function AddIttems2(){
		if (document.Listfrm.cdl.value == ""){
			alert("ī�װ����� ������ �ּ���!");
			document.Listfrm.cdl.focus();
		}
		else if (document.arrFrm.itemidarr.value == ""){
			alert("�����۹�ȣ��  �����ּ���!");
			document.arrFrm.itemidarr.focus();
		}
		else if (confirm(arrFrm.itemidarr.value + '�������� �߰��Ͻðڽ��ϱ�?')){
			arrFrm.itemid.value = arrFrm.itemidarr.value;
			arrFrm.cdl.value = Listfrm.cdl.value;
			arrFrm.mode.value="add";
			arrFrm.submit();
		}
	}

	function RefreshCaMDChoiceRec(){
		if (document.Listfrm.cdl.value == ""){
			alert("ī�װ����� ������ �ּ���!");
			document.Listfrm.cdl.focus();
		}
		 else{
				  var popwin = window.open('','refreshFrm','');
				  popwin.focus();
				  refreshFrm.target = "refreshFrm";
				  refreshFrm.cdl.value = document.Listfrm.cdl.value;
				  refreshFrm.action = "<%=wwwUrl%>/chtml/make_main_tabitem.asp";
				  refreshFrm.submit();
		 }
	}

	// ī�װ��� ����� ����
	function changecontent(){}

</script>
<form name="refreshFrm" method="post">
<input type="hidden" name="cdl">
</form>
<!-- ��� �˻��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="Listfrm" method="get" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		�Ǽ��� :
		<select name='cdl' class="select">
			<option value="">�����ϼ���</option>
			<option value=1 <% if cdl = "1" then response.write " selected" %>>������/���ǽ�</option>
			<option value=2 <% if cdl = "2" then response.write " selected" %>>Ű��Ʈ/���</option>
			<option value=3 <% if cdl = "3" then response.write " selected" %>>����</option>
			<option value=4 <% if cdl = "4" then response.write " selected" %>>�м�</option>
			<option value=5 <% if cdl = "5" then response.write " selected" %>>���̺�/Ű��</option>
			<option value=6 <% if cdl = "6" then response.write " selected" %>>����ä��</option>
		</select>
		&nbsp;/&nbsp;
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
			</td>
		</tr>
		</form>
		<form name="arrFrm" method="post" action="domaintabitem.asp">
		<input type="hidden" name="cdl">
		<input type="hidden" name="mode">
		<input type="hidden" name="itemid">
		<input type="hidden" name="sortNo">
		<tr>
			<td colspan="2" align="right">
				<input type="text" name="itemidarr" value="" size="80" class="input">
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
				<input type="button" value="��������" class="button" onclick="changeSort(arrFrm);">
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
	<td colspan="8">&nbsp;�˻��� ��ǰ�� : <%=oip.FTotalCount%> ��</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td align="center">ī�װ�����</td>
	<td align="center">ItemID</td>
	<td align="center">Image</td>
	<td align="center">��ǰ��</td>
	<td align="center">����</td>
	<td align="center">�������</td>
	<td align="center">ǰ������</td>
</tr>
<% for i=0 to oip.FResultCount-1 %>
<form name="frmBuyPrc<%=i%>" method="post" action="" >
<input type="hidden" name="itemid" value="<%= oip.FItemList(i).FItemID %>">
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center">
		<% if  cStr(oip.FItemList(i).Fcdl) = "1" then
		response.write "������/���ǽ�"
		elseif cStr(oip.FItemList(i).Fcdl) = "2" then
		response.write "Ű��Ʈ/���"
		elseif 	cStr(oip.FItemList(i).Fcdl) = "3" then
		response.write "����"
		elseif 	cStr(oip.FItemList(i).Fcdl) = "4" then
		response.write "�м�"
		elseif 	cStr(oip.FItemList(i).Fcdl) = "5" then
		response.write "���̺�/Ű��"
		elseif 	cStr(oip.FItemList(i).Fcdl) = "6" then
		response.write "����ä��"
		end if
		%>
	</td>
	<td align="center"><img src="<%= oip.FItemList(i).Fimagesmall %>" width="50" height="50"></td>
	<td align="center"><%= oip.FItemList(i).FItemID %></td>
	<td align="center"><%= oip.FItemList(i).FItemname %></td>
	<td align="center"><input type="text" name="sortNo" value="<%= oip.FItemList(i).FsortNo %>" size="3" style="text-align:right;" /></td>
	<td align="center"><%= oip.FItemList(i).Fisusing %></td>
	<td align="center">
		<% if oip.FItemList(i).IsSoldOut then %>
		<font color="red">ǰ��</font>
		<% end if %>
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" align="center">
	<% if oip.HasPreScroll then %>
		<a href="?page=<%= oip.StarScrollPage-1 %>&cdl=<%=cdl%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oip.StarScrollPage to oip.FScrollCount + oip.StarScrollPage - 1 %>
		<% if i>oip.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&cdl=<%=cdl%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oip.HasNextScroll then %>
		<a href="?page=<%= i %>&cdl=<%=cdl%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
set oip = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->