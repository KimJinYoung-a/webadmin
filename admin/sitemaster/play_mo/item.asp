<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/play/play_moCls.asp" -->

<%
dim page, isusing, oitem, i, itemid, itemname, playidx, playcate
	page		= request("page")
	isusing		= request("isusing")
	itemid = request("itemid")
	itemname = request("itemname")
	playidx = request("playidx")
	playcate = request("playcate")
	
	if page="" then page=1
	if isusing = "" then isusing = "Y"
	
set oitem = New CPlayMoContents
	oitem.FCurrPage = page
	oitem.FPageSize=20
	oitem.FRectPlayIdx = playidx
	oitem.FRectitemid = itemid
	oitem.FRectitemname = itemname
	oitem.FRectIsUsing = isusing
	oitem.fnPlayItemList
%>

<script>
//��뿩�� ����
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
					upfrm.idx.value = upfrm.idx.value + frm.idx.value + "," ;
				}
			}
		}

		upfrm.mode.value="chisusing";
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
					upfrm.idx.value = upfrm.idx.value + frm.idx.value + "," ;
					upfrm.orderno.value = upfrm.orderno.value + frm.orderno.value + "," ;
				}
			}
		}
		upfrm.mode.value="ChangeSort";
		upfrm.submit();

	}
}

//��ǰ����
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
					upfrm.idx.value = upfrm.idx.value + frm.idx.value + "," ;
				}
			}
		}
		upfrm.mode.value="delitem";
		upfrm.submit();

	}
}

//��ǰ �˻� �˾� ���� ��ǰ ������ ���� �۾�
function AddIttems(){
	var ret = confirm(arrFrm.itemid.value + '�������� �߰��Ͻðڽ��ϱ�?');
	if (ret){
		arrFrm.itemid.value = arrFrm.itemid.value;
		arrFrm.mode.value="itemadd";
		arrFrm.submit();
	}
}

//��ǰ �˻� �˾�
function popItemWindow(tgf){
	
	if (document.Listfrm.playidx.value == ""){
		alert("PLAY idx���� �����ϴ�.\nâ�� �ݰ� �ٽ� ����� �ּ���!");
		return;
	}

	var popup_item = window.open("/common/pop_CateItemList.asp?target=" + tgf, "popup_item", "width=1024,height=768,scrollbars=yes,resizable=yes");
	popup_item.focus();
}

//��ǰ ���� ���� �߰�
function AddIttems2(){

	if (document.arrFrm.itemidarr.value == ""){
		alert("�����۹�ȣ��  �����ּ���!");
		return;
		document.arrFrm.itemidarr.focus();
	}
	if (confirm(arrFrm.itemidarr.value + '�������� �߰��Ͻðڽ��ϱ�?')){
		arrFrm.itemid.value = arrFrm.itemidarr.value;
		arrFrm.mode.value="itemadd";
		arrFrm.submit();
	}
}

//�˻�
function jsSerach(ipage){
	var frm;
	frm = document.Listfrm;
	
	if(frm.itemid.value!=''){
		if (!IsDouble(frm.itemid.value)){
			alert('��ǰ �ڵ�� ���ڸ� �����մϴ�.');
			frm.itemid.focus();
			return;
		}
	}

	frm.page.value= ipage;
	frm.submit();
}

//���ý� tr �� ����
function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="Listfrm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="playidx" value="<%=playidx%>">
<input type="hidden" name="playcate" value="<%=playcate%>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		��� : <% drawSelectBoxUsingYN "isusing", isusing %>
		��ǰ�ڵ� : <input type="text" name="itemid" value="<%= itemid %>" size=10 maxlength=10>
		��ǰ�� : <input type="text" name="itemname" value="<%= itemname %>" size=30 maxlength=30>
	</td>	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSerach('');">
	</td>
</tr>
</form>
</table>
<br>

<table width="100%" align="center" cellpadding="2" cellspacing="0" class="a">
<form name="arrFrm" method="post" action="item_proc.asp">
<input type="hidden" name="mode">
<input type="hidden" name="itemid">
<input type="hidden" name="idx">
<input type="hidden" name="orderno">
<input type="hidden" name="playidx" value="<%=playidx%>">
<input type="hidden" name="playcate" value="<%=playcate%>">
<tr>
	<td align="left">
		<input type="button" value="���þ����� ����" onClick="delitems(arrFrm)" class="button"> /
		<select name="allusing"  class="select">
			<option value="Y">���� ��� -> Y </option>
			<option value="N">���� ��� -> N </option>
		</select><input type="button" value="����" class="button" onclick="changeUsing(arrFrm);"> /
		<input type="button" value="��������" class="button" onclick="changeSort(arrFrm);">
	</td>	
	<td align="right">
		<input type="text" name="itemidarr" value="" size="80" class="input" onKeyPress="if (event.keyCode == 13){ AddIttems2(); return false;}">
		<input type="button" value="��ǰ �����߰�" onclick="AddIttems2()" class="button">
		<input type="button" value="��ǰ �˻��߰�" onclick="popItemWindow('arrFrm.itemid')" class="button">
	</td>
</tr>
</form>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr bgcolor="#FFFFFF">
	<td colspan="20">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				�˻���� : <b><%= oitem.FTotalCount%></b>
				&nbsp;
				������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>				
			</td>
			<td align="right">
				��ī�װ� ALL �ϰ�쿡��, ����Ʈ�� �ִ� ��ǰ�� ���� ������� ���� �˴ϴ�.		
			</td>			
		</tr>
		</table>
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>�̹���</td>
	<td>ItemID</td>
	<td>��ǰ��</td>
	<td>����<br>����</td>
	<td>���<br>����</td>
	<td>ǰ��<br>����</td>
</tr>
<% if oitem.FResultCount > 0 then %>

<% for i=0 to oitem.FResultCount-1 %>
<form name="frmBuyPrc<%=i%>" method="post" action="" >
<input type="hidden" name="idx" value="<%= oitem.FItemList(i).FIDX %>">
<% if oitem.FItemList(i).fisusing = "Y" then %>
	<tr bgcolor="#FFFFFF">
<% else %>
	<tr bgcolor="#f1f1f1">
<% end if %>
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center"><img src="<%= oitem.FItemList(i).Fimagesmall %>" width="50" height="50"></td>
	<td align="center"><%= oitem.FItemList(i).FItemID %></td>
	<td align="center"><%= oitem.FItemList(i).FItemname %></td>
	<td align="center">
		<input type="text" name="orderno" value="<%= oitem.FItemList(i).forderno %>" size="3" style="text-align:right;" onKeyup="CheckThis(frmBuyPrc<%= i %>)">
	</td>
	<td align="center"><%= oitem.FItemList(i).Fisusing %></td>
	<td align="center">
		<% if oitem.FItemList(i).IsSoldOut then %>
			<font color="red">ǰ��</font>
		<% end if %>
	</td>
</tr>
</form>
<% next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if oitem.HasPreScroll then %>
			<a href="javascript:jsSerach('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
			<% if i>oitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:jsSerach('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if oitem.HasNextScroll then %>
			<a href="javascript:jsSerach('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
</table>

<%
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->