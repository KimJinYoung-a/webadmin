<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/category_md_choicecls.asp"-->
<%
dim page, cdl, cdm, isusing, vIdx, vDisp1
cdl = request("cdl")
cdm = request("cdm")
page = request("page")
isusing = request("isusing")
vIdx = request("idx")
vDisp1 = request("catecode")
if page="" then page=1


dim omd
set omd = New CMDChoice
omd.FCurrPage = page
omd.FPageSize=20
omd.FRectDisp1 = vDisp1
omd.FRectIdx = vIdx
omd.FRectIsUsing = isusing
omd.GetMDChoiceList

dim i
%>
<script language='javascript'>
<!--
function popItemWindow(tgf){
	var popup_item = window.open("/common/pop_CateItemList.asp?target=" + tgf + "&disp=<%=vDisp1%>", "popup_item", "width=1000,height=800,scrollbars=yes,resizable=yes");
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
		upfrm.mode.value="ChangeSort";
		upfrm.submit();

	}
}

function AddIttems(){
	var ret = confirm(arrFrm.itemid.value + '�������� �߰��Ͻðڽ��ϱ�?');
	if (ret){
		arrFrm.itemidarr.value = arrFrm.itemid.value;
		arrFrm.submit();
	}
}

function AddIttems2(){
	if (document.arrFrm.itemidarr.value == ""){
		alert("�����۹�ȣ��  �����ּ���!");
		document.arrFrm.itemidarr.focus();
	}
	else if (confirm(arrFrm.itemidarr.value + '�������� �߰��Ͻðڽ��ϱ�?')){
		arrFrm.itemid.value = arrFrm.itemidarr.value;
		arrFrm.itemsYN.value = "Y";
		arrFrm.submit();
	}
}

//-->
</script>
<!-- ��� �˻��� ���� -->
<form name="Listfrm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="idx" value="<%=vIdx%>">
<input type="hidden" name="catecode" value="<%=vDisp1%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
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
		<form name="arrFrm" method="post" action="doMDChoice.asp" style="margin:0px;">
		<input type="hidden" name="cdl">
		<input type="hidden" name="cdm">
		<input type="hidden" name="mode" value="add">
		<input type="hidden" name="itemid">
		<input type="hidden" name="sortNo">
		<input type="hidden" name="idx" value="<%=vIdx%>">
		<input type="hidden" name="itemsYN">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td colspan="2" align="right">
				<textarea rows="3" cols="10" name="arrItems" id="itemidarr"></textarea>
				<input type="button" value="������ ���� �߰�" onclick="AddIttems2()" class="button">
				&nbsp;<input type="button" value="������ �߰�" onclick="popItemWindow('arrFrm.itemid')" class="button">
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
			<td align="right"></td>
		</tr>
		</table>
		</form>
	</td>
</tr>
</table>
<!-- �׼� �� -->
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FAFAFA" height="22">
	<td colspan="10">&nbsp;�˻��� ��ǰ�� : <%=omd.FTotalCount%> ��</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td align="center">ī�װ�����</td>
	<td align="center">ItemID</td>
	<td align="center">Image</td>
	<td align="center">��ǰ��</td>
	<td align="center">����</td>
	<td align="center">��üID</td>
	<td align="center">����</td>
	<td align="center">�������</td>
	<td align="center">ǰ������</td>
</tr>
<% for i=0 to omd.FResultCount-1 %>
<form name="frmBuyPrc<%=i%>" method="post" action="" >
<input type="hidden" name="itemid" value="<%= omd.FItemList(i).FItemID %>">
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center"><%= omd.FItemList(i).Fcode_nm %>&nbsp;<%= omd.FItemList(i).Fmidcode_nm %></td>
	<td align="center"><img src="<%= omd.FItemList(i).Fimagesmall %>" width="50" height="50"></td>
	<td align="center"><%= omd.FItemList(i).FItemID %></td>
	<td align="center"><%= omd.FItemList(i).FItemname %></td>
	<td align="center">
			<%= FormatNumber(omd.FItemList(i).Forgprice,0) %>
			<%
				'���ΰ�
				if omd.FItemList(i).Fsailyn="Y" then
					Response.Write "<br><font color=#F08050>(��)" & FormatNumber(omd.FItemList(i).Fsailprice,0) & "</font>"
				end if
				'������
				if omd.FItemList(i).FitemCouponYn="Y" then
					Select Case omd.FItemList(i).FitemCouponType
						Case "1"
							Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(omd.FItemList(i).Forgprice*((100-omd.FItemList(i).FitemCouponValue)/100),0) & "</font>"
						Case "2"
							Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(omd.FItemList(i).Forgprice-omd.FItemList(i).FitemCouponValue,0) & "</font>"
					end Select
				end if
			%>
	</td>
	<td align="center"><%= omd.FItemList(i).FMakerId %></td>
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
	<td colspan="10" align="center">
	<% if omd.HasPreScroll then %>
		<a href="?page=<%= omd.StarScrollPage-1 %>&cdl=<%=cdl%>&idx=<%=vIdx%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + omd.StarScrollPage to omd.FScrollCount + omd.StarScrollPage - 1 %>
		<% if i>omd.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&cdl=<%=cdl%>&idx=<%=vIdx%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if omd.HasNextScroll then %>
		<a href="?page=<%= i %>&cdl=<%=cdl%>&idx=<%=vIdx%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[next]</a>
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