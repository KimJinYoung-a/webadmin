<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2010.04.07 ������ ����
'	Description : Favorite Colore ����
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/favoriteColorCls.asp"-->
<%
dim page, Category, colorCD, isusing
Dim oitem, lp , schcolorCD
	category	= request("category")
	colorCD		= request("colorCD")
	page		= request("page")
	isusing		= request("isusing")
	schcolorCD	= request("schcolorCD")
	
	if page="" then page=1

dim oip
	set oip = New CfavoriteColor
	oip.FCurrPage = page
	oip.FPageSize=20
	oip.FRectCategory = category
	oip.FRectColorCD = schcolorCD
	oip.FRectIsUsing = isusing
	oip.GetfavoriteColor

dim i
	
	set oitem = new CItemColor
	oitem.FPageSize = 50
	oitem.FRectUsing = "Y"
	oitem.GetColorList

%>
<script language='javascript'>

	function popItemWindow(tgf){
		if (document.Listfrm.category.value == ""){
			alert("���� ������ �ּ���!");
			document.Listfrm.category.focus();
		}
		else if (document.Listfrm.schcolorCD.value == ""){
			alert("������ ������ �ּ���!");
		}
		else{
			var popup_item = window.open("/common/pop_CateItemList.asp?category=" + document.refreshFrm.category.value + "&target=" + tgf, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
			popup_item.focus();
		}
	}

	function popColorWindow(){
		var popup_item = window.open("/admin/sitemaster/favoriteColor/popManageColorCode.asp", "popup_item", "width=380,height=600,scrollbars=yes,status=no");
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
						upfrm.idx.value = upfrm.idx.value + frm.idx.value + "," ;
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
						upfrm.idx.value = upfrm.idx.value + frm.idx.value + "," ;
					}
				}
			}
			upfrm.category.value = Listfrm.category.value;
			upfrm.colorCD.value = Listfrm.schcolorCD.value;
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
						upfrm.idx.value = upfrm.idx.value + frm.idx.value + "," ;
						upfrm.sortNo.value = upfrm.sortNo.value + frm.sortNo.value + "," ;
					}
				}
			}
			upfrm.category.value = Listfrm.category.value;
			upfrm.colorCD.value = Listfrm.schcolorCD.value;
			upfrm.mode.value="ChangeSort";
			upfrm.submit();
	
		}
	}
	
	function AddIttems(){
		var ret = confirm(arrFrm.itemid.value + '�������� �߰��Ͻðڽ��ϱ�?');
		if (ret){
			arrFrm.itemid.value = arrFrm.itemid.value;
			arrFrm.category.value = Listfrm.category.value;
			arrFrm.colorCD.value = Listfrm.schcolorCD.value;
			arrFrm.mode.value="add";
			arrFrm.submit();
		}
	}
	
	function AddIttems2(){
		if (document.Listfrm.category.value == ""){
			alert("���� ������ �ּ���!");
			document.Listfrm.category.focus();
		}
		else if (document.Listfrm.schcolorCD.value == ""){
			alert("������ ������ �ּ���!");
		}
		else if (document.arrFrm.itemidarr.value == ""){
			alert("�����۹�ȣ��  �����ּ���!");
			document.arrFrm.itemidarr.focus();
		}
		else if (confirm(arrFrm.itemidarr.value + '�������� �߰��Ͻðڽ��ϱ�?')){
			arrFrm.itemid.value = arrFrm.itemidarr.value;
			arrFrm.category.value = Listfrm.category.value;
			arrFrm.colorCD.value = Listfrm.schcolorCD.value;
			arrFrm.mode.value="add";
			arrFrm.submit();
		}
	}

	// �� ����� ���
	function changecontent(){}

	function chgColorChip(ccd,cnt,idx) {
		document.Listfrm.colorCD.value=ccd;
		document.Listfrm.schcolorCD.value=idx;
		for(var i=0;i<=(cnt-1);i++) {
			if(i==ccd) {
				document.getElementById("tbColor"+i).style.backgroundColor="#000000";
			} else {
				document.getElementById("tbColor"+i).style.backgroundColor="#EDEDED";
			}
		}
	}

</script>
<form name="refreshFrm" method="post">
<input type="hidden" name="category">
<input type="hidden" name="colorCD">
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
		<% DrawSelectBoxCateTab "category", category %>&nbsp;/&nbsp;
		������� :
		<select name="isusing" onchange="this.form.submit();" class="select">
			<option value=""  <% if isusing="" then response.write "selected" %>>��ü</option>
			<option value="Y" <% if isusing="Y" then response.write "selected" %>>���</option>
			<option value="N" <% if isusing="N" then response.write "selected" %>>������</option>
		</select><br>
		<%' DrawSelectBoxColoreBar "colorCD",colorCD %>
		<table border="0" cellspacing="3" cellpadding="0">
		<tr>
		<%
			'For i=1 to 20
			if oitem.FResultCount>0 then
				for lp=0 to oitem.FResultCount-1
		%>
			<td onClick="chgColorChip(<%=lp%>,'<%=oitem.FResultCount%>','<%=oitem.FItemList(lp).FcolorCode%>')" style="cursor:pointer">
				<table id="tbColor<%=lp%>" border="0" cellpadding="0" cellspacing="1" bgcolor="<% if cstr(colorCD)=cstr(lp) then %>#000000<% else %>#EDEDED<% end if %>">
				<tr>
					<td bgcolor="#FFFFFF"><img src="<%=oitem.FItemList(lp).FcolorIcon%>" width="15" height="15" hspace="1" vspace="1" border="0"></td>
				</tr>
				</table>
			</td>
		<%
				Next
			End If 
		%>
		<input type="hidden" name="colorCD" value="<%=colorCD%>">
		<input type="hidden" name="schcolorCD" value="<%=schcolorCD%>">
		</tr>
		</table>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
		<form name="arrFrm" method="post" action="doFavoriteColor.asp">
		<input type="hidden" name="category">
		<input type="hidden" name="colorCD">
		<input type="hidden" name="mode">
		<input type="hidden" name="itemid">
		<input type="hidden" name="idx">
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
			<% If C_ADMIN_AUTH then%>
			<td align="right"><input type="button" value="Color �߰�" onclick="popColorWindow();" class="button"></td>
			<% End If %>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FAFAFA" height="22">
	<td colspan="9">&nbsp;�˻��� ��ǰ�� : <%=oip.FTotalCount%> �� &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� ���� 0�� : [���� basic �̹��� ����] �� �Ǻ� ���� ���� ��ǰ 1�� �ʼ��Դϴ�.</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td align="center">��</td>
	<td align="center">�÷�</td>
	<td align="center">�̹���</td>
	<td align="center">ItemID</td>
	<td align="center">��ǰ��</td>
	<td align="center">����</td>
	<td align="center">�������</td>
	<td align="center">ǰ������</td>
</tr>
<% for i=0 to oip.FResultCount-1 %>
<form name="frmBuyPrc<%=i%>" method="post" action="" >
<input type="hidden" name="idx" value="<%= oip.FItemList(i).FIDX %>">
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center">
		<% if  oip.FItemList(i).Fcategory = 1 then
		response.write "Stationary&Persnal"
		elseif oip.FItemList(i).Fcategory = 2 then
		response.write "Home&Living"
		elseif 	oip.FItemList(i).Fcategory = 3 then
		response.write "Fashion&Beauty"
		elseif 	oip.FItemList(i).Fcategory = 4 then
		response.write "Kidult&Hobby"
		elseif 	oip.FItemList(i).Fcategory = 5 then
		response.write "Kids&Baby"
		else
		response.write "�̺з�"
		end if
		%>
	</td>
	<td align="center"><img src="http://fiximage.10x10.co.kr/web2011/common/color/<%= oip.FItemList(i).FcolorCD%>" width="20" height="20"></td>
	<td align="center"><img src="<%= oip.FItemList(i).Fimagesmall %>" width="50" height="50"></td>
	<td align="center"><%= oip.FItemList(i).FItemID %></td>
	<td align="center"><%= oip.FItemList(i).FItemname %></td>
	<td align="center"><input type="text" name="sortNo" value="<%= oip.FItemList(i).FsortNo %>" size="3" style="text-align:right;" /><%If oip.FItemList(i).FsortNo = "0" then%>����<%End if%></td>
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
	<td colspan="9" align="center">
	<% if oip.HasPreScroll then %>
		<a href="?page=<%= oip.StarScrollPage-1 %>&category=<%=category%>&isusing=<%=isusing%>&menupos=<%= menupos %>&colorCD=<%=colorCD%>&schcolorCD=<%=schcolorCD%>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oip.StarScrollPage to oip.FScrollCount + oip.StarScrollPage - 1 %>
		<% if i>oip.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&category=<%=category%>&isusing=<%=isusing%>&menupos=<%= menupos %>&colorCD=<%=colorCD%>&schcolorCD=<%=schcolorCD%>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oip.HasNextScroll then %>
		<a href="?page=<%= i %>&category=<%=category%>&isusing=<%=isusing%>&menupos=<%= menupos %>&colorCD=<%=colorCD%>&schcolorCD=<%=schcolorCD%>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
set oip = Nothing
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
