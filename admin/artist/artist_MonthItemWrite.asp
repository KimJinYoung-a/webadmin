<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��Ƽ��Ʈ�� ������ ���
' History : 2012.03.29 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/artist/artist_class.asp"-->
<%
'// ���� ����
dim oGallery
dim mode, idx, lp
dim page, isusing, gal_div, designerid
mode = request("mode")
If mode = "" Then mode = "add"

idx = request("idx")
page = request("page")
isusing = request("isusing")
gal_div = request("gal_div")
designerid = request("designerid")
%>
<script language="javascript">
function subcheck(){
	var frm=document.inputfrm;

	if(frm.itemid.value==""){
		alert('��ǰ�ڵ带 ����ϼ���');
		frm.itemid.focus();
		return;
	}
	if(frm.comment.value==""){
		alert('�ڸ�Ʈ�� �Է��ϼ���');
		frm.comment.focus();
		return;
	}

	var str_len = frm.comment.value;
	line = str_len.split("\r\n");
	ln = line.length;

	if(ln > 5){
	   alert("�ڸ�Ʈ�� 5�ٱ��� �����մϴ�.");
	  return;
	}

	if(frm.sortNo.value==""){
		alert('������ �Է��ϼ���');
		frm.sortNo.focus();
		return;
	}
	frm.submit();
}

//���ι�� ��ϻ�ǰ ��ǰã��
function popItemWindow(tgf){
	var popup_item = window.open("/common/pop_singleItemSelect.asp?target=" + tgf, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
	popup_item.focus();
}

function f_chk_byte(aro_name,ari_max) { 
	var ls_str = aro_name.value;
	var li_str_len = ls_str.length;
	var li_max = ari_max;
	var i = 0;
	var li_byte = 0;
	var li_len = 0;
	var ls_one_char = "";
	var ls_str2 = "";

	for(i=0; i< li_str_len; i++) {
		ls_one_char = ls_str.charAt(i);
		if (escape(ls_one_char).length > 4) 
		li_byte += 2;
		else 
		li_byte++;
		
		if (li_byte <= li_max) li_len = i + 1;
	}

	if(li_byte > li_max) {
		alert("�ѱ� " + ari_max + "���ڸ� �ʰ� �Է��Ҽ� �����ϴ�. �ʰ��� ������ �ڵ����� ���� �˴ϴ�.");
		ls_str2 = ls_str.substr(0, li_len);
		aro_name.value = ls_str2;
	}
	aro_name.focus(); 
}

</script>
<p>
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<col width="10%"></col>
	<col></col>
	<form name="inputfrm" method="post" action="artist_MonthItem_process.asp">
	<input type="hidden" name="mode" value="<% =mode %>">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="page" value="<%= page %>">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">

	<% if mode="add" then %>
	<tr>
		<td height="30" align="center" bgcolor="#F0F0FD">��ǰ</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="itemid" size=10 readonly>
			<input type="button" class="button" value="ã��" onClick="popItemWindow('inputfrm')">
		</td>
	</tr>
	<tr>
		<td height="30" align="center" bgcolor="#F0F0FD">�ڸ�Ʈ</td>
		<td bgcolor="#FFFFFF"><textarea cols="50" rows="5" name="comment" onkeyup="f_chk_byte(this,235)" style="overflow:hidden"></textarea><br>--���� �������� ������ ��ŭ�� �ѷ����ϴ�.�ʰ�����!!</td>
	</tr>
	<tr>
		<td height="30" align="center" bgcolor="#F0F0FD">����</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="sortNo" size=2 maxlength="1">
		</td>
	</tr>
	<tr>
		<td height="30" align="center" bgcolor="#F0F0FD">�������</td>
		<td bgcolor="#FFFFFF">
			<label><input type="radio" name="isusing" value="Y" checked>Y</label>
			<label><input type="radio" name="isusing" value="N">N</label>
		</td>
	</tr>
	<tr bgcolor="#DDDDFF" >
		<td colspan="2" align="center">
			<input type="button" value=" ���� " onclick="subcheck();"> &nbsp;&nbsp;
			<input type="button" value=" ��� " onclick="history.back();">
		</td>
	</tr>
	<% elseif mode="edit" then
		'// ��� ����
		set oGallery = New cposcode_list
		oGallery.FRectIdx = idx
		oGallery.FArtistMonthItem_one
	%>
	<tr>
		<td width="100" align="center" bgcolor="#F0F0FD" height="30">��ȣ</td>
		<td bgcolor="#FFFFFF"><%=idx%></td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD" height="30">��ǰ</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="itemid" size=10 readonly value="<%=oGallery.FOneItem.fitemid%>">
			<input type="button" class="button" value="ã��" onClick="popItemWindow('inputfrm')">
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">�ڸ�Ʈ</td>
		<td bgcolor="#FFFFFF"><textarea cols="50" rows="5" name="comment" onkeyup="f_chk_byte(this,235)" style="overflow:hidden"><%=oGallery.FOneItem.fcomment%></textarea><br>--���� �������� ������ ��ŭ�� �ѷ����ϴ�.�ʰ�����!!</td>
	</tr>
	<tr>
		<td height="30" align="center" bgcolor="#F0F0FD">����</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="sortNo" size=2 maxlength="1" value="<%=oGallery.FOneItem.FsortNo%>">
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">�������</td>
		<td bgcolor="#FFFFFF">
			<input type="radio" name="isusing" value="Y"<% if oGallery.FOneItem.Fisusing="Y" then Response.Write " checked" %>>Y
			<input type="radio" name="isusing" value="N"<% if oGallery.FOneItem.Fisusing="N" then Response.Write " checked" %>>N
		</td>
	</tr>
	<tr bgcolor="#DDDDFF" >
		<td colspan="2" align="center">
			<input type="button" value=" ���� " onclick="subcheck();"> &nbsp;&nbsp;
			<input type="button" value=" ��� " onclick="history.back();">
		</td>
	</tr>
	<% end if %>
	<%set oGallery = nothing %>
	</form>
</table>
<!-- �׼� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->