<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYItem_imgContentCls.asp" -->
<%
Const CMaxInfoImageCnt = 4
Dim itemid,mode, i
dim oitem
Dim oitemadd

mode = RequestCheckvar(request("mode"),16)
itemid = RequestCheckvar(request("itemid"),10)

set oitem = new CItem
oItem.FRectItemID = itemid

if (itemid<>"") then
    oItem.GetOneItem
end if

set oitemadd = new CInfoImage
oitemadd.getOneInfoImageList itemid


'response.write Ubound(simginfo)
%>
<script language="JavaScript">
<!--

function checkok(frm){
	if (!IsDigit(frm.itemid.value)){
		alert('��ǰ ��ȣ�� ���ڸ� �����մϴ�.');
		frm.itemid.focus();
		return;
	}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
      frm.submit();
    }
}

function DelImage(src, imgname, hid){
	var imgcomp;
	imgcomp = eval("document." + imgname);
	if (hid.checked){
		imgcomp.src = '/images/space.gif';
	}
}

function gopage(itemid){
	document.location="?itemid=" + itemid + "&mode=<%= mode %>&menupos=<%= menupos %>";
}
//-->
</script>
<br>
<div align=center>
<font size=4 color=red><b>�̹��� �������� ���� 300kb</b></font>
</div>
<br>
<form method="post" name="monthly" action="<%= imgFingers %>/linkweb/items/ItemImgContents_process.asp"  enctype="MULTIPART/FORM-DATA" onSubmit="return false;">
<input type="hidden" name="mode" value="<% = mode %>">
<table cellpadding="0" cellspacing="0" border="1" align="center" bordercolordark="White" bordercolorlight="black" class="a">
<tr>
  <td width="100">��ǰ��ȣ</td>
  <td align="center" class="a"><input type="text" name="itemid" value="<%= itemid %>">
  	<input type="button" value="����" onclick="gopage(monthly.itemid.value);" class="button">
  </td>
</tr>
<% if oItem.FResultCount>0 then %>
<tr>
	<td>��ǰ��</td>
	<td><%= oItem.FOneItem.FItemName %></td>
</tr>
<tr>
	<td>����Ʈ �̹���</td>
	<td><img src="<%= oItem.FOneItem.Flistimage %>" width="100"></td>
</tr>

<% for i=0 to CMaxInfoImageCnt -1  %>
<tr>
  <td>�̹���<%= i+1 %></td>
  <td>
	  <table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
		  <td><input type="file" name="img<%= i+1 %>" size="30"></td>
		</tr>
		<tr>
		  <td>
		  <% if (oitemadd.FResultCount>i) then %>
		  <%= oitemadd.FItemList(i).FADDIMAGE %>
		  	<% if oitemadd.FItemList(i).FADDIMAGE<>"" then %>
		  	<input type="checkbox" name="dl_img<%= i+1 %>">����
		  	<% end if %>
		  <% end if %>
		  </td>
		</tr>
	  </table>
  </td>
</tr>
<% next %>
<tr>
	<td align="center" colspan="2" height="30"><input type="button" value="�̹����ֱ�" onclick="checkok(this.form);" class="button">&nbsp;&nbsp;&nbsp;</td>
</tr>
<% else %>
<tr height="40">
    <td align="center" colspan="2">[���� ��ǰ�� �˻� �ϼ���.]</td>
</tr>
<% end if %>
</table>
</form>
<%
set oitem = Nothing
set oitemadd = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->