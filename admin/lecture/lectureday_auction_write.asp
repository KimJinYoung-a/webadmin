<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/lectureday_auctioncls.asp"-->
<%
dim mode,idx

mode = request("mode")
idx = request("idx")

dim editAuction
set editAuction = New CBoardAuction
if idx="" then idx =0
editAuction.GetOneAuction idx

dim i
%>
<script language="javascript">
function AddAuction(frm){
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];

		if ((e.type=="text")) {
			if ((e.name=="itemid") || (e.name=="auctionname") || (e.name=="limitno") || (e.name=="startdate")
				|| (e.name=="finishdate") || (e.name=="supplyer")|| (e.name=="pricestart")|| (e.name=="priceend")
				|| (e.name=="pricefix") ){
				if (e.value.length<1){
					alert('�ʼ� �Է� �����Դϴ�.');
					e.focus();
					return;
				}
			}

			if ((e.name=="itemid") || (e.name=="limitno") || (e.name=="pricestart") || (e.name=="priceend") || (e.name=="pricefix")){
				if (!IsDigit(e.value)){
					alert('���ڸ� �����մϴ�.');
					e.focus();
					return;
				}
			}
		}
	}
	<% if mode="add" then %>
	var ret = confirm('�߰� �Ͻðڽ��ϱ�?');
	<% else %>
	var ret = confirm('���� �Ͻðڽ��ϱ�?');
	<% end if %>
	if (ret) { frm.submit();}
}
</script>
<form name="addfrm" method="post" action="http://partner.10x10.co.kr/admin/lectureday/donewauction.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="idx" value="<% = idx %>">
<table width="600" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
	<td width="100">��ȣ</td>
	<td><%= editAuction.FAuctionList(0).Fidx %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>����̸�</td>
	<td><input type="text" name="auctionname" value="<%= editAuction.FAuctionList(0).Fauctionname %>" size="70" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>������<br>(2005-02-05)</td>
	<% if mode="add" then %>
	<td><input type="text" name="startdate" value="2005-02-05" size="70" class="input_b"></td>
	<% else %>
	<td><input type="text" name="startdate" value="<%= editAuction.FAuctionList(0).Fstartdate %>" size="70" class="input_b"></td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td>������<br>(2002-10-05 23:00:00)</td>
	<% if mode="add" then %>
	<td><input type="text" name="finishdate" value="2005-02-05 00:00:00" size="70" class="input_b"></td>
	<% else %>
	<td><input type="text" name="finishdate" value="<%= editAuction.FAuctionList(0).Ffinishdate %>" size="70" class="input_b"></td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td>�ǸŰ���</td>
	<td><input type="text" name="itemea" value="<%= editAuction.FAuctionList(0).Fitemea %>" size="30" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<% if mode="add" then %>
	<td>�����̹���</td>
	<td><input type="file" name="mainimg" size="50" class="input_b">(����:550����)
	</td>
	<% else %>
	<td>�����̹���</td>
	<td><input type="file" name="mainimg" size="50" class="input_b">(����:550����)<br>
		<input type="checkbox" name="dl_mainimg">���� (<%= editAuction.FAuctionList(0).Fmainimg %>)
	</td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<% if mode="add" then %>
	<td>�߰��̹���1</td>
	<td><input type="file" name="img1" size="50" class="input_b">(����:550����)
	</td>
	<% else %>
	<td>�߰��̹���1</td>
	<td><input type="file" name="img1" size="50" class="input_b">(����:550����)<br>
		<input type="checkbox" name="dl_img1">���� (<%= editAuction.FAuctionList(0).Fimg1 %>)
	</td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<% if mode="add" then %>
	<td>�߰��̹���2</td>
	<td><input type="file" name="img2" size="50" class="input_b">(����:550����)
	</td>
	<% else %>
	<td>�߰��̹���2</td>
	<td><input type="file" name="img2" size="50" class="input_b">(����:550����)<br>
		<input type="checkbox" name="dl_img2">���� (<%= editAuction.FAuctionList(0).Fimg2 %>)
	</td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<% if mode="add" then %>
	<td>�߰��̹���3</td>
	<td><input type="file" name="img3" size="50" class="input_b">(����:550����)
	</td>
	<% else %>
	<td>�߰��̹���3</td>
	<td><input type="file" name="img3" size="50" class="input_b">(����:550����)<br>
		<input type="checkbox" name="dl_img3">���� (<%= editAuction.FAuctionList(0).Fimg3 %>)
	</td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<% if mode="add" then %>
	<td>�߰��̹���4</td>
	<td><input type="file" name="img4" size="50" class="input_b">(����:550����)
	</td>
	<% else %>
	<td>�߰��̹���4</td>
	<td><input type="file" name="img4" size="50" class="input_b">(����:550����)<br>
		<input type="checkbox" name="dl_img4">���� (<%= editAuction.FAuctionList(0).Fimg4 %>)
	</td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<% if mode="add" then %>
	<td>�߰��̹���5</td>
	<td><input type="file" name="img5" size="50" class="input_b">(����:550����)
	</td>
	<% else %>
	<td>�߰��̹���5</td>
	<td><input type="file" name="img5" size="50" class="input_b">(����:550����)<br>
		<input type="checkbox" name="dl_img5">���� (<%= editAuction.FAuctionList(0).Fimg5 %>)
	</td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<% if mode="add" then %>
	<td>����Ʈ�̹���50</td>
	<td><input type="file" name="icon1" size="50" class="input_b">(50*50)
	</td>
	<% else %>
	<td>����Ʈ�̹���50</td>
	<td><input type="file" name="icon1" size="50" class="input_b">(50*50)<br>
		<input type="checkbox" name="dl_icon1">���� (<%= editAuction.FAuctionList(0).Ficon1 %>)
	</td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<% if mode="add" then %>
	<td>����Ʈ�̹���100</td>
	<td><input type="file" name="icon2" size="50" class="input_b">(80*80)
	</td>
	<% else %>
	<td>����Ʈ�̹���100</td>
	<td><input type="file" name="icon2" size="50" class="input_b">(80*80)<br>
		<input type="checkbox" name="dl_icon2">���� (<%= editAuction.FAuctionList(0).Ficon2 %>)
	</td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td>��ǰ����</td>
	<% if mode="add" then %>
	<td><textarea name="itemcontents" rows="10" cols="70" class="input_b"></textarea></td>
	<% else %>
	<td><textarea name="itemcontents" rows="10" cols="70" class="input_b"><%= editAuction.FAuctionList(0).Fitemcontents %></textarea></td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td>���ǻ���</td>
	<% if mode="add" then %>
	<td><textarea name="etc" rows="10" cols="70" class="input_b"></textarea></td>
	<% else %>
	<td><textarea name="etc" rows="10" cols="70" class="input_b"><%= editAuction.FAuctionList(0).Fetc %></textarea></td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td>��žȳ�</td>
	<% if mode="add" then %>
	<td><textarea name="info" rows="10" cols="70" class="input_b"></textarea></td>
	<% else %>
	<td><textarea name="info" rows="10" cols="70" class="input_b"><%= editAuction.FAuctionList(0).Finfo %></textarea></td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td>���۰�</td>
	<% if mode="add" then %>
	<td><input type="text" name="startprice" value="0" size="30" class="input_b"></td>
	<% else %>
	<td><input type="text" name="startprice" value="<%= editAuction.FAuctionList(0).Fstartprice %>" size="30" class="input_b"></td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td>��÷��</td>
	<td><input type="text" name="nakchaluser" value="<%= editAuction.FAuctionList(0).Fnakchaluser %>" size="30" class="input_b">
	<font color=red>(���̵� �ڿ� ���� ���� �Ұ�!)</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>��÷��</td>
	<% if mode="add" then %>
	<td><input type="text" name="nakchalprice" value="0" size="30" class="input_b"></td>
	<% else %>
	<td><input type="text" name="nakchalprice" value="<%= editAuction.FAuctionList(0).Fnakchalprice %>" size="30" class="input_b"></td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td>��뿩��</td>
	<td>
		<input type="radio" name="isusing" value="Y" <% if editAuction.FAuctionList(0).FIsUsing="Y" then response.write "checked" %> >Y
		<input type="radio" name="isusing" value="N" <% if editAuction.FAuctionList(0).FIsUsing<>"Y" then response.write "checked" %> >N
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center">
		<% if mode="add" then %>
		<input type="button" value="�߰�" onClick="AddAuction(addfrm)">
		<% elseif  mode="edit" then %>
		<input type="button" value="����" onClick="AddAuction(addfrm)">
		<% end if %>
	</td>
</tr>
</table>
</form>

<%
set editAuction = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->