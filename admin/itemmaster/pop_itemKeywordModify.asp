<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim itemid
itemid = getNumeric(requestCheckVar(request("itemid"),10))

if itemid = "" then
	response.write "<script>"
	response.write "	alert('��ǰ�ڵ尡 �����ϴ�');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
end if

If IsNumeric(itemid) = false Then
	response.write "<script>"
	response.write "	alert('�߸��� ��ǰ�ڵ��Դϴ�');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
End IF

'==============================================================================
dim oitem

set oitem = new CItem

oitem.FPageSize         = 1
oitem.FCurrPage         = 1
oitem.FRectItemid       = itemid

oitem.GetItemKeywordList

If oitem.FresultCount < 1 Then
	response.write "<script>"
	response.write "	alert('�߸��� ��ǰ�ڵ��Դϴ�');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
End IF

%>

<script language='javascript'>

function SaveItem(frm) {
	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if(ret) {
		frm.submit();
	}
}

function CloseWindow() {
    window.close();
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" Maxlength="9" size="9">
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->

<% if oitem.FResultCount>0 then %>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<form name=frm2 method=post action="itemKeyword_process.asp">
<input type=hidden name=mode value="editone">
<input type=hidden name=itemid value="<%= itemid %>">
<tr>
	<td colspan="2" bgcolor="#FFFFFF">
		<table width="100%" cellspacing=1 cellpadding=1 border="0" class=a bgcolor=#BABABA>
			<tr height="25">
				<td width="120" bgcolor="#DDDDFF">��ǰ��</td>
				<td bgcolor="#FFFFFF">
					<%= oitem.FItemList(0).Fitemname %>
				</td>
			</tr>
			<tr height="25">
				<td width="120" bgcolor="#DDDDFF">Ű����</td>
				<td bgcolor="#FFFFFF">
					<input type="text" class="text" name="keywords" value="<%= oitem.FItemList(0).Fkeywords %>" size="125" maxlength="128" >
				</td>
			</tr>
		</table>
	</td>
</tr>
</form>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center" bgcolor="F4F4F4">
			<input type="button" class="button" value="�����ϱ�" onclick="SaveItem(frm2)">
			&nbsp;
			<input type="button" class="button" value=" �� �� " onclick="CloseWindow()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<% else %>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<tr bgcolor="#FFFFFF">
    <td align="center">[�˻� ����� �����ϴ�.]</td>
</tr>
</table>
<% end if %>
