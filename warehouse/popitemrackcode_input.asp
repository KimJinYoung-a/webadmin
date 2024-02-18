<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/items/itembarcodecls.asp"-->
<%

dim referer
referer = request.ServerVariables("HTTP_REFERER")

dim itemrackcode, itembarcode
itemrackcode = request("itemrackcode")
itembarcode = request("itembarcode")


dim itemgubun, itemid, itemoption

dim sqlStr
if Len(itembarcode)>=12 then
        sqlStr = "select top 1 b.* " + VbCrlf
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option_stock b " + VbCrlf
        sqlStr = sqlStr + " where b.barcode='" + CStr(itembarcode) + "' " + VbCrlf

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
        	itemgubun = rsget("itemgubun")
        	itemid = rsget("itemid")
        	itemoption = rsget("itemoption")
        else
        	itemgubun = Left(itembarcode,2)
        	itemid = CLng(Mid(itembarcode,3,6))
        	itemoption = Right(itembarcode,4)
        end if
        rsget.Close
else
	if Len(itembarcode)=12 then
		itemgubun 	= left(itembarcode,2)
		itemid		= mid(itembarcode,3,6)
		itemoption	= right(itembarcode,4)
	else
		itemgubun = "10"
		itemid = itembarcode
	end if
end if

dim oitembar
set oitembar = new CItemBarCode
oitembar.FRectItemGubun = itemgubun
oitembar.FRectItemID = itemid

if itemid<>"" then
	oitembar.getItemBarcodeInfo
end if
%>
<script language='javascript'>
function SaveRackcode(frm){
	if (frm.itemrackcode.value.length!=4){
		alert('��ǰ ���ڵ� 4�ڸ��� �Է��ϼ���.');
		frm.itemrackcode.focus();
		return;
	}

	if (frm.itemid.value.length<1){
		alert('��ǰ�ڵ尡 �ùٸ��� �ʽ��ϴ�.');
		return;
	}

	if (confirm('��ǰ ���ڵ带 �����Ͻðڽ��ϱ�?')){
		frm.method="post";
		frm.action="itemrackcode_process.asp"
		frm.submit();
	}
}

function research(frm){
	frm.submit();
}

function GetOnLoad(){
	if (document.frmrackcodeinput.itemrackcode.value.length==4){
		document.frmrackcodeinput.itembarcode.select();
		document.frmrackcodeinput.itembarcode.focus();
	}else{
		document.frmrackcodeinput.itemrackcode.select();
		document.frmrackcodeinput.itemrackcode.focus();
	}
}

window.onload=GetOnLoad;
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmrackcodeinput" method="get" >
	<input type="hidden" name="mode" value="ByRackCodeProc">
	<% if oitembar.FResultCount>0 then %>
	<input type="hidden" name="itemid" value="<%= oitembar.FItemList(0).Fitemid %>">
	<% else %>
	<input type="hidden" name="itemid" value="">
	<% end if %>
	<!-- ��ܹ� ���� -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="3">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
				        <font color="red"><strong>���ڵ庰 ��ǰ�Է�</strong></font>
				    </td>
				    <td align="right">

					</td>
				</tr>
			</table>
		</td>
	</tr>
	<!-- ��ܹ� �� -->
	<tr bgcolor="#FFFFFF">
        <td width="60" align="center" bgcolor="<%= adminColor("tabletop") %>">��ǰ���ڵ�</td>
        <td><input type="text" class="text" name="itemrackcode" value="<%= itemrackcode %>" size=4 maxlength=4 AUTOCOMPLETE="off" ></td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td align="center" bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
        <td><input type="text" class="text" name="itembarcode" value="<%= itembarcode %>" size=14 maxlength=14 AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ research(frmrackcodeinput); return false;}">&nbsp;<input type=button value="����" onclick="research(frmrackcodeinput)" ></td>
    </tr>
    <% if oitembar.FResultCount>0 then %>
    <tr bgcolor="#FFFFFF">
        <td align="center" bgcolor="<%= adminColor("tabletop") %>">�귣��</td>
        <td><%= oitembar.FItemList(0).FMakerid %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td align="center" bgcolor="<%= adminColor("tabletop") %>">��ǰID</td>
        <td><%= CHkIIF(oitembar.FItemList(0).FItemGubun="10","�¶���","<font color=blue>����</font>") %>:<%= oitembar.FItemList(0).FItemID %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td align="center" bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
        <td><%= oitembar.FItemList(0).FItemName %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td align="center" bgcolor="<%= adminColor("tabletop") %>">��ǰ����</td>
        <td><%= FormatNumber(oitembar.FItemList(0).FSellcash,0) %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td align="center" bgcolor="<%= adminColor("tabletop") %>">�̹���</td>
        <td><img src="<%= oitembar.FItemList(0).FImageList %>" width="100" height="100"></td>
    </tr>
	<% end if %>



  	</tr>
  	</form>
</table>
<% if (oitembar.FResultCount>0) and (itemid<>"") and (referer<>"") then %>
<script>SaveRackcode(frmrackcodeinput);</script>
<% end if %>
<%
set oitembar = nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->