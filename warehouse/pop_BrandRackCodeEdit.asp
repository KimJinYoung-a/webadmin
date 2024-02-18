<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/stock/rackipgocls.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->

<%
dim makerid, itembarcode, fcs
makerid     = request("makerid")
itembarcode = request("itembarcode")
fcs         = request("fcs")

dim itemgubun, itemid, itemoption
if (Len(itembarcode) = 12) then
    itemgubun = Mid(itembarcode, 1, 2)
    itemid = Mid(itembarcode, 3, 6)
    itemoption = Mid(itembarcode, 9, 4)
end if


''��ǰ�˻�
dim sqlStr, ItemExists, ItemData

if (Len(request("itembarcode")) <> 12) and (itembarcode<>"") then
    ''���� Barcode�˻�
    sqlStr = "select itemgubun, itemid, itemoption from db_item.dbo.tbl_item_option_stock"
    sqlStr = sqlStr & " where barcode='" & itembarcode & "'"
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        itemgubun   = rsget("itemgubun")
        itemid      = rsget("itemid")
        itemoption  = rsget("itemoption")

    end if
	rsget.close

end if


if (itemgubun="10") then
    sqlStr = " select i.itemid, i.makerid, i.itemname, o.optionname , i.sellcash, i.buycash, i.mwdiv "
    sqlStr = sqlStr + " ,i.isusing , i.sellyn, i.limityn, i.danjongyn, i.limitno, i.limitsold"
    sqlStr = sqlStr + " ,o.optlimityn, o.optlimitno, o.optlimitsold"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
	sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o"
	sqlStr = sqlStr + "     on i.itemid=o.itemid"
	sqlStr = sqlStr + " where i.itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and IsNULL(o.itemoption,'0000')='"+ itemoption + "'"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        ItemExists = True
        ItemData   = rsget.getRows
    end if
	rsget.close
elseif (itemgubun<>"") then
    sqlStr = " select i.shopitemid, i.makerid, i.shopitemname, i.shopitemoptionname, i.shopitemprice, i.shopbuyprice, i.centermwdiv "
    sqlStr = sqlStr + " ,i.isusing , i.isusing as sellyn, 'N' as limityn, 'N' as danjongyn, 0 as limitno, 0 as limitsold"
    sqlStr = sqlStr + " ,'N' as optlimityn, 0 as optlimitno, 0 as optlimitsold"
    sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item i"
    sqlStr = sqlStr + " where shopitemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' and itemgubun = '" + CStr(itemgubun) + "' "

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
		ItemExists = True
		ItemData   = rsget.getRows
    end if
	rsget.close

end if

if (ItemExists) then
    makerid = ItemData(1,0)
end if


dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FRectItemGubun = itemgubun
osummarystock.FRectItemID =  itemid
osummarystock.FRectItemOption =  itemoption
if (ItemExists) then
	osummarystock.GetCurrentItemStock
end if

dim opartner
set opartner = new CPartnerUser
opartner.FRectDesignerID = makerid

if (makerid<>"") then
    opartner.GetOnePartnerNUser
end if

%>
<script language='javascript'>

function checkSubmit(frm){
    //if ((frm.prtidx.value.length!=4)||(!IsDigit(frm.prtidx.value))){
    if ((frm.prtidx.value.length != 4) && (frm.prtidx.value.length != 8)) {
        // alert('�귣�� ���ڵ�� ����4�ڸ��Դϴ�.');
        alert('�귣�� ���ڵ�� 4 or 8 �ڸ� �Դϴ�.');
        frm.prtidx.focus();
        frm.prtidx.select();
        return;
    }

    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

function checkSubmitRackBoxNo(frm){
    if (frm.rackboxno.value.length == 0) {
        alert('������ �Է��ϼ���.');
        frm.rackboxno.focus();
        frm.rackboxno.select();
        return;
    }

    if (frm.rackboxno.value*0 != 0) {
        alert('������ ���ڸ� �����մϴ�.');
        frm.rackboxno.focus();
        frm.rackboxno.select();
        return;
    }

    if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.mode.value = "editrackboxno";
        frm.submit();
    }
}

function popItemStock(itemgubun,itemid,itemoption){
    var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709?itemgubun=" + itemgubun + "&itemid=" + itemid + "&itemoption=" + itemoption,"popitemstocklist","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();

}
</script>


<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<!-- ��ܹ� ���� -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="3">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
				        <font color="red">&nbsp;<strong>�귣�巢�ڵ��Է�</strong></font>
				    </td>
				    <td align="right">
						��ǰ�ڵ�: <input type="text" class="text" name="itembarcode" value="<%= itembarcode %>" size="16" maxlength="32">
						<!--
						&nbsp;
						�귣��ID
						<input type="text" name="makerid" value="<%= makerid %>" size="10" maxlength="32">
						-->
        				<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<!-- ��ܹ� �� -->
	</form>

    <form name="frmSubmit" method="post" action="brandRackCode_process.asp" onSubmit="checkSubmit(this); return false;">
    <input type="hidden" name="makerid" value="<%= makerid %>">
    <input type="hidden" name="itembarcode" value="<%= itembarcode %>">
    <input type="hidden" name="mode" value="editprtidx">

    <% if (ItemExists) and  (FALSE) then %>
    <tr bgcolor="#FFFFFF">
        <td width="100" bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
        <td width="100"><a href="javascript:popItemStock('<%= itemgubun %>','<%= itemid %>','<%= itemoption %>');"><%= (ItemData(0,0)) %></a></td>
        <td></td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="<%= adminColor("tabletop") %>">��ǰ��,�ɼ�</td>
        <td><%= db2html(ItemData(2,0)) %></td>
        <td><%= db2html(ItemData(3,0)) %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="<%= adminColor("tabletop") %>">�Һ��ڰ�,����</td>
        <td><%= FormatNumber(ItemData(4,0),0) %></td>
        <td>
            <%= ChkIIF(ItemData(6,0)="M","<font color='#AA3333'>����</font>","") %>
            <%= ChkIIF(ItemData(6,0)="W","<font color='#3333AA'>Ư��</font>","") %>
            <%= ChkIIF(ItemData(6,0)="U","<font color='#000000'>��ü</font>","") %>

            <% if ItemData(5,0)<>0 then %>
                <%= CLng((ItemData(4,0)-ItemData(5,0))/ItemData(4,0)*100) %> %
            <% end if %>
        </td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="<%= adminColor("tabletop") %>">�Ǹű���</td>
        <td>
            <% if (itemoption="0000") then %>
            <!-- �ɼ� ���� -->
                <% if (((ItemData(8,0)="N") or (ItemData(9,0)="N")) or ((ItemData(10,0)="Y") and (ItemData(12,0)-ItemData(13,0)<1))) then %>
                <font color=red>ǰ��</font>
                <% end if %>

                <% if (ItemData(10,0)="Y") then %>
                    <% if ItemData(12,0)-ItemData(13,0)<1 then %>
                    <font color="blue">����(0)</font>
                    <% else %>
                    <font color="blue">����(<%= ItemData(12,0)-ItemData(13,0) %>)</font>
                    <% end if %>
                <% end if %>


            <% else %>
                <% if (((ItemData(8,0)="N") or (ItemData(9,0)="N")) or ((ItemData(14,0)="Y") and (ItemData(15,0)-ItemData(16,0)<1))) then %>
                <font color=red>ǰ��</font>
                <% end if %>

                <% if (ItemData(14,0)="Y") then %>
                    <% if ItemData(15,0)-ItemData(16,0)<1 then %>
                    <font color="blue">����(0)</font>
                    <% else %>
                    <font color="blue">����(<%= ItemData(15,0)-ItemData(16,0) %>)</font>
                    <% end if %>
                <% end if %>
            <% end if %>
        </td>
        <td>
                <%= chkIIf (ItemData(7,0)="Y","<font color='red'>���X</font>","")  %>
                <!-- ���� ���� -->
                <%= chkIIf (ItemData(11,0)="Y","<font color='blue'>����</font>","")  %>
                <%= chkIIf (ItemData(11,0)="S","<font color='green'>�Ͻ�ǰ��</font>","")  %>
        </td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="<%= adminColor("tabletop") %>">�����</td>
        <td><%= osummarystock.FOneItem.GetCheckStockNo %></td>
        <td></td>
    </tr>
    <% end if %>

    <% if (opartner.FResultCount>0) then %>
    <tr bgcolor="#FFFFFF">
        <td width="100" bgcolor="<%= adminColor("tabletop") %>">�귣��ID</td>
        <td width="250"><%= opartner.FOneItem.Fid %></td>
        <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="<%= adminColor("tabletop") %>">�귣���</td>
        <td><%= opartner.FOneItem.FSocName_Kor %></td>
        <td></td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="<%= adminColor("tabletop") %>">��뿩��</td>
        <td><%= opartner.FOneItem.Fisusing %></td>
        <td></td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="<%= adminColor("tabletop") %>">���ڵ�</td>
        <td>
        	<input type="text" class="text" name="prtidx" value="<%= opartner.FOneItem.Fprtidx %>" size="8" maxlength="8">
        	<input type="button" class="button" value="���ڵ� ����" onClick="checkSubmit(frmSubmit);">
            (4 or 8�ڸ� Fix)
        </td>
        <td></td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="<%= adminColor("tabletop") %>">���ڽ�����</td>
        <td>
        	<input type="text" class="text" name="rackboxno" value="<%= opartner.FOneItem.Frackboxno %>" size="4" maxlength="4">
        	<input type="button" class="button" value="���ڽ����� ����" onClick="checkSubmitRackBoxNo(frmSubmit);">
        </td>
        <td></td>
    </tr>
    <% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="3" align="center">[�˻� ����� �����ϴ�.]</td>
    </tr>
    <% end if %>
    </form>
</table>
<script language='javascript'>
function GetOnLoad(){
    <% if (ItemExists) and (opartner.FResultCount>0) and (fcs<>"itembarcode") then %>
    document.frmSubmit.prtidx.focus();
    document.frmSubmit.prtidx.select();
    <% else %>
    document.frm.itembarcode.focus();
    document.frm.itembarcode.select();
    <% end if %>
}
window.onload=GetOnLoad;
</script>
<%
set opartner = Nothing
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
