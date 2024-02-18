<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->
<%
dim shopid, makerid, centermwdiv, itembarcode, usingyn, research, NoZeroStock
dim itemgubun, itemid, itemoption
dim ImgUsing, pagesize

shopid       = RequestCheckVar(request("shopid"),32)
makerid      = RequestCheckVar(request("makerid"),32)
centermwdiv  = RequestCheckVar(request("centermwdiv"),10)
itembarcode  = RequestCheckVar(request("itembarcode"),32)
usingyn      = RequestCheckVar(request("usingyn"),1)
research     = RequestCheckVar(request("research"),2)
ImgUsing     = RequestCheckVar(request("ImgUsing"),1)
NoZeroStock  = RequestCheckVar(request("NoZeroStock"),32)
pagesize  		= RequestCheckVar(request("pagesize"),32)

if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if

if (itembarcode <> "") then
    if Not (fnGetItemCodeByPublicBarcode(itembarcode,itemgubun,itemid,itemoption)) then
        itemgubun   = Left(itembarcode, 2)
        itemid      = CStr(Mid(itembarcode, 3, 6) + 0)
        itemoption  = Right(itembarcode, 4)
    end if
end if

'''if (research="") and (usingyn="") then usingyn="Y"
if (research="") and (ImgUsing="") then ImgUsing="Y"
if (pagesize = "") then pagesize = "100"


dim oOffStock
set oOffStock = new CShopItemSummary
oOffStock.FCurrPage 		= 1
oOffStock.FPageSize 		= pagesize
oOffStock.FRectShopID       = shopid
oOffStock.FRectMakerID      = makerid
oOffStock.FRectCenterMwDiv  = centermwdiv
oOffStock.FRectIsUsing      = usingyn
oOffStock.FRectNoZeroStock  = NoZeroStock
if (itembarcode <> "") then
    oOffStock.FRectItemGubun    = itemgubun
    oOffStock.FRectItemId       = itemid
    oOffStock.FRectItemOption   = itemoption
end if

if ((shopid<>"") and (makerid<>"")) or ((shopid<>"") and (itembarcode<>"")) then
    oOffStock.GetShopItemCurrentSummaryList
end if

dim i
dim totsysstock, totavailstock, totrealstock
%>
<script language='javascript'>
function RefreshPageByImg(ImgUsing){
    document.frm.ImgUsing.value = ImgUsing;
    document.frm.submit();
}
function RefreshPage(){
    document.frm.submit();
}
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="ImgUsing" value="<%=ImgUsing%>">

	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
		    <% if (C_IS_SHOP) then %>
		    <input type="hidden" name="shopid" value="shopid">
            ���� : <%= shopid %>
            <% else %>
		    ���� : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
		    <% end if %>
			�귣�� :
			<% drawSelectBoxDesignerwithName "makerid", makerid %> &nbsp;&nbsp;
			&nbsp;
			ǥ�ð��� :
			<select class="select" name="pagesize">
				<option value="100" <%= CHKIIF(pagesize = "100", "selected", "") %>>100</option>
				<option value="500" <%= CHKIIF(pagesize = "500", "selected", "") %>>500</option>
				<option value="1000" <%= CHKIIF(pagesize = "1000", "selected", "") %>>1000</option>
				<option value="2000" <%= CHKIIF(pagesize = "2000", "selected", "") %>>2000</option>
			</select>
		</td>

		<td rowspan="2" width="220" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button" name="brandstockprint" value="��˻�" onclick="RefreshPage();">
			<% if ImgUsing="Y" then %>
        		<input type="button" class="button" name="brandstockprint" value="�̹������ֱ�" onclick="RefreshPageByImg('N');">
        	<% else %>
        		<input type="button" class="button" name="brandstockprint" value="�̹������̱�" onclick="RefreshPageByImg('Y');">
        	<% end if %>
        	<input type="button" class="button" name="brandstockprint" value="����ϱ�" onclick="window.print();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			��ǰ ��뱸�� : <% drawSelectBoxUsingYN "usingyn", usingyn %> &nbsp;&nbsp;
			<input type="checkbox" name="NoZeroStock" <%= CHKIIF(NoZeroStock="on","checked","") %> > ���0�� ��ǰ �˻� ����.
			<!--
			���͸��Ա��� :
			   <select class="select" name="centermwdiv">
               <option value="">��ü</option>
               <option value="MW" <%= ChkIIF(centermwdiv="MW","selected","") %> >����+��Ź</option>
               <option value="W"  <%= ChkIIF(centermwdiv="W","selected","") %> >��Ź</option>
               <option value="M"  <%= ChkIIF(centermwdiv="M","selected","") %> >����</option>
               <option value="NULL" <%= ChkIIF(centermwdiv="NULL","selected","") %> >������</option>
               </select>
            &nbsp;&nbsp;
            -->
            [����� : <%= now() %>]
		</td>
	</tr>

	</form>
</table>
<!-- �˻� �� -->
<p>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="30">����</td>
    	<td width="40">��ǰID</td>
    	<td width="40">�ɼ�</td>
    	<% if (ImgUsing="Y") then %>
    	<td width="50">�̹���</td>
    	<% end if %>
    	<td>��ǰ��<br>[�ɼǸ�]</td>
    	<!-- td width="40">����<br>����<br>����</td -->
    	<td width="40">����<br>�԰�</td>
    	<td width="40">����<br>��ǰ</td>
    	<td width="40">�귣��<br>�԰�</td>
    	<td width="40">�귣��<br>��ǰ</td>
        <td width="40">����<br>�Ǹ�</td>
        <td width="40">����<br>��ǰ</td>
        <td width="40" bgcolor="F4F4F4">�ý���<br>�����</td>
        <td width="40">��<br>�ǻ�<br>����</td>
        <td width="40" bgcolor="F4F4F4">�ǻ�<br>���</td>
		<td width="40">�����</td>
		<td width="40">��ǰ��</td>
		<td width="40" bgcolor="F4F4F4">����<br>���<br>(����)</td>
        <!-- <td width="40">��<br>����</td>
        <td width="40">��<br>�ҷ�</td>-->
        <!-- <td width="40" bgcolor="F4F4F4">��ȿ<br>���</td-->

        <td width="30">���<br>����</td>
        <td width="100">���</td>
    </tr>
<% if oOffStock.FResultCount<1 then %>
    <tr align="center" bgcolor="#FFFFFF" height="30">
        <% if (shopid="") and (makerid="") then %>
        <td colspan="20" >[ ���� �� �귣�带 ���� �ϼ���. ]</td>
        <% else %>
        <td colspan="20" >[ �˻� ����� �����ϴ�. ]</td>
        <% end if %>
    </tr>
<% else %>
    <% for i=0 to oOffStock.FResultCount - 1 %>
    <%
    totsysstock	    = totsysstock + oOffStock.FItemList(i).FsysstockNo
    totavailstock   = totavailstock + oOffStock.FItemList(i).getAvailStock
    totrealstock    = totrealstock + oOffStock.FItemList(i).FrealstockNo

    %>
    	<% if oOffStock.FItemList(i).Fisusing="Y" then %>
        <tr bgcolor="#FFFFFF" align="center" >
        <% else %>
        <tr bgcolor="#FFFFFF" align="center" >
        <% end if %>
            <td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).FItemGubun %></td>
        	<td style="border-bottom:1px solid black">
        	    <%= oOffStock.FItemList(i).Fitemid %>
        	</td>
        	<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).FItemOption %></td>
        	<% if (ImgUsing="Y") then %>
        	<td style="border-bottom:1px solid black"><img src="<%= oOffStock.FItemList(i).GetImageSmall %>" width=50 height=50> </td>
        	<% end if %>
        	<td align="left" style="border-bottom:1px solid black">
              	<%= oOffStock.FItemList(i).FShopitemname %>
              	<% if oOffStock.FItemList(i).FShopitemoptionName <>"" then %>
              		<br>
              		<font color="blue">[<%= oOffStock.FItemList(i).FShopitemoptionName %>]</font>
              	<% end if %>
            </td>
            <!-- td><%= fnColor(oOffStock.FItemList(i).FCenterMwdiv,"mw") %></td -->
        	<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Flogicsipgono %></td>
        	<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Flogicsreipgono %></td>
        	<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Fbrandipgono %></td>
        	<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Fbrandreipgono %></td>
        	<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Fsellno %></td>
        	<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Fresellno %></td>
        	<td bgcolor="F4F4F4" style="border-bottom:1px solid black"><b><%= oOffStock.FItemList(i).FsysstockNo %></b></td>
        	<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Ferrrealcheckno %></td>
        	<td bgcolor="F4F4F4" style="border-bottom:1px solid black"><b><%= oOffStock.FItemList(i).FrealstockNo %></b></td>
			<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Flogischulgo %></td>
			<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Flogisreturn %></td>
			<td bgcolor="F4F4F4" style="border-bottom:1px solid black"><b><%= oOffStock.FItemList(i).getShopRealStockNoExc %></b></td>
        	<!-- td><%= oOffStock.FItemList(i).Ferrsampleitemno %></td>
        	<td><%= oOffStock.FItemList(i).Ferrbaditemno %></td> -->
        	<!-- td bgcolor="F4F4F4"><b><%= oOffStock.FItemList(i).getAvailStock %></b></td -->

        	<td style="border-bottom:1px solid black">
        	    <% if oOffStock.FItemList(i).Fisusing="N" then %>
        	    <strong><%= oOffStock.FItemList(i).Fisusing %></strong>
        	    <% else %>
        	    <%= oOffStock.FItemList(i).Fisusing %>
        	    <% end if %>
        	</td>
        	<td valign="top" style="border-bottom:1px solid black">
        	<% if (oOffStock.FItemList(i).Ferrsampleitemno<>0) then %>
        	(���� <%= oOffStock.FItemList(i).Ferrsampleitemno*-1 %>)
        	<% else %>
        	&nbsp;
        	<% end if %>
        	</td>

        </tr>
    <% next %>
<% end if %>
</table>

<%
set oOffStock = Nothing
%>






<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
