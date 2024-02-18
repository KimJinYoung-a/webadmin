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
'' �ǻ� ��� �Է�
'' ���� ���� ������ ��� �ڱ� ���常, �귣���ΰ�� ��üƯ���� ����. => Combo, Text ReadOnly
dim NowDate
NowDate = Left(CStr(Now()),10)

dim itemgubun, itemid, itemoption, shopid, itembarcode, makerid
itemgubun  = requestCheckVar(request("itemgubun"),2)
itemid     = requestCheckVar(request("itemid"),9)
itemoption = requestCheckVar(request("itemoption"),4)
itembarcode= requestCheckVar(request("itembarcode"),32)
shopid     = requestCheckVar(request("shopid"),32)

''����
if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if

''��ü
if (C_IS_Maker_Upche) then
    makerid = session("ssBctid")
end if

if (itembarcode <> "") then
    if Not (fnGetItemCodeByPublicBarcode(itembarcode,itemgubun,itemid,itemoption)) then
        if (Len(itembarcode)=12) then
            itemgubun   = Left(itembarcode, 2)
            itemid      = CStr(Mid(itembarcode, 3, 6) + 0)
            itemoption  = Right(itembarcode, 4)
        elseif (Len(itembarcode)=14) then
            itemgubun   = Left(itembarcode, 2)
            itemid      = CStr(Mid(itembarcode, 3, 8) + 0)
            itemoption  = Right(itembarcode, 4)
        end if
    end if
elseif (itemid<>"") then
    if (itemid>=1000000) then
        itembarcode = itemgubun + "" + Format00(8,itemid) + "" + itemoption
    else
        itembarcode = itemgubun + "" + Format00(6,itemid) + "" + itemoption
    end if
end if


'==============================================================================
'��ǰ�⺻����
if itemgubun="" then itemgubun="10"
if itemoption="" then itemoption="0000"

dim ojaegoitem
set ojaegoitem = new COffShopItem
ojaegoitem.FRectItemGubun   = itemgubun
ojaegoitem.FRectItemID      = itemid
ojaegoitem.FRectItemOption  = itemoption
ojaegoitem.FRectShopid      = shopid
if (itemid<>"") then
	ojaegoitem.GetOffOneItem
end if

if (ojaegoitem.FREsultCount<1) then
    response.write "<script>alert('��ǰ�� ���� ���� �ʽ��ϴ�.'); window.close();</script>"
    dbget.Close() : response.end
end if

'==============================================================================
'��ǰ�������(current)
dim ocursummary
set ocursummary = new CShopItemSummary

ocursummary.FRectShopID =  shopid
ocursummary.FRectItemGubun =  itemgubun
ocursummary.FRectItemId =  itemid
ocursummary.FRectItemOption =  itemoption

if itemid<>"" then
	ocursummary.GetShopItemCurrentSummary
end if

dim IsUpcheWitakItem
IsUpcheWitakItem = (ojaegoitem.FOneItem.Fcomm_cd="B012")

%>
<script language='javascript'>
function RecalcuErr(){
	var checkstock = calcufrm.checkstock.value;  // �������.

	//calcufrm.todayerrrealcheckno.value = checkstock-calcufrm.orgrealstock.value ;
	calcufrm.errrealcheckno.value = checkstock - calcufrm.shoprealstock.value ;
}
function SaveSample(){
    <% if (C_IS_Maker_Upche) and (Not IsUpcheWitakItem) then %>
    alert('������ �����ϴ�. - ��üƯ�� ��ǰ�� ��� ���� ����.');
    return;
    <% else %>

    var samplestock = calcufrm.samplestock.value;
	if (!IsInteger(samplestock)){
		alert('���ڸ� �Է��ϼ���.');
		calcufrm.samplestock.focus();
		return;
	}

	if (samplestock*1<0){
	    alert('���� ���� (+)����� �Է��ϼ���.');
		calcufrm.samplestock.focus();
		return;
	}

	if (confirm('���� ��� �����Ͻðڽ��ϱ�?')){
		frmrefresh.mode.value ="OffSampleCheckupdate";
		frmrefresh.samplestock.value = samplestock;
		frmrefresh.submit();
	}
	<% end if %>
}

function SaveErr(){
    <% if (C_IS_Maker_Upche) and (Not IsUpcheWitakItem) then %>
    alert('������ �����ϴ�. - ��üƯ�� ��ǰ�� ��� ���� ����.');
    return;
    <% else %>
	var realstock = calcufrm.checkstock.value - calcufrm.logischulgo.value - calcufrm.logisreturn.value - calcufrm.errsampleitemno.value;
	if (!IsInteger(realstock)){
		alert('���ڸ� �Է��ϼ���.');
		calcufrm.checkstock.focus();
		return;
	}

    //����ľ���
    var nowdate = "<%= NowDate %>";
    var today = new Date();
    var stdate = new Date(calcufrm.stockdate.value.substr(0,4),calcufrm.stockdate.value.substr(5,2)*1-1,calcufrm.stockdate.value.substr(8,2));
    var theBaseDate = new Date(nowdate.substr(0,4),nowdate.substr(5,2)*1-1-1,1);


    if (stdate<theBaseDate){
        alert('����� �ľ����� ' + theBaseDate.toLocaleString().substr(0,11)+ ' ������ ���� �� �� �����ϴ�.');
        return;
    }

    if (stdate>today){
        alert('����� �ľ����� ���� ���ķ� ���� �� �� �����ϴ�.');
        return;
    }

    //alert(today.toLocaleString());
    //alert(stdate.toLocaleString());

	if (confirm('�ǻ������ �����Ͻðڽ��ϱ�?')){
		frmrefresh.mode.value ="Offerrcheckupdate";
		frmrefresh.realstock.value = realstock;
		frmrefresh.stockdate.value = calcufrm.stockdate.value;
		frmrefresh.submit();
	}
	<% end if %>
}

</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
		<!-- ��ܹ� ���� -->
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="2">
				<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
					<tr>
						<td>
							<img src="/images/icon_arrow_down.gif" align="absbottom">
							<font color="red">&nbsp;<strong>���(����)�Է�</strong></font>
						</td>
						<td align="right">
							<% if (C_IS_SHOP) then %>
                            ���� : <%= shopid %>
							<% elseif (C_IS_Maker_Upche) then %>
            		        <!-- ���� ��ü -->
            		        ���� : <% drawSelectBoxOpenOffShop "shopid",shopid %>
							<% else %>
                    	    ���� : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
                    		<% end if %>

							<% if (C_IS_Maker_Upche) then %>
        	                <input type="hidden" name="itembarcode" value="<%= itembarcode %>">
        					<% else %>
						    ��ǰ���ڵ�:
						    <input type=text class="text" name=itembarcode value="<%= itembarcode %>" size=16 maxlength=20 AUTOCOMPLETE="off" <%= ChkIIF(C_ADMIN_USER,"","readonly") %> onKeyPress="if (event.keyCode == 13){ document.frm.submit(); return false;}">
						    &nbsp;
							<% end if %>

        					<input type="button" class="button" value="�˻�" onclick="document.frm.submit();">
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<!-- ��ܹ� �� -->
	</form>
</table>

<p>

	<% if ojaegoitem.FResultCount>0 then %>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#FFFFFF">
    		<td rowspan="5" width="110" valign=top align=center><img src="<%= ojaegoitem.FOneItem.GetImageList %>" width="100" height="100"></td>
      		<td width="60"><b>*��ǰ����</b></td>
      		<td width="300">
      			<% if (Not C_IS_SHOP) then %>
      			<input type="button" class="button" value="����" onclick="popOffItemEdit('<%= itembarcode %>');">
      			<% end if %>
      		</td>
      		<td width="80">�ŷ���� </td>
      		<td colspan=2><%= GetJungsanGubunName(ojaegoitem.FOneItem.FComm_cd) %>
      			<% if (C_ADMIN_USER) then %>
      			[<%= ojaegoitem.FOneItem.FMakerMargin %> -&gt; <%= ojaegoitem.FOneItem.FshopMargin %>]
      			<% end if %>
      		</td>
		</tr>
		<tr bgcolor="#FFFFFF">
      		<td>��ǰ�ڵ�</td>
      		<td><%= ojaegoitem.FOneItem.GetBarCode %></td>
      		<td>�ǸŰ�</td>
      		<td colspan=2>
      			<% if (ojaegoitem.FOneItem.IsOffSaleItem) then %>
      			<strike><%= FormatNumber(ojaegoitem.FOneItem.FShopItemOrgprice,0) %></strike>
      			&nbsp;&nbsp;
      			<%= FormatNumber(ojaegoitem.FOneItem.Fshopitemprice,0) %>
      			<% else %>
      			<%= FormatNumber(ojaegoitem.FOneItem.Fshopitemprice,0) %>
      			<% end if %>
      		</td>
		</tr>
		<tr bgcolor="#FFFFFF">
      		<td>�귣��ID</td>
      		<td><%= ojaegoitem.FOneItem.FMakerid %></td>
      		<% if (C_IS_Maker_Upche) or (C_ADMIN_USER) then %>
      		<td>���԰�(��ü)</td>
      		<td colspan=2>
      			<% if ojaegoitem.FOneItem.IsShopContractExists then %>
      			<%= FormatNumber(ojaegoitem.FOneItem.GetOfflineBuycash,0) %>
      			<% end if %>
      		</td>
      		<% elseif (C_IS_SHOP) then %>
      		<td>���ް�(SHOP)</td>
      		<td colspan=2>
      			<% if ojaegoitem.FOneItem.IsShopContractExists then %>
      			<%= FormatNumber(ojaegoitem.FOneItem.GetOfflineSuplycash,0) %>
      			<% end if %>
      		</td>
      		<% else %>
      		<td></td>
      		<td colspan=2></td>
      		<% end if %>
		</tr>
		<tr bgcolor="#FFFFFF">
      		<td>��ǰ��</td>
      		<td>
      			<%= ojaegoitem.FOneItem.FShopItemName %>
      			<% if (ojaegoitem.FOneItem.FShopItemOptionName<>"") then %>
      			<font color="blue">[<%= ojaegoitem.FOneItem.FShopItemOptionName %>]</font>
      			<% end if %>
      		</td>
      		<% if (C_ADMIN_USER) then %>
      		<td>���ް�(SHOP)</td>
      		<td colspan=2>
      			<% if ojaegoitem.FOneItem.IsShopContractExists then %>
      			<%= FormatNumber(ojaegoitem.FOneItem.GetOfflineSuplycash,0) %>
      			<% end if %>
      		</td>
      		<% else %>
			<td></td>
      		<td colspan=2></td>
			<% end if %>
		</tr>
	</table>

	<p>
		<table width="100%" align="center" cellpadding="2" cellspacing="1" bgcolor="#BABABA" class="a">
			<form name=calcufrm >
			<input type="hidden" name="orgrealstock" value="<%= ocursummary.FOneItem.FrealstockNo %>">
			<input type="hidden" name="orgerrrealcheckno" value="<%= ocursummary.FOneItem.Ferrrealcheckno %>">
			<input type="hidden" name="availsysstock" value="<%= ocursummary.FOneItem.getAvailStock %>">
			<input type="hidden" name="shoprealstock" value="<%= ocursummary.FOneItem.getShopRealStock %>">
			<input type="hidden" name="logischulgo" value="<%= ocursummary.FOneItem.Flogischulgo %>">
			<input type="hidden" name="logisreturn" value="<%= ocursummary.FOneItem.Flogisreturn %>">
			<input type="hidden" name="errsampleitemno" value="<%= ocursummary.FOneItem.Ferrsampleitemno %>">
			<tr align="center" bgcolor="#DDDDFF">
    		<td width="60">���԰�<br>(�ٹ�����)</td>
    		<td width="60">�ѹ�ǰ<br>(�ٹ�����)</td>
    		<td width="60">���԰�<br>(��ü)</td>
    		<td width="60">�ѹ�ǰ<br>(��ü)</td>
    		<td width="60">���Ǹ�</td>
    		<td width="60">�ѹ�ǰ</td>
    		<td width="65" bgcolor="F4F4F4">�ý������</td>
    		<td width="60">����</td>
    		<td width="65" bgcolor="F4F4F4">�ǻ����
    			<br>(Sys+����)
    		</td>
    		<td width="60" bgcolor="F4F4F4">����</td>
    		<!-- <td width="55">�ҷ�</td> -->
    		<td bgcolor="F4F4F4">�������</td>
    		<td width="90" bgcolor="F4F4F4">
				��ȿ���
    			<br>(�ǻ����+����)
    		</td>
			<td width="60">�����</td>
			<td width="60">��ǰ��</td>
			<td width="65" bgcolor="F4F4F4">�������<br />(����)</td>
			<td >�ǻ�����Է�</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="25" align=center>
    		<td><%= ocursummary.FOneItem.Flogicsipgono %></td>
    		<td><%= ocursummary.FOneItem.Flogicsreipgono %></td>
    		<td><%= ocursummary.FOneItem.Fbrandipgono %></td>
    		<td><%= ocursummary.FOneItem.Fbrandreipgono %></td>
    		<td><%= ocursummary.FOneItem.Fsellno %></td>
    		<td><%= ocursummary.FOneItem.Fresellno %></td>
    		<td bgcolor="F4F4F4"><b><%= ocursummary.FOneItem.Fsysstockno %></b></td>
    		<td><%= ocursummary.FOneItem.Ferrrealcheckno %></td>
    		<td bgcolor="F4F4F4"><b><%= ocursummary.FOneItem.Frealstockno %></b></td>
    		<td><%= ocursummary.FOneItem.Ferrsampleitemno %></td>
    		<!-- <td><%= ocursummary.FOneItem.Ferrbaditemno %></td> -->
    		<td><input type="text" name="samplestock" value="<%= ocursummary.FOneItem.Ferrsampleitemno*-1 %>" size="4" maxlength="7" style="border:1px #999999 solid; text-align=center" ></td> <!-- ǥ�ø� +�� �� -->
    		<td bgcolor="F4F4F4"><%= ocursummary.FOneItem.getAvailStock %></td>
			<td><%= ocursummary.FOneItem.Flogischulgo %></td>
			<td><%= ocursummary.FOneItem.Flogisreturn %></td>
			<td bgcolor="F4F4F4"><b><%= ocursummary.FOneItem.getShopRealStock %></b></td>
			<td bgcolor="#FFDDDD"><input type="text" name="checkstock" value="<%= ocursummary.FOneItem.getShopRealStock %>" size="4" maxlength="7" style="border:1px #999999 solid; text-align=center" onKeyUp="RecalcuErr();"></td>
    </tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
        <td colspan="7" align="right"></td>
        <td><input type="text" name="errrealcheckno" value="0"  size="4" maxlength="7" readonly style="background:#CCCCCC; border:1px #999999 solid; text-align=center"></td>
        <td></td>
        <td></td>
        <td><input type="button" class="button" value="��������" onclick="SaveSample();" ></td>
        <td colspan="4"></td>
        <td bgcolor="#FFDDDD">
            <input type="text" class="text" name="stockdate" value="<%= NowDate %>" size=9 readonly ><a href="javascript:calendarOpen(calcufrm.stockdate);">
        	<img src="/images/calicon.gif" border="0" align="absmiddle" height=21 /></a>
			<br />
			<input type="button" class="button" value="�ǻ��������" onclick="SaveErr();" />
        </td>
    </tr>
    </form>
</table>

<form name=frmrefresh method=post action="/common/offshop/shop_stockrefresh_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="realstock" value="">
<input type="hidden" name="samplestock" value="">
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="<%= itemoption %>">
<input type="hidden" name="stockdate" value="">
</form>
<% end if %>
<%
set ojaegoitem = Nothing
set ocursummary= Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
