<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ ���
' Hieditor : 2009.11.17 ������ ����
'			 2011.05.06 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->
<!-- #include virtual="/lib/classes/stock/shopbatchstockcls.asp"-->

<%
dim StockDate ,shopid , jobkey ,centermwdiv ,idx ,usingyn ,oshopBatch ,oOffStock ,i
dim totsysstock, totavailstock, totrealstock ,BasicMonth, IsExpireEdit  
	StockDate = Left(CStr(Now()),10)
	idx = RequestCheckVar(request("idx"),10)
	usingyn = RequestCheckVar(request("usingyn"),10)
	centermwdiv = RequestCheckVar(request("centermwdiv"),10)

set oshopBatch = new CShopOrder
	oshopBatch.FRectidx=idx
	
	if (idx<>"") then
	    oshopBatch.GetOneShopBatchOrder
	end if

if (oshopBatch.FResultCount>0) then
    shopid = oshopBatch.FOneItem.Fjobshopid
    jobkey = oshopBatch.FOneItem.Fjobkey
    StockDate = Left(oshopBatch.FOneItem.FShopRegDate,10)
end if 

set oOffStock = new CShopItemSummary
	oOffStock.FRectShopID = shopid
	oOffStock.FRectMakerID = ""
	oOffStock.FRectBatchIdx = idx
	
	if (shopid<>"") and (idx<>"") then
	    oOffStock.GetShopCurrentStockByBatchJob 
	end if

BasicMonth = CStr(DateSerial(Year(now()),Month(now())-1,1))
%>

<script language='javascript'>

function popShopCurrentStock(shopid,itemgubun,itemid,itemoption){
    var popwin = window.open('/common/offshop/shop_itemcurrentstock.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popShopCurrentStock','width=900,height=600,resizable=yes,scrollbars=yes');
    popwin.focus();
}

function popOffItemEdit(ibarcode){
    <% if C_IS_SHOP then %>
        return;
    <% elseif C_IS_Maker_Upche then %>
        var popwin = window.open('/designer/offshop/popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
        popwin.focus();
    <% else %>
	    var popwin = window.open('/admin/offshop/popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
	    popwin.focus();
	<% end if %>
}


function popOffErrInput(shopid,itemgubun,itemid,itemoption){
    var popwin = window.open('/common/offshop/popOffrealerrinput.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popAdmOffrealerrinput','width=900,height=460,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function PopOFFBrandStockSheet(){
    
    var shopid = document.frm.shopid.value;
    var makerid = document.frm.makerid.value;
    var centermwdiv = document.frm.centermwdiv.value;
    var usingyn= document.frm.usingyn.value;
    
    if ((shopid.length<1)||(makerid.length<1)){
        alert('���� ����� �귣�带 ������ ����� �ּ���.');
        return;
    }
    
    var popwin;
    
    popwin = window.open('/common/pop_offbrandstockprint.asp?shopid=' + shopid + '&makerid=' + makerid + '&centermwdiv=' + centermwdiv + '&usingyn=' + usingyn ,'pop_offbrandstockprint','width=1000,height=600,scrollbars=yes,resizable=yes')
    popwin.focus();
}

function RealStockInputArr(){
   
    
    var frm = document.frmArr;
    var ischecked = false;
    var i = 0;
    var stockdate = frmStockDt.stockdate.value;
    
    //2���� �������� �ǻ��Է� �Ұ�.
    if (stockdate<'<%= BasicMonth %>'){
		alert('�δ� ���� ��¥�δ� ����ľ����� ��� �� �� �����ϴ�.');
		return;
	}
	
    if (!frm.cksel) return;
    
    if (frm.cksel.length){
        for (i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                ischecked = true;
                if (!IsInteger(frm.Arrrealstock[i].value)){
                    alert('������ �����մϴ�.');
                    frm.Arrrealstock[i].focus();
                    return;
                }
            }
        }
    }else{
        if (frm.cksel.checked){
            ischecked = true;
            if (!IsInteger(frm.Arrrealstock.value)){
                alert('������ �����մϴ�.');
                frm.Arrrealstock.focus();
                return;
            }
        }
    }
    
    if (!(ischecked)){
        alert('���õ� ��ǰ�� �����ϴ�.');
        return;
    }
    
    if (confirm('�ǻ� ��� ���� �Ͻðڽ��ϱ�?')){
        frm.stockdate.value = stockdate;
        frm.submit();
    }
}

function CheckThis(i){
    var frm = document.frmArr;
    if (frm.cksel.length){
        frm.cksel[i].checked = true;
        AnCheckClick(frm.cksel[i]);
    }else{
        frm.cksel.checked = true;
        AnCheckClick(frm.cksel);
    }
}

function switchCheckBox(comp){
    var frm = comp.form;

	if(frm.cksel.length>1){
		for(i=0;i<frm.cksel.length;i++){
		    if (!frm.cksel[i].disabled){
    			frm.cksel[i].checked = comp.checked;
    			AnCheckClick(frm.cksel[i]);
			}
		}
	}else{
	    if (!frm.cksel.disabled){
    		frm.cksel.checked = comp.checked;
    		AnCheckClick(frm.cksel);
    	}
	}
}

</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left" >
	    ���� : <%= shopid %>
	    &nbsp;&nbsp;
	    �۾���ȣ : <%= jobkey %>
	</td>
	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		��뱸�� : <% drawSelectBoxUsingYN "usingyn", usingyn %> &nbsp;&nbsp;
		���͸��Ա��� :
		   <select class="select" name="centermwdiv">
           <option value="">��ü</option>
           <option value="MW" <%= ChkIIF(centermwdiv="MW","selected","") %> >����+Ư��</option>
           <option value="W"  <%= ChkIIF(centermwdiv="W","selected","") %> >Ư��</option>
           <option value="M"  <%= ChkIIF(centermwdiv="M","selected","") %> >����</option>
           <option value="NULL" <%= ChkIIF(centermwdiv="NULL","selected","") %> >������</option>
           </select>
           
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" >
<form name="frmStockDt">
<tr height="30">
	<td align="left">
	    �˻���� : <%= oOffStock.FResultCount %> (�ִ� 2,000��)
		<% if C_ADMIN_AUTH=true then %>
    	<!--
        <input type="button" class="button" value="�귣�� ��ü ���ΰ�ħ" onclick="RefreshIpchulStock();">
        -->
        <% end if %>
        
	</td>
	<td align="right">
	    <input type="text" class="text" name="stockdate" value="<%= StockDate %>" size=11 readonly ><a href="javascript:calendarOpen(frmStockDt.stockdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21>
		<input type="button" class="button" name="stock_sheet_print" value="���� ��ǰ �ǻ���� �ϰ��Է�" onclick="RealStockInputArr();"> 
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmArr" method="post" action="/common/offshop/shop_stockrefresh_process.asp">
<input type="hidden" name="mode" value="ArrOfferrcheckupdate">
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="stockdate" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"><input type="checkbox" name="chkAll" onClick="switchCheckBox(this);"></td>
    <td width="30">����</td>
	<td width="40">��ǰID</td>
	<td width="40">�ɼ�</td>
	<td width="50">�̹���</td>
	<td>��ǰ��<br>[�ɼǸ�]</td>
	<td width="40">����<br>����<br>����</td>
	<td width="40">����<br>�԰�<br>��ǰ</td>
	<td width="40">�귣��<br>�԰�<br>��ǰ</td>
    <td width="40">����<br>�Ǹ�<br>��ǰ</td>
    <td width="40" bgcolor="F4F4F4">�ý���<br>�����</td>
    <td width="40">��<br>�ǻ�<br>����</td>
    <!-- <td width="40" bgcolor="F4F4F4">�ǻ�<br>���</td> 
    <td width="40">��<br>����</td>
    <td width="40">��<br>�ҷ�</td> -->
    <td width="40" bgcolor="F4F4F4">�ǻ�<br>���</td>
    
    <td width="30">���<br>����</td>
    <!--
    <td width="30">�Ǹ�<br>����</td>
    <td width="30">����<br>����</td>
    <td width="50">����<br>����</td>
    -->
    <td width="40">�ǻ�<br>���</td>
    <td width="40">����<br>�ǻ�<br>�Է�</td>
</tr>
<% if oOffStock.FResultCount<1 then %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td colspan="20" >[ �˻� ����� �����ϴ�. ]</td>
</tr>
<% else %>
<% for i=0 to oOffStock.FResultCount - 1 %>
<%
totsysstock	    = totsysstock + oOffStock.FItemList(i).FsysstockNo
totavailstock   = totavailstock + oOffStock.FItemList(i).getAvailStock
totrealstock    = totrealstock + oOffStock.FItemList(i).FrealstockNo
%>
<% if oOffStock.FItemList(i).Fisusing="Y" then %>
<tr bgcolor="#FFFFFF" align="center">
<% else %>
<tr bgcolor="#EEEEEE" align="center">
<% end if %>
    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" value="<%= i %>">
    <input type="hidden" name="Arritemgubun" value="<%= oOffStock.FItemList(i).FItemGubun %>">
    <input type="hidden" name="Arritemid" value="<%= oOffStock.FItemList(i).FItemID %>">
    <input type="hidden" name="Arritemoption" value="<%= oOffStock.FItemList(i).FItemOption %>">
    </td>
    <td><%= oOffStock.FItemList(i).FItemGubun %></td>
	<td>
	    <% if (C_ADMIN_USER or C_IS_Maker_Upche) then %>
	    <a href="javascript:popOffItemEdit('<%= oOffStock.FItemList(i).getBarcode %>');"><%= oOffStock.FItemList(i).Fitemid %></a>
	    <% else %>
	    <%= oOffStock.FItemList(i).Fitemid %>
	    <% end if %>
	</td>
	<td><%= oOffStock.FItemList(i).FItemOption %></td>
	<td><img src="<%= oOffStock.FItemList(i).GetImageSmall %>" width=50 height=50> </td>
	<td align="left">
      	<a href="javascript:popShopCurrentStock('<%= oOffStock.FItemList(i).FShopid %>','<%= oOffStock.FItemList(i).Fitemgubun %>','<%= oOffStock.FItemList(i).FItemID %>','<%= oOffStock.FItemList(i).FItemOption %>');"><%= oOffStock.FItemList(i).FShopitemname %></a>
      	<% if oOffStock.FItemList(i).FShopitemoptionName <>"" then %>
      		<br>
      		<font color="blue">[<%= oOffStock.FItemList(i).FShopitemoptionName %>]</font>
      	<% end if %>
    </td>
    <td><%= fnColor(oOffStock.FItemList(i).FCenterMwdiv,"mw") %></td>
	<td><%= oOffStock.FItemList(i).Flogicsipgono + oOffStock.FItemList(i).Flogicsreipgono %></td>
	<td><%= oOffStock.FItemList(i).Fbrandipgono + oOffStock.FItemList(i).Fbrandreipgono %></td>
	<td><%= oOffStock.FItemList(i).Fsellno + oOffStock.FItemList(i).Fresellno %></td>
	<td bgcolor="F4F4F4"><b><%= oOffStock.FItemList(i).FsysstockNo %></b></td>
	<td><%= oOffStock.FItemList(i).Ferrrealcheckno %></td>
	<!-- <td bgcolor="F4F4F4"><b><%= oOffStock.FItemList(i).FrealstockNo %></td> 
	<td><%= oOffStock.FItemList(i).Ferrsampleitemno %></td>
	<td><%= oOffStock.FItemList(i).Ferrbaditemno %></td> -->
	<td bgcolor="F4F4F4"><b><font color="<%= ChkIIF(oOffStock.FItemList(i).getAvailStock<0,"#FF0000","#000000") %>"><%= oOffStock.FItemList(i).getAvailStock %></font></b></td>
	
	<td><%= oOffStock.FItemList(i).Fisusing %></td>
	<!--
    <td></td>
	<td></td>
	<td></td>
	-->
	<td><input type="text" class="text" name="Arrrealstock" value="<%= oOffStock.FItemList(i).FBatchItemNo %>" size="4" maxlength="4" AUTOCOMPLETE="off" style="text-align=center" onKeyDown="CheckThis('<%= i %>');"></td>
	<td>
		<input type="button" class="button" value="�ǻ�" onclick="popOffErrInput('<%= shopid %>','<%= oOffStock.FItemList(i).Fitemgubun %>','<%= oOffStock.FItemList(i).Fitemid %>','<%= oOffStock.FItemList(i).Fitemoption %>');">
	</td>
</tr>
<% next %>
<% end if %>
</form>
</table>

<%
set oshopBatch = Nothing
set oOffStock = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->