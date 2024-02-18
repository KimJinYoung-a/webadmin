<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->
<%
response.write "������� �޴� -  [OFF]����_�������>>[����]�������귣�� ���ٶ�"
response.end

dim shopid, makerid, centermwdiv, itembarcode, research
dim itemgubun, itemid, itemoption, grpType
dim params
shopid       = RequestCheckVar(request("shopid"),32)
makerid      = RequestCheckVar(request("makerid"),32)
centermwdiv  = RequestCheckVar(request("centermwdiv"),10)
itembarcode  = RequestCheckVar(request("itembarcode"),32)
research     = RequestCheckVar(request("research"),2)
grpType      = RequestCheckVar(request("grpType"),10)

dim fromdate,todate
fromdate = request("fromdate")
todate   = request("todate")

params       = "menupos=1379&research=on&page=&shopid="&shopid&"&makerid="&makerid&"&yyyy1="&Left(fromdate,4)&"&mm1="&Mid(fromdate,6,2)&"&yyyy2="&Left(dateAdd("d",-1,todate),4)&"&mm2="&Mid(dateAdd("d",-1,todate),6,2)&"&grpType="&grpType



	
if (C_IS_SHOP) or (C_IS_Maker_Upche) then
    response.write "������ �����ϴ�."
    dbget.close() : response.end
end if

dim oOffStock
set oOffStock = new CShopItemSummary
oOffStock.FRectShopID   = shopid
oOffStock.FRectMakerID  = makerid
oOffStock.FRectErrType  = "D"
oOffStock.FRectStartDate = fromdate
oOffStock.FRectEndDate   = todate
oOffStock.FRectGroupType = grpType

oOffStock.GetOFFErrItemSummaryGroupByItem

Dim i, TotErrrealcheckno
%>
<script languag='javascript'>
function reCalcuLoss(comp,i){
    var frm = comp.form;
    
    if (frm.cksel.length){
        frm.SUBTTLrealcheckErrRemain[i].value = frm.realcheckErr[i].value*1+frm.AssignrealcheckErr[i].value*1;
        frm.SUBTTLshopsuplycash[i].value = frm.AssignrealcheckErr[i].value*1*frm.shopsuplycash[i].value*1;
    }else{
        frm.SUBTTLrealcheckErrRemain.value = frm.realcheckErr.value*1+frm.AssignrealcheckErr.value*1;
        frm.SUBTTLshopsuplycash.value = frm.AssignrealcheckErr.value*1*frm.shopsuplycash.value*1;
    }
    
    summaryTotal(frm);
}

function summaryTotal(frm){
    
    var ttlsum = 0;
    var itemcnt = 0;
    var remaincnt = 0;
    if (frm.cksel.length){
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                itemcnt+=frm.AssignrealcheckErr[i].value*1;
                remaincnt+=frm.SUBTTLrealcheckErrRemain[i].value*1;
                ttlsum+=frm.AssignrealcheckErr[i].value*1*frm.shopsuplycash[i].value*1;
            }
        }
    }else{
        if (frm.cksel.checked){
            itemcnt+=frm.AssignrealcheckErr.value*1;
            remaincnt+=frm.SUBTTLrealcheckErrRemain.value*1;
            ttlsum+=frm.AssignrealcheckErr.value*1*frm.shopsuplycash.value*1;
        }
    }
    frm.TTLrealcheckErr.value = itemcnt;
    frm.TTLrealcheckErrRemain.value = remaincnt;
    frm.TTLshopsuplycash.value = ttlsum;
}
    
function chkALL(comp){
    var frm = comp.form;
    if (frm.cksel.length){
        for (var i=0;i<frm.cksel.length;i++){
            frm.cksel[i].checked=comp.checked;
            AnCheckClick(frm.cksel[i]);
        }
    }else{
        frm.cksel.checked=comp.checked;
        AnCheckClick(frm.cksel);
    }
    summaryTotal(frm);
}

function AssignErrLoss(){
    var frm = document.frmArr;
    frm.lossDate.value = document.frmStockDt.stockdate.value;
    
    if (!chkExesits(frm.cksel)){
        alert('���� ������ �����ϴ�.');
        return;  
    }
    
    if (confirm('���� �ν� ��� �ݿ� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

function chkExesits(comp){
    var frm = comp.form;
    
    if (comp.length){
        for (var i=0;i<comp.length;i++){
            if (comp[i].checked){
                return true;
            }
        }
    }else{
        if (comp.checked){
            return true;
        }
    }
    return false;
}

function AssignMeaipPro(){
   var frm = document.frmArr;
   var pro = document.frmStockDt.assignPro.value;
   if (!chkExesits(frm.cksel)){
        alert('���� ������ �����ϴ�.');
        return;  
   }
   
   if (frm.cksel.length){
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                frm.shopsuplycash[i].value = frm.Orgshopsuplycash[i].value*1*pro/100*1;
                frm.SUBTTLshopsuplycash[i].value = frm.AssignrealcheckErr[i].value*1*frm.shopsuplycash[i].value*1;
            }
        }
   }else{
        frm.cksel.checked=comp.checked;
        frm.shopsuplycash.value = frm.Orgshopsuplycash.value*1*pro/100*1;
        frm.SUBTTLshopsuplycash.value = frm.AssignrealcheckErr.value*1*frm.shopsuplycash.value*1;
   }
   
   summaryTotal(frm);
}

function AssignMeaipProbySell(){
   var frm = document.frmArr;
   var pro = document.frmStockDt.assignProSell.value;
   if (!chkExesits(frm.cksel)){
        alert('���� ������ �����ϴ�.');
        return;  
   }
   
   if (frm.cksel.length){
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                frm.shopsuplycash[i].value = frm.OrgshopSellcash[i].value*1*pro/100*1;
                frm.SUBTTLshopsuplycash[i].value = frm.AssignrealcheckErr[i].value*1*frm.shopsuplycash[i].value*1;
            }
        }
   }else{
        frm.cksel.checked=comp.checked;
        frm.shopsuplycash.value = frm.OrgshopSellcash.value*1*pro/100*1;
        frm.SUBTTLshopsuplycash.value = frm.AssignrealcheckErr.value*1*frm.shopsuplycash.value*1;
   }
   
   summaryTotal(frm);
}
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
		    ���� : <%= shopid %> &nbsp;&nbsp;
    		�귣��ID : <%= makerid %> &nbsp;&nbsp;
			<br>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    (����)����� : <%= fromdate %> ~ <%= DateAdd("d",-1,todate) %>
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
		    �ݿ��ܰ� ���� ���� ���԰��� <input type="text" class="text" name="assignPro" value="100" size="3">%
			<input type="button" class="button" value="�ݿ�" onClick="AssignMeaipPro();">
			�ݿ��ܰ� ���� ���� �ǸŰ��� <input type="text" class="text" name="assignProSell" value="100" size="3">%
			<input type="button" class="button" value="�ݿ�" onClick="AssignMeaipProbySell();">
		     
		</td>
		<td align="right">
		    �ν���� �ݿ���
		    <input type="text" class="text" name="stockdate" value="<%= DateAdd("d",-1,todate) %>" size=11 readonly ><a href="javascript:calendarOpen(frmStockDt.stockdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		    (���� �� �� ��������)
			<input type="button" class="button" name="stock_sheet_print" value="���� ��ǰ �ν� ��� �ݿ�" onclick="AssignErrLoss();"> 
		</td>
	</tr>
	</form>
</table>
<!-- �׼� �� -->

<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmArr" method="post" action="shopErrorLoss_Process.asp">
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="lossDate" value="">
<input type="hidden" name="params" value="<%= params %>">

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"><input type="checkbox" name="ckAll" onClick="chkALL(this);"></td>
	<td width="80">��ǰ�ڵ�</td>
	<td width="200">��ǰ�� <font color="blue">[�ɼǸ�]</font></td>
	<td width="70">�� �ǸŰ�</td>
	<td width="70">���� ���԰�</td>
	<td width="70">���� ���԰�</td>
	<td width="70">���� �հ�</td>
	<td width="100">�ν� �ݿ���</td>
	<td width="100">�ν� �ݿ��ܰ�</td>
	<td width="40">����<br>����</td>
	<td >�ν������</td>
</tr>
<% for i=0 to oOffStock.FResultcount -1 %>
<% TotErrrealcheckno = TotErrrealcheckno + oOffStock.FItemList(i).Ferrrealcheckno %>
<tr bgcolor="#FFFFFF">
    <td><input type="checkbox" name="cksel" value="<%= i %>" onClick="reCalcuLoss(this,<%= i %>);AnCheckClick(this);">
    <input type="hidden" name="itemgubun" value="<%= oOffStock.FItemList(i).Fitemgubun %>">
    <input type="hidden" name="itemid" value="<%= oOffStock.FItemList(i).Fitemid %>">
    <input type="hidden" name="itemoption" value="<%= oOffStock.FItemList(i).Fitemoption %>">
    <input type="hidden" name="shopitemprice" value="<%= oOffStock.FItemList(i).Fshopitemprice %>">
    <input type="hidden" name="shopbuyprice" value="<%= oOffStock.FItemList(i).Fshopbuyprice %>">
    
    </td>
    <td><%= oOffStock.FItemList(i).getBarcode %></td>
    <td><%= oOffStock.FItemList(i).Fshopitemname %>
    <% if oOffStock.FItemList(i).Fshopitemoptionname<>"" then %>
        <font color="blue">[<%= oOffStock.FItemList(i).Fshopitemoptionname %>]</font>
    <% end if %>
    </td>
    <td align="right">
    
    <input type="hidden" name="OrgshopSellcash" value="<%= oOffStock.FItemList(i).Fshopitemprice %>">
    <%= FormatNumber(oOffStock.FItemList(i).Fshopitemprice,0) %></td>
    <td align="right"><%= FormatNumber(oOffStock.FItemList(i).Fshopbuyprice,0) %></td>
    <td align="right"><%= FormatNumber(oOffStock.FItemList(i).fshopsuplycash,0) %></td>
    <td align="center"><%= FormatNumber(oOffStock.FItemList(i).Ferrrealcheckno,0) %></td>
    <td align="center">
    <input type="hidden" name="realcheckErr" value="<%= oOffStock.FItemList(i).Ferrrealcheckno %>">
    <input type="text" name="AssignrealcheckErr" value="<%= oOffStock.FItemList(i).Ferrrealcheckno*-1 %>" class="text" size="5"  style="text-align=center" onKeyUp="reCalcuLoss(this,<%= i %>)"></td>
    <td align="center">
    <input type="hidden" name="Orgshopsuplycash" value="<%= oOffStock.FItemList(i).fshopsuplycash %>">
    <input type="text" name="shopsuplycash" value="<%= oOffStock.FItemList(i).fshopsuplycash %>" class="text" size="9"  style="text-align=right" onKeyUp="reCalcuLoss(this,<%= i %>)">
    </td>
    <td align="center"><input type="text" name="SUBTTLrealcheckErrRemain" value="0" class="text" size="5"  style="text-align=center;border=0" READONLY ></td>
	<td align="center"><input type="text" name="SUBTTLshopsuplycash" value="<%= oOffStock.FItemList(i).fshopsuplycash*oOffStock.FItemList(i).Ferrrealcheckno*-1 %>" class="text" size="9"  style="text-align=right;border=0" READONLY ></td>
</tr>
<% next %>
<tr bgcolor="#DDFFFF">
    <td>�հ�</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td align="center"><%= FormatNumber(TotErrrealcheckno,0) %></td>
    <td align="center"><input type="text" name="TTLrealcheckErr" value="" class="text" size="5"  style="text-align=center;border=0" READONLY ></td>
    <td></td>
    <td align="center"><input type="text" name="TTLrealcheckErrRemain" value="" class="text" size="5"  style="text-align=center;border=0" READONLY ></td>
    <td align="center"><input type="text" name="TTLshopsuplycash" value="" class="text" size="9"  style="text-align=right;border=0" READONLY ></td>
</tr>
</form>
</table>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
