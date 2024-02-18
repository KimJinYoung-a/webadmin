<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  오프라인매장 재고파악
' History : 2010.04.02 한용민 수정 
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/stock_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->
<!-- #include virtual="/lib/classes/stock/shopbatchstockcls.asp"-->
<% 
response.charset = "euc-kr"
Dim isFirstInput, isWorkingtate
Dim shopid, makerid, itembarcode, stTakingIdx, ostTakingIdx, vidx
Dim mode, actStr

shopid       = RequestCheckVar(request("shopid"),32)
makerid      = RequestCheckVar(request("makerid"),32)
itembarcode  = RequestCheckVar(request("itembarcode"),32)
stTakingIdx  = RequestCheckVar(request("stTakingIdx"),10)
mode         = RequestCheckVar(request("mode"),32)
actStr       = RequestCheckVar(request("actStr"),128)
vidx         = RequestCheckVar(request("vidx"),10)

if (stTakingIdx="") then stTakingIdx="0"
ostTakingIdx = stTakingIdx

dim oOffStockTaking
set oOffStockTaking = new CStockTaking
oOffStockTaking.FRectShopID       = shopid
oOffStockTaking.FRectMakerID      = makerid

if (stTakingIdx<>"0") then
    oOffStockTaking.FRectIdx = stTakingIdx
    oOffStockTaking.getOneStockTaking
elseif ((shopid<>"") and (makerid<>"")) then
    oOffStockTaking.getRecentStockTaking
    
    if (oOffStockTaking.FResultCount>0) then 
        stTakingIdx = oOffStockTaking.FOneItem.FstTakingIdx
    end if
end if

if (oOffStockTaking.FResultCount>0) then
    isWorkingtate= (CStr(oOffStockTaking.FOneItem.FstStatus)=0)
end if

Dim ErrStr, ErrNo, inputedStNo
ErrNo = 0
if (itembarcode<>"") then
    if Not (oOffStockTaking.AddByBarcode(stTakingIdx, itembarcode, 1,inputedStNo )) then
        ErrStr = oOffStockTaking.getLastErrStr
        ErrNo  = oOffStockTaking.getLastErrNo
    end if
end if
oOffStockTaking.FRectIdx = stTakingIdx    

if (mode<>"actSave") then
    oOffStockTaking.getStockTakingDetailWithList
end if

dim i

if (CStr(ostTakingIdx)="0") and (CStr(stTakingIdx)<>"0") then 
    isFirstInput = true
    isWorkingtate  = true
end if

%>
<!-- Spread -->
<% if (mode<>"actSave") then %>
<form name="frmArr" method="post" action="/common/offshop/shop_stockrefresh_process.asp">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="mode" value="ArrOffStockTakingupdate">
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="stTakingIdx" value="<%= stTakingIdx %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"></td>
    <td width="30">구분</td>
	<td width="40">상품ID</td>
	<td width="40">옵션</td>
	<td width="50">이미지</td>
	<td>상품명<br>[옵션명]</td>
	<td>판매가</td>
	<td width="50" >사용<br>구분</td>
    <td width="60" bgcolor="F4F4F4">현 실사<br>재고</td>
    <td width="60">재고파악</td>
</tr>
<% if oOffStockTaking.FResultCount<1 then %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td colspan="21" >[ 검색 결과가 없습니다. ]</td>
</tr>
<% else %>
<%
dim totalSysRealStock, totalRealStock, totalSum
dim fRow
for i=0 to oOffStockTaking.FResultCount - 1 
%>
<%
totalSysRealStock = totalSysRealStock + oOffStockTaking.FItemList(i).Frealstockno
totalRealStock    = totalRealStock    + oOffStockTaking.FItemList(i).FstNo
totalSum          = totalSum          + oOffStockTaking.FItemList(i).FstNo
%>
	<% if (itembarcode<>"") and ((oOffStockTaking.FItemList(i).getBarcode=itembarcode) or (oOffStockTaking.FItemList(i).getPublicBarcode=itembarcode)) then %>
    <% fRow = i %>
    <tr bgcolor="#FFFFFF" align="center">
    <% else %>
    <tr bgcolor="#FFFFFF" align="center">
    <% end if %>
        <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" value="<%= i %>" <%= CHKIIF(isWorkingtate,"","disabled") %> >
        <input type="hidden" name="Arritemgubun" value="<%= oOffStockTaking.FItemList(i).FItemGubun %>">
        <input type="hidden" name="Arritemid" value="<%= oOffStockTaking.FItemList(i).FItemID %>">
        <input type="hidden" name="Arritemoption" value="<%= oOffStockTaking.FItemList(i).FItemOption %>">
        <input type="hidden" name="tenBarcoed" value="<%= oOffStockTaking.FItemList(i).getBarcode %>">
        <input type="hidden" name="publicBarcoed" value="<%= TRIM(oOffStockTaking.FItemList(i).Fextbarcode) %>">
        <input type="hidden" name="prechkNo" value="<%= NULL2Zero(oOffStockTaking.FItemList(i).FstNo) %>">
        </td>
        <td><%= oOffStockTaking.FItemList(i).FItemGubun %></td>
    	<td>
    	    <% if (C_ADMIN_USER or C_IS_Maker_Upche) then %>
    	    <a href="javascript:popOffItemEdit('<%= oOffStockTaking.FItemList(i).getBarcode %>');"><%= oOffStockTaking.FItemList(i).Fitemid %></a>
    	    <% else %>
    	    <%= oOffStockTaking.FItemList(i).Fitemid %>
    	    <% end if %>
    	</td>
    	<td><%= oOffStockTaking.FItemList(i).FItemOption %></td>
    	<td><img src="<%= oOffStockTaking.FItemList(i).GetImageSmall %>" width=50 height=50> </td>
    	<td align="left">
          	<a href="javascript:popShopCurrentStock('<%= Shopid %>','<%= oOffStockTaking.FItemList(i).Fitemgubun %>','<%= oOffStockTaking.FItemList(i).FItemID %>','<%= oOffStockTaking.FItemList(i).FItemOption %>');"><%= oOffStockTaking.FItemList(i).FShopitemname %></a>
          	<% if oOffStockTaking.FItemList(i).FShopitemoptionName <>"" then %>
          		<br>
          		<font color="blue">[<%= oOffStockTaking.FItemList(i).FShopitemoptionName %>]</font>
          	<% end if %>
        </td>
    	<td><%= FormatNumber(oOffStockTaking.FItemList(i).fshopitemprice,0) %></td>  
    	<td><%= oOffStockTaking.FItemList(i).Fisusing  %></td>             
    	<td><%= FormatNumber(oOffStockTaking.FItemList(i).Frealstockno,0) %></td>        	     
    	<td ><input type="text" class="<%= CHKIIF(NULL2Zero(oOffStockTaking.FItemList(i).FstNo)<>0,"text","text_ro") %>" id="Arrrealstock" name="Arrrealstock" size="3" value="<%= NULL2Zero(oOffStockTaking.FItemList(i).FstNo) %>" onKeyUp="CheckThis(<%=i%>)"></td>
    	
    </tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
    <td colspan=4>상품건수</td>
    <td ><%= i %></td>
    <td colspan=3></td>
    <td ><%= totalSysRealStock %></td>
	<td ><%= totalRealStock %></td>
</tr>
<% end if %>
</table> 
</form>
<% else %>
<form name="frmArr2" method="post" action="/common/offshop/shop_stockrefresh_process.asp">
<input type="hidden" id="Arrrealstock" name="Arrrealstock" value="">
</form>
<% end if %>

<%
set oOffStockTaking = Nothing
 
if (itembarcode<>"") then
    if (CStr(ErrNo)<>"0") then 
        response.write  "<script>playding('chord');</script>"
    else
        if (mode="actSave") then
            if (vidx="") then vidx=-1
          ''response.write  "<script>alert('"&inputedStNo&"');</script>"
            response.write  "<script>"&actStr&";setSvrSt("&vidx&","&inputedStNo&");</script>"
        else
            response.write  "<script>focusRow(frmArr.Arrrealstock["&fRow&"]);</script>"
        end if
    
        if (isFirstInput) then
            ''작업번호 refresh
            ''response.write "<script>alert('작업번호 refresh')</script>"
            response.write "<script>location.reload();</script>"
        end if
    end if
end if

if ErrStr<>"" then response.write "<script>alert('"& ErrStr &"');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" --> 