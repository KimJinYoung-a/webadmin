<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  오프라인매장 재고현황
' History : 2010.04.02 한용민 수정 
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->
<!-- #include virtual="/lib/classes/stock/shopbatchstockcls.asp"-->

<%
dim SType : SType= RequestCheckVar(request("SType"),32)


dim idx : idx = RequestCheckVar(request("idx"),32)
Dim shopid, jobkey, joborderno, StockDate
dim makerid

IF (SType="batch") then
    dim oshopBatch
    makerid = RequestCheckVar(request("makerid"),32)
    set oshopBatch = new CShopOrder
    	oshopBatch.FRectidx=idx
    	
    	if (idx<>"") then
    	    oshopBatch.GetOneShopBatchOrder
    	end if
    
    if (oshopBatch.FResultCount>0) then
        shopid = oshopBatch.FOneItem.Fjobshopid
        jobkey = oshopBatch.FOneItem.Fjobkey
        joborderno = oshopBatch.FOneItem.Forderno
        StockDate = Left(oshopBatch.FOneItem.FShopRegDate,10)
    end if 
    set oshopBatch= Nothing
ELSEIF (SType="stTaking") then
    Dim oOffStockTaking
    set oOffStockTaking = new CStockTaking
    oOffStockTaking.FRectIdx = idx
    if (idx<>"") then
        oOffStockTaking.getOneStockTaking
    end if
    
    if (oOffStockTaking.FResultCount>0) then
        shopid = oOffStockTaking.FOneItem.Fshopid
        makerid = oOffStockTaking.FOneItem.Fmakerid
        StockDate = Left(oOffStockTaking.FOneItem.FStockDate,10)
    end if 
    set oOffStockTaking= Nothing
ELSE
    response.write "<script>alert('재고파악 구분이 지정되지 않았습니다.');</script>"
    dbget.Close() : response.end
End IF

dim itemgubun, itemid, itemoption


dim oOffStock
set oOffStock = new CShopItemSummary
oOffStock.FRectShopID       = shopid
oOffStock.FRectMakerID      = makerid
oOffStock.FRectBatchIdx     = idx
if ((shopid<>"") and (makerid<>"") and (idx<>"")) then
    IF (SType="stTaking") then
        oOffStock.GetShopCurrentStockByStockTaking
    ELSE
        oOffStock.GetShopCurrentStockByBatchJobByBrand
    END IF
end if

dim i
dim totsysstock, totavailstock, totrealstock    

dim IsUpcheWitakItem
if (makerid<>"") and (shopid<>"") then
    IsUpcheWitakItem = (GetShopBrandContract(shopid,makerid)="B012")
end if
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
    <% if (C_IS_Maker_Upche) and (Not IsUpcheWitakItem) then %>
        alert('권한이 없습니다. - 업체위탁 상품만 재고 수정 가능.');
        return; //업체위탁 상품인 경우?
    <% else %>
        var popwin = window.open('/common/offshop/popOffrealerrinput.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popAdmOffrealerrinput','width=900,height=460,scrollbars=yes,resizable=yes');
	    popwin.focus();
	<% end if %>
}

function PopOFFBrandStockSheet(){
    
    var shopid = document.frm.shopid.value;
    var makerid = document.frm.makerid.value;
    var centermwdiv = "";//document.frm.centermwdiv.value;
    var usingyn= document.frm.usingyn.value;
    
    if ((shopid.length<1)||(makerid.length<1)){
        alert('먼저 매장과 브랜드를 선택후 출력해 주세요.');
        return;
    }
    
    var popwin;
    
    popwin = window.open('/common/pop_offbrandstockprint.asp?shopid=' + shopid + '&makerid=' + makerid + '&centermwdiv=' + centermwdiv + '&usingyn=' + usingyn ,'pop_offbrandstockprint','width=1000,height=600,scrollbars=yes,resizable=yes')
    popwin.focus();
}

function RealStockInputArr(){
    <% if (C_IS_Maker_Upche) and (Not IsUpcheWitakItem) then %>
        alert('권한이 없습니다. - 업체위탁 상품만 재고 수정 가능.');
        return;
    <% end if %>
    
    var frm = document.frmArr;
    var ischecked = false;
    var i = 0;
    var stockdate = frmStockDt.stockdate.value;
    
    if (!frm.cksel) return;
    
    if (frm.cksel.length){
        for (i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                ischecked = true;
                if (!IsInteger(frm.Arrrealstock[i].value)){
                    alert('정수만 가능합니다.');
                    frm.Arrrealstock[i].focus();
                    return;
                }
                
                if (frm.Arrrealstock[i].value*1<0){
                    alert('재고는 마이너스는 불가 합니다.');
                    frm.Arrrealstock[i].focus();
                    return;
                }
            }
        }
    }else{
        if (frm.cksel.checked){
            ischecked = true;
            if (!IsInteger(frm.Arrrealstock.value)){
                alert('정수만 가능합니다.');
                frm.Arrrealstock.focus();
                return;
            }
        }
    }
    
    if (!(ischecked)){
        alert('선택된 상품이 없습니다.');
        return;
    }
    
    if (confirm('실사 재고를 저장 하시겠습니까?')){
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

function cpStock(i,ival){
    var frm = document.frmArr;
    if (frm.Arrrealstock.length){
        frm.Arrrealstock[i].value = ival;
        frm.cksel[i].checked = !(frm.cksel[i].checked);
        var comp=frm.cksel[i];
    }else{
        frm.Arrrealstock.value = ival;
        frm.cksel.checked = !(frm.cksel.checked);
        var comp=frm.cksel;
    }
    
    AnCheckClick(comp);
    
}

function MiALLZero(){
    var frm = document.frmArr;
    if (!frm.cksel) return;
    
    if (frm.cksel.length){
        for (i=0;i<frm.cksel.length;i++){
            if (!frm.cksel[i].checked){
                frm.Arrrealstock[i].value = 0;
                CheckThis(i);
                
            }
        }
    }else{
        if (!frm.cksel.checked){
            frm.Arrrealstock.value = 0;
            CheckThis(0);
        }
    }
}

function nextStockStep(nVal){
    var frm = document.frmup;
    if (document.frmStockDt.stockdate){
        var stockdate = document.frmStockDt.stockdate.value;
    }else{
        var stockdate = "NULL"
    }
    frm.stStatus.value = nVal;
    frm.stockdate.value = stockdate;
    
    if (nVal==3){
        var confirmStr = "재고 반영 요청 하시겠습니까?";
    }else if (nVal==0){
        var confirmStr = "재고파악중 상태로 변경 하시겠습니까?";
    }else{
        var confirmStr = "수정 하시겠습니까?";
    }
    
    if (confirm(confirmStr)){
        frm.submit();
    }
}
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="SType" value="<%= SType %>">
	<input type="hidden" name="makerid" value="<%= makerid %>">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="idx" value="<%= idx %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
    		    <input type="hidden" name="shopid" value="<%= shopid %>">
    		    매장 : <%= shopid %>
		    
		        <input type="hidden" name="makerid" value="<%= makerid %>">
    			브랜드 :
    			<%= makerid %> &nbsp;&nbsp;
			<br>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			* 실사파악재고 부분을 더블클릭 하시면 실사 값으로 지정됩니다.
		</td>
	</tr>
	
	</form>
</table>
<!-- 검색 끝 -->


<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" > 
    <form name="frmStockDt">
	<tr height="30">
		<td align="left">
			<% if C_ADMIN_AUTH=true then %>
        	<!--
	        <input type="button" class="button" value="브랜드 전체 새로고침" onclick="RefreshIpchulStock();">
	        -->
	        <% end if %>
        	<!-- input type="button" class="button" name="stock_sheet_print" value="재고파악SHEET출력" onclick="javascript:PopOFFBrandStockSheet();" -->
		    <% if (C_IS_Maker_Upche) and (Not IsUpcheWitakItem) then %>
                (업체위탁 계약 매장만 재고 수정 가능)
            <% end if %> 
            <% if (SType="stTaking") then %>
            <input type="button" class="button" value="재고파악중 상태로 변경" onClick="nextStockStep(0);">
            &nbsp;
            <input type="button" class="button" value="미 선택 내역 0 처리" onClick="MiALLZero();">
            <% end if %>
		</td>
		<td align="right">
		    재고파악일 : <input type="text" class="text" name="stockdate" value="<%= StockDate %>" size=11 readonly ><a href="javascript:calendarOpen(frmStockDt.stockdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21>
			<input type="button" class="button" name="stock_sheet_print" value="선택 상품 실사재고 일괄입력" onclick="RealStockInputArr();"> 
		</td>
	</tr>
	</form>
</table>
<!-- 액션 끝 -->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frmArr" method="post" action="/common/offshop/shop_stockrefresh_process.asp">
    <input type="hidden" name="mode" value="ArrOfferrcheckupdate">
    <input type="hidden" name="shopid" value="<%= shopid %>">
    <input type="hidden" name="makerid" value="<%= makerid %>">
    <input type="hidden" name="SType" value="<%= SType %>">
    <input type="hidden" name="idx" value="<%= idx %>">
    <input type="hidden" name="stTakingIdx" value="<%= idx %>">
    <input type="hidden" name="stockdate" value="">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="20"></td>
        <td width="30">구분</td>
    	<td width="40">상품ID</td>
    	<td width="40">옵션</td>
    	<td width="50">이미지</td>
    	<td>상품명<br>[옵션명]</td>
    	<td>매장<br>매입가</td>
    	<td>판매가</td>
    	<!-- td width="40">센터<br>매입<br>구분</td -->
    	<td width="40">센터<br>입고<br>반품</td>
    	<td width="40">브랜드<br>입고<br>반품</td>
        <td width="40">매장<br>판매<br>반품</td>
        <td width="40" bgcolor="F4F4F4">시스템<br>총재고</td>
        <td width="40">총<br>실사<br>오차</td>
        <td width="40" bgcolor="F4F4F4">현재<br>실사<br>재고</td>
        <td width="40">실사<br>입력</td>
        <td width="40" bgcolor="F4F4F4"><strong>실사<br>파악<br>재고</strong></td>
        <td width="30">사용<br>여부</td>
        <td width="40">개별<br>입력</td>
    </tr>
<% if oOffStock.FResultCount<1 then %>
    <tr align="center" bgcolor="#FFFFFF" height="30">
        <% if (shopid="") and (makerid="") then %>
        <td colspan="21" >[ 매장 및 브랜드를 선택 하세요. ]</td>
        <% else %>
        <td colspan="21" >[ 검색 결과가 없습니다. ]</td>
        <% end if %>
    </tr>
<% else %>
    <%
    dim totalBuycash ,totalshopitemprice , totallogicsipgono , totalbrandreipgono ,totalresellno, totalbuycashSum
    dim totalsysstockNo , totalerrrealcheckno , totalAvailStock, totalerrsampleitemno, totalRealStock, totbatchItemNo
    for i=0 to oOffStock.FResultCount - 1 
    %>
    <%
    totalBuycash = totalBuycash + oOffStock.FItemList(i).GetOfflineSuplycash
    totalshopitemprice = totalshopitemprice + oOffStock.FItemList(i).fshopitemprice
    totallogicsipgono = totallogicsipgono + oOffStock.FItemList(i).Flogicsipgono + oOffStock.FItemList(i).Flogicsreipgono
    totalbrandreipgono = totalbrandreipgono + oOffStock.FItemList(i).Fbrandipgono + oOffStock.FItemList(i).Fbrandreipgono
    totalresellno = totalresellno + oOffStock.FItemList(i).Fsellno + oOffStock.FItemList(i).Fresellno
    totalsysstockNo = totalsysstockNo + oOffStock.FItemList(i).FsysstockNo
    totalerrrealcheckno = totalerrrealcheckno + oOffStock.FItemList(i).Ferrrealcheckno
    
    totalRealStock       = totalRealStock + oOffStock.FItemList(i).Frealstockno
    totalerrsampleitemno = totalerrsampleitemno + oOffStock.FItemList(i).Ferrsampleitemno
    totalAvailStock = totalAvailStock + oOffStock.FItemList(i).getAvailStock
    
    totalbuycashSum = totalbuycashSum + oOffStock.FItemList(i).Frealstockno*oOffStock.FItemList(i).GetOfflineSuplycash
    totbatchItemNo  = totbatchItemNo + NULL2Zero(oOffStock.FItemList(i).FbatchItemNo)
    
    %>
    	<% if Not isNULL(oOffStock.FItemList(i).FbatchItemNo) then %>
        <tr bgcolor="#FFFFFF" align="center" class="H">
        <% else %>
        <tr bgcolor="#FFFFFF" align="center">
        <% end if %>
            <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" value="<%= i %>" <%= CHKIIF(Not IsNULL(oOffStock.FItemList(i).FbatchItemNo),"checked","") %> >
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
        	<td><%= FormatNumber(oOffStock.FItemList(i).GetOfflineSuplycash,0) %></td>
        	<td><%= FormatNumber(oOffStock.FItemList(i).fshopitemprice,0) %></td>        	            
            <!-- td><%= fnColor(oOffStock.FItemList(i).FCenterMwdiv,"mw") %></td -->
        	<td><%= oOffStock.FItemList(i).Flogicsipgono + oOffStock.FItemList(i).Flogicsreipgono %></td>
        	<td><%= oOffStock.FItemList(i).Fbrandipgono + oOffStock.FItemList(i).Fbrandreipgono %></td>
        	<td><%= oOffStock.FItemList(i).Fsellno + oOffStock.FItemList(i).Fresellno %></td>
        	<td bgcolor="F4F4F4"><b><%= oOffStock.FItemList(i).FsysstockNo %></b></td>
        	<td><%= oOffStock.FItemList(i).Ferrrealcheckno %></td>
        	<td bgcolor="F4F4F4"><b><font color="<%= ChkIIF(oOffStock.FItemList(i).Frealstockno<0,"#FF0000","#000000") %>"><%= oOffStock.FItemList(i).Frealstockno %></font></b></td>
        	<td>
        	    <% if isNULL(oOffStock.FItemList(i).FbatchItemNo) then %>
            	    <% if (FALSE) then %>
            	    <input type="text" class="text" name="Arrrealstock" value="<%= ChkIIF(oOffStock.FItemList(i).Frealstockno<1,0,oOffStock.FItemList(i).Frealstockno) %>" size="4" maxlength="4" AUTOCOMPLETE="off" style="text-align=center" onKeyDown="CheckThis('<%= i %>');">
            	    <% else %>
            	    <input type="text" class="text" name="Arrrealstock" value="<%= oOffStock.FItemList(i).Frealstockno %>" size="4" maxlength="4" AUTOCOMPLETE="off" style="text-align=center" onKeyDown="CheckThis('<%= i %>');">
            	    <% end if %>
        	    <% else %>
        	    <input type="text" class="text" name="Arrrealstock" value="<%= oOffStock.FItemList(i).FbatchItemNo %>" size="4" maxlength="4" AUTOCOMPLETE="off" style="text-align=center" onKeyDown="CheckThis('<%= i %>');">
        	    <% end if %>
        	</td>
        	<td ondblclick="cpStock(<%= i %>,<%= NULL2Zero(oOffStock.FItemList(i).FbatchItemNo) %>);"><%= oOffStock.FItemList(i).FbatchItemNo %></td>
        	<td>
        	    <% if oOffStock.FItemList(i).Fisusing="N" then %>
        	    <strong><%= oOffStock.FItemList(i).Fisusing %></strong>
        	    <% else %>
        	    <%= oOffStock.FItemList(i).Fisusing %>
        	    <% end if %>
        	</td>
        	<td>
        		<input type="button" class="button" value="실사" onclick="popOffErrInput('<%= shopid %>','<%= oOffStock.FItemList(i).Fitemgubun %>','<%= oOffStock.FItemList(i).Fitemid %>','<%= oOffStock.FItemList(i).Fitemoption %>');">
        	</td>
        </tr>
    <% next %>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td colspan=6></td>
    	<td><%= FormatNumber(totalBuycash,0) %></td>
    	<td><%= FormatNumber(totalshopitemprice,0) %></td>
    	<!-- td></td -->
    	<td><%= FormatNumber(totallogicsipgono,0) %></td>
    	<td><%= FormatNumber(totalbrandreipgono,0) %></td>
        <td><%= FormatNumber(totalresellno,0) %></td>
        <td><%= FormatNumber(totalsysstockNo,0) %></td>
        <td><%= FormatNumber(totalerrrealcheckno,0) %></td>
        <td><%= FormatNumber(totalRealStock,0) %></td>
        <td></td>
        <td><%= FormatNumber(totbatchItemNo,0) %></td>
        <td></td>
        <td></td>
    </tr>    
<% end if %>
    </form>
</table> 

<%
set oOffStock = Nothing
%>

<form name="frmup" method="post" action="/common/offshop/shop_stockrefresh_process.asp">
<input type="hidden" name="mode" value="stockTakingNext">
<input type="hidden" name="stTakingIdx" value="<%= idx %>">
<input type="hidden" name="stStatus" value="">
<input type="hidden" name="stockdate" value="">
</form>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" --> 