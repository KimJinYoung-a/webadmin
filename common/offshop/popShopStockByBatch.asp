<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프샾 재고
' Hieditor : 2009.11.17 서동석 생성
'			 2011.05.06 한용민 수정
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
        alert('먼저 매장과 브랜드를 선택후 출력해 주세요.');
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
    
    //2개월 이전으로 실사입력 불가.
    if (stockdate<'<%= BasicMonth %>'){
		alert('두달 이전 날짜로는 재고파악일을 사용 할 수 없습니다.');
		return;
	}
	
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
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left" >
	    매장 : <%= shopid %>
	    &nbsp;&nbsp;
	    작업번호 : <%= jobkey %>
	</td>
	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		사용구분 : <% drawSelectBoxUsingYN "usingyn", usingyn %> &nbsp;&nbsp;
		센터매입구분 :
		   <select class="select" name="centermwdiv">
           <option value="">전체</option>
           <option value="MW" <%= ChkIIF(centermwdiv="MW","selected","") %> >매입+특정</option>
           <option value="W"  <%= ChkIIF(centermwdiv="W","selected","") %> >특정</option>
           <option value="M"  <%= ChkIIF(centermwdiv="M","selected","") %> >매입</option>
           <option value="NULL" <%= ChkIIF(centermwdiv="NULL","selected","") %> >미지정</option>
           </select>
           
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
	    검색결과 : <%= oOffStock.FResultCount %> (최대 2,000건)
		<% if C_ADMIN_AUTH=true then %>
    	<!--
        <input type="button" class="button" value="브랜드 전체 새로고침" onclick="RefreshIpchulStock();">
        -->
        <% end if %>
        
	</td>
	<td align="right">
	    <input type="text" class="text" name="stockdate" value="<%= StockDate %>" size=11 readonly ><a href="javascript:calendarOpen(frmStockDt.stockdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21>
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
<input type="hidden" name="stockdate" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"><input type="checkbox" name="chkAll" onClick="switchCheckBox(this);"></td>
    <td width="30">구분</td>
	<td width="40">상품ID</td>
	<td width="40">옵션</td>
	<td width="50">이미지</td>
	<td>상품명<br>[옵션명]</td>
	<td width="40">센터<br>매입<br>구분</td>
	<td width="40">센터<br>입고<br>반품</td>
	<td width="40">브랜드<br>입고<br>반품</td>
    <td width="40">매장<br>판매<br>반품</td>
    <td width="40" bgcolor="F4F4F4">시스템<br>총재고</td>
    <td width="40">총<br>실사<br>오차</td>
    <!-- <td width="40" bgcolor="F4F4F4">실사<br>재고</td> 
    <td width="40">총<br>샘플</td>
    <td width="40">총<br>불량</td> -->
    <td width="40" bgcolor="F4F4F4">실사<br>재고</td>
    
    <td width="30">사용<br>여부</td>
    <!--
    <td width="30">판매<br>여부</td>
    <td width="30">한정<br>여부</td>
    <td width="50">단종<br>여부</td>
    -->
    <td width="40">실사<br>재고</td>
    <td width="40">개별<br>실사<br>입력</td>
</tr>
<% if oOffStock.FResultCount<1 then %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td colspan="20" >[ 검색 결과가 없습니다. ]</td>
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
		<input type="button" class="button" value="실사" onclick="popOffErrInput('<%= shopid %>','<%= oOffStock.FItemList(i).Fitemgubun %>','<%= oOffStock.FItemList(i).Fitemid %>','<%= oOffStock.FItemList(i).Fitemoption %>');">
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