<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/itembarcodecls.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->

<%
dim itembarcode
dim itemgubun,itemid,itemoption
dim actType, makerid

itembarcode  = requestCheckVar(request("itembarcode"),20)
itemgubun 	 = requestCheckVar(request("itemgubun"),2)
itemid		 = requestCheckVar(request("itemid"),10)
itemoption	 = requestCheckVar(request("itemoption"),4)
actType      = requestCheckVar(request("actType"),10)

if (C_IS_Maker_Upche) then
    makerid = session("ssBctId")
end if

if (Len(itembarcode)=12) then
	itemgubun 	= left(itembarcode,2)
	itemid		= CLng(mid(itembarcode,3,6))
	if (itemoption="") then itemoption = right(itembarcode,4)
	itembarcode = itemgubun + Format00(6,itemid) + itemoption
elseif (Len(itembarcode)=14) then
	itemgubun 	= left(itembarcode,2)
	itemid		= CLng(mid(itembarcode,3,8))
	if (itemoption="") then itemoption = right(itembarcode,4)
	itembarcode = itemgubun + Format00(8,itemid) + itemoption
elseif (Len(itembarcode)<>0) and (itemid<>"") then
	itemgubun = "10"
	itemid = itembarcode
	if (itemoption="") then itemoption  = "0000"
elseif (Len(itembarcode)>6) then
    '''바코드인경우 검색후 상품코드 가져옴.
    call fnGetItemCodeByPublicBarcode(itembarcode, itemgubun, itemid, itemoption)
else
    itemgubun = "10"
    if (itemid="") then itemid = itembarcode
    if (itemoption="") then itemoption  = "0000"

    if (itemid>=1000000) then
        itembarcode = itemgubun + Format00(8,itemid) + itemoption
    else
        itembarcode = itemgubun + Format00(6,itemid) + itemoption
    end if
end if




dim oitem
set oitem = new CItem
oitem.FRectMakerid= makerid
oitem.FRectItemID = itemid

if itemid<>"" then
	oitem.GetOneItem
end if

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if


dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FRectItemGubun = itemgubun
osummarystock.FRectItemID =  itemid
osummarystock.FRectItemOption =  itemoption
if itemid<>"" then
	osummarystock.GetCurrentItemStock
end if


dim otodayerritem
set otodayerritem = new CSummaryItemStock
otodayerritem.FRectItemGubun = itemgubun
otodayerritem.FRectItemID =  itemid
otodayerritem.FRectItemOption =  itemoption
if itemid<>"" then
    otodayerritem.GetTodayErrItem
end if

dim difftime
if (osummarystock.FResultcount>0) then
    difftime = ABS(datediff("h",osummarystock.FOneItem.Flastupdate,now()))
end if

dim i
dim IsVaildCode
IsVaildCode = False
if (oitemoption.FResultCount>0) then
    for i=0 to oitemoption.FResultCount-1
        if (oitemoption.FITemList(i).FItemOption=itemoption) then
            IsVaildCode = (oitem.FResultCount>0)
            exit For
        end if
    next
else
    IsVaildCode = (oitem.FResultCount>0) and (itemoption="0000")
end if


dim ErrMsg, sqlStr
dim SelectedOptionStr
dim stockReipgoDate

sqlStr = "select top 1 stockReipgoDate from [db_item].[dbo].tbl_item_option_Stock"
sqlStr = sqlStr & " where itemgubun='" & itemgubun & "'"
sqlStr = sqlStr & " and itemid=" & itemid
sqlStr = sqlStr & " and itemoption='" & itemoption & "'"

rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
    stockReipgoDate = rsget("stockReipgoDate")
end if
rsget.Close

%>
<script language='javascript'>

var delCliked = false;

function ClearVal(comp){
    comp.value='';
    delCliked = true;
}

//재입고일 설정.
function SaveStockReipgoDate(frm){
    var nowDate = "<%= Left(Now(),10) %>";

    if ((frm.stockReipgoDate.value.length<1)&&(!delCliked)){
        alert('재입고 예정일을 설정해 주세요.');
        return;
    }



    if (frm.stockReipgoDate.value.length<1){
        if (confirm('재입고 예정일이 삭제 됩니다. 계속 하시겠습니까?')){
            frm.mode.value = "stockreipgodate";
            frm.submit();
        }
    }else{
        if (frm.stockReipgoDate.value<nowDate){
            alert('재입고 예정일은 금일 이후로 설정하세요.');
            return;
        }

        if (confirm('재입고 예정일을 설정 하시겠습니까?')){
            frm.mode.value = "stockreipgodate";
            frm.submit();
        }
    }
}

//단종 설정
function SaveDanjongSoldOut(frm){
    //단종설정은 한정판매인경우만 가능함
	if (frm.isEditValid.value==""){
		alert('한정 판매인 경우만 재고부족,단종품절, MD품절로 설정 할 수 있습니다.');
		//frm.limityn[0].focus();
		return;
	}

    if (confirm('단종 품절 처리 진행하시겠습니까?')){
        frm.mode.value = "danjong";
        frm.submit();
    }
}

//MD품절 설정
function SaveMdSoldOut(frm){
    //단종설정은 한정판매인경우만 가능함
	if (frm.isEditValid.value==""){
		alert('한정 판매인 경우만 재고부족,단종품절, MD품절로 설정 할 수 있습니다.');
		//frm.limityn[0].focus();
		return;
	}

    if (confirm('MD 품절 처리 진행하시겠습니까?')){
        frm.mode.value = "mssoldout";
        frm.submit();
    }
}



function GetOnLoad(){
	<% if Not IsVaildCode then %>
    	<% if oitemoption.FResultCount>0 then %>
    	alert('상품코드가 정확하지 않습니다. 옵션 선택후 재검색 하세요.');
    	<% else %>
    	alert('상품코드가 정확하지 않습니다. 재검색 하세요.');
    	<% end if %>
	document.frm.itembarcode.select();
	document.frm.itembarcode.focus();
	<% end if %>
}
window.onload=GetOnLoad;

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="actType" value="<%= actType %>">
	<!-- 상단바 시작 -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="2">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
				        <font color="red">&nbsp;
				        <strong>
				        <% if (actType="D") then %>
				        단종설정
				        <% elseif (actType="R") then %>
				        재입고예정일 설정
				        <% end if %>
				        </strong></font>
				    </td>
				    <td align="right">
						상품코드:
						<% if (C_IS_Maker_Upche) then %>
						<input type=text class="text_ro" ReadOnly name=itembarcode value="<%= itembarcode %>" size=16 maxlength=16 UTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ Research(frmbar); return false;}">
						<% else %>
						<input type=text class="text" name=itembarcode value="<%= itembarcode %>" size=16 maxlength=16 UTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ Research(frmbar); return false;}">
						<% end if %>
						&nbsp;

						<% if oitemoption.FResultCount>0 then %>

						<select class="select" name="itemoption">
						<option value="0000">----
						<% for i=0 to oitemoption.FResultCount-1 %>
						<option value="<%= oitemoption.FITemList(i).FItemOption %>" <% if itemoption=oitemoption.FITemList(i).FItemOption then response.write "selected" %> ><%= oitemoption.FITemList(i).FOptionName %>
						<% next %>
						</select>
						<% end if %>

        				<input type="button" class="button" value="검색" onclick="document.frm.submit();">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<!-- 상단바 끝 -->
	</form>
</table>

<p>
<% if (oitem.FResultCount<1) then %>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
        <td align="center">[검색 결과가 없습니다.]</td>
    </tr>
</table>
<% else %>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
    	<td rowspan=<%= 6 + oitemoption.FResultCount -1 %> width="110" valign=top align=center><img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100"></td>
      	<td width="60">상품코드</td>
      	<td width="300">
      		10<Strong><%= CHKIIF(oitem.FOneItem.FItemID>=1000000,Format00(8,oitem.FOneItem.FItemID),Format00(6,oitem.FOneItem.FItemID)) %></Strong><%= itemoption %>
      	</td>
      	<td width="60"></td>
      	<td colspan=2></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>브랜드ID</td>
      	<td><%= oitem.FOneItem.FMakerid %></td>
      	<td>판매여부</td>
      	<td colspan=2><%= fnColor(oitem.FOneItem.FSellyn,"yn") %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>상품명</td>
      	<td><%= oitem.FOneItem.FItemName %></td>
      	<td>사용여부</td>
      	<td colspan=2><%= fnColor(oitem.FOneItem.FIsUsing,"yn") %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>판매가</td>
      	<td>
      		<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
      		&nbsp;&nbsp;
      		<font color="<%= mwdivColor(oitem.FOneItem.FMwDiv) %>"><%= oitem.FOneItem.getMwDivName %></font>
      	    <% if oitem.FOneItem.FSellcash<>0 then %>
			<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
			<% end if %>
			&nbsp;&nbsp;
			<!-- 할인여부/쿠폰적용여부 -->
			<% if (oitem.FOneItem.FSailYn="Y") then %>
			    <font color=red>
			    <% if (oitem.FOneItem.Forgprice<>0) then %>
			        <%= CLng((oitem.FOneItem.Forgprice-oitem.FOneItem.Fsellcash)/oitem.FOneItem.Forgprice*100) %> %
			    <% end if %>
			     할인
			    </font>
			<% end if %>

			<% if (oitem.FOneItem.Fitemcouponyn="Y") then %>

			    <font color=green><%= oitem.FOneItem.GetCouponDiscountStr %> 쿠폰
			    (<%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %>)</font>
			<% end if %>

      	</td>
      	<td>단종여부</td>
      	<td colspan=2>
      		<% if oitem.FOneItem.Fdanjongyn="Y" then %>
			<font color="#33CC33">단종</font>
			<% elseif oitem.FOneItem.Fdanjongyn="M" then %>
			<font color="#CC3333">MD품절</font>
			<% elseif oitem.FOneItem.Fdanjongyn="S" then %>
			<font color="#33CC33">일시품절</font>
			<% else %>
			생산중
			<% end if %>
		</td>
    </tr>
     <% if oitemoption.FResultCount>1 then %>
	    <!-- 옵션이 있는경우 -->
	    <% for i=0 to oitemoption.FResultCount -1 %>
	    	<% if oitemoption.FITemList(i).Fitemoption=itemoption then %>
	    	<% SelectedOptionStr = "<font color=blue>[" & oitemoption.FITemList(i).FOptionName & "]</font>" %>
	    	<tr bgcolor="#FFFFFF">
	    		<td>옵션명</td>
		      	<td><%= oitemoption.FITemList(i).FOptionName %> (<%= fnColor(oitemoption.FITemList(i).FOptIsUsing,"yn") %>)</td>
		      	<td>한정여부</td>
		      	<td><%= fnColor(oitem.FOneItem.Flimityn,"yn") %> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
		      	<td>한정 비교재고 (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
		    </tr>
		    <% end if %>
		<% next %>
	<% else %>
    	<tr bgcolor="#FFFFFF">
	      	<td>옵션명</td>
	      	<td>-</td>
	      	<td>한정여부</td>
	      	<td><%= fnColor(oitem.FOneItem.Flimityn,"yn") %> (<%= oitem.FOneItem.GetLimitEa %>)</font></td>
	      	<td>한정 비교재고 (<b><%= oitem.FOneItem.GetLimitStockNo %></b>)</td>
	    </tr>
    <% end if %>

</table>
<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name=frm2 method=post action="itemdanjong_process.asp">
    <input type=hidden name=mode value="">
    <input type=hidden name=itemgubun value="<%= itemgubun %>">
    <input type=hidden name=itemid value="<%= itemid %>">
    <input type=hidden name=itemoption value="<%= itemoption %>">

    <% if (actType="D") and (oitem.FOneItem.FLimityn<>"Y") then %>
    <input type=hidden name=isEditValid value="">
    <% else %>
    <input type=hidden name=isEditValid value="on">
    <% end if %>

<!-- 실시간 업데이트 됨
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td colspan="16" align=right>최종업데이트 : <%= osummarystock.FOneItem.Flastupdate %> </td>
    </tr>
-->
    <tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="50">총입고/반품</td>
    	<td width="50">총판매/반품</td>
		<td width="50">샾출고/반품</td>
		<td width="50">기타출고/반품</td>
		<td width="50">CS<br>출고/반품</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">시스템<br>총재고</td>
		<td width="50">총불량</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">시스템<br>유효재고</td>
		<td width="50">총실사<br>오차</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">실사<br>재고</td>
		<td width="50">ON상품<br>준비</td>
		<td width="50">OFF상품<br>준비</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">재고파악<br>재고</td>
		<td width="50">ON결제<br>완료</td>
		<td width="50">ON주문<br>접수</td>
		<td width="50">OFF주문<br>접수</td>
    </tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td rowspan="2"><%= osummarystock.FOneItem.Ftotipgono %></td>
    	<td rowspan="2"><%= -1*osummarystock.FOneItem.Ftotsellno %></td>
    	<td rowspan="2"><%= osummarystock.FOneItem.Foffchulgono + osummarystock.FOneItem.Foffrechulgono %></td>
    	<td rowspan="2"><%= osummarystock.FOneItem.Fetcchulgono + osummarystock.FOneItem.Fetcrechulgono %></td>
    	<td rowspan="2"><%= osummarystock.FOneItem.Ferrcsno %></td>
    	<td rowspan="2" bgcolor="<%= adminColor("gray") %>"><%= osummarystock.FOneItem.Ftotsysstock %></td>
    	<td rowspan="2"><%= osummarystock.FOneItem.Ferrbaditemno %></td>
    	<td rowspan="2" bgcolor="<%= adminColor("gray") %>"><%= osummarystock.FOneItem.Favailsysstock %></td>
    	<td rowspan="2" ><%= osummarystock.FOneItem.Ferrrealcheckno %></td>
    	<td rowspan="2" bgcolor="<%= adminColor("gray") %>"><%= osummarystock.FOneItem.Frealstock %></td>
    	<td><%= osummarystock.FOneItem.Fipkumdiv5 %></td>
    	<td><%= osummarystock.FOneItem.Foffconfirmno %></td>
    	<td rowspan="2" bgcolor="<%= adminColor("gray") %>"><%= osummarystock.FOneItem.GetCheckStockNo %></td>
    	<td><%= osummarystock.FOneItem.Fipkumdiv4 %></td>
    	<td><%= osummarystock.FOneItem.Fipkumdiv2 %></td>
    	<td><%= osummarystock.FOneItem.Foffjupno %></td>

    </tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td colspan="2"><%= osummarystock.FOneItem.Fipkumdiv5 + osummarystock.FOneItem.Foffconfirmno %></td>
    	<td colspan="3"><%= osummarystock.FOneItem.Fipkumdiv4 + osummarystock.FOneItem.Fipkumdiv2 + osummarystock.FOneItem.Foffjupno %></td>

    </tr>


</table>
<p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
<tr>
    <td align="center">
        <% if (actType="D") then %>
            <% if (oitemoption.FResultCount>0) then %>
            (단종 설정시 모든 옵션 포함 단종 설정됩니다)
            <br>
            <% end if %>
            <input type="button" class="button" value="단종품절설정" onclick="SaveDanjongSoldOut(frm2)">
            <% if Not (C_IS_Maker_Upche) then %>
            &nbsp;&nbsp;
            <input type="button" class="button" value="MD품절설정" onclick="SaveMdSoldOut(frm2)">
            <% end if %>
        <% elseif (actType="R") then %>
          <% if ErrMsg<>"" then %>
            <%= ErrMsg %>
          <% else %>
          <%= SelectedOptionStr %> 재입고 예정일 : <input type="text" class="text" name="stockReipgoDate" size="10" value="<%= stockReipgoDate %>" readOnly>
          <a href="javascript:calendarOpen(frm2.stockReipgoDate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
  	      <a href="javascript:ClearVal(frm2.stockReipgoDate);"><img src="/images/icon_delete2.gif" width="20" border="0" align="absmiddle"></a>
          <br><br>
          <input type="button" class="button" value="재입고 예정일 설정" onclick="SaveStockReipgoDate(frm2)">
          <% end if %>
        <% end if %>
    </td>
</tr>
</form>
</table>



<% end if %>
<%
set otodayerritem = Nothing
set osummarystock = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->