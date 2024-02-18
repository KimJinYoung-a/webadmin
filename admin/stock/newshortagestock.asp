<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/newshortagestockcls.asp"-->
<%
const C_STOCK_DAY=7

''아래 두 페이지는 검색조건을 동일하게 맞춰야 한다.
''/admin/stock/newshortagestock.asp
''/admin/newstorage/popjumunitemNew.asp

dim page, mode, makerid, shopid,itemid, research
dim onlynotupchebeasong, onlyusingitem, onlyusingitemoption, onlynotdanjong, soldoutover7days, onlysell, onlynottempdanjong
dim onlynotmddanjong, includepreorder, skiplimitsoldout
dim onoffgubun, idx, shortagetype, onlystockminus
dim changemakerid
dim purchasetype
dim itemgubun, itemname
dim chkMinusStockGubun, minusStockGubun
dim mwdiv, excmkr, onlyOn

shopid = requestCheckVar(("shopid"),32)
page = requestCheckVar(request("page"),32)
mode = requestCheckVar(request("mode"),32)
itemid = requestCheckVar(request("itemid"),32)
research = requestCheckVar(request("research"),32)
onlynotupchebeasong = requestCheckVar(request("onlynotupchebeasong"),32)
onlyusingitem = requestCheckVar(request("onlyusingitem"),32)
onlyusingitemoption = requestCheckVar(request("onlyusingitemoption"),32)
onlynotdanjong = requestCheckVar(request("onlynotdanjong"),32)
soldoutover7days = requestCheckVar(request("soldoutover7days"),32)
onoffgubun = requestCheckVar(request("onoffgubun"),32)
idx = requestCheckVar(request("idx"),32)
shortagetype = requestCheckVar(request("shortagetype"),32)
onlysell = requestCheckVar(request("onlysell"),32)
onlynottempdanjong = requestCheckVar(request("onlynottempdanjong"),32)
onlynotmddanjong = requestCheckVar(request("onlynotmddanjong"),32)
includepreorder = requestCheckVar(request("includepreorder"),32)
skiplimitsoldout = requestCheckVar(request("skiplimitsoldout"),32)
onlystockminus = requestCheckVar(request("onlystockminus"),32)
purchasetype = requestCheckVar(request("purchasetype"),32)
itemgubun = requestCheckVar(request("itemgubun"),32)
itemname = requestCheckVar(request("itemname"),128)
chkMinusStockGubun = requestCheckVar(request("chkMinusStockGubun"),32)
minusStockGubun = requestCheckVar(request("minusStockGubun"),32)
mwdiv = requestCheckVar(request("mwdiv"),32)
excmkr = requestCheckVar(request("excmkr"),32)
onlyOn = requestCheckVar(request("onlyOn"),32)

changemakerid = "Y"
if (changemakerid = "") then
	changemakerid = requestCheckVar(request("changemakerid"),32)
end if

makerid = request("makerid")
if (makerid = "") then
	makerid = requestCheckVar(request("suplyer"),32)
end if


if (research<>"on") then
	excmkr = "Y"
    shortagetype = "14day"
    onlynotmddanjong = "on"
    includepreorder = "on"
end if

if (research<>"on") and (onlynotupchebeasong = "") then
	onlynotupchebeasong = "on"
end if

if (research<>"on") and (onlyusingitem = "") then
	onlyusingitem = "on"
end if

if (research<>"on") and (onlyusingitemoption="") then
	onlyusingitemoption = "on"
end if

if (research<>"on") and (onlynotdanjong = "") then
	onlynotdanjong = "on"
end if

if (research<>"on") and (onoffgubun="") then
	onoffgubun = "online"
end if

if (research<>"on") and (itemgubun="") then
	itemgubun = "10"
end if


if page="" then page=1
if mode="" then mode="bybrand"

'상품코드 유효성 검사(2008.07.31;허진원)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

dim oshortagestock
set oshortagestock  = new CShortageStock
oshortagestock.FPageSize = 50
oshortagestock.FCurrPage = page

oshortagestock.FRectOnlySell			= onlysell
oshortagestock.FRectOnlyUsingItem		= onlyusingitem
oshortagestock.FRectOnlyUsingItemOption	= onlyusingitemoption
oshortagestock.FRectOnlyNotUpcheBeasong	= onlynotupchebeasong

oshortagestock.FRectShortage7days		= chkIIF(shortagetype="7day","on","")
oshortagestock.FRectShortage14days		= chkIIF(shortagetype="14day","on","")
oshortagestock.FRectShortageRealStock	= chkIIF(shortagetype="5under","on","")
oshortagestock.FRectOnlyNotDanjong		= onlynotdanjong
oshortagestock.FRectOnlyNotTempDanjong	= onlynottempdanjong
oshortagestock.FRectOnlyNotMDDanjong	= onlynotmddanjong
oshortagestock.FRectIncludePreOrder		= includepreorder
oshortagestock.FRectSkipLimitSoldOut	= skiplimitsoldout
oshortagestock.FRectOnlyStockMinus		= onlystockminus

oshortagestock.FRectPurchaseType		= purchasetype

oshortagestock.FRectMakerid				= makerid
oshortagestock.FRectItemId				= itemid
'oshortagestock.FRectItemOption			= makerid

oshortagestock.FRectItemGubun			= itemgubun

if (chkMinusStockGubun = "Y") then
	oshortagestock.FRectMinusStockGubun			= minusStockGubun
end if

if (itemname <> "") then
	if (makerid <> "") then
		oshortagestock.FRectItemName			= itemname
	else
		response.write "<script>alert('먼저 브랜드를 선택하세요.');</script>"
	end if
end if

oshortagestock.FRectMWDiv				= mwdiv
oshortagestock.FRectExcMkr				= excmkr
oshortagestock.FRectOnlyOn				= onlyOn

if (itemgubun = "10") then
	oshortagestock.GetShortageItemListOnline
else
	oshortagestock.GetShortageItemListOffline
end if



dim i, shopsuplycash, buycash
dim IsAvailDelete



'==============================================================================
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, nowdate, iStartDate, iEndDate

'재입고예정일
'오늘기준 +- 일주일은 검정색 표시 / 그 이외 회색표시
if (yyyy1="") then
    nowdate = Left(CStr(DateAdd("d",now(),-7)),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

    nowdate = Left(CStr(DateAdd("d",now(),+7)),10)
	yyyy2 = Left(nowdate,4)
	mm2   = Mid(nowdate,6,2)
	dd2   = Mid(nowdate,9,2)
end if

iStartDate  = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
iEndDate    = Left(CStr(DateSerial(yyyy2,mm2,dd2)),10)

%>

<script language='javascript'>
function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ChangeReqDay(frm){
	if (!(IsDigit(frm.maxsellday.value))){
		alert('숫자만 가능합니다.');
		return;
	}

	if (confirm('필요 재고 기준일을 변경하시겠습니까?')){
		frm.submit();
	}
}

function Research(page){
	document.frm.page.value= page;
	document.frm.submit();
}

function DeleteStockLog(itemgubun,itemid,itemoption){
    if (confirm('삭제 하시겠습니까?')){
        frmdelstock.target="_blank";
        frmdelstock.itemgubun.value = itemgubun;
        frmdelstock.itemid.value = itemid;
        frmdelstock.itemoption.value = itemoption;
        frmdelstock.submit();
    }
}

function search(frm){
	/*
	if ((frm.makerid.value.length<1)){
		if ((frm.mode[0].checked)&&(frm.designer.value.length<1)){
			alert('브랜드를 선택 하세요.');
			frm.designer.focus();
			return;
		}
	}
	*/
	document.frm.page.value = 1;
	frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="page" value="<%= page %>">
	<% if (changemakerid <> "Y") then %>
	<input type="hidden" name="makerid" value="<%= makerid %>" >
	<% else %>
	<input type="hidden" name="changemakerid" value="Y" >
	<% end if %>
	<input type="hidden" name="shopid" value="<%= shopid %>" >
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<% if (changemakerid <> "Y") then %>
			브랜드 : <b><%= makerid %></b>
			<% else %>
			브랜드 : <% drawSelectBoxDesignerwithName "makerid", makerid %>
			<% end if %>
			&nbsp;
			|
			&nbsp;
			구분 :
			<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
			<!--
			<select class="select" name="itemgubun">
				<option value="10" <% if (itemgubun = "10") then %>selected<% end if %> >온라인(10)</option>
				<option value="90" <% if (itemgubun = "90") then %>selected<% end if %> >오프라인(90)</option>
				<option value="70" <% if (itemgubun = "70") then %>selected<% end if %> >사은품 등(70)</option>
				<option value="80" <% if (itemgubun = "80") then %>selected<% end if %> >사은품 등(80)</option>
				<option value="XX" <% if (itemgubun = "XX") then %>selected<% end if %> >기타</option>
			</select>
			-->
			&nbsp;
			|
			&nbsp;
			<input type=checkbox name="onlyusingitem" <% if onlyusingitem = "on" then response.write "checked" %> >사용상품만
			<input type=checkbox name="onlyusingitemoption" <% if onlyusingitemoption = "on" then response.write "checked" %> >사용옵션만
			<input type=checkbox name="onlysell" <% if onlysell = "on" then response.write "checked" %> >판매상품만
			<input type=checkbox name="onlynotupchebeasong" <% if onlynotupchebeasong = "on" then response.write "checked" %> >업체배송제외
		</td>

		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:search(frm);">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
            부족구분:
            <input type="radio" name="shortagetype" value="" <% if shortagetype="" then response.write "checked" %> >전체
            <input type="radio" name="shortagetype" value="7day" <% if shortagetype="7day" then response.write "checked" %> ><%= C_STOCK_DAY %>일후재고부족
			<input type="radio" name="shortagetype" value="14day" <% if shortagetype="14day" then response.write "checked" %> ><%= C_STOCK_DAY*2 %>일후재고부족
            <input type="radio" name="shortagetype" value="5under" <% if shortagetype="5under" then response.write "checked" %> >실사유효재고 5이하
			&nbsp;
			|
			&nbsp;
			<input type=checkbox name="onlynotdanjong" <% if onlynotdanjong = "on" then response.write "checked" %> >단종제외(옵션포함)
			<input type=checkbox name="onlynottempdanjong" <% if onlynottempdanjong = "on" then response.write "checked" %> >일시품절제외(옵션포함)
			<input type=checkbox name="onlynotmddanjong" <% if onlynotmddanjong = "on" then response.write "checked" %> >MD품절제외(옵션포함)


		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size=8 maxlength=7>
			&nbsp;
			상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size=16 maxlength=16>
			&nbsp;
			|
			&nbsp;
            <input type=checkbox name="includepreorder" <% if includepreorder = "on" then response.write "checked" %> >기주문포함부족만
            <!--
            <input type=checkbox name="skiplimitsoldout" <% if skiplimitsoldout = "on" then response.write "checked" %> >한정&판매중지제외
            -->
            <input type=checkbox name="onlystockminus" <% if onlystockminus = "on" then response.write "checked" %> >실사유효재고마이너스만
			&nbsp;
			|
			&nbsp;
			거래구분 :
			<select class="select" name="mwdiv">
				<option value="">-선택-</option>
				<option value="M" <% if (mwdiv = "M") then %>selected<% end if %> >매입</option>
				<option value="W" <% if (mwdiv = "W") then %>selected<% end if %> >특정</option>
				<option value="U" <% if (mwdiv = "U") then %>selected<% end if %> >업체</option>
				<option value="Z" <% if (mwdiv = "Z") then %>selected<% end if %> >미지정</option>
			</select>
			&nbsp;
			<% if (FALSE) then %>
			구매유형 : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", CHKIIF(purchasetype="", "1", purchasetype), "" %> <!-- 수정함. by eastone -->
			<% else %>
			구매유형 : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
		    <% end if %>
			&nbsp;
			|
			&nbsp;
			<input type="checkbox" name="chkMinusStockGubun" value="Y" <%if (chkMinusStockGubun = "Y") then %>checked<% end if %> >
			재고구분 :
			<select class="select" name="minusStockGubun">
				<option value="real" <%if (minusStockGubun = "real") then %>selected<% end if %> >실사유효재고</option>
				<option value="check" <%if (minusStockGubun = "check") then %>selected<% end if %> >재고파악재고</option>
				<option value="may" <%if (minusStockGubun = "may") then %>selected<% end if %> >예상재고</option>
			</select>
			마이너스만
			&nbsp;
			|
			&nbsp;
			<input type="checkbox" class="checkbox" name="excmkr" value="Y" <%= CHKIIF(excmkr="Y", "checked", "")%> > 아이띵소제외
			&nbsp;
			|
			&nbsp;
			<input type="checkbox" class="checkbox" name="onlyOn" value="Y" <%= CHKIIF(onlyOn="Y", "checked", "")%> > ON 판매만
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

* 브랜드 : <%= oshortagestock.FTotalMakeridCount %> / * 14일 부족재고(SKU) : <%= oshortagestock.FTotalCount %> / * 14일 부족재고(PCS) : <%= oshortagestock.FTotalPieceCount %>

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="22">
			검색결과 : <b><%= oshortagestock.FResultCount %></b>
			&nbsp;
			(최대검색건수 : <%= oshortagestock.FTotalCount %>)
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>브랜드ID</td>
		<td width="30">구분</td>
		<td width="40">상품<br>코드</td>
		<td width="50">이미지</td>
		<td>상품명<font color="blue">[옵션명]</font></td>
		<td width="50">마진</td>
		<td width="35">입고<br>반품<br>(B)</td>
		<td width="35">ON<br>판매<br>(D)</td>
		<td width="35">OFF<br>출고<br>(C)</td>
		<td width="35">기타<br>출고<br>(C)</td>
		<td width="35">CS<br>출고<br>(C)</td>
		<td width="35">누적<br>불량<br>(S)</td>
		<td width="35">누적<br>오차<br>(E)</td>
		<td width="35" bgcolor="#F3F3FF"><b>실사<br>유효<br>재고<br>(V)</b></td>
		<td width="35" bgcolor="#F3F3FF"><b>재고<br>파악<br>재고</b></td>
		<td width="35" bgcolor="#F3F3FF"><b>예상<br>재고</b></td>

		<td width="40">ON<br>(7일)<br>판매</td>
		<td width="40">OFF<br>(7일)<br>판매</td>

		<td width="50" bgcolor="#F3F3FF"><b>총(<%= C_STOCK_DAY %>일)<br>필요<br>수량</b></td>
		<td width="30">출고이전<br>필요수량 <!-- OFF<br>주문 --></td>
		<td width="30" bgcolor="#F3F3FF"><b>부족<br>수량</b></td>
		<td width="70">비고</td>
	</tr>
<% for i=0 to oshortagestock.FResultCount -1 %>
<%
    IsAvailDelete = (oshortagestock.FItemList(i).Ftotipgono=0) and (oshortagestock.FItemList(i).FtotSellNo=0) and (oshortagestock.FItemList(i).Fshortageno=0) and (oshortagestock.FItemList(i).Frealstock=0) and (oshortagestock.FItemList(i).Fpreorderno=0)
%>

	<% if oshortagestock.FItemList(i).IsInvalidOption then %>
	<tr align="center" bgcolor="#CCCCCC">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><a href="/admin/newstorage/orderinput.asp?suplyer=<%= oshortagestock.FItemList(i).FMakerID %>" target="iorderinput"><%= oshortagestock.FItemList(i).FMakerID %></a></td>
		<td><%= oshortagestock.FItemList(i).Fitemgubun %></td>
		<td><a href="javascript:PopItemSellEdit('<%= oshortagestock.FItemList(i).FItemID %>');"><%= oshortagestock.FItemList(i).FItemID %></a></td>
    	<td width="50" align=center><img src="<%= oshortagestock.FItemList(i).FimageSmall %>" width=50 height=50></td>
		<td align="left">
			<a href="/admin/stock/itemcurrentstock.asp?itemid=<%= oshortagestock.FItemList(i).FItemID %>&itemoption=<%= oshortagestock.FItemList(i).FItemOption %>" target=_blank ><%= oshortagestock.FItemList(i).FItemName %></a>
			<% if oshortagestock.FItemList(i).FItemOption <> "0000" then %>
				<% if oshortagestock.FItemList(i).Foptionusing="Y" then %>
					<br><font color="blue">[<%= oshortagestock.FItemList(i).FItemOptionName %>]</font>
				<% else %>
					<br><font color="#AAAAAA">[<%= oshortagestock.FItemList(i).FItemOptionName %>]</font>
				<% end if %>
			<% end if %>
		</td>
		<td>
			<font color="<%= oshortagestock.FItemList(i).getMwDivColor %>"><%= oshortagestock.FItemList(i).getMwDivName %></font><br>
			<% if oshortagestock.FItemList(i).Fsellcash<>0 then %>
			<%= 100-(CLng(oshortagestock.FItemList(i).Fbuycash/oshortagestock.FItemList(i).Fsellcash*10000)/100) %> %
			<% end if %>
		</td>
		<td><%= oshortagestock.FItemList(i).Ftotipgono %></td>
		<td><%= oshortagestock.FItemList(i).FtotSellNo %></td>
		<td><%= oshortagestock.FItemList(i).Foffchulgono + oshortagestock.FItemList(i).Foffrechulgono %></td>
		<td><%= oshortagestock.FItemList(i).Fetcchulgono + oshortagestock.FItemList(i).Fetcrechulgono %></td>
		<td></td>
		<td><%= oshortagestock.FItemList(i).Ferrbaditemno %></td>
		<td>
			<% if oshortagestock.FItemList(i).Ferrrealcheckno<0 then %>
			<font color="#cc3333"><%= oshortagestock.FItemList(i).Ferrrealcheckno %></font>
			<% else %>
				<%= oshortagestock.FItemList(i).Ferrrealcheckno %>
			<% end if %>
		</td>
		<td bgcolor="#F3F3FF"><b><%= oshortagestock.FItemList(i).Frealstock %></b></td>
		<td bgcolor="#F3F3FF"><b><%= oshortagestock.FItemList(i).GetCheckStockNo %></b></td>
		<td bgcolor="#F3F3FF"><b><%= oshortagestock.FItemList(i).GetMaystock %></b></td>

		<td><%= oshortagestock.FItemList(i).Fsell7days %></td>
		<td><%= oshortagestock.FItemList(i).Foffchulgo7days %></td>

		<td bgcolor="#F3F3FF"><b><%= oshortagestock.FItemList(i).Frequireno %></b></td>
		<td>
		    <!-- 출고이전 필요수량 -->
		    <%= (oshortagestock.FItemList(i).Fipkumdiv5 + oshortagestock.FItemList(i).Foffconfirmno+oshortagestock.FItemList(i).Fipkumdiv4 + oshortagestock.FItemList(i).Fipkumdiv2 + oshortagestock.FItemList(i).Foffjupno)*-1 %>
		</td>
		<td bgcolor="#F3F3FF"><b><%= oshortagestock.FItemList(i).Fshortageno %></b></td>
		<td>
			<%= fnColor(oshortagestock.FItemList(i).Fdanjongyn,"dj") %>
            <% if oshortagestock.FItemList(i).Foptdanjongyn="S" then %>
			<font color="#3333CC">옵션부족</font>
			<% end if %>
            <% if oshortagestock.FItemList(i).Foptdanjongyn="Y" then %>
			<font color="#33CC33">옵션단종</font><br>
			<% end if %>
            <% if oshortagestock.FItemList(i).Foptdanjongyn="M" then %>
			<font color="#CC3333">옵션MD</font><br>
			<% end if %>
			<br>
			<!-- 재고부족의 경우 재입고예정일 표시 -->
			<% if (oshortagestock.FItemList(i).Fdanjongyn = "S") or (oshortagestock.FItemList(i).Foptdanjongyn = "S") then %>
				<% if ((Not IsNull(oshortagestock.FItemList(i).FreipgoMayDate)) and (Left(oshortagestock.FItemList(i).FreipgoMayDate, 10) >= iStartDate) and (Left(oshortagestock.FItemList(i).FreipgoMayDate, 10) <= iEndDate)) then %>
					<%= Left(oshortagestock.FItemList(i).FreipgoMayDate, 10) %><br>
				<% elseif (Not IsNull(oshortagestock.FItemList(i).FreipgoMayDate)) then %>
					<font color="gray"><%= Left(oshortagestock.FItemList(i).FreipgoMayDate, 10) %></font><br>
				<% end if %>
			<% end if %>
			<% if oshortagestock.FItemList(i).Foptionusing="N" then %>
			<font color="red">옵션x</font><br>
			<% end if %>

			<% if oshortagestock.FItemList(i).IsSoldOut then %>
			<font color="red">품절</font><br>
			<% end if %>
			<% if oshortagestock.FItemList(i).Flimityn="Y" then %>
			<font color="blue">한정(<%= oshortagestock.FItemList(i).GetLimitStr %>)</font><br>
			<% end if %>

			<% if oshortagestock.FItemList(i).Fpreorderno<>0 then %>
			기주문:
	    		<% if oshortagestock.FItemList(i).Fpreorderno<>oshortagestock.FItemList(i).Fpreordernofix then response.write "</br>" + CStr(oshortagestock.FItemList(i).Fpreorderno) + " -> " %>
			<%= oshortagestock.FItemList(i).Fpreordernofix %>
			<% end if %>

			<% if IsAvailDelete then %>
			<a href="javascript:DeleteStockLog('10','<%= oshortagestock.FItemList(i).FItemID %>','<%= oshortagestock.FItemList(i).FItemOption %>');"><img src="/images/icon_delete.gif" border="0"></a>
			<% end if %>
		</td>
	</tr>
<% next %>
</table>



<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<% if oshortagestock.HasPreScroll then %>
		<a href="javascript:Research('<%= oshortagestock.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oshortagestock.StartScrollPage to oshortagestock.FScrollCount + oshortagestock.StartScrollPage - 1 %>
			<% if i>oshortagestock.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:Research('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oshortagestock.HasNextScroll then %>
			<a href="javascript:Research('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<%
set oshortagestock = Nothing
%>
<form name="frmdelstock" method="post" action="doshortagestockrefresh.asp">

<input type="hidden" name="mode" value="dellog">
<input type="hidden" name="itemgubun">
<input type="hidden" name="itemid">
<input type="hidden" name="itemoption">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
