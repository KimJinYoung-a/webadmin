<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  오프라인 [재고]월별 오차(로스) 내역
' History : 2009.04.07 이상구 생성
'			2010.04.02 한용민 수정
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->

<h3>작업중</h3>
<%
dim shopid, makerid, centermwdiv, itembarcode, research
dim itemgubun, itemid, itemoption, grpType
dim comm_cd

shopid       = RequestCheckVar(request("shopid"),32)
makerid      = RequestCheckVar(request("makerid"),32)
centermwdiv  = RequestCheckVar(request("centermwdiv"),10)
itembarcode  = RequestCheckVar(request("itembarcode"),32)
research     = RequestCheckVar(request("research"),2)
grpType      = RequestCheckVar(request("grpType"),10)
comm_cd      = RequestCheckVar(request("comm_cd"),32)

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim fromdate,todate

if request("yyyy1")<>"" then
	yyyy1 = RequestCheckVar(request("yyyy1"),4)
	mm1 = RequestCheckVar(request("mm1"),2)
	dd1 = "01"
else
	yyyy1 = Left(DateAdd("m",-1,now()),4)
	mm1 = Mid(DateAdd("m",-1,now()),6,2)
	dd1 = "01"
end if

if request("yyyy2")<>"" then
	yyyy2 = RequestCheckVar(request("yyyy2"),4)
	mm2 = RequestCheckVar(request("mm2"),2)
	dd2 = "01"
else
	yyyy2 = Left(DateAdd("m",-1,now()),4)
	mm2 = Mid(DateAdd("m",-1,now()),6,2)
	dd2 = "01"
end if

fromdate = CStr(DateSerial(yyyy1, mm1, dd1))
todate = CStr(DateAdd("m",1,DateSerial(yyyy2, mm2, dd2)))

if (grpType="") then grpType="M"

''매장
if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if

''업체
if (C_IS_Maker_Upche) then
    makerid = session("ssBctid")
end if

if (itembarcode <> "") then
    if Not (fnGetItemCodeByPublicBarcode(itembarcode,itemgubun,itemid,itemoption)) then
        itemgubun   = Left(itembarcode, 2)
        itemid      = CStr(Mid(itembarcode, 3, 6) + 0)
        itemoption  = Right(itembarcode, 4)
    end if
end if

dim oOffStock
set oOffStock = new CShopItemSummary
	oOffStock.FRectShopID   = shopid
	oOffStock.FRectMakerID  = makerid
	oOffStock.FRectItemGubun= itemgubun
	oOffStock.FRectItemID   = itemid
	oOffStock.FRectItemOption= itemoption
	oOffStock.FRectErrType  = "D"
	oOffStock.FRectStartDate = fromdate
	oOffStock.FRectEndDate   = todate
	oOffStock.FRectGroupType = grpType
	oOffStock.FRectComm_cd      = comm_cd
	oOffStock.GetOFFErrItemSummary

Dim i, TotErrrealcheckno, TotRealSellno, Totshopitemprice, TotSuplyPrice
dim yyyy, mm
%>
<script type='text/javascript'>

function lossProcess(fromdate,todate,shopid,makerid,grpType){
    var param = "shopid="+shopid+"&makerid="+makerid+"&fromdate="+fromdate+"&todate="+todate+"&grpType="+grpType;
    var popwin= window.open('shopErrLossmaker.asp?' + param,'shopErrLossmaker','width=1100,ehight=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function lossProcess1(fromdate,todate,shopid,makerid,grpType){
    //http://webadmin.10x10.co.kr/admin/offshop/stock/OutItemListByBrand.asp?menupos=1&research=on&page=&shopid=streetshop011&makerid=100atom&cType=L&CLDiv=L&LstYYYYMM=2012-12&ipchulcode=
    var param = "shopid="+shopid+"&makerid="+makerid+"&cType=L&CLDiv=L&LstYYYYMM="+todate;
    var popwin= window.open('http://webadmin.10x10.co.kr/admin/offshop/stock/OutItemListByBrand.asp?' + param,'OutItemListByBrand','width=1100,ehight=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}




function PopErrList(shopid, yyyy, mm, paramakerid) {
	var tmpmakerid = "<%= makerid %>";
	var itembarcode = "<%= itembarcode %>";
	var makerid = "";
	var yyyy1 = <%= yyyy1 %>;
	var mm1 = <%= mm1 %>;
	var dd1 = "01";

	var yyyy2 = <%= yyyy2 %>;
	var mm2 = <%= mm2 %>;
	var dd2 = "<%= Day(DateAdd("d", -1, DateSerial(yyyy2, (CLng(mm2) + 1), 1))) %>";

	var v = new Date(yyyy, (mm + 1), 1);
	var vv = new Date(v - 1);
	var lastday = vv.getDate();

	var grpType = "<%= grpType %>";

	if (tmpmakerid != ''){
		makerid = tmpmakerid;
	}else{
		makerid = paramakerid;
	}

	var u;
	if (grpType == "M") {
		u = "/admin/stock/off_baditem_list.asp?menupos=1076&shopid=" + shopid + "&makerid=" + makerid + "&itembarcode=" + itembarcode + "&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=1&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + lastday + "&errType=D";
	} else {
		u = "/admin/stock/off_baditem_list.asp?menupos=1076&shopid=" + shopid + "&makerid=" + makerid + "&itembarcode=" + itembarcode + "&yyyy1=" +yyyy1 + "&mm1=" +mm1 + "&dd1=" +dd1 + "&yyyy2=" +yyyy2 + "&mm2=" +mm2 + "&dd2=" +dd2 + "&errType=D";
	}

    var popwin= window.open(u,'PopErrList','width=1100,ehight=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
		    <% if (C_IS_SHOP) then %>
    		    <input type="hidden" name="shopid" value="<%= shopid %>">
    		    매장 : <%= shopid %>
		    <% elseif (C_IS_Maker_Upche) then %>
    		    <!-- 계약된 업체 -->
    		    매장 : <% drawSelectBoxOpenOffShop "shopid",shopid %>
		    <% else %>
		        매장 : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
		    <% end if %>

		    <% if (C_IS_Maker_Upche) then %>
		        <input type="hidden" name="makerid" value="<%= makerid %>">
		    <% else %>
    			브랜드 :
    			<% drawSelectBoxDesignerwithName "makerid", makerid %> &nbsp;&nbsp;
			<% end if %>

			<!-- 카테고리 :  -->
			상품바코드 :
			<input type="text" class="text" name="itembarcode" value="<%= itembarcode %>" size="20" maxlength="32">
			<br>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    (오차)등록일 : <% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %>
		    &nbsp;&nbsp;
			구분:
        	<input type="radio" name="grpType" value="M" <%= Chkiif(grpType = "M","checked","") %> > 월별
        	<input type="radio" name="grpType" value="S" <%= Chkiif(grpType = "S","checked","") %> > 합계
			&nbsp;&nbsp;
			<% if (shopid <> "") then %>
			매장매입구분 :
			<% drawSelectBoxOFFJungsanCommCDmulti "comm_cd",comm_cd %>
			<% end if %>
		</td>
	</tr>

	</form>
</table>
<!-- 검색 끝 -->
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmList">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="200">기간</td>
		<td width="110">매장ID</td>
		<td width="130">브랜드ID</td>
		<td width="130">현 정산구분</td>
		<td width="100">오차</td>
		<td width="100">오차 <br>(현 판매가)</td>
		<% IF (C_IS_FRN_SHOP) then %>
		<td width="100">오차 <br>(매장 매입가)</td>
		<% ELSE %>
		<td width="100">오차 <br>(본사 매입가)</td>
		<% END IF %>
		<td >비고</td>
    </tr>
	<% for i=0 to oOffStock.FResultCount - 1 %>
	<%
	TotErrrealcheckno = TotErrrealcheckno + oOffStock.FItemList(i).Ferrrealcheckno
	''TotRealSellno     = TotRealSellno  -1*(oOffStock.FItemList(i).FSellno- oOffStock.FItemList(i).FReSellno)
	Totshopitemprice  = Totshopitemprice + (oOffStock.FItemList(i).Fshopitemprice)
	IF (C_IS_FRN_SHOP) then
	    TotSuplyPrice     = TotSuplyPrice + (oOffStock.FItemList(i).Fshopbuyprice)
	ELSE
	    TotSuplyPrice     = TotSuplyPrice + (oOffStock.FItemList(i).fshopsuplycash)
    END IF

	if IsNull(oOffStock.FItemList(i).Fmakerid) then
		oOffStock.FItemList(i).Fmakerid = ""
	end if

	%>
    <tr align="center" bgcolor="#FFFFFF">
		<td><%= oOffStock.FItemList(i).Fyyyymmdd %></td>
		<td>
			<%
			yyyy = ""
			mm = ""
			if (Len(oOffStock.FItemList(i).Fyyyymmdd) = 7) then
				yyyy = Left(oOffStock.FItemList(i).Fyyyymmdd, 4)
				mm = Right(oOffStock.FItemList(i).Fyyyymmdd, 2)
			end if
			%>
			<a href="javascript:PopErrList('<%= oOffStock.FItemList(i).Fshopid %>', '<%= yyyy %>', '<%= mm %>','<%= replace(oOffStock.FItemList(i).Fmakerid,"전체","") %>')"><%= oOffStock.FItemList(i).Fshopid %></a>
		</td>
		<td><%= oOffStock.FItemList(i).Fmakerid %></td>
		<td><%= oOffStock.FItemList(i).Fcomm_name %></td>
		<td><%= FormatNumber(oOffStock.FItemList(i).Ferrrealcheckno,0) %></td>
		<td><%= FormatNumber(NULL2Zero(oOffStock.FItemList(i).Fshopitemprice),0) %></td>
		<td>
		    <% IF (C_IS_FRN_SHOP) then %>
		    <%= FormatNumber(NULL2Zero(oOffStock.FItemList(i).Fshopbuyprice),0) %>
		    <% ELSE %>
		    <%= FormatNumber(NULL2Zero(oOffStock.FItemList(i).fshopsuplycash),0) %>
		    <% END IF %>
		</td>
		<td>
			<% if C_ADMIN_AUTH or C_OFF_AUTH then %>
				<% if  (Shopid<>"") then %> <!-- (makerid<>"") and -->
		    		<% if (grpType="S") then %>
		    			<input type="button" class="button_s" value="로스처리1" onClick="lossProcess1('<%= fromdate %>','<%= todate %>','<%= shopid %>','<%= oOffStock.FItemList(i).Fmakerid %>','<%=grpType%>');">

		    		<% else %>
		    			<input type="button" class="button_s" value="로스처리" onClick="lossProcess('<%= oOffStock.FItemList(i).Fyyyymmdd + "-01" %>','<%= DateAdd("m",1,oOffStock.FItemList(i).Fyyyymmdd + "-01") %>','<%= shopid %>','<%= oOffStock.FItemList(i).Fmakerid %>','<%=grpType%>');">
		    		<% end if %>
				<% end if %>
			<% end if %>
		</td>
    </tr>
   	<% next %>
	<tr align="center" bgcolor="#FFFFFF">
	  <td>Total</td>
	  <td colspan="3"></td>
	  <td><%= FormatNumber(TotErrrealcheckno,0) %></td>
	  <td><%= FormatNumber(Totshopitemprice,0) %></td>
	  <td><%= FormatNumber(TotSuplyPrice,0) %></td>
	  <td></td>

	</tr>
</form>
</table>

<%
set oOffStock = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
