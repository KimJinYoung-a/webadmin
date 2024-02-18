<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 매출상품 상세 공용페이지 NO 페이징 버전
' History : 2009.04.07 서동석 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls_datamart.asp"-->
<%
dim shopid , datefg , i ,makerid ,yyyy1,mm1,dd1,yyyy2,mm2,dd2, toDate,fromDate
dim totitemno ,totalsum ,totsuplysum ,totsellsum ,oldlist ,offgubun ,vOffCateCode ,offmduserid
dim vOffMDUserID ,vPurchaseType ,ordertype ,itemid ,itemname ,extbarcode ,reload, buyergubun
dim inc3pl, commCd
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	menupos = request("menupos")
	shopid = request("shopid")
	datefg = request("datefg")
	makerid = request("makerid")
	oldlist = request("oldlist")
	offgubun = request("offgubun")
	vOffCateCode = request("offcatecode")
	vOffMDUserID = request("offmduserid")
	vPurchaseType = requestCheckVar(request("purchasetype"),2)
	ordertype = request("ordertype")
	itemid = request("itemid")
	itemname = request("itemname")
	extbarcode = request("extbarcode")
	reload = request("reload")
	buyergubun = request("buyergubun")
    inc3pl = request("inc3pl")
    commCd = request("commCd")
    
if datefg = "" then datefg = "maechul"
if shopid<>"" then offgubun=""
if ordertype="" then ordertype="totalprice"
if reload <> "on" and offgubun = "" then offgubun = "95"

if (yyyy1="") then
	'fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now())))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

'C_IS_Maker_Upche = TRUE
'C_IS_SHOP = TRUE

'/매장
if (C_IS_SHOP) then

	'//직영점일때
	if C_IS_OWN_SHOP then

		'/어드민권한 점장 미만
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if
else
	'/업체
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")	'"7321"
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if

''두타쪽 매출조회 권한
Dim isFixShopView
IF (session("ssBctID")="doota01") then
    shopid="streetshop014"
    C_IS_SHOP = TRUE
    isFixShopView = TRUE
ENd If

dim ooffsell
set ooffsell = new COffShopSellReport
	ooffsell.FRectOldData = oldlist
	ooffsell.FRectShopid = shopid
	ooffsell.FRectNormalOnly = "on"
	ooffsell.frectdatefg = datefg
    ooffsell.FRectTerms = ""
    ooffsell.FRectStartDay = fromDate
    ooffsell.FRectEndDay = toDate
    ooffsell.FRectDesigner = makerid
	ooffsell.FRectOffgubun = offgubun
	ooffsell.frectoffcatecode = vOffCateCode
	ooffsell.frectoffmduserid = vOffMDUserID
	ooffsell.FRectBrandPurchaseType = vPurchaseType
	ooffsell.FRectOrdertype = ordertype
	ooffsell.FRectitemid = itemid
	ooffsell.FRectitemname = itemname
	ooffsell.FRectextbarcode = extbarcode
	ooffsell.FRectbuyergubun = buyergubun
	ooffsell.FRectInc3pl = inc3pl
	ooffsell.FRectCommCD = commCd
    ooffsell.GetDaylySellItemList

totitemno = 0
totalsum =0
totsuplysum = 0
totsellsum = 0
%>

<script language="javascript">

function frmsubmit(){

	if(frm.itemid.value!=''){
		if (!IsDouble(frm.itemid.value)){
			alert('상품코드는 숫자만 가능합니다.');
			frm.itemid.focus();
			return;
		}
	}

	frm.submit();
}

function pop_exceldown(){
	frm.action='/admin/offshop/todayselldetail_excel.asp';
	frm.method='post';
	frm.target='view';
	frm.submit()
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="reload" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* 기간 : <% drawmaechuldatefg "datefg" ,datefg ,""%>
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >3년이전
				&nbsp;&nbsp;
				<%
				'직영/가맹점
				if (C_IS_SHOP) then
				%>
					<% if (not C_IS_OWN_SHOP and shopid <> "") or (isFixShopView) then %>
						* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
					<% end if %>
				<% else %>
					* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
				<% end if %>
				<p>
				* 매장구분 : <% drawoffshop_commoncode "offgubun", offgubun, "shopdivithinkso", "", "", " onchange='frmsubmit();'" %>
				&nbsp;&nbsp;
				* 카테고리 : <% SelectBoxBrandCategory "offcatecode", vOffCateCode %>
				&nbsp;&nbsp;
				* 담당MD : <% drawSelectBoxCoWorker_OnOff "offmduserid", vOffMDUserID, "off" %>
				&nbsp;&nbsp;
				* 구매유형 : 
				<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
				&nbsp;&nbsp;
				* 매입구분 : <% drawSelectBoxOFFJungsanCommCD "commCd",commCd %>
				<p>
				* 상품코드 : <input type="text" name="itemid" value="<%=itemid %>" size=10 maxlength=10>
				&nbsp;&nbsp;
				* 상품명 : <input type="text" name="itemname" value="<%=itemname %>" size=20 maxlength=20>
				&nbsp;&nbsp;
				* 공용바코드 : <input type="text" name="extbarcode" value="<%=extbarcode %>" size=15 maxlength=15>
				&nbsp;&nbsp;
				* 국적구분: <% drawoffshop_commoncode "buyergubun", buyergubun, "buyergubun", "MAIN", "", " onchange='frmsubmit();'" %>
				<p>
	            <b>* 매출처구분</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
				<% if C_IS_Maker_Upche then %>
					&nbsp;&nbsp;
					* 브랜드 : <%= makerid %><input type="hidden" name="makerid" value="<%= makerid %>">
				<% else %>
					&nbsp;&nbsp;
					* 브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
				<% end if %>
			</td>
		</tr>
		</table>
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit();">
	</td>
</tr>
</table>
<!-- 표 상단바 끝-->

<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
    	※ 정산은 주문일 기준으로 정산 됩니다.
    </td>
    <td align="right">
    	<input type="button" value="엑셀출력" onclick="pop_exceldown();" class="button_s">
    	<% drawordertype "ordertype" ,ordertype ," onchange='frmsubmit();'" ,"I"  %>
    </td>
</tr>
</form>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="25">
	<td colspan="20">
		검색결과 : <b><%=ooffsell.FresultCount%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>바코드</td>
	<td>공용바코드</td>
	<td>브랜드</td>
	<td>상품명(옵션명)</td>
	<% if (NOT C_InspectorUser) then %>
	<td>판매액</td>
    <% end if %>
	<td>매출액</td>

	<% if not(C_IS_SHOP) then %>
		<td>매입액</td>
	<% end if %>

	<td>판매수량</td>
	<td>비고</td>
</tr>
<%
if ooffsell.FresultCount > 0 then

for i=0 to ooffsell.FresultCount-1

totitemno = totitemno + ooffsell.FItemList(i).Fitemno
totalsum = totalsum + ooffsell.FItemList(i).FSubTotal
totsellsum = totsellsum + ooffsell.FItemList(i).fsellsum
totsuplysum = totsuplysum + ooffsell.FItemList(i).fsuplysum
%>
<tr bgcolor="#FFFFFF" align="center">
	<td><%= ooffsell.FItemList(i).GetBarCode %></td>
	<td><%= ooffsell.FItemList(i).fextbarcode %></td>
	<td><%= ooffsell.FItemList(i).FMakerID %></td>
	<td align="left">
		<%= ooffsell.FItemList(i).FItemName %>
		<% if ooffsell.FItemList(i).FItemOptionName <> "" then %>
			(<%=ooffsell.FItemList(i).FItemOptionName%>)
		<% end if %>
	</td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).fsellsum,0) %></td>
    <% end if %>
	<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(ooffsell.FItemList(i).Fsubtotal,0) %></td>

	<% if not(C_IS_SHOP) then %>
		<td align="right"><%= FormatNumber(ooffsell.FItemList(i).fsuplysum,0) %></td>
	<% end if %>

	<td><%= ooffsell.FItemList(i).Fitemno %></td>
	<td align="center"><%= ooffsell.FItemList(i).fjcomm_cd %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan=4><b>총계</b></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right"><%= FormatNumber(totsellsum,0) %></td>
    <% end if %>
	<td align="right"><%= FormatNumber(totalsum,0) %></td>

	<% if not(C_IS_SHOP) then %>
		<td align="right"><%= FormatNumber(totsuplysum,0) %></td>
	<% end if %>

	<td><%= FormatNumber(totitemno,0) %></td>
	<td></td>
</tr>
<% else %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="20">등록된 내용이 없습니다.</td>
</tr>
<% end if %>
</table>
<iframe id="view" name="view" width=0 height=0 frameborder="0" scrolling="no"></iframe>

<%
set ooffsell = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->