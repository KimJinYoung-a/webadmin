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
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls_datamart.asp"-->
<%
dim shopid , datefg , i ,makerid ,yyyy1,mm1,dd1,yyyy2,mm2,dd2, toDate,fromDate
dim totitemno ,totalsum ,totsuplysum ,totsellsum ,oldlist ,offgubun ,vOffCateCode ,offmduserid
dim vOffMDUserID ,vPurchaseType ,ordertype ,itemid ,itemname ,extbarcode ,reload, buyergubun
dim inc3pl, commCd, arrlist, chkImg
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	menupos = requestCheckVar(request("menupos"),10)
	shopid = requestCheckVar(request("shopid"),32)
	datefg = requestCheckVar(request("datefg"),32)
	makerid = requestCheckVar(request("makerid"),32)
	oldlist = requestCheckVar(request("oldlist"),10)
	offgubun = requestCheckVar(request("offgubun"),32)
	vOffCateCode = requestCheckVar(request("offcatecode"),32)
	vOffMDUserID = requestCheckVar(request("offmduserid"),32)
	vPurchaseType = requestCheckVar(request("purchasetype"),2)
	ordertype = requestCheckVar(request("ordertype"),32)
	itemid = requestCheckVar(request("itemid"),10)
	itemname = requestCheckVar(request("itemname"),124)
	extbarcode = requestCheckVar(request("extbarcode"),32)
	reload = requestCheckVar(request("reload"),2)
	buyergubun = requestCheckVar(request("buyergubun"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    commCd = requestCheckVar(request("commCd"),32)
	chkImg		= requestCheckvar(request("chkImg"),1)

if datefg = "" then datefg = "maechul"
if shopid<>"" then offgubun=""
if ordertype="" then ordertype="totalprice"
if reload <> "on" and offgubun = "" then offgubun = "95"
if chkImg ="" then chkImg = 0

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
	ooffsell.FPageSize = 3000
	ooffsell.FCurrPage = 1
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
    arrlist = ooffsell.GetDaylySellItemList_getrows

totitemno = 0
totalsum =0
totsuplysum = 0
totsellsum = 0
%>

<script type='text/javascript'>

function frmsubmit(){
	frm.method='get';
	frm.target='';

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
	frm.action='/admin/offshop/todayselldetail_datamart_excel.asp';
	frm.method='post';
	frm.target='view';
	frm.submit()
	frm.method='get';
	frm.target='';
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
				<br><br>
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
				<br><br>
				* 상품코드 : <input type="text" name="itemid" value="<%=itemid %>" size=10 maxlength=10>
				&nbsp;&nbsp;
				* 상품명 : <input type="text" name="itemname" value="<%=itemname %>" size=20 maxlength=20>
				&nbsp;&nbsp;
				* 공용바코드 : <input type="text" name="extbarcode" value="<%=extbarcode %>" size=15 maxlength=15>
				&nbsp;&nbsp;
				* 국적구분: <% drawoffshop_commoncode "buyergubun", buyergubun, "buyergubun", "MAIN", "", " onchange='frmsubmit();'" %>
				<br><br>
	            <b>* 매출처구분</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
				<% if C_IS_Maker_Upche then %>
					&nbsp;&nbsp;
					* 브랜드 : <%= makerid %><input type="hidden" name="makerid" value="<%= makerid %>">
				<% else %>
					&nbsp;&nbsp;
					* 브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
				<% end if %>
			    &nbsp;&nbsp;
			    <input type="checkbox" name="chkImg" value="1" <%if chkImg = 1 then%>checked<%end if%>>상품이미지 보기
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
		검색결과 : <b><%=ooffsell.FresultCount%></b> ※최대 3,000건까지 노출 됩니다.(엑셀다운은 10,000건 제한)
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% IF chkImg = 1 then %>
		<td width=50>이미지</td>
	<% end if %>
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
if isarray(arrlist) then

for i=0 to ubound(arrlist,2)

totitemno = totitemno + arrlist(3,i)
totalsum = totalsum + arrlist(0,i)
totsellsum = totsellsum + arrlist(2,i)
totsuplysum = totsuplysum + arrlist(1,i)
%>
<tr bgcolor="#FFFFFF" align="center">
	<% IF chkImg = 1 then %>
		<td>	
			<% if arrlist(17,i)<>"" and not(isnull(arrlist(17,i))) then %>
				<img src="<%= webImgUrl %>/image/small/<%= GetImageSubFolderByItemid(arrlist(11,i)) %>/<%= arrlist(17,i) %>" width="50" height="50" border="0">
			<% else %>
				<img src="<%= webImgUrl %>/offimage/offsmall/i<%= arrlist(10,i) %>/<%= GetImageSubFolderByItemid(arrlist(11,i)) %>/<%= arrlist(18,i) %>" width="50" height="50" border="0">
			<% end if %>
		</td>
	<% end if %>
	<td><%= arrlist(10,i) %><%= BF_GetFormattedItemId(arrlist(11,i)) %><%= arrlist(12,i) %></td>
	<td><%= arrlist(15,i) %></td>
	<td><%= arrlist(13,i) %></td>
	<td align="left">
		<%= arrlist(8,i) %>
		<% if arrlist(9,i) <> "" then %>
			(<%=arrlist(9,i)%>)
		<% end if %>
	</td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right"><%= FormatNumber(arrlist(2,i),0) %></td>
    <% end if %>
	<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(arrlist(0,i),0) %></td>

	<% if not(C_IS_SHOP) then %>
		<td align="right"><%= FormatNumber(arrlist(1,i),0) %></td>
	<% end if %>

	<td><%= arrlist(3,i) %></td>
	<td align="center"><%= arrlist(16,i) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="center">
	<% IF chkImg = 1 then %>
		<td colspan=5><b>총계</b></td>
	<% else %>
		<td colspan=4><b>총계</b></td>
	<% end if %>
	
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