<% @codepage="949" language="VBScript" %>
<% option explicit %>
<%
 Session.CodePage = 949
 Response.ChaRset = "EUC-KR"
'###########################################################
' Description : 샵별패킹내역(박스별) 엑셀 다운로드
' Hieditor : 2018.08.30 정태훈 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_balju.asp"-->

<%
dim page, shopid, chulgoyn, showdeleted, showmichulgo, michulgoreason ,statecd, itemid, brandid
dim day5chulgo, shortchulgo, tempshort, danjong, etcshort ,research, i, shopdiv, baljucode ,tmpcartonboxbarcode
dim innerboxno, innerboxsongjangno, cartoonboxno, cartonboxsongjangno ,innerboxbarcode , cartonboxbarcode
dim yyyy1,mm1 , dd1, yyyy2, mm2, dd2, fromDate, toDate ,siteSeq
dim dateType, tplgubun
	page = request("page")
	shopid = request("shopid")
	chulgoyn = request("chulgoyn")
	showdeleted = request("showdel")		'웹서버 웹나이트가 파라미터중 delete 문구가 있는 경우 막는다.
	showmichulgo = request("showmichulgo")
	michulgoreason = request("michulgoreason")
	statecd = request("statecd")
	itemid = request("itemid")
	brandid = request("brandid")
	shopdiv = request("shopdiv")
	baljucode = request("baljucode")
	day5chulgo = request("day5chulgo")
	shortchulgo = request("shortchulgo")
	tempshort = request("tempshort")
	danjong = request("danjong")
	etcshort = request("etcshort")
	research = request("research")
	innerboxno 			= request("innerboxno")
	innerboxsongjangno 	= request("innerboxsongjangno")
	innerboxbarcode = request("innerboxbarcode")
	cartoonboxno 		= request("cartoonboxno")
	cartonboxsongjangno = request("cartonboxsongjangno")
	cartonboxbarcode = request("cartonboxbarcode")
	dateType = requestCheckVar(request("dateType"),1)
	tplgubun = requestCheckVar(request("tplgubun"),16)

siteSeq = "10"
if (page = "") then
	page = 1
end if

if (research = "") then
	showdeleted = "N"
	michulgoreason = "all"
end if

michulgoreason = "|"
if (day5chulgo = "Y") then
	'5일내출고
	michulgoreason = michulgoreason + "5|"
end if
if (shortchulgo = "Y") then
	'재고부족
	michulgoreason = michulgoreason + "S|"
end if
if (tempshort = "Y") then
	'일시품절
	michulgoreason = michulgoreason + "T|"
end if
if (danjong = "Y") then
	'단종
	michulgoreason = michulgoreason + "D|"
end if
if (etcshort = "Y") then
	'기타
	michulgoreason = michulgoreason + "E|"
end if

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

if (yyyy1="") then
	yyyy1 = Cstr(Year(now()))
	mm1 = Cstr(Month(now()))
	dd1 = Cstr(day(now()))
end if

if (yyyy2="") then
	yyyy2 = Cstr(Year(now()))
	mm2 = Cstr(Month(now()))
	dd2 = Cstr(day(now()))
end if

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim oshopbalju
set oshopbalju = new CShopBalju
	oshopbalju.FRectFromDate = fromDate
	oshopbalju.FRectToDate = toDate
	oshopbalju.FRectDateType = dateType
	oshopbalju.FRectBaljuId = shopid
	oshopbalju.FRectItemid = itemid
	oshopbalju.FRectBrandid = brandid
	oshopbalju.FRectShopdiv = shopdiv
	oshopbalju.FRectBaljucode = baljucode
	oshopbalju.FRectBoxno = innerboxno
	oshopbalju.FRectCartonBoxno = cartoonboxno
	oshopbalju.FRectBoxsongjangno = innerboxsongjangno
	oshopbalju.FRectCartonBoxsongjangno = cartonboxsongjangno
	oshopbalju.frectinnerboxbarcode = innerboxbarcode
	oshopbalju.frectcartonboxbarcode = cartonboxbarcode
	oshopbalju.FtplGubun = tplgubun

	if (statecd = "A") then
		oshopbalju.FRectChulgoYN = "N"
	else
		oshopbalju.FRectStatecd = statecd
	end if

	oshopbalju.FRectShowDeleted = "N"
	'oshopbalju.FRectMichulgoReason = michulgoreason
	oshopbalju.FCurrPage = page
	oshopbalju.Fpagesize = 10000
	''oshopbalju.GetShopBaljuByBox
	oshopbalju.GetShopBaljuByBoxNEW

Sub SearchDeliverCompany(selectedId)
	dim query1
    query1 = " select top 100 divcd,divname from [db_order].[dbo].tbl_songjang_div where isUsing='Y' and divcd='" + selectedId + "'"
    rsget.Open query1,dbget,1
    if  not rsget.EOF  then
        response.write replace(db2html(rsget("divname")),"'","")
    end if
    rsget.close
End Sub

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>
<html>
<head>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
<meta http-equiv="content-type" content="text/html; charset=euc-kr">
</head>
<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="<%= adminColor("tabletop") %>" align="center">
        <td>샵아이디</td>
        <td>매장명</td>
        <td width=80>발주일</td>
        <td width=60>Inner<br>박스번호</td>
        <td width=60>Inner<br>박스중량</td>
        <td width=60>Carton<br>박스번호</td>
        <td width=60>Carton<br>박스중량</td>
        <td>발주코드</td>
        <td>주문코드</td>
        <td>공급가</td>
        <td>출고상태</td>
        <td>출고일</td>
        <td>Inner<br>운송장번호</td>
        <td>Carton<br>택배사</td>
        <td>Carton<br>운송장번호</td>
    </tr>
	<% if oshopbalju.FResultCount >0 then %>
	<% for i=0 to oshopbalju.FResultcount-1 %>
    <tr bgcolor="#FFFFFF">
        <td align="center"><%= oshopbalju.FItemList(i).Fbaljuid %></td>
        <td align="center"><%= oshopbalju.FItemList(i).Fbaljuname %></td>
        <td align="center"><%= oshopbalju.FItemList(i).Fbaljudate %></td>
        <td align="center">
            <%
            if (oshopbalju.FItemList(i).Fboxno <> "0") then
                response.write oshopbalju.FItemList(i).Fboxno
            end if
            %>
        </td>
        <td align="center">
            <%
            if (oshopbalju.FItemList(i).Finnerboxweight <> "") then
                oshopbalju.FItemList(i).Finnerboxweight = FormatNumber(oshopbalju.FItemList(i).Finnerboxweight, 2)
            end if
            %>
            <%= oshopbalju.FItemList(i).Finnerboxweight %>
        </td>
        <td align="center">
            <%= oshopbalju.FItemList(i).Fcartoonboxno %>
        </td>
        <td align="center">
            <%
            if (oshopbalju.FItemList(i).Fcartoonboxweight <> "") then
                oshopbalju.FItemList(i).Fcartoonboxweight = FormatNumber(oshopbalju.FItemList(i).Fcartoonboxweight, 2)
            end if
            %>
            <%= oshopbalju.FItemList(i).Fcartoonboxweight %>
        </td>
        <td align="center"><%= oshopbalju.FItemList(i).Fbaljunum %></td>
        <td align="center"><%= oshopbalju.FItemList(i).Fbaljucode %></td>
        <td align="center"><%= FormatNumber(oshopbalju.FItemList(i).Ftotsuplycash, 0) %></td>
        <td align="center">
            <font color="<%= oshopbalju.FItemList(i).GetStateColor %>"><%= oshopbalju.FItemList(i).GetStateName %></font>
        </td>
        <td align="center"><%= oshopbalju.FItemList(i).Fchulgodate %></td>
        <td align="center" class="txt">
            <%= Cstr(oshopbalju.FItemList(i).Fboxsongjangno) %>
        </td>
        <td align="center">
            <% SearchDeliverCompany oshopbalju.FItemList(i).Fcartonboxsongjangdiv %>
        </td>
        <td align="center" class="txt">
            <%= Cstr(oshopbalju.FItemList(i).Fcartonboxsongjangno) %>
        </td>
    </tr>
	<% next %>
	<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align=center>[ 검색결과가 없습니다. ]</td>
	</tr>
	<% end if %>
</table>
</html>
<%
Set oshopbalju = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
