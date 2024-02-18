<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 온라인상품 엑셀다운로드
' Hieditor : 2019.10.31 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" --> 
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, mwdiv, limityn, vatyn, sailyn, itemcouponyn, overSeaYn, itemdiv
dim cdl, cdm, cds, showminusmagin, marginup, margindown, dispCate, pojangok, vPurchasetype, sDt, eDt, reserveItemTp
dim page, infodivYn, infodiv, deliverytype, sortDiv, i,bufStr
    itemid      = requestCheckvar(request("itemid"),1500)
    itemname    = requestCheckvar(request("itemname"),64)
    makerid     = requestCheckvar(request("makerid"),32)
    sellyn      = requestCheckvar(request("sellyn"),10)
    usingyn     = requestCheckvar(request("usingyn"),10)
    danjongyn   = requestCheckvar(request("danjongyn"),10)
    mwdiv       = requestCheckvar(request("mwdiv"),10)
    limityn     = requestCheckvar(request("limityn"),10)
    vatyn       = requestCheckvar(request("vatyn"),10)
    sailyn      = requestCheckvar(request("sailyn"),10)
    itemcouponyn = requestCheckvar(request("itemcouponyn"),10)
    overSeaYn   = requestCheckvar(request("overSeaYn"),10)
    itemdiv     = requestCheckvar(request("itemdiv"),10)
    deliverytype= requestCheckvar(request("deliverytype"),10)
    pojangok	= requestCheckvar(request("pojangok"),10)
    vPurchasetype = request("purchasetype")
    reserveItemTp	= requestCheckvar(request("reserveItemTp"),1)
    page = requestCheckvar(request("page"),10)
    cdl = requestCheckvar(request("cdl"),10)
    cdm = requestCheckvar(request("cdm"),10)
    cds = requestCheckvar(request("cds"),10)
    dispCate = requestCheckvar(request("disp"),16)
    showminusmagin = request("showminusmagin")
    marginup = request("marginup")
    margindown = request("margindown")

    sDt     = requestCheckvar(request("sDt"),10)
    eDt     = requestCheckvar(request("eDt"),10)
    sortDiv	= requestCheckvar(request("sortDiv"),5)

if sortDiv="" then sortDiv="new"
infodiv  = request("infodiv")
infodivYn  = requestCheckvar(request("infodivYn"),10)

If infodiv <> "" Then
	infodivYn = "Y"	
End If

If marginup <> "" AND IsNumeric(marginup) = False Then
	rw "<script>alert('마진값(이상)이 잘못되었습니다. - "&marginup&"');history.back();</script>"
	dbget.close()
	Response.End
End If
If margindown <> "" AND IsNumeric(margindown) = False Then
	rw "<script>alert('마진값(이하)이 잘못되었습니다. - "&margindown&"');history.back();</script>"
	dbget.close()
	Response.End
End If

if (page="") then page=1

if itemid<>"" then
	dim iA ,arrTemp,arrItemid
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

dim oitem
set oitem = new CItem
    oitem.FPageSize         = 5000
    oitem.FCurrPage         = page
    oitem.FRectMakerid      = makerid
    oitem.FRectItemid       = itemid
    oitem.FRectItemName     = itemname
    oitem.FRectSellYN       = sellyn
    oitem.FRectIsUsing      = usingyn
    oitem.FRectDanjongyn    = danjongyn
    oitem.FRectLimityn      = limityn
    oitem.FRectMWDiv        = mwdiv
    oitem.FRectVatYn        = vatyn
    oitem.FRectSailYn       = sailyn
    oitem.FRectCouponYN		= itemcouponyn
    oitem.FRectIsOversea	= overSeaYn
    oitem.FRectCate_Large   = cdl
    oitem.FRectCate_Mid     = cdm
    oitem.FRectCate_Small   = cds
    oitem.FRectDispCate		= dispCate
    oitem.FRectItemDiv      = itemdiv
    oitem.FRectMinusMigin = showminusmagin
    oitem.FRectMarginUP = marginup
    oitem.FRectMarginDown = margindown
    oitem.FRectInfodivYn    = infodivYn
    oitem.FRectInfodiv    = infodiv 
    oitem.FRectDeliverytype = deliverytype
    oitem.FRectSortDiv		= sortDiv
    oitem.FRectPurchasetype = vPurchasetype
    oitem.FRectdispcateviewyn = "Y"
    oitem.FRectStartDate = sDt
    oitem.FRectEndDate = eDt
    oitem.FRectreserveItemTp		= reserveItemTp
    oitem.GetItemList

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_ITEM" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '버퍼사용여부
%>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>

<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>상품코드</td>
	<td>브랜드</td>
	<td>상품명</td>
	<td>소비자가</td>
    <td>ISBN13</td>
    <td>표시브랜드</td>
	<td>매입가</td>
	<td>마진</td>
	<td>할인율</td>
	<td>할인가</td>
	<td>할인매입가</td>
	<td>할인마진</td>
	<td>쿠폰할인율</td>
	<td>쿠폰적용판매가</td>
	<td>쿠폰적용매입가</td>
	<td>쿠폰적용마진</td>
	<td>거래구분</td>
	<td>배송구분</td>
	<td>판매여부</td>
	<td>사용여부</td>
	<td>한정여부</td>
	<td>한정수량</td>
    <td>기본전시카테고리코드</td>
    <td>등록일</td>
    <td>마지막수정일</td>
    <td>판매시작일</td>
</tr>
<% if oitem.FresultCount>0 then %>
	<% for i=0 to oitem.FresultCount -1 %>
	<tr bgcolor="#FFFFFF" align="center">
		<td bgcolor="#e1e1e1"><%= oitem.FItemList(i).Fitemid %></td>
		<td align="left" bgcolor="#e1e1e1" class="txt"><%= oitem.FItemList(i).Fmakerid %></td>
		<td align="left"><%= replace(db2html(oitem.FItemList(i).Fitemname),","," ") %></td>
		<td><%= oitem.FItemList(i).Forgprice %></td>
        <td class="txt"><%= oitem.FItemList(i).Fisbn13 %></td>
        <td class="txt"><%= oitem.FItemList(i).ffrontmakerid %></td>
		<td bgcolor="#e1e1e1"><%= oitem.FItemList(i).Forgsuplycash %></td>
		<td bgcolor="#e1e1e1"><%= fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).Forgprice,1) %></td>
		<td bgcolor="#e1e1e1">
            <% if oitem.FItemList(i).Fsailyn="Y" then %>
                <%= CLng((oitem.FItemList(i).Forgprice-oitem.FItemList(i).Fsailprice)/oitem.FItemList(i).Forgprice*100) & "%" %>
            <% else %>
                0%
            <% end if %>
        </td>
		<td bgcolor="#e1e1e1">
            <% if oitem.FItemList(i).Fsailyn="Y" then %>
                <%= oitem.FItemList(i).Fsailprice %>
            <% end if %>
        </td>
		<td bgcolor="#e1e1e1">
            <% if oitem.FItemList(i).Fsailyn="Y" then %>
                <%= oitem.FItemList(i).Fsailsuplycash %>
            <% end if %>
        </td>
		<td bgcolor="#e1e1e1">
            <% if oitem.FItemList(i).Fsailyn="Y" then %>
                <%= fnPercent(oitem.FItemList(i).Fsailsuplycash,oitem.FItemList(i).Fsailprice,1) %>
            <% end if %>
        </td>
		<td bgcolor="#e1e1e1">
            <%
            ' 쿠폰할인율
			if oitem.FItemList(i).FitemCouponYn="Y" then
            %>
                <% if oitem.FItemList(i).FitemCouponType =1 or oitem.FItemList(i).FitemCouponType =2 then %>
                    <%= CLng((oitem.FItemList(i).Forgprice-oitem.FItemList(i).GetCouponAssignPrice)/oitem.FItemList(i).Forgprice*100) & "%" %>
                <% else %>
                    0%
                <% end if %>
            <% else %>
                0%
            <% end if %>
        </td>
		<td bgcolor="#e1e1e1">
            <% if oitem.FItemList(i).FitemCouponYn="Y" then %>
                <% if oitem.FItemList(i).FitemCouponType =1 or oitem.FItemList(i).FitemCouponType =2 then %>
                    <%= oitem.FItemList(i).GetCouponAssignPrice() %>
                <% end if %>
            <% end if %>
        </td>
		<td bgcolor="#e1e1e1">
            <% if oitem.FItemList(i).FitemCouponYn="Y" then %>
                <% if oitem.FItemList(i).FitemCouponType =1 or oitem.FItemList(i).FitemCouponType =2 then %>
                    <% if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then %>
                        <%= oitem.FItemList(i).Forgsuplycash %>
                    <% else %>
                        <%= oitem.FItemList(i).Fcouponbuyprice %>
                    <% end if %>
                <% end if %>
            <% end if %>
        </td>
		<td bgcolor="#e1e1e1">
            <% if oitem.FItemList(i).FitemCouponYn="Y" then %>
                <% Select Case oitem.FItemList(i).FitemCouponType %>
                <% Case "1" %>
                    <% if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then %>
                        <%= fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).GetCouponAssignPrice(),1) %>
                    <% else %>
                        <%= fnPercent(oitem.FItemList(i).Fcouponbuyprice,oitem.FItemList(i).GetCouponAssignPrice(),1) %>
                    <% end if %>
                <% Case "2" %>
                    <% if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then %>
                        <%= fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).GetCouponAssignPrice(),1) %>
                    <% else %>
                        <%= fnPercent(oitem.FItemList(i).Fcouponbuyprice,oitem.FItemList(i).GetCouponAssignPrice(),1) %>
                    <% end if %>
                <% end Select %>
            <% end if %>
        </td>
		<td align="left" bgcolor="#e1e1e1"><%= mwdivName(oitem.FItemList(i).Fmwdiv) %></td>
		<td align="left" bgcolor="#e1e1e1"><%= getBeadalDivname(oitem.FItemList(i).Fdeliverytype) %></td>
		<td bgcolor="#e1e1e1"><%= oitem.FItemList(i).Fsellyn %></td>
		<td bgcolor="#e1e1e1"><%= oitem.FItemList(i).Fisusing %></td>
		<td bgcolor="#e1e1e1"><%= oitem.FItemList(i).Flimityn %></td>
		<td bgcolor="#e1e1e1">
            <% if  oitem.FItemList(i).Flimityn ="Y" then %>
                <%= (oitem.FItemList(i).Flimitno-oitem.FItemList(i).Flimitsold ) %>
            <% end if %>
        </td>
		<td align="left" bgcolor="#e1e1e1" class="txt"><%= oitem.FItemList(i).fcatecode %></td>
        <td align="left" bgcolor="#e1e1e1" class="txt"><%= dateconvert(oitem.FItemList(i).Fregdate) %></td>
        <td align="left" bgcolor="#e1e1e1" class="txt"><%= dateconvert(oitem.FItemList(i).Flastupdate) %></td>
        <td align="left" bgcolor="#e1e1e1" class="txt"><%= dateconvert(oitem.FItemList(i).fsellSTDate) %></td>
	</tr>
	<%
    if i mod 500 = 0 then
        Response.Flush		' 버퍼리플래쉬
    end if
    next
    %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="27" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>

</table>

<%
Set oitem = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->