<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인상품 등록
' Hieditor : 2009.04.07 서동석 생성
'			 2010.06.07 한용민 수정
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
dim designer,page,usingyn , research,pricediff,imageview, pricelow ,itemgubun, itemid, itemname
dim cdl, cdm, cds ,onexpire ,i, PriceDiffExists , IsDirectIpchulContractExistsBrand ,publicbarcode, excelsize
dim centermwdiv, onlineMwDiv, readonlyyn, isupcheitemreg
	onlineMwDiv  	= RequestCheckVar(request("onlineMwDiv"),1)
	designer    = RequestCheckVar(request("designer"),32)
	page        = RequestCheckVar(request("page"),9)
	usingyn     = RequestCheckVar(request("usingyn"),1)
	research    = RequestCheckVar(request("research"),9)
	pricediff   = RequestCheckVar(request("pricediff"),9)
	pricelow    = RequestCheckVar(request("pricelow"),9)
	imageview   = RequestCheckVar(request("imageview"),9)
	onexpire    = RequestCheckVar(request("onexpire"),9)
	itemgubun   = RequestCheckVar(request("itemgubun"),2)
	itemid      = RequestCheckVar(request("itemid"),1500)
	itemname    = RequestCheckVar(request("itemname"),32)
	publicbarcode    = RequestCheckVar(request("publicbarcode"),20)
	cdl         = RequestCheckVar(request("cdl"),3)
	cdm         = RequestCheckVar(request("cdm"),3)
	cds         = RequestCheckVar(request("cds"),3)
	centermwdiv = RequestCheckVar(request("centermwdiv"),3)
	excelsize = RequestCheckVar(request("excelsize"),3)
	if excelsize="" then excelsize=1
	if research<>"on" then usingyn="Y"

readonlyyn = "N"
isupcheitemreg = false

if C_ADMIN_USER then

'/매장
elseif (C_IS_SHOP) then
	'//직영점일때
	if C_IS_OWN_SHOP then
	else
	end if

	readonlyyn = "Y"
else
	'/업체일 경우 아이디 박아넣음
	if C_IS_Maker_Upche then
		designer = session("ssBctId")
		IsDirectIpchulContractExistsBrand = fnIsDirectIpchulContractExistsBrand(designer)
		isupcheitemreg = getupcheitemregyn(designer)
	end if

	readonlyyn = "Y"
end if

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

dim ioffitem
set ioffitem  = new COffShopItem
	ioffitem.FPageSize = 3500
	ioffitem.FCurrPage = excelsize
	ioffitem.FRectDesigner = designer
	ioffitem.FRectOnlyUsing = usingyn
	ioffitem.FRectItemgubun = itemgubun
	ioffitem.FRectItemID = itemid
	ioffitem.FRectItemName = html2db(itemname)
	ioffitem.FRectCDL = cdl
	ioffitem.FRectCDM = cdm
	ioffitem.FRectCDS = cds
	ioffitem.FRectOnlineExpiredItem = onexpire
	ioffitem.FRectpublicbarcode = publicbarcode
    ioffitem.FRectCenterMwdiv = centermwdiv
	ioffitem.FRectOnlineMwDiv = onlineMwDiv

	if pricediff="on" then
	    ioffitem.FRectPriceRow = pricelow
		ioffitem.GetOffShopPriceDiffItemList
	else
		ioffitem.GetOffNOnLineShopItemList
	end if

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>

<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td width="70">브랜드ID</td>
	<td width="90">상품코드</td>
	<td>상품명</td>
	<td>옵션명</td>
	<td width="60">소비자가</td>
	<td width="60">판매가</td>
	<td width="40">할인율(%)</td>
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td width="60">매입가</td>
		<td width="60">매장공급가</td>
		<td width="30">매입마진</td>
		<td width="30">공급마진</td>
	<% end if %>
	<td width="30">ON매입구분</td>
	<td width="80">
		센터매입구분
	</td>
	<td width="30">ON판매</td>
	<td width="30">ON단종</td>
	<td width="100">범용바코드</td>

	<% if C_ADMIN_USER or C_IS_Maker_Upche then %>
		<td width="60">사용 여부</td>
	<% end if %>
	<% if C_ADMIN_USER then %>
		<td width="50">ON/OFF가격연동</td>
	<% end if %>

	<td width="100">대카테</td>
	<td width="100">중카테</td>
	<td width="100">소카테</td>
	<td>비고</td>
</tr>
<% if ioffitem.FresultCount>0 then %>
	<% for i=0 to ioffitem.FresultCount -1 %>
		<tr bgcolor="#FFFFFF">
		<td ><%= ioffitem.FItemlist(i).FMakerID %></td>
		<td>
			<%= ioffitem.FItemlist(i).Fitemgubun %>-<%= CHKIIF(ioffitem.FItemlist(i).Fshopitemid>=1000000,Format00(8,ioffitem.FItemlist(i).Fshopitemid),Format00(6,ioffitem.FItemlist(i).Fshopitemid)) %>-<%= ioffitem.FItemlist(i).Fitemoption %>
		</td>
		<td>
			<%= ioffitem.FItemlist(i).FShopItemName %>
		</td>
		<td>
			<%= ioffitem.FItemlist(i).FShopitemOptionname %>

			<% if ioffitem.FItemlist(i).FOnlineOptaddprice<>0 then %>
			    옵션추가금액: <%= FormatNumber(ioffitem.FItemlist(i).FOnlineOptaddprice,0) %>
			<% end if %>
		</td>
	    <% PriceDiffExists = false %>
	    <% if C_ADMIN_USER then %>
			    <% if (ioffitem.FItemlist(i).FItemGubun="10") then %>
				    <% if (ioffitem.FItemlist(i).FOnlineitemorgprice+ ioffitem.FItemlist(i).FOnlineOptaddprice<>ioffitem.FItemlist(i).FShopItemOrgprice) or (ioffitem.FItemlist(i).FOnLineItemprice+ ioffitem.FItemlist(i).FOnlineOptaddprice<>ioffitem.FItemlist(i).FShopItemprice) then %>
					    <% PriceDiffExists = true %>
				    <% end if %>
			    <% end if %>
		<% end if %>
	    <td align="right" bgcolor="#e1e1e1">
	        <%= ioffitem.FItemlist(i).FShopItemOrgprice %>
	    </td>
		<td align="right" bgcolor="#e1e1e1">
		    <%= ioffitem.FItemlist(i).FShopItemprice %>
		</td>
		<td align="center" >
	        <% if (ioffitem.FItemlist(i).FShopItemOrgprice<>0) then %>
	            <% if ioffitem.FItemlist(i).FShopItemOrgprice<>ioffitem.FItemlist(i).FShopItemprice then %>
					OFF:<font color="#FF3333"><%= CLng((ioffitem.FItemlist(i).FShopItemOrgprice-ioffitem.FItemlist(i).FShopItemprice)/ioffitem.FItemlist(i).FShopItemOrgprice*100*100)/100 %></font>
	            <% end if %>
		    <% end if %>

		    <% if (ioffitem.FItemlist(i).FOnlineitemorgprice<>0) then %>
		        <% if ioffitem.FItemlist(i).FOnlineitemorgprice<>ioffitem.FItemlist(i).FOnLineItemprice then %>
					ON:<font color="#FF3333"><%= CLng((ioffitem.FItemlist(i).FOnlineitemorgprice-ioffitem.FItemlist(i).FOnLineItemprice)/ioffitem.FItemlist(i).FOnlineitemorgprice*100*100)/100 %></font>
	            <% end if %>
		    <% end if %>
		</td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td align="right" bgcolor="#e1e1e1">
				<%= ioffitem.FItemlist(i).Fshopsuplycash %>
			</td>
			<td align="right" bgcolor="#e1e1e1">
				<%= ioffitem.FItemlist(i).Fshopbuyprice %>
			</td>
			<td align="right" >
				<% if (ioffitem.FItemlist(i).FShopItemprice<>0) and (ioffitem.FItemlist(i).Fshopsuplycash<>0) then %>
					<font color="blue"><%= CLng((ioffitem.FItemlist(i).FShopItemprice-ioffitem.FItemlist(i).Fshopsuplycash)/ioffitem.FItemlist(i).FShopItemprice*100) %>%</font>
				<% end if %>
			</td>
			<td align="right" >
				<% if (ioffitem.FItemlist(i).FShopItemprice<>0) and (ioffitem.FItemlist(i).Fshopbuyprice<>0) then %>
					<font color="blue"><%= CLng((ioffitem.FItemlist(i).FShopItemprice-ioffitem.FItemlist(i).Fshopbuyprice)/ioffitem.FItemlist(i).FShopItemprice*100) %>%</font>
				<% end if %>
		    </td>
		<% end if %>

		<td align="center" ><%= ioffitem.FItemlist(i).FmwDiv %></td>
	    <td align="center" bgcolor="#e1e1e1">
	    	<% if ioffitem.FItemlist(i).Fstockitemid = 0 or C_ADMIN_AUTH or C_OFF_AUTH then %>
				<% =CHKIIF(ioffitem.FItemlist(i).Fcentermwdiv="M","매입","위탁")%>
		    <% else %>
		    	<%= ioffitem.FItemlist(i).Fcentermwdiv %>
			<% end if %>

	        <% if (ioffitem.FItemlist(i).FmwDiv="W" or ioffitem.FItemlist(i).FmwDiv="M") and (ioffitem.FItemlist(i).FmwDiv<>ioffitem.FItemlist(i).FCenterMwDiv) then %>
	            <font color='red'>온라인과상이</font></strong>
	        <% end if %>
	    </td>
	    <td align="center" ><%= fnColor(ioffitem.FItemlist(i).Fsellyn,"sellyn") %></td>
	    <td align="center" ><%= fnColor(ioffitem.FItemlist(i).FonLineDanjongyn,"dj") %></td>
		<td align="right" bgcolor="#e1e1e1" class="txt">
			<%= ioffitem.FItemlist(i).FextBarcode %>
		</td>

		<% if C_ADMIN_USER or C_IS_Maker_Upche then %>
			<td align="left" bgcolor="#e1e1e1">
				<% if ioffitem.FItemlist(i).Fisusing="Y" then %>
					Y
				<% else %>
					N
				<% end if %>
			</td>
		<% end if %>

		<% if C_ADMIN_USER then %>
			<td align="center" bgcolor="#e1e1e1">
				<% if ioffitem.FItemlist(i).fonofflinkyn="Y" then response.write "Y" %><% if ioffitem.FItemlist(i).fonofflinkyn="N" then response.write "N" %>
			</td>
		<% end if %>

		<td align="center">
			<%= ioffitem.FItemlist(i).FCateCDLName %>
		</td>
		<td align="center">
			<%= ioffitem.FItemlist(i).FCateCDMName %>
		</td>
		<td align="center">
			<%= ioffitem.FItemlist(i).FCateCDSName %>
		</td>
		<td align="center">
		</td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>

</table>

<%
Set ioffitem = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->