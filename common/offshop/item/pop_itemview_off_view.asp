<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 오프라인 상품 등록 통합
' Hieditor : 2011.10.18 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionCheckReferrerOnly.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->

<%
dim itemgubun,itemid, itemoption, barcode ,i ,makerid ,ioffitem ,opartner ,ooffontract ,IsOnlineItem
dim editmode , CenterMwDiv ,offList ,offSmall ,OnlineSailYn , IsDirectIpchulContractExistsBrand
dim shopitemname ,shopitemoptionname ,cd1 ,cd2 ,cd3 ,cd1_name ,cd2_name ,cd3_name ,orgsellprice ,shopitemprice
dim shopsuplycash ,shopbuyprice ,isusing ,vatinclude ,extbarcode ,imageList ,offmain ,OnlineOrgprice
dim OnlineBuycash, mwDiv ,OnlineSellcash ,regdate ,updt

barcode	  = request("barcode")

editmode = FALSE

if barcode <> "" and not(isnull(barcode)) then
	itemgubun 	= BF_GetItemGubun(barcode)
	itemid 		= BF_GetItemId(barcode)
	itemoption 	= BF_GetItemOption(barcode)

	set ioffitem  = new COffShopItem
		ioffitem.FRectItemgubun = itemgubun
		ioffitem.FRectItemId = itemid
		ioffitem.FRectItemOption = itemoption
		ioffitem.GetOffOneItem

	if ioffitem.FResultCount > 0 then
		makerid = ioffitem.FOneItem.Fmakerid
		Barcode = ioffitem.FOneItem.GetBarcode
		shopitemname = ioffitem.FOneItem.Fshopitemname
		shopitemoptionname = ioffitem.FOneItem.Fshopitemoptionname
		cd1 = ioffitem.FOneItem.FCateCDL
		cd2 = ioffitem.FOneItem.FCateCDM
		cd3 = ioffitem.FOneItem.FCateCDS
		cd1_name = ioffitem.FOneItem.FCateCDLName
		cd2_name = ioffitem.FOneItem.FCateCDMName
		cd3_name = ioffitem.FOneItem.FCateCDSName
		orgsellprice = ioffitem.FOneItem.FShopItemOrgprice
		shopitemprice = ioffitem.FOneItem.Fshopitemprice
		shopsuplycash = ioffitem.FOneItem.Fshopsuplycash
		shopbuyprice = ioffitem.FOneItem.Fshopbuyprice
		ItemGubun = ioffitem.FOneItem.FItemGubun
		isusing = ioffitem.FOneItem.Fisusing
		CenterMwDiv = ioffitem.FOneItem.FCenterMwDiv
		vatinclude = ioffitem.FOneItem.Fvatinclude
		extbarcode = ioffitem.FOneItem.Fextbarcode
		imageList = ioffitem.FOneItem.FimageList
		offmain = ioffitem.FOneItem.FOffImgMain
		offList = ioffitem.FOneItem.FOffImgList
		offSmall = ioffitem.FOneItem.FOffImgSmall
		OnlineSailYn = ioffitem.FOneItem.FOnlineSailYn
		OnlineOrgprice = ioffitem.FOneItem.FOnlineOrgprice
		OnlineBuycash = ioffitem.FOneItem.FOnlineBuycash
		mwDiv = ioffitem.FOneItem.FmwDiv
		OnlineSellcash = ioffitem.FOneItem.FOnlineSellcash
		regdate = ioffitem.FOneItem.Fregdate
		updt = ioffitem.FOneItem.Fupdt
	else
		response.write "<script language='javascript'>"
		response.write "	alert('해당되는 상품이 없습니다');"
		'response.write "	self.close();"
		response.write "</script>"
		dbget.close()	:	response.end
	end if

	IsOnlineItem = (itemgubun="10")

else
	response.write "<script language='javascript'>"
	response.write "	alert('해당되는 상품이 없습니다');"
	'response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.end
end if

set opartner = new CPartnerUser
    opartner.FRectDesignerID = makerid

    if makerid <> "" then
    	opartner.GetOnePartnerNUser
    end if

set ooffontract = new COffContractInfo
    ooffontract.FRectDesignerID = makerid

    if makerid <> "" then
		ooffontract.GetPartnerOffContractInfo
	end if

%>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>이미지</td>
	<td bgcolor="#FFFFFF">
		<% if IsOnlineItem then %>
			<img src="<%= imageList %>" width="50" height="50">
		<% else %>
   				<% IF offmain <> "" THEN %>
	   				<img src="<%=offmain%>" border="0" width=400 height=400> 400x400
   				<% END IF %>
   				<% if offlist <> "" then %>
   					<img src="<%=offlist%>" border="0" width=100 height=100> 100x100
   				<% end if %>
   				<% if offsmall <> "" then %>
   					<img src="<%=offsmall%>" border="0" width=50 height=50> 50x50
   				<% end if %>
		<% end if %>
	</td>
</tr>
</table>

<%
set opartner = Nothing
set ooffontract = Nothing
set ioffitem = Nothing
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
