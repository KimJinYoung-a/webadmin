<%
session.codePage = 65001
%>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 오프라인 바코드 출력
' Hieditor : 2016.12.15 한용민 생성
'###########################################################
%>
<%
dim shopid ,useforeigndata , ipchul, baljucode, masteridx
dim prdname, itemgubun
dim shopitemname ,cdl, cdm, cds, cartonboxno, maxcartonboxno, boxno, maxboxno
dim currentstockexist, realstockonemore, shopitemnameinserted, displayrealstockno, isupcheitemreg, IsDirectIpchulContractExistsBrand
	baljucode 		= requestCheckVar(request("baljucode"),32)
	masteridx = requestCheckVar(request("masteridx"),32)
	listgubun 		= requestCheckVar(request("listgubun"), 32)
	prdcode 		= requestCheckVar(request("prdcode"),32)
	prdname 		= requestCheckVar(request("prdname"),32)
	itemid 			= requestCheckVar(request("itemid"),255)
	generalbarcode 	= requestCheckVar(request("generalbarcode"),32)
	makerid = requestCheckVar(request("makerid"),32)
	shopid 		= requestCheckVar(request("shopid"),32)
	shopitemname 	= RequestCheckVar(request("shopitemname"),32)
	cdl         	= RequestCheckVar(request("cdl"),3)
	cdm         	= RequestCheckVar(request("cdm"),3)
	cds         	= RequestCheckVar(request("cds"),3)
	page 			= requestCheckVar(request("page"),32)
	itembarcodearr = request("itembarcodearr")
	printpriceyn 	= requestCheckVar(request("printpriceyn"),1)
	isforeignprint 	= requestCheckVar(request("isforeignprint"),32)
	currentstockexist 		= requestCheckVar(request("currentstockexist"),32)
	realstockonemore 		= requestCheckVar(request("realstockonemore"),32)
	shopitemnameinserted 	= requestCheckVar(request("shopitemnameinserted"),32)
	displayrealstockno 		= requestCheckVar(request("displayrealstockno"),32)
	ipchul 		= requestCheckVar(request("ipchul"),32)
	research 		= requestCheckVar(request("research"),32)
	makeriddispyn 			= requestCheckVar(request("makeriddispyn"),32)
	papername = requestCheckVar(request("papername"),2)
	itemcopydispyn = requestCheckVar(request("itemcopydispyn"),1)
	itemoptionyn = requestCheckVar(request("itemoptionyn"),1)
	cartonboxno = requestCheckVar(request("cartonboxno"),32)
	boxno = requestCheckVar(request("boxno"),32)
	itemgubun = requestCheckVar(request("itemgubun"),2)

isdispsql=true
isdispconfirm=true
isupcheitemreg = false

if C_ADMIN_USER then

'/매장일경우 본인 매장만 사용가능
elseif (C_IS_SHOP) then
	'/가맹점 일경우
	'if getoffshopdiv(C_STREETSHOPID) = "3" then
		'isdispconfirm=false
	'end if
	'/어드민권한 점장 미만
	'if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	'end if
else
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")
		IsDirectIpchulContractExistsBrand = fnIsDirectIpchulContractExistsBrand(makerid)
		isupcheitemreg = getupcheitemregyn(makerid)
	else
		if (Not C_ADMIN_USER) then

		else

		end if
	end if
end if

if itemid<>"" then
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
if displayrealstockno = "" and research <> "on" then
	displayrealstockno = "Y"
end if

if printpriceyn = "" then printpriceyn = "R"
if listgubun = "" then listgubun = "ITEM"
if makeriddispyn = "" then makeriddispyn = "Y"
if papername = "" then papername = "BQ"
if itemcopydispyn = "" then itemcopydispyn = "Y"
if itemoptionyn = "" then itemoptionyn = "Y"
if page = "" then page = 1
'if (masteridx = "") then masteridx = 0 end if
useforeigndata = "N"
currencyunit = "KRW"
currencyChar = "￦"
iPageSize=50

'/패킹리스트
if listgubun = "PACKING" then
	set oproduct = new CStorageDetail

	set ocstoragemaster = new CStorageMaster
		ocstoragemaster.FRectCompanyId = companyid
		ocstoragemaster.FRectMasterIdx = masteridx

	if masteridx<>"" then
		barcodetypestring = "오프라인 주문"
		ocstoragemaster.GetOneStorageMaster
		if (ocstoragemaster.FOneItem.Fisforeignorder = "Y") then
			barcodetypestring = barcodetypestring + "(해외)"
		end if

		if C_ADMIN_USER then
		elseif (C_IS_SHOP = true) then
			if ((ocstoragemaster.FOneItem.Flocationidfrom <> C_STREETSHOPID) and (ocstoragemaster.FOneItem.Flocationidto <> C_STREETSHOPID)) then
				response.write "<script>alert('"& CTX_The_wrong_approach &"');</script>"
				session.codePage = 949
				response.end
			end if
		end if

		ocstoragemaster.FRectShopId = ocstoragemaster.FOneItem.Flocationidto

		IsOneOrderOnly = False

		oproduct.FRectMakerid = makerid
		oproduct.FRectItemid       = itemid
		oproduct.FRectItemName     = html2db(prdname)
		oproduct.FRectPrdCode = prdcode
		oproduct.FRectGeneralBarcode = generalbarcode
		oproduct.FRectisforeignprint = isforeignprint
		oproduct.FRectSellYN       = sellyn
		oproduct.FRectIsUsing      = usingyn
		oproduct.FRectCompanyId = companyid
		oproduct.FRectMasterIdx = masteridx
		oproduct.FRectIsForeignOrder = ocstoragemaster.FOneItem.Fisforeignorder
		oproduct.FRectForeignOrderShopid = ocstoragemaster.FOneItem.Fforeignordershopid
		oproduct.FRectShopId = ocstoragemaster.FOneItem.Flocationidto
		oproduct.FRectBoxNo = boxno
		oproduct.FRectShopItemName = html2db(shopitemname)
		oproduct.FRectCurrentStockExist = currentstockexist
		oproduct.FRectRealStockOneMore = realstockonemore
		oproduct.FRectShopItemNameInserted = shopitemnameinserted
		oproduct.FRectCDL = cdl
		oproduct.FRectCDM = cdm
		oproduct.FRectCDS = cds
		oproduct.FRectitemgubun = itemgubun
		oproduct.frectitembarcodearr = itembarcodearr

		'상품종류가 3000 가지를 넘기면 문제가 생긴다.
		oproduct.FPageSize = 3000

		if masteridx<>"" then
			''oproduct.FRectCartonBoxNo = cartonboxno

			maxboxno = oproduct.GetMaxBoxNoByBox
			''maxcartonboxno = oproduct.GetMaxCartonBoxNo(ocstoragemaster.FOneItem.Flocationidto, ocstoragemaster.GetPackingDayList)

			if (boxno <> "") then
				cartonboxno = oproduct.GetCartonBoxNo(ocstoragemaster.FOneItem.Flocationidto, ocstoragemaster.GetPackingDayList, boxno)
			end if

		''rw ocstoragemaster.GetPackingDayList
			if (ocstoragemaster.GetPackingDayList = "") then
				IsOneOrderOnly = True
				oproduct.Getjumundetaillist
			else
				oproduct.GetjumundetaillistByBox
			end if
		end if

		divcd = ocstoragemaster.FOneItem.Fdivcd
		locationidfrom = ocstoragemaster.FOneItem.Flocationidfrom
		locationnamefrom = ocstoragemaster.FOneItem.Flocationnamefrom
		locationidto = ocstoragemaster.FOneItem.Flocationidto
		locationnameto = ocstoragemaster.FOneItem.Flocationnameto

		''' 추가..;;
		set olocation = new CLocation
		olocation.FRectCompanyId = companyid
		olocation.FRectlocationid = locationidto

		if (locationidto <> "") then
			olocation.GetOneLocation

			'useforeigndata = olocation.FOneItem.Fuseforeigndata
			'if (isforeignprint = "") then
			'	isforeignprint = useforeigndata
			'end if
			currencyunit = olocation.FOneItem.Fcurrencyunit
			currencyChar = olocation.FOneItem.FcurrencyChar
		end if
		Set olocation= Nothing

		if isarray(getarrinnerbox(siteSeq ,ocstoragemaster.GetPackingDayList ,ocstoragemaster.FRectShopId ,boxno,"","")) then
			innerboxweight = getarrinnerbox(siteSeq ,ocstoragemaster.GetPackingDayList ,ocstoragemaster.FRectShopId ,boxno,"","")(4,0)
			innerboxidx = getarrinnerbox(siteSeq ,ocstoragemaster.GetPackingDayList ,ocstoragemaster.FRectShopId ,boxno,"","")(8,0)
			cartonboxweight = getarrinnerbox(siteSeq ,ocstoragemaster.GetPackingDayList ,ocstoragemaster.FRectShopId ,boxno,"","")(9,0)
		end if

		shopseq = gettenshopidx(ocstoragemaster.FRectShopId)
		isforeignprint = ocstoragemaster.FOneItem.Fisforeignorder
	end if

'/주문리스트
elseif listgubun = "JUMUN" then
	set oproduct = new CStorageDetail

	set ocstoragemaster = new CStorageMaster
		ocstoragemaster.FRectCompanyId = companyid
		ocstoragemaster.FRectMasterIdx = masteridx
		ocstoragemaster.frectbaljucode = baljucode
		
		if masteridx<>"" or baljucode<>"" then
			ocstoragemaster.GetOneStorageMaster
		end if

	if ocstoragemaster.FTotalCount > 0 then
		barcodetypestring = "오프라인 주문"
		if (ocstoragemaster.FOneItem.Fisforeignorder = "Y") then
			barcodetypestring = barcodetypestring + "(해외)"
		end if

		if C_ADMIN_USER then
		elseif (C_IS_SHOP = true) then
			if ((ocstoragemaster.FOneItem.Flocationidfrom <> C_STREETSHOPID) and (ocstoragemaster.FOneItem.Flocationidto <> C_STREETSHOPID)) then
				response.write "<script>alert('"& CTX_The_wrong_approach &"');</script>"
				session.codePage = 949
				response.end
			end if
		end if

		ocstoragemaster.FRectShopId = ocstoragemaster.FOneItem.Flocationidto
		baljucode = ocstoragemaster.FOneItem.Fordercode

		if ocstoragemaster.FOneItem.fforeign_statecd<>"" then
			IsForeignOrder=true

			if ocstoragemaster.FOneItem.fforeign_statecd="7" then
				IsForeign_confirmed = true
			end if
		else
			IsForeign_confirmed = true
		end if
		if ocstoragemaster.FOneItem.FStatecd=" " then
			jumunwait = true	'/주문서작성중
		end if
		if (ocstoragemaster.FOneItem.FStatecd>"5") then
			isfixed = true
		else
			isfixed = false
		end if

		IsOneOrderOnly = False

		oproduct.FRectMakerid = makerid
		oproduct.FRectItemid       = itemid
		oproduct.FRectItemName     = html2db(prdname)
		oproduct.FRectPrdCode = prdcode
		oproduct.FRectGeneralBarcode = generalbarcode
		oproduct.FRectisforeignprint = isforeignprint
		oproduct.FRectSellYN       = sellyn
		oproduct.FRectIsUsing      = usingyn
		oproduct.FRectCompanyId = companyid
		oproduct.FRectMasterIdx = masteridx
		oproduct.FRectbaljucode = baljucode
		oproduct.FRectIsForeignOrder = ocstoragemaster.FOneItem.Fisforeignorder
		oproduct.FRectForeignOrderShopid = ocstoragemaster.FOneItem.Fforeignordershopid
		oproduct.FRectShopId = ocstoragemaster.FOneItem.Flocationidto
		oproduct.FRectBoxNo = boxno
		oproduct.FRectShopItemName = html2db(shopitemname)
		oproduct.FRectCurrentStockExist = currentstockexist
		oproduct.FRectRealStockOneMore = realstockonemore
		oproduct.FRectShopItemNameInserted = shopitemnameinserted
		oproduct.FRectCDL = cdl
		oproduct.FRectCDM = cdm
		oproduct.FRectCDS = cds
		oproduct.FRectitemgubun = itemgubun
		oproduct.frectitembarcodearr = itembarcodearr

		'상품종류가 3000 가지를 넘기면 문제가 생긴다.
		oproduct.FPageSize = 3000

		if masteridx<>"" or baljucode<>"" then
			maxboxno = oproduct.GetMaxBoxNo

			IsOneOrderOnly = True
			oproduct.Getjumundetaillist()
		end if

		divcd = ocstoragemaster.FOneItem.Fdivcd
		locationidfrom = ocstoragemaster.FOneItem.Flocationidfrom
		locationnamefrom = ocstoragemaster.FOneItem.Flocationnamefrom
		locationidto = ocstoragemaster.FOneItem.Flocationidto
		locationnameto = ocstoragemaster.FOneItem.Flocationnameto

		''' 추가..;;
		set olocation = new CLocation
		olocation.FRectCompanyId = companyid
		olocation.FRectlocationid = locationidto

		if (locationidto <> "") then
			olocation.GetOneLocation

			'useforeigndata = olocation.FOneItem.Fuseforeigndata
			'if (isforeignprint = "") then
			'	isforeignprint = useforeigndata
			'end if
			currencyunit = olocation.FOneItem.Fcurrencyunit
			currencyChar = olocation.FOneItem.FcurrencyChar
		end if
		Set olocation= Nothing

		if isarray(getarrinnerbox(siteSeq ,ocstoragemaster.GetPackingDayList ,ocstoragemaster.FRectShopId ,boxno,"","")) then
			innerboxweight = getarrinnerbox(siteSeq ,ocstoragemaster.GetPackingDayList ,ocstoragemaster.FRectShopId ,boxno,"","")(4,0)
			innerboxidx = getarrinnerbox(siteSeq ,ocstoragemaster.GetPackingDayList ,ocstoragemaster.FRectShopId ,boxno,"","")(8,0)
			cartonboxweight = getarrinnerbox(siteSeq ,ocstoragemaster.GetPackingDayList ,ocstoragemaster.FRectShopId ,boxno,"","")(9,0)
		end if

		shopseq = gettenshopidx(ocstoragemaster.FRectShopId)
		isforeignprint = ocstoragemaster.FOneItem.Fisforeignorder
	end if

'/상품리스트, 주문리스트 일경우
else
	set oproduct = new CProduct
		oproduct.FCurrpage = page
		oproduct.FPageSize = 21
		oproduct.FRectLocationId = shopid				'이동처
		oproduct.FRectLocationIdMaker = makerid
		oproduct.FRectPrdCode = prdcode
		oproduct.FRectItemID = itemid
		oproduct.FRectPrdName = html2db(prdname)
		oproduct.FRectGeneralBarcode = generalbarcode
		''oproduct.FRectUseYN = "Y"                         ''사용구분 상관없이 전체 표시
		oproduct.FRectCDL = cdl
		oproduct.FRectCDM = cdm
		oproduct.FRectCDS = cds
		oproduct.FRectShopItemName = html2db(shopitemname)
		oproduct.FRectCurrentStockExist = currentstockexist
		oproduct.FRectRealStockOneMore = realstockonemore
		oproduct.FRectShopItemNameInserted = shopitemnameinserted
		oproduct.frectipchul = ipchul
		oproduct.FRectitemgubun = itemgubun
		oproduct.frectitembarcodearr = itembarcodearr

		'/상품리스트
		if listgubun = "ITEM" then
			if shopid<>"" and (makerid<>"" or prdname<>"" or prdcode<>"" or itemid<>"" or generalbarcode<>"") then
				oproduct.GetProductListOffline()
			else
				isdispsql=false
			end if

		'/주문리스트
		elseif listgubun = "UPCHEJUMUN" then

			if ipchul <> "" then
				oproduct.GetipchulListOffline()
			end if
		end if

	set olocation = new CLocation
		olocation.FRectlocationid = shopid

		if (shopid <> "") then
			olocation.GetOneLocation

			useforeigndata = olocation.FOneItem.Fuseforeigndata
			if (isforeignprint = "") then
				isforeignprint = useforeigndata
			end if
			currencyunit = olocation.FOneItem.Fcurrencyunit
			currencyChar = olocation.FOneItem.FcurrencyChar
		end if
end if

wd = 80
ht = 80
qt = "M"

%>
<script type="text/javascript">

function reg(page){
	frm.page.value=page;
	frm.action='';
	frm.target='';
	frm.method="post"
	frm.submit();
}

</script>

<table align="left" valign="top" cellpadding="0" cellspacing="0" border="0">

<% if not(isdispconfirm) then %>
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<font color="red"><strong>오프라인 상품정보 조회 권한이 없습니다.</strong></font>
		</td>
	</tr>
<% elseif oproduct.FresultCount > 0 then %>
	<%
	'/물류코드, 범용바코드
	if papername="T" or papername="G" then
	%>
		<tr bgcolor='#FFFFFF'>
			<% for i=0 to oproduct.FresultCount-1 %>
			<% tmptdcnt = tmptdcnt + 1 %>
			<td style='width:208px; height:133px;' valign='top'>
				<table align='left' cellpadding='0' cellspacing='0' border='0' valign='top' width='100%' height='100%'>
				<tr valign='top' align='left'>
					<td height=20>
						<% if makeriddispyn="Y" then %>
							<strong><span class="currencychar10X10" style='font-size:18px;'><%= oproduct.FItemList(i).fsocname %></span></strong>
						<% end if %>
					</td>			
				</tr>
				<tr valign='top' align='left'>
					<td height=20 style="vetical-align:top; line-height:6px;">
						<% if makeriddispyn="Y" then %>
							<strong><span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).fsocname_kor %></span></strong>
						<% end if %>
					</td>
				</tr>
				<tr valign='top' align='left'>
					<td style="vetical-align:top; line-height:10px;">
						<% if isforeignprint = "Y" then %>
							<span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Flcitemname %></span>

							<% if itemoptionyn="Y" then %>
								<% if oproduct.FItemList(i).Flcitemoptionname <> "" then %>
									<br><span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Flcitemoptionname %></span>
								<% end if %>
							<% end if %>
						<% else %>
							<span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Fprdname %></span>

							<% if itemoptionyn="Y" then %>
								<% if oproduct.FItemList(i).Fitemoptionname <> "" then %>
									<br><span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Fitemoptionname %></span>
								<% end if %>
							<% end if %>
						<% end if %>
					</td>
				</tr>

				<% if itemcopydispyn="Y" then %>
					<tr valign='top' align='left'>
						<td style="vetical-align:top; line-height:6px;">
							<span class="currencychar10X10" style='font-size:6px;'><%= chrbyte(oproduct.FItemList(i).fitemcopy,70,"Y") %></span>
						</td>
					</tr>
				<% end if %>

				<tr valign='top' align='left'>
					<td height=40>
						<table align='left' cellpadding='0' cellspacing='0' border='0' valign='top' width='100%' height='100%'>
						<tr style='padding-bottom:5px'>
							<td align='left' valign='bottom' style="vetical-align:bottom; line-height:15px;">
								<% if printpriceyn = "Y" or printpriceyn = "C" or printpriceyn="R" or printpriceyn="S" then %>
									<% if isforeignprint = "Y" then %>
										<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= round(oproduct.FItemList(i).Flcprice,2) %></span></strong>
									<% else %>
										<%
										'//할인가 표시
										if printpriceyn="C" then
										%>
											<%
											'/할인 처리
											if oproduct.FItemList(i).Fsaleyn="Y" and oproduct.FItemList(i).Fcustomerprice>oproduct.FItemList(i).Flcprice then
											%>
												<strong><span class="currencychardefault" style='text-decoration:line-through; font-size:8px;'><%= currencychar %></span><span class="currencychar10X10" style='text-decoration:line-through; font-size:8px;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span></strong>
												<br><strong><span class="currencychardefault" style='font-size:15px; color:red;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px; color:red;'><%= FormatNumber(oproduct.FItemList(i).Flcprice,0) %></span></strong>
											<% else %>
												<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span></strong>
											<% end if %>
										<%
										'/판매가 표시
										elseif printpriceyn="R" then
										%>
											<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fsellprice,0) %></span></strong>
										<%
										'/심플금액표시
										elseif printpriceyn="S" then
										%>
											<strong><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fsellprice,0) %></span></strong>
										<%
										'//소비자가 표시
										else
										%>
											<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span></strong>
										<% end if %>
									<% end if %>
								<% end if %>
							</td>
							<td align='right' valign='bottom' style='padding-right:8px'>
								<%
								'/물류코드
								if papername="T" then
								%>
									<img src="http://company.10x10.co.kr/barcode/barcode.asp?image=3&type=25&height=25&barwidth=1&ClearAreaLeft=0&ClearAreaRight=0&ClearAreaTop=0&ClearAreaBottom=0&ClearAreaMiddle=0&Size=6&data=<%= BF_MakeTenBarcode(oproduct.FItemList(i).Fitemgubun, oproduct.FItemList(i).Fitemid, oproduct.FItemList(i).Fitemoption) %>&caption=<%= BF_GetItemGubun(oproduct.FItemList(i).Fitemgubun) %>-<%= BF_GetFormattedItemId(oproduct.FItemList(i).Fitemid) %>-<%= BF_GetItemOption(oproduct.FItemList(i).Fitemoption) %>" alt="<%= BF_MakeTenBarcode(oproduct.FItemList(i).Fitemgubun, oproduct.FItemList(i).Fitemid, oproduct.FItemList(i).Fitemoption) %>" />
								<%
								'/범용바코드
								elseif papername="G" then
								%>
									<% if oproduct.FItemList(i).Fgeneralbarcode<>"" then %>
										<img src="http://company.10x10.co.kr/barcode/barcode.asp?image=3&type=1&height=25&barwidth=1&ClearAreaLeft=0&ClearAreaRight=0&ClearAreaTop=0&ClearAreaBottom=0&ClearAreaMiddle=0&Size=6&data=<%= oproduct.FItemList(i).Fgeneralbarcode %>&caption=<%= oproduct.FItemList(i).Fgeneralbarcode %>" alt="<%= oproduct.FItemList(i).Fgeneralbarcode %>" />
									<% end if %>
								<% end if %>
							</td>
						</tr>
						</table>
					</td>			
				</tr>
				</table>
			</td>
			<%
			'/ 세로칸 중간에 공백 줌
			if tmptdcnt=1 or tmptdcnt=2 then
				response.write "<td style='width:19px; height:133px;'>&nbsp;</td>"
			end if
			%>
			<%
			'/ 3개 넘으면 줄내림
			if tmptdcnt >= 3 then
				tmptrcnt = tmptrcnt + 1

				if (oproduct.FresultCount/3) <> tmptrcnt then
					response.write "</tr>" & vbcrlf
					response.write "<tr bgcolor='#FFFFFF'><td colspan=5 style='height:15px;'>&nbsp;</td></tr>" & vbcrlf
					response.write "<tr bgcolor='#FFFFFF'>" & vbcrlf
				end if

				tmptdcnt = 0
			end if
			%>
			<% next %>
		</tr>

	<%
	'/(물류코드,이미지), (범용바코드,이미지)
	elseif papername="M" or papername="N" then
	%>
		<tr bgcolor='#FFFFFF'>
			<% for i=0 to oproduct.FresultCount-1 %>
			<%
			'/이미지
			if papername="M" or papername="N" then
				if oproduct.FItemList(i).Fitemgubun="10" then
					IF application("Svr_Info")="Dev" THEN
						imgPath = oproduct.FItemList(i).Fmainimageurl
					else
						imgPath = getThumbImgFromURL(oproduct.FItemList(i).Fmainimageurl,wd,ht,"true","false")
					end if
				else
					IF application("Svr_Info")="Dev" THEN
						imgPath = oproduct.FItemList(i).Fmainimageurl
					else
						imgPath = getThumbImgFromURL(oproduct.FItemList(i).Fmainimageurl,wd,ht,"true","false")
					end if
				end if
			end if
			%>
			<% tmptdcnt = tmptdcnt + 1 %>
			<td style='width:208px; height:133px;' valign='top'>
				<table align='left' cellpadding='0' cellspacing='0' border='0' valign='top' width='100%' height='100%'>
				<tr valign='top' align='left'>
					<td height=98>
						<table align='left' cellpadding='0' cellspacing='0' border='0' valign='top' width='100%' height='100%'>
						<tr valign='top' align='left'>
							<td height=12>
								<% if makeriddispyn="Y" then %>
									<strong><span class="currencychar10X10" style='font-size:11px;'><%= oproduct.FItemList(i).fsocname %></span></strong>
								<% end if %>
							</td>

							<%
							'/이미지
							if imgPath <> "" then
							%>
								<td align='right' valign='top' width=85 rowspan=4 style='padding-right:5px'>
									<img src="<%= imgPath %>" width="<%= wd %>" height="<%= ht %>" />
								</td>
							<% end if %>
						</tr>
						<tr valign='top' align='left'>
							<td height=10 style="vetical-align:top;">
								<% if makeriddispyn="Y" then %>
									<strong><span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).fsocname_kor %></span></strong>
								<% end if %>
							</td>
						</tr>
						<tr align='left' valign='top'>
							<td style="vetical-align:top;">
								<% if isforeignprint = "Y" then %>
									<span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Flcitemname %></span>

									<% if itemoptionyn="Y" then %>
										<% if oproduct.FItemList(i).Flcitemoptionname <> "" then %>
											<br><span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Flcitemoptionname %></span>
										<% end if %>
									<% end if %>
								<% else %>
									<span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Fprdname %></span>

									<% if itemoptionyn="Y" then %>
										<% if oproduct.FItemList(i).Fitemoptionname <> "" then %>
											<br><span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Fitemoptionname %></span>
										<% end if %>
									<% end if %>
								<% end if %>
							</td>
						</tr>

						<tr align='left' valign='top'>
							<td height=20>
								<table align='left' cellpadding='0' cellspacing='0' border='0' valign='top' width='100%' height='100%'>
								<tr align='left' valign='bottom'>
									<% if printpriceyn = "Y" or printpriceyn = "C" or printpriceyn="R" or printpriceyn="S" then %>
										<% if isforeignprint = "Y" then %>
											<td style="vetical-align:bottom;">
												<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= round(oproduct.FItemList(i).Flcprice,2) %></span></strong>
											</td>
										<% else %>
											<%
											'//할인가 표시
											if printpriceyn="C" then
											%>
												<%
												'/할인 처리
												if oproduct.FItemList(i).Fsaleyn="Y" and oproduct.FItemList(i).Fcustomerprice>oproduct.FItemList(i).Flcprice then
												%>
													<td style='text-decoration:line-through;' style="vetical-align:bottom;">
														<strong><span class="currencychardefault" style='font-size:8px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:8px;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span></strong>
													</td>
													<td style="vetical-align:bottom;">
														<strong><span class="currencychardefault" style='font-size:15px; color:red;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px; color:red;'><%= FormatNumber(oproduct.FItemList(i).Flcprice,0) %></span></strong>
													</td>
												<% else %>
													<td style="vetical-align:bottom;">
														<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span></strong>
													</td>
												<% end if %>
											<%
											'/판매가 표시
											elseif printpriceyn="R" then
											%>
												<td style="vetical-align:bottom;">
													<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fsellprice,0) %></span></strong>
												</td>
											<%
											'/심플금액표시
											elseif printpriceyn="S" then
											%>
												<td style="vetical-align:bottom;">
													<strong><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fsellprice,0) %></span></strong>
												</td>
											<%
											'//소비자가 표시
											else
											%>
												<td style="vetical-align:bottom;">
													<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span></strong>
												</td>
											<% end if %>
										<% end if %>
									<% end if %>
								</tr>
								</table>
							</td>
						</tr>					
						</table>
					</td>			
				</tr>
				<tr>
					<td align='right' valign='top' style='padding-right:8px' height=35>
						<%
						'/물류코드
						if papername="M" then
						%>
							<img src="http://company.10x10.co.kr/barcode/barcode.asp?image=3&type=25&height=25&barwidth=1&ClearAreaLeft=0&ClearAreaRight=0&ClearAreaTop=0&ClearAreaBottom=0&ClearAreaMiddle=0&Size=6&data=<%= BF_MakeTenBarcode(oproduct.FItemList(i).Fitemgubun, oproduct.FItemList(i).Fitemid, oproduct.FItemList(i).Fitemoption) %>&caption=<%= BF_GetItemGubun(oproduct.FItemList(i).Fitemgubun) %>-<%= BF_GetFormattedItemId(oproduct.FItemList(i).Fitemid) %>-<%= BF_GetItemOption(oproduct.FItemList(i).Fitemoption) %>" alt="<%= BF_MakeTenBarcode(oproduct.FItemList(i).Fitemgubun, oproduct.FItemList(i).Fitemid, oproduct.FItemList(i).Fitemoption) %>" />
						<%
						'/범용바코드
						elseif papername="N" then
						%>
							<% if oproduct.FItemList(i).Fgeneralbarcode<>"" then %>
								<img src="http://company.10x10.co.kr/barcode/barcode.asp?image=3&type=1&height=25&barwidth=1&ClearAreaLeft=0&ClearAreaRight=0&ClearAreaTop=0&ClearAreaBottom=0&ClearAreaMiddle=0&Size=6&data=<%= oproduct.FItemList(i).Fgeneralbarcode %>&caption=<%= oproduct.FItemList(i).Fgeneralbarcode %>" alt="<%= oproduct.FItemList(i).Fgeneralbarcode %>" />
							<% end if %>
						<% end if %>
					</td>
				</tr>		
				</table>
			</td>
			<%
			'/ 세로칸 중간에 공백 줌
			if tmptdcnt=1 or tmptdcnt=2 then
				response.write "<td style='width:19px; height:133px;'>&nbsp;</td>"
			end if
			%>
			<%
			'/ 3개 넘으면 줄내림
			if tmptdcnt >= 3 then
				tmptrcnt = tmptrcnt + 1

				if (oproduct.FresultCount/3) <> tmptrcnt then
					response.write "</tr>" & vbcrlf
					response.write "<tr bgcolor='#FFFFFF'><td colspan=5 style='height:15px;'>&nbsp;</td></tr>" & vbcrlf
					response.write "<tr bgcolor='#FFFFFF'>" & vbcrlf
				end if

				tmptdcnt = 0
			end if
			%>
			<% next %>
		</tr>

	<%
	'/장바구니 쇼카드(QR코드), QR코드, 이미지, 쇼카드만
	'elseif papername="BQ" or papername="Q" or papername="I" or papername="" then
	else
	%>
		<tr bgcolor='#FFFFFF'>
			<% for i=0 to oproduct.FresultCount-1 %>
			<%
			'/ 장바구니 쇼카드(QR코드)
			if papername="BQ" then
				if oproduct.FItemList(i).Fitemgubun="10" then
					'msg = "http://m.10x10.co.kr/offshop/view/category_prd.asp?barcode=" & BF_MakeTenBarcode(oproduct.FItemList(i).Fitemgubun,oproduct.FItemList(i).Fitemid,oproduct.FItemList(i).Fitemoption)
					msg = "http://m.10x10.co.kr/offshop/view/category_prd.asp?barcode=" & oproduct.FItemList(i).Fprdbarcode

					'// 구글 Chart API - QRCode URL (반드시 UTF-8로 전송)
					imgPath = "http://chart.apis.google.com/chart?cht=qr&chl=" & URLEncodeUTF8(msg) & "&choe=UTF-8&chs=" & wd & "x" & ht & "&chld=" & qt & "|1"
				else
					IF application("Svr_Info")="Dev" THEN
						imgPath = oproduct.FItemList(i).Fmainimageurl
					else
						imgPath = getThumbImgFromURL(oproduct.FItemList(i).Fmainimageurl,wd,ht,"true","false")
					end if
				end if

			'/QR코드
			elseif papername="Q" then
				if oproduct.FItemList(i).Fitemgubun="10" then
					'msg = "http://m.10x10.co.kr/category/category_itemPrd.asp?itemid=" & oproduct.FItemList(i).Fitemid
					msg = "http://m.10x10.co.kr/offshop/view/iteminfo.asp?itemid=" & oproduct.FItemList(i).Fitemid

					'// 구글 Chart API - QRCode URL (반드시 UTF-8로 전송)
					imgPath = "http://chart.apis.google.com/chart?cht=qr&chl=" & URLEncodeUTF8(msg) & "&choe=UTF-8&chs=" & wd & "x" & ht & "&chld=" & qt & "|1"
				else
					IF application("Svr_Info")="Dev" THEN
						imgPath = oproduct.FItemList(i).Fmainimageurl
					else
						imgPath = getThumbImgFromURL(oproduct.FItemList(i).Fmainimageurl,wd,ht,"true","false")
					end if
				end if

			'/이미지
			elseif papername="I" then
				if oproduct.FItemList(i).Fitemgubun="10" then
					IF application("Svr_Info")="Dev" THEN
						imgPath = oproduct.FItemList(i).Fmainimageurl
					else
						imgPath = getThumbImgFromURL(oproduct.FItemList(i).Fmainimageurl,wd,ht,"true","false")
					end if
				else
					IF application("Svr_Info")="Dev" THEN
						imgPath = oproduct.FItemList(i).Fmainimageurl
					else
						imgPath = getThumbImgFromURL(oproduct.FItemList(i).Fmainimageurl,wd,ht,"true","false")
					end if
				end if

			'/쇼카드만
			else
				imgPath = ""
			end if
			%>
			<% tmptdcnt = tmptdcnt + 1 %>
			<td style='width:208px; height:133px;' valign='top'>
				<table align='left' cellpadding='0' cellspacing='0' border='0' valign='top' width='100%' height='100%'>
				<tr valign='top' align='left'>
					<td height=20>
						<% if makeriddispyn="Y" then %>
							<strong><span class="currencychar10X10" style='font-size:18px;'><%= oproduct.FItemList(i).fsocname %></span></strong>
						<% end if %>
					</td>			
				</tr>
				<tr valign='top' align='left'>
					<td height=20 style="vetical-align:top; line-height:6px;">
						<% if makeriddispyn="Y" then %>
							<strong><span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).fsocname_kor %></span></strong>
						<% end if %>
					</td>
				</tr>
				<tr valign='top' align='left'>
					<td height=93>
						<table align='left' cellpadding='0' cellspacing='0' border='0' valign='top' width='100%' height='100%'>
						<tr align='left' valign='top'>
							<td style="vetical-align:top; line-height:10px;">
								<% if isforeignprint = "Y" then %>
									<span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Flcitemname %></span>

									<% if itemoptionyn="Y" then %>
										<% if oproduct.FItemList(i).Flcitemoptionname <> "" then %>
											<br><span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Flcitemoptionname %></span>
										<% end if %>
									<% end if %>
								<% else %>
									<span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Fprdname %></span>

									<% if itemoptionyn="Y" then %>
										<% if oproduct.FItemList(i).Fitemoptionname <> "" then %>
											<br><span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Fitemoptionname %></span>
										<% end if %>
									<% end if %>
								<% end if %>
							</td>

							<%
							'/QR코드, 이미지
							if imgPath <> "" then
							%>
								<td align='right' valign='top' width=85 rowspan=3 style='padding-right:5px'>
									<img src="<%= imgPath %>" width="<%= wd %>" height="<%= ht %>" />
								</td>
							<% end if %>
						</tr>

						<% if itemcopydispyn="Y" then %>
							<tr align='left' valign='top'>
								<td style="vetical-align:top; line-height:6px;">
									<%
									'/QR코드, 이미지
									if imgPath <> "" then
									%>
										<span class="currencychar10X10" style='font-size:6px;'><%= chrbyte(oproduct.FItemList(i).fitemcopy,40,"Y") %></span>
									<%
									'/쇼카드만
									else
									%>
										<span class="currencychar10X10" style='font-size:6px;'><%= chrbyte(oproduct.FItemList(i).fitemcopy,70,"Y") %></span>
									<% end if %>
								</td>
							</tr>
						<% end if %>

						<tr align='left' valign='top'>
							<td height=20>
								<table align='left' cellpadding='0' cellspacing='0' border='0' valign='top' width='100%' height='100%'>
								<tr align='left' valign='bottom' style='padding-bottom:5px'>
									<% if printpriceyn = "Y" or printpriceyn = "C" or printpriceyn="R" or printpriceyn="S" then %>
										<% if isforeignprint = "Y" then %>
											<td style="vetical-align:bottom; line-height:15px;">
												<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= round(oproduct.FItemList(i).Flcprice,2) %></span></strong>
											</td>
										<% else %>
											<%
											'//할인가 표시
											if printpriceyn="C" then
											%>
												<%
												'/할인 처리
												if oproduct.FItemList(i).Fsaleyn="Y" and oproduct.FItemList(i).Fcustomerprice>oproduct.FItemList(i).Flcprice then
												%>
													<td style='text-decoration:line-through;' style="vetical-align:bottom; line-height:15px;">
														<strong><span class="currencychardefault" style='font-size:8px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:8px;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span></strong>
													</td>
													<td style="vetical-align:bottom; line-height:15px;">
														<strong><span class="currencychardefault" style='font-size:15px; color:red;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px; color:red;'><%= FormatNumber(oproduct.FItemList(i).Flcprice,0) %></span></strong>
													</td>
												<% else %>
													<td style="vetical-align:bottom; line-height:15px;">
														<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span></strong>
													</td>
												<% end if %>
											<%
											'/판매가 표시
											elseif printpriceyn="R" then
											%>
												<td style="vetical-align:bottom; line-height:15px;">
													<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fsellprice,0) %></span></strong>
												</td>
											<%
											'/심플금액표시
											elseif printpriceyn="S" then
											%>
												<td style="vetical-align:bottom; line-height:15px;">
													<strong><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fsellprice,0) %></span></strong>
												</td>
											<%
											'//소비자가 표시
											else
											%>
												<td style="vetical-align:bottom; line-height:15px;">
													<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span></strong>
												</td>
											<% end if %>
										<% end if %>
									<% end if %>
								</tr>
								</table>
							</td>
						</tr>
						</table>
					</td>			
				</tr>
				</table>
			</td>
			<%
			'/ 세로칸 중간에 공백 줌
			if tmptdcnt=1 or tmptdcnt=2 then
				response.write "<td style='width:19px; height:133px;'>&nbsp;</td>"
			end if
			%>
			<%
			'/ 3개 넘으면 줄내림
			if tmptdcnt >= 3 then
				tmptrcnt = tmptrcnt + 1

				if (oproduct.FresultCount/3) <> tmptrcnt then
					response.write "</tr>" & vbcrlf
					response.write "<tr bgcolor='#FFFFFF'><td colspan=5 style='height:15px;'>&nbsp;</td></tr>" & vbcrlf
					response.write "<tr bgcolor='#FFFFFF'>" & vbcrlf
				end if

				tmptdcnt = 0
			end if
			%>
			<% next %>
		</tr>

	<% end if %>

	<tr bgcolor="FFFFFF">
		<td colspan="10" align="center">
	       	<% if oproduct.HasPreScroll then %>
				<font size=1><a href="javascript:reg(<%=oproduct.StartScrollPage-1%>)">[pre]</a></font>
			<% else %>
				<font size=1>[pre]</font>
			<% end if %>
			<% for i = 0 + oproduct.StartScrollPage to oproduct.StartScrollPage + oproduct.FScrollCount - 1 %>
				<% if (i > oproduct.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oproduct.FCurrPage) then %>
					<font color="red" size=1><b><%= i %></b></font>
				<% else %>
					<a href="javascript:reg(<%=i%>);" class="list_link"><font color="#000000" size=1><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oproduct.HasNextScroll then %>
				<font size=1><a href="javascript:reg(<%=i%>);">[next]</a></font>
			<% else %>
				<font size=1>[next]</font>
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<% if not(isdispsql) then %>
				<font color="red"><strong>검색 조건(매장[필수],브랜드,상품명,물류코드,상품코드,범용바코드)을 입력 하셔야 검색이 됩니다.</strong></font>
			<% else %>
				[검색결과가 없습니다.]
			<% end if %>
		</td>
	</tr>
<% end if %>

</table>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="onoffgubun" value="<%= onoffgubun %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<input type="hidden" name="itembarcodearr" value="<%= itembarcodearr %>">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="prdname" value="<%= prdname %>">
<input type="hidden" name="prdcode" value="<%= prdcode %>">
<input type="hidden" name="itemid" value="<%=replace(itemid,",",chr(10))%>">
<input type="hidden" name="generalbarcode" value="<%= generalbarcode %>">
<input type="hidden" name="displayrealstockno" value="<%= displayrealstockno %>">
<input type="hidden" name="shopitemname" value="<%= shopitemname %>">
<input type="hidden" name="currentstockexist" value="<%= currentstockexist %>">
<input type="hidden" name="realstockonemore" value="<%= realstockonemore %>">
<input type="hidden" name="shopitemnameinserted" value="<%= shopitemnameinserted %>">
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="cdl" value="<%= cdl %>">
<input type="hidden" name="cdm" value="<%= cdm %>">
<input type="hidden" name="cds" value="<%= cds %>">
<input type="hidden" name="isforeignprint" value="<%= isforeignprint %>">
<input type="hidden" name="printpriceyn" value="<%= printpriceyn %>">
<input type="hidden" name="makeriddispyn" value="<%= makeriddispyn %>">
<input type="hidden" name="listgubun" value="<%= listgubun %>">
<input type="hidden" name="ipchul" value="<%= ipchul %>">
<input type="hidden" name="papername" value="<%= papername %>">
</form>

<%
set oproduct = nothing
set olocation = nothing

Function AddSpace(byval str)
	if ((str = "") or (IsNull(str))) then
		AddSpace = "&nbsp;"
	else
		AddSpace = str
	end if
End Function
%>
