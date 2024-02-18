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
' Description : 온라인 바코드 출력
' Hieditor : 2016.12.15 한용민 생성
'/////////////////// 이파일 수정시 밑에 파일도 모두 동일하게 같이 고쳐야 한다. ////////////////////////
' SCM : /common/barcode/inc_barcodeprint_off.asp
' 		/partner/common/barcode/inc_barcodeprint_off.asp
' LOGICS : /v2/common/barcode/inc_barcodeprint_off.asp
'###########################################################
%>
<%
dim shopid ,useforeigndata , ipchul, baljucode, masteridx
dim prdname, location_name, itemgubun
dim shopitemname ,cdl, cdm, cds, cartonboxno, maxcartonboxno, boxno, maxboxno, i_address1, i_address2, i_telephone
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
	printpriceyn 	= requestCheckVar(request("printpriceyn"),1)
	isforeignprint 	= requestCheckVar(request("isforeignprint"),32)
	currentstockexist 		= requestCheckVar(request("currentstockexist"),32)
	realstockonemore 		= requestCheckVar(request("realstockonemore"),32)
	shopitemnameinserted 	= requestCheckVar(request("shopitemnameinserted"),32)
	displayrealstockno 		= requestCheckVar(request("displayrealstockno"),32)
	printername = requestCheckVar(request("printername"),32)
	ipchul 		= requestCheckVar(request("ipchul"),32)
	research 		= requestCheckVar(request("research"),32)
	makeriddispyn 			= requestCheckVar(request("makeriddispyn"),32)
	papername = requestCheckVar(request("papername"),2)
	itemcopydispyn = requestCheckVar(request("itemcopydispyn"),1)
	itemoptionyn = requestCheckVar(request("itemoptionyn"),1)
	titledispyn = requestCheckVar(request("titledispyn"),1)
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
		if Trim(arrTemp(iA))<>"" then
			arrItemid = arrItemid & Trim(getNumeric(arrTemp(iA))) & ","
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
if printername = "" then printername = "TEC_B-FV4_45x22"
if listgubun = "" then listgubun = "ITEM"
if makeriddispyn = "" then makeriddispyn = "Y"
if papername = "" then papername = "BQ"
if itemcopydispyn = "" then itemcopydispyn = "Y"
if itemoptionyn = "" then itemoptionyn = "Y"
if titledispyn = "" then titledispyn = "Y"
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

			useforeigndata = olocation.FOneItem.Fuseforeigndata
			if (isforeignprint = "") then
				isforeignprint = useforeigndata
			end if
			currencyunit = olocation.FOneItem.Fcurrencyunit
			currencyunit_pos = olocation.FOneItem.Fcurrencyunit_pos
			currencyChar = olocation.FOneItem.FcurrencyChar
			i_address1 = olocation.FOneItem.Faddress
			i_address2 = olocation.FOneItem.fmanager_address
			location_name = olocation.FOneItem.flocation_name
			i_telephone = olocation.FOneItem.ftel
		end if
		Set olocation= Nothing

		if isarray(getarrinnerbox(siteSeq ,ocstoragemaster.GetPackingDayList ,ocstoragemaster.FRectShopId ,boxno,"","")) then
			innerboxweight = getarrinnerbox(siteSeq ,ocstoragemaster.GetPackingDayList ,ocstoragemaster.FRectShopId ,boxno,"","")(4,0)
			innerboxidx = getarrinnerbox(siteSeq ,ocstoragemaster.GetPackingDayList ,ocstoragemaster.FRectShopId ,boxno,"","")(8,0)
			cartonboxweight = getarrinnerbox(siteSeq ,ocstoragemaster.GetPackingDayList ,ocstoragemaster.FRectShopId ,boxno,"","")(9,0)
		end if

		shopseq = gettenshopidx(ocstoragemaster.FRectShopId)
		isforeignprint = ocstoragemaster.FOneItem.Fisforeignorder
	else
		isdispsql=false
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
		else
			isdispsql=false
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

			useforeigndata = olocation.FOneItem.Fuseforeigndata
			if (isforeignprint = "") then
				isforeignprint = useforeigndata
			end if
			currencyunit = olocation.FOneItem.Fcurrencyunit
			currencyunit_pos = olocation.FOneItem.Fcurrencyunit_pos
			currencyChar = olocation.FOneItem.FcurrencyChar
			i_address1 = olocation.FOneItem.Faddress
			i_address2 = olocation.FOneItem.fmanager_address
			location_name = olocation.FOneItem.flocation_name
			i_telephone = olocation.FOneItem.ftel
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
		oproduct.FPageSize = iPageSize
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

		'/상품리스트
		if listgubun = "ITEM" then
			if shopid<>"" and (makerid<>"" or prdname<>"" or prdcode<>"" or itemid<>"" or generalbarcode<>"") then
				oproduct.GetProductListOffline()
			else
				isdispsql=false
			end if

		'/주문리스트
		elseif listgubun = "UPCHEJUMUN" then
			if shopid<>"" then	
				if ipchul <> "" then
					oproduct.GetipchulListOffline()
				else
					isdispsql=false
				end if
			else
				isdispsql=false
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
			currencyunit_pos = olocation.FOneItem.Fcurrencyunit_pos
			currencyChar = olocation.FOneItem.FcurrencyChar
			i_address1 = olocation.FOneItem.Faddress
			i_address2 = olocation.FOneItem.fmanager_address
			location_name = olocation.FOneItem.flocation_name
			i_telephone = olocation.FOneItem.ftel
		end if
end if

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

<%
'/주문리스트, 패킹리스트만 뿌린다.
if listgubun = "JUMUN" or listgubun = "PACKING" then
%>
	//carton box 인덱스 출력
	function cartonboxindexprint() {
		var showdomainyn;

		<% if ocstoragemaster.FRectShopId = "" then %>
			alert('매장을 선택해 주세요.');
			return;
		<% end if %>

		<% if ocstoragemaster.GetPackingDayList = "" then %>
			alert('출고지시일을 선택해 주세요.');
			return;
		<% end if %>

		<% if cartonboxno = "" or cartonboxno = "0" then %>
			alert('박스번호를 선택해 주세요.');
			return;
		<% end if %>

		<% if cartonboxweight = "" then %>
			alert('CARTONBOX 무게를 입력해 주세요.');
			return;
		<% end if %>

		showdomainyn	= frm.titledispyn.value;

		var paperwidth = frm.paperwidth.value;
		var paperheight = frm.paperheight.value;
		var papermargin = frm.papermargin.value;
		var heightoffset = 0;

	    var shopid; var shopname; var packingdate; var cartonboxno; var cartonboxweight; var prdcode; var prdbarcode;
		shopid = '<%=ocstoragemaster.FRectShopId%>';
		shopname = '                <%= locationnameto %>';
		packingdate = '<%=ocstoragemaster.GetPackingDayList%>';
		cartonboxno = '<%=cartonboxno%>';
		cartonboxweight = '<%=cartonboxweight%>';
		prdcode = '                      <%= Format00(2,siteseq) & "-" & Format00(6,shopseq) & "-" & left(ocstoragemaster.GetPackingDayList,4) & mid(ocstoragemaster.GetPackingDayList,6,2) & right(ocstoragemaster.GetPackingDayList,2) & "-" & Format00(3,cartonboxno) %>';
		prdbarcode = '<%= Format00(2,siteseq) & Format00(6,shopseq) & left(ocstoragemaster.GetPackingDayList,4) & mid(ocstoragemaster.GetPackingDayList,6,2) & right(ocstoragemaster.GetPackingDayList,2) & Format00(3,cartonboxno) %>';

		//TEC B-FV4		//2016.11.24 한용민 생성
		if (frm.printername.value=='TEC_B-FV4_80x50'){
			//if (TEC_DO3.IsDriver != 1){ alert("TEC B-FV4 드라이버를 설치해 주세요!!"); return; }
			if (confirm("cartonbox 인덱스를 출력합니다.\n\nTEC B-FV4 로 출력하시겠습니까?") == true) {
				TOSHIBA_PAPERWIDTH = 800;
				TOSHIBA_PAPERHEIGHT = 500;
				TOSHIBA_PAPERMARGIN = 3;
				TOSHIBA_HEIGHTOFFSET = heightoffset;
				TOSHIBA_DOMAINNAME = '           CARTON BOX INDEX               ';
				TOSHIBA_SHOWDOMAINYN = showdomainyn;

				printTOSHIBAcartonboxLabel(shopid, shopname, packingdate, cartonboxno, cartonboxweight, prdcode, prdbarcode);
			}

		}else if (frm.printername.value=='TTP-243_80x50'){
			if (confirm("cartonbox 인덱스를 출력합니다.\n\nTTP-243 로 출력하시겠습니까?") == true) {
				TTP_PAPERWIDTH = paperwidth;
				TTP_PAPERHEIGHT = paperheight;
				TTP_PAPERMARGIN = papermargin;
				TTP_HEIGHTOFFSET = heightoffset;
				TTP_DOMAINNAME = '           CARTON BOX INDEX               ';
				TTP_SHOWDOMAINYN = showdomainyn;

				printTTPcartonboxLabel(shopid, shopname, packingdate, cartonboxno, cartonboxweight, prdcode, prdbarcode);
			}
		}else {
		    alert("TTP-243(구)나 TEC B-FV4 드라이버를 설치해 주세요");
		}
		return;
	}

	//inner box 인덱스 출력
	function innerboxindexprint() {
		var showdomainyn;

		<% if ocstoragemaster.FRectShopId = "" then %>
			alert('매장을 선택해 주세요.');
			return;
		<% end if %>

		<% if ocstoragemaster.GetPackingDayList = "" then %>
			alert('출고지시일을 선택해 주세요.');
			return;
		<% end if %>

		<% if boxno = "" or boxno = "0" then %>
			alert('박스번호를 선택해 주세요.');
			return;
		<% end if %>

		<% if innerboxweight = "" then %>
			alert('INNERBOX 무게를 입력해 주세요.');
			return;
		<% end if %>

		showdomainyn	= frm.titledispyn.value;

		var paperwidth = frm.paperwidth.value;
		var paperheight = frm.paperheight.value;
		var papermargin = frm.papermargin.value;
		var heightoffset = 0;

	    var shopid; var shopname; var packingdate; var innerboxno; var innerboxweight; var prdcode; var prdbarcode;
		shopid = '<%=ocstoragemaster.FRectShopId%>';
		shopname = '                <%= locationnameto %>';
		packingdate = '<%=ocstoragemaster.GetPackingDayList%>';
		innerboxno = '<%=boxno%>';
		innerboxweight = '<%=innerboxweight%>';
		prdcode = '                      <%= Format00(2,siteseq) & "-" & Format00(6,shopseq) & "-" & left(ocstoragemaster.GetPackingDayList,4) & mid(ocstoragemaster.GetPackingDayList,6,2) & right(ocstoragemaster.GetPackingDayList,2) & "-" & Format00(3,boxno) %>';
		prdbarcode = '<%= Format00(2,siteseq) & Format00(6,shopseq) & left(ocstoragemaster.GetPackingDayList,4) & mid(ocstoragemaster.GetPackingDayList,6,2) & right(ocstoragemaster.GetPackingDayList,2) & Format00(3,boxno) %>';

		//TEC B-FV4		//2016.11.24 한용민 생성
		if (frm.printername.value=='TEC_B-FV4_80x50'){
			//if (TEC_DO3.IsDriver != 1){ alert("TEC B-FV4 드라이버를 설치해 주세요!!"); return; }
			if (confirm("Innerbox 인덱스를 출력합니다.\n\nTEC B-FV4 로 출력하시겠습니까?") == true) {
				TOSHIBA_PAPERWIDTH = paperwidth;
				TOSHIBA_PAPERHEIGHT = paperheight;
				TOSHIBA_PAPERMARGIN = papermargin;
				TOSHIBA_HEIGHTOFFSET = heightoffset;
				TOSHIBA_DOMAINNAME = '                INNER BOX INDEX               ';
				TOSHIBA_SHOWDOMAINYN = showdomainyn;

				printTOSHIBAinnerboxLabel(shopid, shopname, packingdate, innerboxno, innerboxweight, prdcode, prdbarcode);
			}

		}else if (frm.printername.value=='TTP-243_80x50'){
			if (confirm("Innerbox 인덱스를 출력합니다.\n\nTTP-243 로 출력하시겠습니까?") == true) {
				TTP_PAPERWIDTH = paperwidth;
				TTP_PAPERHEIGHT = paperheight;
				TTP_PAPERMARGIN = papermargin;
				TTP_HEIGHTOFFSET = heightoffset;
				TTP_DOMAINNAME = '                INNER BOX INDEX               ';
				TTP_SHOWDOMAINYN = showdomainyn;

				printTTPinnerboxLabel(shopid, shopname, packingdate, innerboxno, innerboxweight, prdcode, prdbarcode);
			}
		}else {
		    alert("TTP-243(구)나 TEC B-FV4 드라이버를 설치해 주세요");
		}
		return;
	}
<% end if %>

//상품수정
function pop_itemedit_off_edit(ibarcode){
    <% if C_IS_Maker_Upche and Not(IsDirectIpchulContractExistsBrand) and Not isupcheitemreg then %>
        alert('권한이 없습니다. - 매장 직접 입고 브랜드만 수정 가능합니다.');
    	return;
	<% else %>
		var pop_itemedit_off_edit = window.open('/common/offshop/item/pop_itemedit_off_edit.asp?barcode=' + ibarcode,'pop_itemedit_off_edit','width=1024,height=768,resizable=yes,scrollbars=yes');
		pop_itemedit_off_edit.focus();
    <% end if %>
}

function reg(page){
//	if(frm.itemid.value!=""){
//		if (!IsDouble(frm.itemid.value)){
//			alert("상품코드는 숫자만 가능합니다.");
//			frm.itemid.focus();
//			return;
//		}
//	}

	if(frm.prdcode.value!=""){
		if ( GetByteLength(frm.prdcode.value) < 10 ){
			alert("물류코드를 정확하게 입력하세요.");
			frm.prdcode.focus();
			return;
		}
	}

	frm.page.value=page;
	frm.action='';
	frm.target='';
	frm.method="get"
	frm.submit();
}

function SearchByItemId(frm) {
	frm.prdcode.value = "";
	frm.action='';
	frm.target='';
	frm.method="get"
	frm.submit();
}

function SearchByPrdcode(frm) {
	frm.itemid.value = "";
	frm.action='';
	frm.target='';
	frm.method="get"
	frm.submit();
}

function ipchulview(v,onload){
	if (onload =="ONLOAD"){
		if (frm.ipchul.disabled){
			frm.listgubun[1].disabled = true;
		}else{
			if (v == "ITEM" || v == "JUMUN" || v == "PACKING"){
				frm.ipchul.disabled = true;
			} else if (v == "UPCHEJUMUN") {
				frm.ipchul.disabled = false;
			}
		}
	}else{
		if (v == "ITEM" || v == "JUMUN" || v == "PACKING"){
			frm.ipchul.disabled = true;
		} else if (v == "UPCHEJUMUN") {
			frm.ipchul.disabled = false;
		}
	}
}

function SelectCk(opt){
	$(document.frmList.cksel).prop('checked',opt.checked);
}

function CheckThis(tn){
	var cksel = $("#frmList #cksel"+tn);
	cksel.prop("checked", true);
}

function jsSetSelectBoxColor() {
	var frm = document.frm;

	frm.printpriceyn.style.background = "";
	frm.makeriddispyn.style.background = "";
	frm.isforeignprint.style.background = "";

	if (frm.printpriceyn.value == "N") {
		frm.printpriceyn.style.background = "orange";
	}

	if (frm.makeriddispyn.value == "N") {
		frm.makeriddispyn.style.background = "orange";
	}

	if (frm.isforeignprint.value == "Y") {
		frm.isforeignprint.style.background = "orange";
	}
}

// 폼텍 바코드 출력
function CssFORMTECBarcodeprint(barcodetype) {
	frmList.barcodetype.value=barcodetype;

	if(!$(frmList.cksel).is(':checked')) {
		alert('선택된 상품이 없습니다.');
		return;
	}

	var browser = navigator.userAgent.toLowerCase();
	if ( -1 != browser.indexOf('chrome') ){
		frmList.target = "FrameCKP" ;
		frmList.action = "/common/barcode/CssBarcodeprint_A4.asp" ;
		frmList.submit() ;
	}else if ( -1 != browser.indexOf('trident') ){	// 익스플로러 출력 기능을 제거 했더니 바로 업체에서 문의옴.. 익스플로러 모드 가 편하다고 출력되게 해달라고..
		try{
			AddArr();
		}catch (e) {
			alert("- 도구 > 인터넷 옵션 > 보안 탭 > 신뢰할 수 있는 사이트 선택\n   1. 사이트 버튼 클릭 > 사이트 추가\n   2. 사용자 지정 수준 클릭 > 스크립팅하기 안전하지 않은 것으로 표시된 ActiveX 컨트롤 (사용)으로 체크\n\n※ 위 설정은 프린트 기능을 사용하기 위함임");
		}
	}else{
		alert("사용하시는 브라우저를 확인해주세요.\n- 크로미움 엔진을 사용한 브라우저(Chrome, Edge, Whale 등)에서 출력하실 수 있습니다.");
	}
}

function AddData(itemid, itemoption, prdname, prdoptionname, socname, itemprice, itemtype, fixedno){
	iaxobject.AddData(itemid, itemoption, prdname, prdoptionname, socname, itemprice, itemtype, fixedno);
}

//AddData(v,'0000','아이템명','옵션명','브랜드',3000,'T','5')
function AddArr(){
	var makeriddisp;
	var printprice; var showpriceyn; var saleyn;

	iaxobject.ClearItem();
	//iaxobject.setTitleVisible(true);

	$("input[name='cksel']:checked").each(function(){
		var vid = $(this).val()-1; // 체크id

		//브랜드표시
		if (frm.makeriddispyn.value != 'N'){
			makeriddisp = makeriddisp = $(frmList.socname).eq(vid).val();
		}else{
			makeriddisp = '';
		}

		//가격표시
		switch (frm.printpriceyn) {
			case 'C':	//할인가표시
				if(frmList.saleyn.value=="Y") {
					//할인가
					printprice = $(frmList.saleprice).eq(vid).val().trim();
				} else {
					//소비자가
					printprice = $(frmList.customerprice).eq(vid).val().trim();
				}
				break;
			case 'R':	//판매가표시
				printprice = $(frmList.sellprice).eq(vid).val().trim();
				break;
			default:
				//소비자가 표시
				printprice = $(frmList.customerprice).eq(vid).val().trim();
				break;
		}

		// 데이터 추가
		if ($(frmList.itemid).eq(vid).val()*1>=1000000){
			AddData($(frmList.itemid).eq(vid).val(),
				$(frmList.itemoption).eq(vid).val(),
				$(frmList.prdname).eq(vid).val(),
				$(frmList.prdoptionname).eq(vid).val(),
				makeriddisp, printprice,
				$(frmList.itemgubun).eq(vid).val()*10,
				$(frmList.fixedno).eq(vid).val());
		}else{
			AddData($(frmList.itemid).eq(vid).val(),
				$(frmList.itemoption).eq(vid).val(),
				$(frmList.prdname).eq(vid).val(),
				$(frmList.prdoptionname).eq(vid).val(),
				makeriddisp, printprice,
				$(frmList.itemgubun).eq(vid).val(),
				$(frmList.fixedno).eq(vid).val());
		}

	});

	iaxobject.ShowFrm();
}

// 상품 바코드 출력
function CssBarcodeprint(barcodetype) {
	var isforeignprint = frm.isforeignprint.value;
	var currencychar="";
	frmList.barcodetype.value=barcodetype;

	if (isforeignprint == "N") {
		currencychar = "￦";
	} else {
		currencychar = "<%= currencyChar %>";
	}
	frmList.currencychar.value=currencychar;

	if(!$(frmList.cksel).is(':checked')) {
		alert('선택된 상품이 없습니다.');
		return;
	}

	var browser = navigator.userAgent.toLowerCase();
	if ( -1 != browser.indexOf('chrome') ){
		frmList.target = "FrameCKP" ;
		frmList.action = "/common/barcode/CssBarcodeprint_45x22.asp" ;
		frmList.submit() ;

	}else{
		alert("사용하시는 브라우저를 확인해주세요.\n- 크로미움 엔진을 사용한 브라우저(Chrome, Edge, Whale 등)에서 출력하실 수 있습니다.");
	}
}

// 쥬얼리 바코드 출력
function jewelleryCssBarcodePrint(barcodetype) {
	var isforeignprint = frm.isforeignprint.value;
	var currencychar="";
	frmList.barcodetype.value=barcodetype;

	if (isforeignprint == "N") {
		currencychar = "￦";
	} else {
		currencychar = "<%= currencyChar %>";
	}
	frmList.currencychar.value=currencychar;

	if(!$(frmList.cksel).is(':checked')) {
		alert('선택된 상품이 없습니다.');
		return;
	}

	var browser = navigator.userAgent.toLowerCase();
	if ( -1 != browser.indexOf('chrome') ){
		frmList.target = "FrameCKP" ;
		frmList.action = "/common/barcode/CssBarcodeprint_35x15.asp" ;
		frmList.submit() ;

	}else{
		alert("사용하시는 브라우저를 확인해주세요.\n- 크로미움 엔진을 사용한 브라우저(Chrome, Edge, Whale 등)에서 출력하실 수 있습니다.");
	}
}

// 인덱스 바코드 출력
function IndexCssBarcodePrint() {
	var isforeignprint = frm.isforeignprint.value;
	var currencychar="";

	if (isforeignprint == "N") {
		currencychar = "￦";
	} else {
		currencychar = "<%= currencyChar %>";
	}
	frmList.currencychar.value=currencychar;

	if(!$(frmList.cksel).is(':checked')) {
		alert('선택된 상품이 없습니다.');
		return;
	}

	var browser = navigator.userAgent.toLowerCase();
	if ( -1 != browser.indexOf('chrome') ){
		frmList.target = "FrameCKP" ;
		frmList.action = "/common/barcode/CssBarcodeprint_80x50.asp" ;
		frmList.submit() ;

	}else{
		alert("사용하시는 브라우저를 확인해주세요.\n- 크로미움 엔진을 사용한 브라우저(Chrome, Edge, Whale 등)에서 출력하실 수 있습니다.");
	}
}

//인덱스 출력
function IndexBarcodePrint() {
	var arr = new Array();
	var isforeignprint; var printpriceval; var domainname; var showdomainyn; var ttptype;
	var found = false;
	var prdname; var itemoptionname; var customerprice; var barcodetype;
	var skipnotinserted; var printno; var shopbrandyn; var currencychar; var showpriceyn;
	var saleprice; var saleyn; var socname, brandrackcode, itemrackcode, itemoptionrackcode, subitemrackcode;

	isforeignprint = document.frm.isforeignprint.value;
	skipnotinserted = false;

	shopbrandyn = frm.makeriddispyn.value;
	if (shopbrandyn=="") shopbrandyn="Y";
	ttptype			= "TTP-243_80x50";
	barcodetype		= "2";								// T or G or 2(텐바이텐 바코드 or 범용바코드 or 텐텐바코드_범용바코드)
	var paperwidth = frm.paperwidth.value;
	var paperheight = frm.paperheight.value;
	var papermargin = frm.papermargin.value;
	var heightoffset = 0;
	showpriceyn = frm.printpriceyn.value;

	if (isforeignprint == "N") {
		currencychar = "￦";
	} else {
		currencychar = "<%= currencyChar %>";
	}
	domainname		= "www.10x10.co.kr";
	showdomainyn	= frm.titledispyn.value;

	for (var i = 0; ; i++) {
		itembarcode = document.getElementById("itembarcode_" + i);
		chk = document.getElementById("cksel" + i);

		if (itembarcode == undefined) {
			break;
		}

		if (chk.checked != true) {
			continue;
		}

		found = true;

		if (chk.checked == true) {
			saleprice = document.getElementById("saleprice_" + i).value.trim();
			saleyn = document.getElementById("saleyn_" + i).value.trim();

			//해외 상품명
			if (isforeignprint == "Y") {
				itemname = document.getElementById("itemname_foreign_" + i).value.trim();
				itemoptionname = document.getElementById("itemoptionname_foreign_" + i).value.trim();
				customerprice = document.getElementById("customerprice_foreign_" + i).value.trim();

			//국내 상품명
			} else {
				itemname = document.getElementById("itemname_" + i).value.trim();
				itemoptionname = document.getElementById("itemoptionname_" + i).value.trim();

				//할인가 표시
				if (showpriceyn=='C'){
					//할인
					if (saleyn=='Y'){
						customerprice = saleprice;

					//소비자가
					}else{
						customerprice = document.getElementById("customerprice_" + i).value.trim();
					}

				//판매가 표시
				}else if (showpriceyn=='R'){
					customerprice = document.getElementById("sellprice_" + i).value.trim();

				//소비자가 표시
				}else{
					customerprice = document.getElementById("customerprice_" + i).value.trim();
				}
			}

			itembarcode = document.getElementById("itembarcode_" + i).value.trim();
			publicbarcode = document.getElementById("publicbarcode_" + i).value.trim();
			itembarcode = itembarcode + "_" + publicbarcode;

			makerid = document.getElementById("makerid_" + i).value.trim();
			socname = document.getElementById("socname_" + i).value.trim();
			//printno = document.getElementById("printno_" + i).value.trim();
			printno = 1;
			brandrackcode = document.getElementById("prtidx_" + i).value.trim();
			itemrackcode = document.getElementById("itemrackcode_" + i).value.trim();
			subitemrackcode = document.getElementById("subitemrackcode_" + i).value.trim();

			if (printno*1 != 0) {
				var v = new BarcodeDataClass_index(itembarcode, socname, itemname, itemoptionname, customerprice, printno, '', '', '', brandrackcode, itemrackcode, itemoptionrackcode, subitemrackcode);
				arr.push(v);
			}
		}
	}

	if (found == false) {
		alert("선택된 상품이 없습니다.");
		return;
	}

	//TEC B-FV4		//2016.11.24 한용민 생성
	if (frm.printername.value=='TEC_B-FV4_80x50'){
		//if (TEC_DO3.IsDriver != 1){ alert("TEC B-FV4 드라이버를 설치해 주세요!!"); return; }
		if (confirm("선택 상품의 인덱스를 출력합니다.\n\nTEC B-FV4 로 출력하시겠습니까?") == true) {
			TOSHIBA_DOMAINNAME = domainname;
			TOSHIBA_SHOWDOMAINYN = showdomainyn;
			TOSHIBA_PAPERWIDTH = 800;
			TOSHIBA_PAPERHEIGHT = 500;
			TOSHIBA_PAPERMARGIN = 3;
			TOSHIBA_SHOWPRICEYN = showpriceyn;
			TOSHIBA_currencyChar = currencychar;
			TOSHIBA_SHOPBRANDYN = shopbrandyn;
			TOSHIBA_BARCODETYPE = barcodetype;

			printTOSHIBAMultiItemLabel(arr);
		}

	// /js/barcode.js 참조
	}else if (initTTPprinter(ttptype, barcodetype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, heightoffset) == true) {
		if (confirm("선택 상품의 인덱스를 출력합니다.\n\nTTP-243 로 출력하시겠습니까?") == true) {
			printTTPMultiItemLabel(arr);
		}

	}else {
	    alert("TTP-243(구)나 TEC B-FV4 드라이버를 설치해 주세요");
	}
	return;
}

//상품코드 출력
function BarcodePrint(btype) {
	var arrbarcode = new Array();
	var chk, itembarcode, makerid, itemname, itemoptionname, customerprice, printno; var ttptype;
	var found = false;
	var isforeignprint; var currencychar; var domainname; var showdomainyn; var showpriceyn; var shopbrandyn; var publicbarcode;
	var saleprice; var saleyn; var socname; var socname_kor;
	var catename;

	isforeignprint = frm.isforeignprint.value;

	shopbrandyn = frm.makeriddispyn.value;
	if (shopbrandyn=="") shopbrandyn="Y";
	ttptype			= "TTP-243_45x22";
	var paperwidth = frm.paperwidth.value;
	var paperheight = frm.paperheight.value;
	var papermargin = frm.papermargin.value;
	var heightoffset = 0;
	showpriceyn = frm.printpriceyn.value;

	if (isforeignprint == "N") {
		currencychar = "￦";
	} else {
		currencychar = "<%= currencyChar %>";
	}
	domainname		= "www.10x10.co.kr";
	showdomainyn	= frm.titledispyn.value;

	for (var i = 0; ; i++) {
		itembarcode = document.getElementById("itembarcode_" + i);
		chk = document.getElementById("cksel" + i);

		if (itembarcode == undefined) {
			break;
		}

		if (chk.checked != true) {
			continue;
		}

		found = true;

		saleprice = document.getElementById("saleprice_" + i).value.trim();
		saleyn = document.getElementById("saleyn_" + i).value.trim();

		//해외 상품명
		if (isforeignprint == "Y") {
			itemname = document.getElementById("itemname_foreign_" + i).value.trim();
			itemoptionname = document.getElementById("itemoptionname_foreign_" + i).value.trim();
			customerprice = document.getElementById("customerprice_foreign_" + i).value.trim();

		//국내 상품명
		} else {
			itemname = document.getElementById("itemname_" + i).value.trim();
			itemoptionname = document.getElementById("itemoptionname_" + i).value.trim();

			//할인가 표시
			if (showpriceyn=='C'){
				//할인
				if (saleyn=='Y'){
					customerprice = saleprice;

				//소비자가
				}else{
					customerprice = document.getElementById("customerprice_" + i).value.trim();
				}

			//판매가 표시
			}else if (showpriceyn=='R'){
				customerprice = document.getElementById("sellprice_" + i).value.trim();

			//소비자가 표시
			}else{
				customerprice = document.getElementById("customerprice_" + i).value.trim();
			}
		}

		itembarcode = document.getElementById("itembarcode_" + i).value.trim();
		publicbarcode = document.getElementById("publicbarcode_" + i).value.trim();
		itembarcode = itembarcode + "_" + publicbarcode;
		makerid = document.getElementById("makerid_" + i).value.trim();
		socname = document.getElementById("socname_" + i).value.trim();
		socname_kor = document.getElementById("socname_kor_" + i).value.trim();
		printno = document.getElementById("printno_" + i).value.trim();

		var v = new BarcodeDataClass(itembarcode, makerid, itemname, itemoptionname, customerprice, printno, '', socname, socname_kor);
		arrbarcode.push(v);
	}

	if (found == false) {
		alert("선택된 상품이 없습니다.");
		return;
	}

	//TEC B-FV4		//2016.11.24 한용민 생성
	if (frm.printername.value=='TEC_B-FV4_45x22'){
		//if (TEC_DO3.IsDriver != 1){ alert("TEC B-FV4 드라이버를 설치해 주세요!!"); return; }
		if (confirm("선택 상품의 바코드를 출력합니다.\n\nTEC B-FV4 로 출력하시겠습니까?") == true) {
			TOSHIBA_PAPERWIDTH = paperwidth;
			TOSHIBA_PAPERHEIGHT = paperheight;
			TOSHIBA_PAPERMARGIN = papermargin;
			TOSHIBA_SHOWDOMAINYN = showdomainyn;
			TOSHIBA_DOMAINNAME = domainname;
			TOSHIBA_SHOWPRICEYN = showpriceyn;
			TOSHIBA_currencyChar = currencychar;
			TOSHIBA_SHOPBRANDYN = shopbrandyn;
			TOSHIBA_BARCODETYPE = btype;

			printTOSHIBAMultiBarcode(arrbarcode);
		}

	// /js/barcode.js 참조
	}else if (initTTPprinter(ttptype, btype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, 0) == true) {
		if (confirm("선택 상품의 바코드를 출력합니다.\n\nTTP-243 로 출력하시겠습니까?") == true) {
			printTTPMultiBarcode(arrbarcode);
		}

	}else {
	    alert("TTP-243(구)나 TEC B-FV4 드라이버를 설치해 주세요");
	}
	return;
}

//쥬얼리 바코드 출력
function jewellery_BarcodePrint(btype) {
	var arrbarcode = new Array();
	var chk, itembarcode, makerid, itemname, itemoptionname, customerprice, printno; var ttptype;
	var found = false;
	var isforeignprint; var currencychar; var domainname; var showdomainyn; var showpriceyn; var shopbrandyn; var publicbarcode;
	var saleprice; var saleyn; var socname;

	isforeignprint = frm.isforeignprint.value;

	shopbrandyn = frm.makeriddispyn.value;
	if (shopbrandyn=="") shopbrandyn="Y";
	ttptype			= "TTP-243_35x15";
	var paperwidth = frm.paperwidth.value;
	var paperheight = frm.paperheight.value;
	var papermargin = frm.papermargin.value;
	var heightoffset = 0;
	showpriceyn = frm.printpriceyn.value;

	if (isforeignprint == "N") {
		currencychar = "￦";
	} else {
		currencychar = "<%= currencyChar %>";
	}
	domainname		= "www.10x10.co.kr";
	showdomainyn	= frm.titledispyn.value;

	for (var i = 0; ; i++) {
		itembarcode = document.getElementById("itembarcode_" + i);
		chk = document.getElementById("cksel" + i);

		if (itembarcode == undefined) {
			break;
		}

		if (chk.checked != true) {
			continue;
		}

		found = true;

		saleprice = document.getElementById("saleprice_" + i).value.trim();
		saleyn = document.getElementById("saleyn_" + i).value.trim();

		//해외 상품명
		if (isforeignprint == "Y") {
			itemname = document.getElementById("itemname_foreign_" + i).value.trim();
			itemoptionname = document.getElementById("itemoptionname_foreign_" + i).value.trim();
			customerprice = document.getElementById("customerprice_foreign_" + i).value.trim();

		//국내 상품명
		} else {
			itemname = document.getElementById("itemname_" + i).value.trim();
			itemoptionname = document.getElementById("itemoptionname_" + i).value.trim();

			//할인가 표시
			if (showpriceyn=='C'){
				//할인
				if (saleyn=='Y'){
					customerprice = saleprice;

				//소비자가
				}else{
					customerprice = document.getElementById("customerprice_" + i).value.trim();
				}

			//판매가 표시
			}else if (showpriceyn=='R'){
				customerprice = document.getElementById("sellprice_" + i).value.trim();

			//소비자가 표시
			}else{
				customerprice = document.getElementById("customerprice_" + i).value.trim();
			}
		}

		itembarcode = document.getElementById("itembarcode_" + i).value.trim();
		publicbarcode = document.getElementById("publicbarcode_" + i).value.trim();
		itembarcode = itembarcode + "_" + publicbarcode;
		makerid = document.getElementById("makerid_" + i).value.trim();
		socname = document.getElementById("socname_" + i).value.trim();
		printno = document.getElementById("printno_" + i).value.trim();

		var v = new TTPBarcodeDataClass(itembarcode, socname, itemname, itemoptionname, customerprice, printno);
		arrbarcode.push(v);
	}

	if (found == false) {
		alert("선택된 상품이 없습니다.");
		return;
	}

	//TEC B-FV4		//2016.11.24 한용민 생성
	if (frm.printername.value=='TEC_B-FV4_35x15'){
		//if (TEC_DO3.IsDriver != 1){ alert("TEC B-FV4 드라이버를 설치해 주세요!!"); return; }
		if (confirm("선택 상품의 쥬얼리 바코드를 출력합니다.\n\nTEC B-FV4 로 출력하시겠습니까?") == true) {
			TOSHIBA_PAPERWIDTH = paperwidth;
			TOSHIBA_PAPERHEIGHT = paperheight;
			TOSHIBA_PAPERMARGIN = papermargin;
			TOSHIBA_SHOWDOMAINYN = showdomainyn;
			TOSHIBA_DOMAINNAME = domainname;
			TOSHIBA_SHOWPRICEYN = showpriceyn;
			TOSHIBA_currencyChar = currencychar;
			TOSHIBA_SHOPBRANDYN = shopbrandyn;
			TOSHIBA_BARCODETYPE = btype;

			printTOSHIBAjewelleryMultiBarcode(arrbarcode);
		}

	// /js/barcode.js 참조
	}else if (initTTPprinter(ttptype, btype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, 0) == true) {
		if (confirm("선택 상품의 쥬얼리 바코드를 출력합니다.\n\nTTP-243 로 출력하시겠습니까?") == true) {
			printTTPjewelleryMultiBarcode(arrbarcode);
		}

	}else {
	    alert("TTP-243(구)나 TEC B-FV4 드라이버를 설치해 주세요");
	}
	return;
}

//해외 바코드 출력
function foreign_BarcodePrint(btype) {
	var arrbarcode = new Array();
	var chk, itembarcode, makerid, itemname, itemoptionname, customerprice, printno; var ttptype;
	var found = false;
	var isforeignprint; var currencychar; var currencyunit_pos; var domainname; var showdomainyn; var showpriceyn; var shopbrandyn; var publicbarcode;
	var saleprice; var saleyn; var socname;
	var currencyunit; var catename; var sourcearea; var itemsource; var itemsize_10x10;
	var manufacturer; var m_address1; var m_address2; var m_telephone; var vimport; var i_telephone; var i_address1; var i_address2;
	var i_telephone;

	isforeignprint = frm.isforeignprint.value;

	shopbrandyn		= frm.makeriddispyn.value;
	ttptype			= "TTP-243_45x45";
	var paperwidth = frm.paperwidth.value;
	var paperheight = frm.paperheight.value;
	var papermargin = frm.papermargin.value;
	var heightoffset = 0;
	showpriceyn = frm.printpriceyn.value;

	if (isforeignprint == "N") {
		currencychar = "￦";
		currencyunit = "KRW";
		currencyunit_pos = "WON";
	} else {
		currencychar = "<%= currencyChar %>";
		currencyunit = "<%= currencyunit %>";
		currencyunit_pos = "<%= currencyunit_pos %>"
	}
	domainname		= "www.10x10.co.kr";
	showdomainyn	= frm.titledispyn.value;

	for (var i = 0; ; i++) {
		itembarcode = document.getElementById("itembarcode_" + i);
		chk = document.getElementById("cksel" + i);

		if (itembarcode == undefined) {
			break;
		}

		if (chk.checked != true) {
			continue;
		}

		found = true;

		saleprice = document.getElementById("saleprice_" + i).value.trim();
		saleyn = document.getElementById("saleyn_" + i).value.trim();

		//해외 상품명
		if (isforeignprint == "Y") {
			itemname = document.getElementById("itemname_foreign_" + i).value.trim();
			itemoptionname = document.getElementById("itemoptionname_foreign_" + i).value.trim();
			customerprice = document.getElementById("customerprice_foreign_" + i).value.trim();

			if (currencyunit_pos=='TWD'){
				catename = document.getElementById("catename_cn_bun3_" + i).value.trim();
			} else {
				catename = document.getElementById("catename_eng3_" + i).value.trim();
			}

			sourcearea =  document.getElementById("sourcearea_en_" + i).value.trim();
			itemsource =  document.getElementById("itemsource_en_" + i).value.trim();
			itemsize =  document.getElementById("itemsize_en_" + i).value.trim();

		//국내 상품명
		} else {
			itemname = document.getElementById("itemname_" + i).value.trim();
			itemoptionname = document.getElementById("itemoptionname_" + i).value.trim();

			//할인가 표시
			if (showpriceyn=='C'){
				//할인
				if (saleyn=='Y'){
					customerprice = saleprice;

				//소비자가
				}else{
					customerprice = document.getElementById("customerprice_" + i).value.trim();
				}

			//판매가 표시
			}else if (showpriceyn=='R'){
				customerprice = document.getElementById("sellprice_" + i).value.trim();

			//소비자가 표시
			}else{
				customerprice = document.getElementById("customerprice_" + i).value.trim();
			}

			catename = document.getElementById("catename3_" + i).value.trim();
			sourcearea =  document.getElementById("sourcearea_10x10_" + i).value.trim();
			itemsource =  document.getElementById("itemsource_10x10_" + i).value.trim();
			itemsize =  document.getElementById("itemsize_10x10_" + i).value.trim();
		}

		itembarcode = document.getElementById("itembarcode_" + i).value.trim();
		publicbarcode = document.getElementById("publicbarcode_" + i).value.trim();
		itembarcode = itembarcode + "_" + publicbarcode;
		makerid = document.getElementById("makerid_" + i).value.trim();
		socname = document.getElementById("socname_" + i).value.trim();
		printno = document.getElementById("printno_" + i).value.trim();
		manufacturer = "TENBYTEN"	//제조사
		m_address1 = "14F(GyoYukDong)"	//주소
		m_address2 = "57, Daehak-ro, Jongno-gu Seoul, Korea [03082]"	//주소
		m_telephone = "+82 2 554 2033"
		vimport =  '<%= location_name %>'
		i_address1 = '<%= i_address1 %>'
		i_address2 = '<%= i_address2 %>'
		i_telephone =  '<%= i_telephone %>'

		var v = new BarcodeDataClass_foreign(itembarcode, socname, itemname, itemoptionname, customerprice, printno, catename
		, sourcearea, itemsource, itemsize, manufacturer, m_address1, m_address2, m_telephone, vimport, i_address1, i_address2
		, i_telephone);
		arrbarcode.push(v);
	}

	if (found == false) {
		alert("선택된 상품이 없습니다.");
		return;
	}

	//TEC B-FV4		//2016.11.24 한용민 생성
	if (frm.printername.value=='TEC_B-FV4_45x45'){
		//if (TEC_DO3.IsDriver != 1){ alert("TEC B-FV4 드라이버를 설치해 주세요!!"); return; }
		if (confirm("선택 상품의 해외 바코드를 출력합니다.\n\nTEC B-FV4 로 출력하시겠습니까?") == true) {
			TOSHIBA_PAPERWIDTH = paperwidth;
			TOSHIBA_PAPERHEIGHT = paperheight;
			TOSHIBA_PAPERMARGIN = papermargin;
			TOSHIBA_SHOWDOMAINYN = showdomainyn;
			TOSHIBA_DOMAINNAME = domainname;
			TOSHIBA_SHOWPRICEYN = showpriceyn;
			TOSHIBA_currencyChar = currencychar;
			TOSHIBA_SHOPBRANDYN = shopbrandyn;
			TOSHIBA_BARCODETYPE = btype;
			TOSHIBA_isforeignprint = isforeignprint;
			TOSHIBA_currencyunit = currencyunit;
			TOSHIBA_currencyunit_pos = currencyunit_pos;

			//해외 상품명
			if (isforeignprint == "Y") {
				if (currencyunit_pos=='TWD'){
					TOSHIBA_brand_str = '品名';		//품명
					TOSHIBA_origin_str = '産地';	//산지
					TOSHIBA_material_str = '材質';		//재질
					TOSHIBA_standard_str = '規格';		//규격
					TOSHIBA_manufacturer_str = '委製商';	//위제상
					TOSHIBA_address_str = '地址';	//지지
					TOSHIBA_import_str = '進口商';		//진구상
					TOSHIBA_telephone_str = '電話';		//전화
				} else {
					TOSHIBA_brand_str = 'Brand';
					TOSHIBA_origin_str = 'Origin';
					TOSHIBA_material_str = 'Material';
					TOSHIBA_standard_str = 'Size';
					TOSHIBA_manufacturer_str = 'Manufacturer';
					TOSHIBA_address_str = 'Address';
					TOSHIBA_import_str = 'Import';
					TOSHIBA_telephone_str = 'Tel';
				}

			//국내 상품명
			} else {
				// 바코드 공용 js파일에 기본값 그대로 뿌림.
			}

			printTOSHIBAforeignBarcode(arrbarcode);
		}

	// /js/barcode.js 참조
	}else if (initTTPprinter(ttptype, btype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, 0) == true) {
		if (confirm("선택 상품의 해외 바코드를 출력합니다.\n\nTTP-243 로 출력하시겠습니까?") == true) {
			TTP_isforeignprint = isforeignprint;
			TTP_currencyunit = currencyunit;
			TTP_currencyunit_pos = currencyunit_pos;

			//해외 상품명
			if (isforeignprint == "Y") {
				if (currencyunit_pos=='TWD'){
					TTP_brand_str = '品名';		//품명
					TTP_origin_str = '産地';	//산지
					TTP_material_str = '材質';		//재질
					TTP_standard_str = '規格';		//규격
					TTP_manufacturer_str = '委製商';	//위제상
					TTP_address_str = '地址';	//지지
					TTP_import_str = '進口商';		//진구상
					TTP_telephone_str = '電話';		//전화
				} else {
					TTP_brand_str = 'Brand';
					TTP_origin_str = 'Origin';
					TTP_material_str = 'Material';
					TTP_standard_str = 'Size';
					TTP_manufacturer_str = 'Manufacturer';
					TTP_address_str = 'Address';
					TTP_import_str = 'Import';
					TTP_telephone_str = 'Tel';
				}

			//국내 상품명
			} else {
				// 바코드 공용 js파일에 기본값 그대로 뿌림.
			}

			printTTPMultiforeignBarcode(arrbarcode);
		}

	}else {
	    alert("TTP-243(구)나 TEC B-FV4 드라이버를 설치해 주세요");
	}
	return;
}

String.prototype.trim = function() {
    return this.replace(/^\s+|\s+$/g, "");
};

//쇼카드 출력
function paperBarcodePrint(onoffgubun) {
	var chk; var itembarcode; var itembarcodearr=''; var papername='';
	var found = false;
	var papername = frm.papername.value;

//	if (papername==''){
//		alert('인쇄 하실 쇼카드를 선택해 주세요.');
//		return;
//	}

	for (var i = 0; ; i++) {
		itembarcode = document.getElementById("itembarcode_" + i);
		chk = document.getElementById("cksel" + i);

		if (itembarcode == undefined) {
			break;
		}

		if (chk.checked != true) {
			continue;
		}

		found = true;

		itembarcodearr = itembarcodearr + document.getElementById("itembarcode_" + i).value.trim() + ','
	}

	if (found == false) {
		alert("선택된 상품이 없습니다.");
		return;
	}

	alert("[필수]인쇄될 쇼카드가 뜨면, 마우스 오른쪽 버튼을 눌러 인쇄를 클릭해 주세요.\n\n상단에 있는 쇼카드출력설명서 대로, 설정후에 인쇄하셔야 간격이 정상적으로 인쇄됩니다.");

	frm.itembarcodearr.value=itembarcodearr;
	frm.action='/common/barcode/paperbarcodeprint_on_off_multi.asp';
	frm.target='_blank';
	frm.method="post"
	frm.submit();
}

function IndexSudongBarcodePrint(){
	var popwin = window.open('/common/barcode/sudongindexprint.asp?menupos=<%=menupos%>','IndexSudong','width=1024,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<OBJECT
	  id=iaxobject
	  classid="clsid:5D776FEA-8C6B-4C53-8EC3-3585FC040BDB"
	  codebase="https://scm.10x10.co.kr/common/cab/tenbarPrint.cab#version=1,0,0,29"
	  width=0
	  height=0
	  align=center
	  hspace=0
	  vspace=0
>
</OBJECT>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="onoffgubun" value="<%= onoffgubun %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<input type="hidden" name="itembarcodearr" value="">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<% if (C_IS_Maker_Upche) then %>
			* 브랜드 : <%= makerid %>
			<input type="hidden" name="makerid" value="<%= makerid %>">
			&nbsp;
		<% else %>
			* 브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
			&nbsp;
		<% end if %>

		* 상품명 : <input type="text" class="text" name="prdname" value="<%= prdname %>" size="32" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
		&nbsp;
		* 물류코드 :
		<input type="text" class="text" name="prdcode" value="<%= prdcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
		&nbsp;
		* 상품코드 : 
		<textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea><!--onKeyPress="if (event.keyCode == 13) 	reg('');"-->
	</td>

	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="reg('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 범용바코드 :
		<input type="text" class="text" name="generalbarcode" value="<%= generalbarcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">

		<% if listgubun = "ITEM" or listgubun = "JUMUN" or listgubun = "PACKING" then %>
			<% if not(C_IS_Maker_Upche) then %>
				&nbsp;
				<input type="checkbox" class="checkbox" name="displayrealstockno" value="Y" onClick="reg('');" <% if (displayrealstockno = "Y") then %>checked<% end if %>>재고로 수량설정
			<% end if %>
		<% elseif listgubun = "UPCHEJUMUN" then %>
			&nbsp;
			<input type="checkbox" class="checkbox" name="displayrealstockno" value="Y" onClick="reg('');" <% if (displayrealstockno = "Y") then %>checked<% end if %>>
			입출고내역으로 수량설정
		<% end if %>

		<% if not(C_IS_Maker_Upche) then %>
			* 매장별상품명 :
			<input type="text" class="text" name="shopitemname" value="<%= shopitemname %>" size="24" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
			&nbsp;
			<input type="checkbox" class="checkbox" name="currentstockexist" value="Y" onClick="reg('');" <% if (currentstockexist = "Y") then %>checked<% end if %>> 입고내역존재상품
			&nbsp;
			<input type="checkbox" class="checkbox" name="realstockonemore" value="Y" onClick="reg('');" <% if (realstockonemore = "Y") then %>checked<% end if %>> 재고존재상품
			&nbsp;
			<input type="checkbox" class="checkbox" name="shopitemnameinserted" value="Y" onClick="reg('');" <% if (shopitemnameinserted = "Y") then %>checked<% end if %>> 매장별상품명등록상품만
			&nbsp;
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<% 
		'/상품리스트와 업체주문리스트만
		if listgubun = "ITEM" or listgubun = "UPCHEJUMUN" then
		%>
			<%
			'직영/가맹점
			if (C_IS_SHOP) then
			%>
				<% if C_IS_OWN_SHOP then %>
					* 매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
				<% else %>
					* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
				<% end if %>
			<% else %>
				<% if (C_IS_Maker_Upche) then %>
					* 매장 : <% drawBoxDirectIpchulOffShopByMakerchfg "shopid", shopid, makerid, " onchange='reg("""");'", " 'B011','B012','B013'" %>
				<% else %>
					* 매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
				<% end if %>
			<% end if %>
		<% end if %>

		&nbsp;&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox_utf8.asp"-->
		&nbsp;&nbsp;
		* 상품구분:<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
	</td>
</tr>
</table>
<!-- 검색 끝 -->

<br>
<!-- #include virtual="/common/barcode/inc_setting_menu_barcodeprint.asp" -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<br><font color="red">리스트기준 :</font>
		<input type="radio" name="listgubun" value="ITEM" onClick="ipchulview(this.value,''); reg('');" <% if listgubun = "ITEM" then response.write " checked" %>>상품리스트

		<input type="radio" name="listgubun" value="UPCHEJUMUN" onClick="ipchulview(this.value,''); reg('');" <% if listgubun = "UPCHEJUMUN" then response.write " checked" %>>업체주문리스트
		<% if shopid <> "" then %>
			<% drawipchulmaster "ipchul",ipchul,shopid ,makerid ," onchange=reg('');","" %>
			<script language="javascript">
				ipchulview("<%=listgubun%>","ONLOAD");
			</script>
		<% else %>
			<select name="ipchul" disabled><option value=''>매장을 선택하세요</option></select>
		<% end if %>

		<input type="radio" name="listgubun" value="JUMUN" onClick="ipchulview(this.value,''); reg('');" <% if listgubun = "JUMUN" then response.write " checked" %>>주문리스트
		<input type="radio" name="listgubun" value="PACKING" onClick="ipchulview(this.value,''); reg('');" <% if listgubun = "PACKING" then response.write " checked" %>>패킹리스트
	</td>
	<td align="right"></td>
</tr>
</table>

<% if listgubun = "PACKING" or listgubun = "JUMUN" then %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="30" bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("gray") %>">매장</td>
		<td width="300"><%= ocstoragemaster.FRectShopId %></td>
		<td width="120" bgcolor="<%= adminColor("gray") %>">매장명</td>
		<td>
			<%= locationnameto %>
		</td>
	</tr>

	<% 'if (IsOneOrderOnly <> True) then %>
	<% if listgubun = "PACKING" then %>
		<tr height="30" bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%= adminColor("gray") %>">출고지시일</td>
			<td width="300" colspan=3><%= AddSpace(ocstoragemaster.GetPackingDayList) %></td>
		</tr>
		<tr height="30" bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%= adminColor("gray") %>">Inner박스&nbsp;번호</td>
			<td width="300">
				<select name="boxno" onChange="frmmaster.submit()">
					<option value="">ALL</option>
					<% If Not IsNULL(maxboxno) then %>
					<% for i=1 to maxboxno %>
						<option value="<%= i %>" <%if (CStr(boxno) = CStr(i)) then %>selected<% end if %>>No. <%= i %></option>
					<% next %>
					<% end if %>
				</select>
			</td>
			<td width="120" bgcolor="<%= adminColor("gray") %>">Inner박스&nbsp;IDX</td>
			<td>
				<% if (boxno <> "") then %>
					10<% if shopseq<>"" then %>-<%= format00(6,shopseq) %><% end if %>-<%= Replace(Replace(ocstoragemaster.GetPackingDayList, "-", ""), "-", "") %>-<%= format00(3,boxno) %>

					<% if printername = "TTP-243_80x50" or printername = "TEC_B-FV4_80x50" then %>
						<br><input type="button" value="Inner박스바코드출력" onclick="innerboxindexprint()" class="button">
					<% end if %>
				<% end if %>
			</td>
		</tr>
		<tr height="30" bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%= adminColor("gray") %>">Carton박스 번호</td>
			<td width="300">
				<% If False and Not IsNULL(maxcartonboxno) then %>
				<select name="cartonboxno" onChange="frmmaster.boxno.value = ''; frmmaster.submit()">
					<option value="">ALL</option>
					<% for i=1 to maxcartonboxno %>
					<option value="<%= i %>" <%if (CStr(cartonboxno) = CStr(i)) then %>selected<% end if %>>No. <%= i %></option>
					<% next %>
				</select>
				<% else %>
				<%= cartonboxno %>
				<% end if %>
			</td>
			<td width="120" bgcolor="<%= adminColor("gray") %>">Carton박스&nbsp;IDX</td>
			<td>
				<% if (cartonboxno <> "") then %>
					10<% if shopseq<>"" then %>-<%= format00(6,shopseq) %><% end if %>-<%= Replace(Replace(ocstoragemaster.GetPackingDayList, "-", ""), "-", "") %>-<%= format00(3,cartonboxno) %>

					<% if printername = "TTP-243_80x50" or printername = "TEC_B-FV4_80x50" then %>
						<br><input type="button" value="Carton박스바코드출력" onclick="cartonboxindexprint()" class="button">
					<% end if %>
				<% end if %>
			</td>
		</tr>
		<tr height="30" bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%= adminColor("gray") %>">관련주문</td>
			<td width="300">
				<%= AddSpace(ocstoragemaster.GetOrderCodeList) %>
			</td>
			<td width="120" bgcolor="<%= adminColor("gray") %>">주문IDX</td>
			<td >
				<input type="text" name="masteridx" value="<%= masteridx %>" size="10" maxlength=10 onKeyPress="if (event.keyCode == 13) reg('');">
			</td>
		</tr>
	<% else %>
		<tr height="30" bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%= adminColor("gray") %>">주문코드</td>
			<td colspan=4>
				<input type="text" name="baljucode" value="<%= baljucode %>" size="12" maxlength=12 onKeyPress="if (event.keyCode == 13) reg('');">
			</td>
			<!--<td width="120" bgcolor="<%'= adminColor("gray") %>">주문IDX</td>
			<td >
				<input type="text" name="masteridx" value="<%'= masteridx %>" size="10" maxlength=10>
			</td>-->
		</tr>
	<% end if %>

	</table>
<% end if %>

<br>
<!-- #include virtual="/common/barcode/inc_button_barcodeprint.asp" -->
</form>

<form name="frmList" id="frmList" method="POST" tyle="margin:0px;">
<input type="hidden" name="barcodetype" value="">
<input type="hidden" name="currencychar" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="25">
        검색결과 : <b><%= oproduct.FTotalcount %></b>
        &nbsp;
        <b><%= page %> / <%= oproduct.FTotalpage %></b>
        &nbsp;&nbsp;※ 상품명에 특수문자가 있는 경우 검색되지 않습니다.
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=20>
		<input type="checkbox" name="ckall" onClick="SelectCk(this);" id="ckall">
	</td>

	<% if listgubun = "PACKING" then %>
		<td>박스<br>번호</td>
		<td>주문코드</td>
	<% end if %>

	<td width=50>이미지</td>
	<td width=120>물류코드<br><font color=blue>[범용바코드]</font></td>
	<td>브랜드</td>
	<td>
		상품명<font color=blue>[옵션명]</font>
		<% if (useforeigndata = "Y") then %>
			<br>샵별상품명<font color=blue>[샵별옵션명]</font>
		<% end if %>
	</td>
	<td width=60>
		소비자가
		<% if (useforeigndata = "Y") then %>
			<br>[샵별금액]
		<% end if %>
	</td>
	<td width=60>
		판매가
	</td>
	<td width=60>
		할인가
	</td>
	<td width=60>수량</td>
	<td width=80>비고</td>
</tr>

<% if not(isdispconfirm) then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center">
			<font color="red"><strong>오프라인 상품정보 조회 권한이 없습니다.</strong></font>
		</td>
	</tr>
<% elseif oproduct.FresultCount > 0 then %>
<% for i=0 to oproduct.FresultCount-1 %>
<input type="hidden" name="location_name" value="<%= replace(oproduct.FItemList(i).Flocation_name,"""","") %>">
<input type="hidden" id="makerid_<%= i %>" name="locationid" value="<%= oproduct.FItemList(i).Flocationid %>">
<input type="hidden" id="socname_<%= i %>" name="socname" value="<%= replace(oproduct.FItemList(i).fsocname,"""","") %>">
<input type="hidden" id="socname_kor_<%= i %>" name="socname_kor" value="<%= replace(oproduct.FItemList(i).fsocname_kor,"""","") %>">
<input type="hidden" id="itemgubun_<%= i %>" name="itemgubun" value="<%= oproduct.FItemList(i).Fitemgubun %>">
<input type="hidden" id="itemid_<%= i %>" name="itemid" value="<%= oproduct.FItemList(i).Fitemid %>">
<input type="hidden" id="itemoption_<%= i %>" name="itemoption" value="<%= oproduct.FItemList(i).Fitemoption %>">
<input type="hidden" id="itembarcode_<%= i %>" name="prdcode" value="<%= BF_MakeTenBarcode(oproduct.FItemList(i).Fitemgubun, oproduct.FItemList(i).Fitemid, oproduct.FItemList(i).Fitemoption) %>">
<input type="hidden" id="publicbarcode_<%= i %>" name="generalbarcode" value="<%= replace(oproduct.FItemList(i).Fgeneralbarcode,"""","") %>">
<input type="hidden" id="itemname_<%= i %>" name="prdname" value="<%= replace(oproduct.FItemList(i).Fprdname,"""","") %>">
<input type="hidden" id="itemoptionname_<%= i %>" name="prdoptionname" value="<%= replace(oproduct.FItemList(i).Fitemoptionname,"""","") %>">
<input type="hidden" id="customerprice_<%= i %>" name="customerprice" value="<%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %>">
<input type="hidden" id="sellprice_<%= i %>" name="sellprice" value="<%= FormatNumber(oproduct.FItemList(i).Fsellprice,0) %>">
<input type="hidden" id="itemname_foreign_<%= i %>" name="prdname_foreign" value="<%= replace(oproduct.FItemList(i).Flcitemname,"""","") %>">
<input type="hidden" id="itemoptionname_foreign_<%= i %>" name="prdoptionname_foreign" value="<%= replace(oproduct.FItemList(i).Flcitemoptionname,"""","") %>">
<input type="hidden" id="customerprice_foreign_<%= i %>" name="customerprice_foreign" value="<%= round(oproduct.FItemList(i).Flcprice,2) %>">
<input type="hidden" id="saleprice_<%= i %>" name="saleprice" value="<%= FormatNumber(oproduct.FItemList(i).Flcprice,0) %>">
<input type="hidden" id="saleyn_<%= i %>" name="saleyn" value="<% if oproduct.FItemList(i).Fsaleyn="Y" and oproduct.FItemList(i).Fcustomerprice>oproduct.FItemList(i).Flcprice then %>Y<% else %>N<% end if %>">
<input type="hidden" id="catename1_<%= i %>" name="catename1" value="<%= replace(oproduct.FItemList(i).fcatename1,"""","") %>">
<input type="hidden" id="catename2_<%= i %>" name="catename2" value="<%= replace(oproduct.FItemList(i).fcatename2,"""","") %>">
<input type="hidden" id="catename_cn_gan2_<%= i %>" name="catename_cn_gan2" value="<%= replace(oproduct.FItemList(i).fcatename_cn_gan2,"""","") %>">
<input type="hidden" id="catename_cn_bun2_<%= i %>" name="catename_cn_bun2" value="<%= replace(oproduct.FItemList(i).fcatename_cn_bun2,"""","") %>">
<input type="hidden" id="catename3_<%= i %>" name="catename3" value="<%= replace(oproduct.FItemList(i).fcatename3,"""","") %>">
<input type="hidden" id="catename_cn_gan3_<%= i %>" name="catename_cn_gan3" value="<%= replace(oproduct.FItemList(i).fcatename_cn_gan3,"""","") %>">
<input type="hidden" id="catename_cn_bun3_<%= i %>" name="catename_cn_bun3" value="<%= replace(oproduct.FItemList(i).fcatename_cn_bun3,"""","") %>">
<input type="hidden" id="catename_eng2_<%= i %>" name="catename_eng2" value="<%= replace(oproduct.FItemList(i).fcatename_eng2,"""","") %>">
<input type="hidden" id="catename_eng3_<%= i %>" name="catename_eng3" value="<%= replace(oproduct.FItemList(i).fcatename_eng3,"""","") %>">
<input type="hidden" id="sourcearea_10x10_<%= i %>" name="sourcearea_10x10" value="<%= replace(oproduct.FItemList(i).fsourcearea_10x10,"""","") %>">
<input type="hidden" id="sourcearea_en_<%= i %>" name="sourcearea_en" value="<%= replace(oproduct.FItemList(i).fsourcearea_en,"""","") %>">
<input type="hidden" id="itemsource_10x10_<%= i %>" name="itemsource_10x10" value="<%= replace(oproduct.FItemList(i).fitemsource_10x10,"""","") %>">
<input type="hidden" id="itemsource_en_<%= i %>" name="itemsource_en" value="<%= replace(oproduct.FItemList(i).fitemsource_en,"""","") %>">
<input type="hidden" id="itemsize_10x10_<%= i %>" name="itemsize_10x10" value="<%= replace(oproduct.FItemList(i).fitemsize_10x10,"""","") %>">
<input type="hidden" id="itemsize_en_<%= i %>" name="itemsize_en" value="<%= replace(oproduct.FItemList(i).fitemsize_en,"""","") %>">
<input type="hidden" id="prtidx_<%= i %>" name="prtidx" value="<%= oproduct.FItemList(i).fprtidx %>">
<input type="hidden" id="itemrackcode_<%= i %>" name="itemrackcode" value="<%= replace(oproduct.FItemList(i).fitemrackcode,"""","") %>">
<input type="hidden" id="itemoptionrackcode_<%= i %>" name="itemoptionrackcode" value="">
<input type="hidden" id="subitemrackcode_<%= i %>" name="subitemrackcode" value="<%= replace(oproduct.FItemList(i).fsubitemrackcode,"""","") %>">

<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" id="cksel<%= i %>" name="cksel" value="<%=i+1%>" onClick="AnCheckClick(this);"></td>

	<% if listgubun = "PACKING" then %>
		<td><%= AddSpace(oproduct.FItemList(i).Fboxno) %></td>
		<td><%= AddSpace(oproduct.FItemList(i).Fbaljucode) %></td>
	<% end if %>

	<td><img src="<%= oproduct.FItemList(i).Fmainimageurl %>" width=50 height=50></td>
	<td>
	    <% if C_IS_Maker_Upche and Not(IsDirectIpchulContractExistsBrand) and Not isupcheitemreg then %>
			<a href="javascript:jsAlertNoAuth('권한이 없습니다. - 매장 직접 입고 브랜드만 수정 가능합니다');" onfocus="this.blur()">
		<% Else %>
			<a href="#" onclick="pop_itemedit_off_edit('<%= oproduct.FItemList(i).Fprdcode %>'); return false;" onfocus="this.blur()">
		<% End If %>

		<%= AddSpace(oproduct.FItemList(i).Fprdcode) %>

		<% if oproduct.FItemList(i).Fgeneralbarcode <> "" then %>
			<br><font color=blue>[<%= AddSpace(oproduct.FItemList(i).Fgeneralbarcode) %>]</font>
		<% end if %>
		</a>
	</td>
	<td>
		<%= oproduct.FItemList(i).Flocationid %>
	</td>
	<td align="left">
		<%= AddSpace(oproduct.FItemList(i).Fprdname) %>
		<% if oproduct.FItemList(i).Fitemoptionname <> "" then %>
			<font color=blue>[<%= AddSpace(oproduct.FItemList(i).Fitemoptionname) %>]</font>
		<% end if %>
		<% if (useforeigndata = "Y") then %>
			<br><%= AddSpace(oproduct.FItemList(i).Flcitemname) %><font color=blue>[<%= AddSpace(oproduct.FItemList(i).Flcitemoptionname) %>]</font>
		<% end if %>
	</td>
	<td align="right">
		<%= currencychar %>&nbsp;<%= AddSpace(FormatNumber(oproduct.FItemList(i).Fcustomerprice,0)) %>

		<% if (useforeigndata = "Y") then %>
			<br><%= AddSpace(oproduct.FItemList(i).Flcprice) %>&nbsp;<%= currencyChar %>
		<% end if %>
	</td>
	<td align="right">
		<%= currencychar %>&nbsp;<%= AddSpace(FormatNumber(oproduct.FItemList(i).Fsellprice,0)) %>
	</td>
	<td align="right">
		<%
		'/할인 처리
		if oproduct.FItemList(i).Fsaleyn="Y" and oproduct.FItemList(i).Fcustomerprice>oproduct.FItemList(i).Flcprice then
		%>
			<font color='red'><%= currencychar %>&nbsp;<%= AddSpace(FormatNumber(oproduct.FItemList(i).Flcprice,0)) %></font>
		<% end if %>
	</td>
	<td>
		<% if listgubun = "ITEM" then %>
			<input type="text" id="printno_<%= i %>" class="text" name="fixedno" value="<% if (displayrealstockno = "Y") then %><%= oproduct.FItemList(i).Frealstockno %><% else %>1<% end if %>" size="4" maxlength="4" onKeyPress="CheckThis(<%= i %>);">
		<% elseif listgubun = "UPCHEJUMUN" then %>
			<input type="text" id="printno_<%= i %>" class="text" name="fixedno" value="<% if (displayrealstockno = "Y") then %><%= oproduct.FItemList(i).Fitemno %><% else %>1<% end if %>" size="4" maxlength="4" onKeyPress="CheckThis(<%= i %>);">
		<% elseif listgubun = "JUMUN" then %>
			<% if isfixed then %>
				<input type="text" id="printno_<%= i %>" class="text" name="fixedno" value="<% if (displayrealstockno = "Y") then %><%= oproduct.FItemList(i).Ffixedno %><% else %>1<% end if %>" size="4" maxlength="4" onFocus="this.select()" onKeyPress="CheckThis(<%= i %>);">
			<% else %>
				<input type="text" id="printno_<%= i %>" class="text" name="fixedno" value="<% if (displayrealstockno = "Y") then %><%= oproduct.FItemList(i).Frequestedno %><% else %>1<% end if %>" size="4" maxlength="4" onFocus="this.select()" onKeyPress="CheckThis(<%= i %>);">
			<% end if %>
		<% elseif listgubun = "PACKING" then %>
			<input type="text" id="printno_<%= i %>" class="text" name="fixedno" value="<% if (displayrealstockno = "Y") then %><%= oproduct.FItemList(i).Ffixedno %><% else %>1<% end if %>" size="4" maxlength="4" onFocus="this.select()" onKeyPress="CheckThis(<%= i %>);">
		<% end if %>
	</td>
	<td>
		<!--<input type="checkbox" name="IsSellPricePrint" checked>가격출력-->
	</td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
       	<% if oproduct.HasPreScroll then %>
			<span class="list_link"><a href="javascript:reg(<%=oproduct.StartScrollPage-1%>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oproduct.StartScrollPage to oproduct.StartScrollPage + oproduct.FScrollCount - 1 %>
			<% if (i > oproduct.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oproduct.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:reg(<%=i%>);" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if oproduct.HasNextScroll then %>
			<span class="list_link"><a href="javascript:reg(<%=i%>);">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center">
			<% if not(isdispsql) then %>
				<font color="red"><strong>
				<% if listgubun = "ITEM" or listgubun = "UPCHEJUMUN" then %>
					상품리스트는 검색 조건(매장[필수],브랜드[필수],상품명,물류코드,상품코드,범용바코드)을 입력 하셔야 검색이 됩니다.
					<Br><br>업체주문리스트는 검색조건(매장)을 입력 하시고 주문을 선택하셔야 검색이 됩니다.
				<% elseif listgubun = "JUMUN" or listgubun = "PACKING" then %>
					주문리스트,패킹리스트는 검색조건(주문코드)을 입력 하셔야 검색이 됩니다.
				<% end if %>
				</strong></font>
			<% else %>
				[검색결과가 없습니다.]
			<% end if %>
		</td>
	</tr>
<% end if %>
</table>
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
