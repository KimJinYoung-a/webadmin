<%@ language=vbscript %>
<% option Explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description : 세금계산서 발행정보
' History : 서동석 생성
'			2022.10.31 한용민 수정(위하고 세금계산서 발행 api 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/offshop/etcmeachulcls.asp"-->
<!-- #include virtual="/cscenter/lib/TaxSheetFunc.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<%
dim orderserial, issuetype, etcmeachulIdx, taxidx, previssuecount, mode, supplyBusiIdx, busiIdx, IsIssusOK
dim ordercancelyn, chulgoyear, sqlStr, errMSG, i, tplcompanyid, tplcompanyname, tplgroupid, tplbillUserID
dim userid, busiNo, busiSubNo, busiName, busiCEOName, busiAddr, busiType, busiItem, repName, repEmail, repTel, confirmYn
dim groupID, itemname, chulgoPrice, taxtype
	IsIssusOK = True
	orderserial 	= requestcheckvar(trim(request("orderserial")),11)
	issuetype 		= trim(request("issuetype"))
	chulgoyear 		= trim(request("chulgoyear"))
	taxidx			= requestcheckvar(getNumeric(trim(request("taxidx"))),10)
	etcmeachulIdx	= trim(request("idx"))		'// 기타매출코드

previssuecount = 0

dim oTax, oTaxCheck
set oTax = new CTax
oTax.FRecttaxIdx = taxidx

if (taxidx = "") then
	mode = "new"
	oTax.GetTaxEmptyOne

	if (orderserial <> "") or (issuetype = "orderserial") then
		oTax.FOneItem.ForderIdx 	= orderserial
		oTax.FOneItem.Ftaxtype		= "Y"
		oTax.FOneItem.FtotalTax		= 0

		groupID 	= request("groupID")		'// 공급자 그룹코드
		itemname 	= request("itemname")
		chulgoPrice = request("chulgoPrice")
		taxtype 	= request("taxtype")
		busiIdx		= request("busiIdx")

		if (chulgoyear <> "") and (chulgoyear >= "2014") and (groupID = "") then
			'// 발행요청 등록
			oTax.FOneItem.Fbilldiv 		= "11"

			oTax.FOneItem.FsellBizCd 		= "0000000101"		'// 온라인(공통)
			oTax.FOneItem.Fselltype 		= "20166"			'// B2C
			oTax.FOneItem.Ftaxissuetype		= "C"

			oTax.FOneItem.FconsignYN	= "N"
			oTax.FOneItem.FissueMethod	= "WEHAGO"
		elseif (groupID <> "") then
			'// 업체별 계산서 발행
			oTax.FOneItem.Fbilldiv 		= "11"

			oTax.FOneItem.FsellBizCd 		= "0000000101"		'// 온라인(공통)
			oTax.FOneItem.Fselltype 		= "20166"			'// B2C
			oTax.FOneItem.Ftaxissuetype		= "C"

			oTax.FOneItem.FconsignYN	= "N"
			oTax.FOneItem.FissueMethod	= "WEHAGO"

			if (groupID <> "G00456") then
				oTax.FOneItem.FconsignYN	= "Y"
				oTax.FOneItem.FissueMethod	= "eSero"
			end if
		else
			'// 2013년도 출고분 계산서 발행
			oTax.FOneItem.Fbilldiv 		= "01"

			oTax.FOneItem.FsellBizCd 		= "0000000101"		'// 온라인(공통)
			oTax.FOneItem.Fselltype 		= "20166"			'// B2C
			oTax.FOneItem.Ftaxissuetype		= "C"

			oTax.FOneItem.FconsignYN	= "N"
			oTax.FOneItem.FissueMethod	= "WEHAGO"
		end if

		oTax.FOneItem.Fuserid		= session("ssBctId")

		'// --------------------------------------------------------------------
		dim ojumun
		set ojumun = new COrderMaster

		if (orderserial <> "") then
			ojumun.FRectOrderSerial = orderserial
			ojumun.QuickSearchOrderMaster
		end if

		if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
			ojumun.FRectOldOrder = "on"
			ojumun.QuickSearchOrderMaster
		end if

		if (ojumun.FResultCount < 1) and (errMSG = "") then
			errMSG = "잘못된 주문번호 입니다."
		else
			oTax.FOneItem.FisueDate = getMayTaxDate(ojumun.FOneItem.Fipkumdate)
		end if

		'// --------------------------------------------------------------------
		dim oTaxPrevIssue
		set oTaxPrevIssue = new CTax

		if (errMSG = "") then
			oTax.FCurrPage = 1
			oTax.FPageSize = 100
			''oTax.FRectsearchBilldiv = "01"				'소비자매출
			oTax.FRectsearchKey = "t.userid"
			oTax.FRectDelYn = "N"

			if (ojumun.FOneItem.FUserID <> "") then
				oTax.FRectsearchString = ojumun.FOneItem.FUserID
				userid = ojumun.FOneItem.FUserID
			else
				oTax.FRectsearchString = "----"
			end if

			oTax.GetTaxList

			previssuecount = oTax.FTotalCount
		end if

		''기발행 세금계산서인지 체크
		if errMSG = "" and (oTax.FOneItem.ForderIdx <> "") then
			set oTaxCheck = new CTax

			oTaxCheck.FRectsearchKey = " t.orderserial "
			oTaxCheck.FRectsearchString = CStr(oTax.FOneItem.ForderIdx)
			oTaxCheck.FRectDelYn = "N"

			if (groupID <> "") then
				oTaxCheck.FRectSupplyGroupID = groupID
			end if

			oTaxCheck.GetTaxList

			if oTaxCheck.FResultCount > 0 and (oTax.FOneItem.Fbilldiv = "01" or groupID <> "") then
				if oTaxCheck.FTaxList(0).FisueYn="Y" then
					if (errMSG = "") then
						errMSG = "이미 발행된 세금계산서가 있습니다.\n\n재발행 하시려면 거래처에 [취소요청]후 매출 세금계산서 목록에서 [삭제]후 발행하셔야 합니다"
					end if
				else
					if (errMSG = "") then
						errMSG = "발행대기중인 세금계산서가 있습니다.\n\n재발행 하시려면 매출 세금계산서 목록에서 [삭제]후 발행하셔야 합니다"
					end if
				end if
			end if
		end if

		sqlStr =	"select ( select " &_
				"			Case " &_
				"				When count(idx)>1 Then max(itemname) + '외 ' + Cast((count(idx)-1) as varchar) + '건' " &_
				"				Else max(itemname) " &_
				"			End " &_
				"		from db_order.[dbo].tbl_order_detail " &_
				"		where orderserial='" & orderserial & "' and itemid<>0 and cancelyn='N' group by orderserial " &_
				"	) as itemname " &_
				"	, subtotalprice, accountdiv, IsNull(sumPaymentEtc, 0) as sumPaymentEtc " &_
				"from db_order.[dbo].tbl_order_master " &_
				"Where orderserial = '" & orderserial & "'"

		'response.write sqlStr &"<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		if Not(rsget.EOF or rsget.BOF) then
			oTax.FOneItem.Fitemname = rsget("itemname")

			if (chulgoyear <> "") and (chulgoyear >= "2014") and (groupID <> "") then
				oTax.FOneItem.Fitemname = "업체별 상품"

				if request("itemname") <> "" then
					oTax.FOneItem.Fitemname = request("itemname")
					oTax.FOneItem.FtotalPrice = request("chulgoPrice")
					oTax.FOneItem.Ftaxtype = request("taxtype")
				end if
			elseif (CStr(rsget("accountdiv")) = "7") or (CStr(rsget("accountdiv")) = "20") then
				'무통장, 실시간이체 : 전체금액
				oTax.FOneItem.FtotalPrice = rsget("subtotalprice")
			else
				'보조결제금액만
				oTax.FOneItem.FtotalPrice = rsget("sumPaymentEtc")
			end if
		end if
		rsget.close

		if (groupID <> "") then
			dim ogroupSupply
			set ogroupSupply = new CPartnerGroup
			ogroupSupply.FRectGroupid = groupID
			ogroupSupply.GetOneGroupInfo

			if (groupID = "G00456") then
				'// 공급자(텐바이텐)
				oTax.FOneItem.FsupplyBusiNo			= TENBYTEN_SOCNO
				oTax.FOneItem.FsupplyBusiName		= TENBYTEN_SOCNAME
				oTax.FOneItem.FsupplyBusiCEOName	= TENBYTEN_CEONAME
				oTax.FOneItem.FsupplyBusiAddr		= TENBYTEN_SOCADDR
				oTax.FOneItem.FsupplyBusiType		= TENBYTEN_SOCSTATUS
				oTax.FOneItem.FsupplyBusiItem		= TENBYTEN_SOCEVENT
				oTax.FOneItem.FsupplyRepName		= TENBYTEN_MANAGERNAME
				oTax.FOneItem.FsupplyRepTel			= TENBYTEN_MANAGERPHONE
				oTax.FOneItem.FsupplyRepEmail		= TENBYTEN_MANAGERMAIL
			else
				'// 공급자(업체)
				oTax.FOneItem.FsupplybusiNo			= ogroupSupply.FOneItem.Fcompany_no
				oTax.FOneItem.FsupplybusiName		= ogroupSupply.FOneItem.FCompany_name
				oTax.FOneItem.FsupplybusiCEOName	= ogroupSupply.FOneItem.Fceoname
				oTax.FOneItem.FsupplybusiAddr		= ogroupSupply.FOneItem.Fcompany_address & " " & ogroupSupply.FOneItem.Fcompany_address2
				oTax.FOneItem.FsupplybusiType		= ogroupSupply.FOneItem.Fcompany_uptae
				oTax.FOneItem.FsupplybusiItem		= ogroupSupply.FOneItem.Fcompany_upjong
				oTax.FOneItem.FsupplyrepName		= ogroupSupply.FOneItem.Fjungsan_name
				oTax.FOneItem.FsupplyrepTel			= ogroupSupply.FOneItem.Fjungsan_hp
				oTax.FOneItem.FsupplyrepEmail		= ogroupSupply.FOneItem.Fjungsan_email
			end if


		end if

		if (busiIdx <> "") then
			sqlStr = " select top 1 "
			sqlStr = sqlStr + " 	b.busiNo, b.busiSubNo, b.busiName, b.busiCEOName, b.busiAddr, b.busiType, b.busiItem, b.repName, b.repEmail, b.repTel "
			sqlStr = sqlStr + " from db_order.dbo.tbl_busiInfo b "
			sqlStr = sqlStr + " where b.busiidx = " + CStr(busiIdx) + " "
			rsget.Open sqlStr, dbget, 1

			if Not(rsget.EOF or rsget.BOF) then
				'// 공급받는자
				oTax.FOneItem.FbusiNo = rsget("busiNo")
				oTax.FOneItem.FbusiSubNo = rsget("busiSubNo")
				oTax.FOneItem.FbusiName = rsget("busiName")
				oTax.FOneItem.FbusiCEOName = rsget("busiCEOName")
				oTax.FOneItem.FbusiAddr = rsget("busiAddr")
				oTax.FOneItem.FbusiType = rsget("busiType")
				oTax.FOneItem.FbusiItem = rsget("busiItem")
				oTax.FOneItem.FrepName = rsget("repName")
				oTax.FOneItem.FrepEmail = rsget("repEmail")
				oTax.FOneItem.FrepTel = rsget("repTel")
			end if
			rsget.close
		end if

	end if

	if (issuetype = "etcmeachul") and (etcmeachulIdx <> "") then

		'// --------------------------------------------------------------------
		'// 기타매출
		dim oetcmeachul
		set oetcmeachul = new CEtcMeachul
		oetcmeachul.FRectidx = etcmeachulIdx
		oetcmeachul.getOneEtcMeachul

		oTax.FOneItem.FsellBizCd	= oetcmeachul.FOneItem.Fbizsection_cd
		oTax.FOneItem.Fselltype		= oetcmeachul.FOneItem.Fselltype
		oTax.FOneItem.Ftaxissuetype	= "E"
		oTax.FOneItem.Ftaxtype		= "Y"
		oTax.FOneItem.Fbilldiv		= "51"
		oTax.FOneItem.FconsignYN	= "N"
		oTax.FOneItem.FissueMethod	= "WEHAGO"

		oTax.FOneItem.Fuserid		= session("ssBctId")

		'// --------------------------------------------------------------------
		'삽아이디에서 그룹코드 추출
		dim opartner
		set opartner = new CPartnerUser

		opartner.FCurrpage = 1
		opartner.FPageSize = 100
		opartner.FRectDesignerID = oetcmeachul.FOneItem.Fshopid
		opartner.FRectUserDiv = "all"

		opartner.GetPartnerNUserCList

		'// --------------------------------------------------------------------
		'그룹코드에서 세금계산서/정산담당자 정보 추출
		dim ogroup
		set ogroup = new CPartnerGroup

		if (opartner.FResultCount > 0) then
			ogroup.FRectGroupid = opartner.FPartnerList(0).FGroupID
			ogroup.GetOneGroupInfo
		else
			ogroup.FResultCount = 0
		end if

		if (opartner.FResultCount < 1) then
			errMSG = "잘못된 브랜드입니다."
		elseif (ogroup.FResultCount < 1) then
			errMSG = "그룹코드가 지정되어 있지 않은 업체입니다."
		else
			'// 공급받는자
			if (ogroup.FOneItem.Fcompany_no <> "211-87-00620") then
				if ogroup.FOneItem.fBIZ_NO<>"" then
					oTax.FOneItem.FbusiNo		= ogroup.FOneItem.fBIZ_NO
				else					
					oTax.FOneItem.FbusiNo		= ogroup.FOneItem.Fcompany_no
				end if
				if ogroup.FOneItem.fCUST_NM<>"" then
					oTax.FOneItem.FbusiName		= ogroup.FOneItem.fCUST_NM
				else					
					oTax.FOneItem.FbusiName		= ogroup.FOneItem.FCompany_name
				end if
				if ogroup.FOneItem.fCEO_NM<>"" then
					oTax.FOneItem.FbusiCEOName	= ogroup.FOneItem.fCEO_NM
				else					
					oTax.FOneItem.FbusiCEOName	= ogroup.FOneItem.Fceoname
				end if
				if ogroup.FOneItem.faddr<>"" then
					oTax.FOneItem.FbusiAddr		= ogroup.FOneItem.faddr		' ogroup.FOneItem.fPOST_CD
				else					
					oTax.FOneItem.FbusiAddr		= ogroup.FOneItem.Fcompany_address & " " & ogroup.FOneItem.Fcompany_address2
				end if
				if ogroup.FOneItem.fBSCD<>"" then
					oTax.FOneItem.FbusiType		= ogroup.FOneItem.fBSCD
				else					
					oTax.FOneItem.FbusiType		= ogroup.FOneItem.Fcompany_uptae
				end if
				if ogroup.FOneItem.fINTP<>"" then
					oTax.FOneItem.FbusiItem		= ogroup.FOneItem.fINTP
				else					
					oTax.FOneItem.FbusiItem		= ogroup.FOneItem.Fcompany_upjong
				end if
				
				oTax.FOneItem.FrepName		= ogroup.FOneItem.Fjungsan_name

				if ogroup.FOneItem.fTEL_NO<>"" then
					oTax.FOneItem.FrepTel		= ogroup.FOneItem.fTEL_NO
				else					
					oTax.FOneItem.FrepTel		= ogroup.FOneItem.Fjungsan_hp
				end if
				if ogroup.FOneItem.fEMAIL<>"" then
					oTax.FOneItem.FrepEmail		= ogroup.FOneItem.fEMAIL
				else					
					oTax.FOneItem.FrepEmail		= ogroup.FOneItem.Fjungsan_email
				end if			
			end if

			oTax.FOneItem.FconfirmYn	= "Y"
			oTax.FOneItem.FtotalPrice	= oetcmeachul.FOneItem.Ftotalsum
			oTax.FOneItem.FtotalTax		= 0
			oTax.FOneItem.Fitemname		= oetcmeachul.FOneItem.Ftitle
			oTax.FOneItem.ForderIdx		= etcmeachulIdx
		end if

		'// --------------------------------------------------------------------
		'' 삽아이디에서 3PL 업체인지 확인
		if (Is3PLShopid(oetcmeachul.FOneItem.Fshopid) = True) then
			Call Get3PLUpcheInfoByShopid(oetcmeachul.FOneItem.Fshopid, tplcompanyid, tplcompanyname, tplgroupid, tplbillUserID)

			dim otplgroup
			set otplgroup = new CPartnerGroup

			otplgroup.FRectGroupid = tplgroupid
			otplgroup.GetOneGroupInfo

			if (otplgroup.FResultCount < 1) then
				errMSG = "3PL그룹코드가 지정되어 있지 않은 업체입니다."
			else
				'// 공급자(3PL)
				oTax.FOneItem.FsupplyBusiNo			= otplgroup.FOneItem.Fcompany_no
				oTax.FOneItem.FsupplyBusiName		= otplgroup.FOneItem.FCompany_name
				oTax.FOneItem.FsupplyBusiCEOName	= otplgroup.FOneItem.Fceoname
				oTax.FOneItem.FsupplyBusiAddr		= otplgroup.FOneItem.Fcompany_address & " " & otplgroup.FOneItem.Fcompany_address2
				oTax.FOneItem.FsupplyBusiType		= otplgroup.FOneItem.Fcompany_uptae
				oTax.FOneItem.FsupplyBusiItem		= otplgroup.FOneItem.Fcompany_upjong
				oTax.FOneItem.FsupplyRepName		= otplgroup.FOneItem.Fjungsan_name
				oTax.FOneItem.FsupplyRepTel			= otplgroup.FOneItem.Fjungsan_hp
				oTax.FOneItem.FsupplyRepEmail		= otplgroup.FOneItem.Fjungsan_email

				oTax.FOneItem.Fbilldiv = "99"
				oTax.FOneItem.FsupplyConfirmYn = "Y"
			end if
		else
			'// 공급자(텐바이텐)
			oTax.FOneItem.FsupplyBusiNo			= TENBYTEN_SOCNO
			oTax.FOneItem.FsupplyBusiName		= TENBYTEN_SOCNAME
			oTax.FOneItem.FsupplyBusiCEOName	= TENBYTEN_CEONAME
			oTax.FOneItem.FsupplyBusiAddr		= TENBYTEN_SOCADDR
			oTax.FOneItem.FsupplyBusiType		= TENBYTEN_SOCSTATUS
			oTax.FOneItem.FsupplyBusiItem		= TENBYTEN_SOCEVENT
			oTax.FOneItem.FsupplyRepName		= TENBYTEN_MANAGERNAME
			oTax.FOneItem.FsupplyRepTel			= TENBYTEN_MANAGERPHONE
			oTax.FOneItem.FsupplyRepEmail		= TENBYTEN_MANAGERMAIL

			oTax.FOneItem.FsupplyConfirmYn = "Y"
		end if

		''기발행 세금계산서인지 체크
		if errMSG = "" and (oTax.FOneItem.ForderIdx <> "") then
			set oTaxCheck = new CTax

			oTaxCheck.FRectsearchKey = " t.orderidx "
			oTaxCheck.FRectsearchString = CStr(oTax.FOneItem.ForderIdx)
			oTaxCheck.FRectDelYn = "N"

			oTaxCheck.GetTaxList

			if oTaxCheck.FResultCount > 0 then
				if oTaxCheck.FTaxList(0).FisueYn="Y" then
					if (errMSG = "") then
						errMSG = "이미 발행된 세금계산서가 있습니다.\n\n재발행 하시려면 거래처에 [취소요청]후 매출 세금계산서 목록에서 [삭제]후 발행하셔야 합니다"
					end if
				else
					if (errMSG = "") then
						errMSG = "발행대기중인 세금계산서가 있습니다.\n\n재발행 하시려면 매출 세금계산서 목록에서 [삭제]후 발행하셔야 합니다"
					end if
				end if
			end if
		end if
	end if

	if (oTax.FOneItem.FTaxType = "") then
		oTax.FOneItem.FTaxType = "Y"
	end if

	if (oTax.FOneItem.FsupplyBusiNo = "") then
		''oTax.FOneItem.FsupplyBusiNo			= TENBYTEN_SOCNO
		''oTax.FOneItem.FsupplyBusiName		= TENBYTEN_SOCNAME
		''oTax.FOneItem.FsupplyBusiCEOName	= TENBYTEN_CEONAME
		''oTax.FOneItem.FsupplyBusiAddr		= TENBYTEN_SOCADDR
		''oTax.FOneItem.FsupplyBusiType		= TENBYTEN_SOCSTATUS
		''oTax.FOneItem.FsupplyBusiItem		= TENBYTEN_SOCEVENT
		''oTax.FOneItem.FsupplyRepName		= TENBYTEN_MANAGERNAME
		''oTax.FOneItem.FsupplyRepTel			= TENBYTEN_MANAGERPHONE
		''oTax.FOneItem.FsupplyRepEmail		= TENBYTEN_MANAGERMAIL

		oTax.FOneItem.FsupplyConfirmYn = "Y"
	end if
else
	mode = "view"
	oTax.GetTaxRead

	'response.write oTax.FOneItem.Fbilldiv
	'response.write oTax.FOneItem.Ftplcompanyid

	if (oTax.FOneItem.Fbilldiv = "99") then
		if Not IsNull(oTax.FOneItem.Ftplcompanyid) then
			Call Get3PLUpcheInfo(oTax.FOneItem.Ftplcompanyid, tplcompanyname, tplgroupid, tplbillUserID)
		end if
	end if
end if

if (errMSG <> "") then
	IsIssusOK = False
end if

'수익부서목록
Dim clsBS, arrBizList
Set clsBS = new CBizSection
	clsBS.FUSE_YN = "Y"
	clsBS.FOnlySub = "Y"
	clsBS.FSale = "N"
	arrBizList = clsBS.fnGetBizSectionList
Set clsBS = nothing
%>
<script type='text/javascript'>

var errMSG = "<%= errMSG %>";

function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

function setRegisterInfo() {
	var TENBYTEN_SOCNO = "<%= TENBYTEN_SOCNO %>";
	var TENBYTEN_SOCNAME = "<%= TENBYTEN_SOCNAME %>";
	var TENBYTEN_CEONAME = "<%= TENBYTEN_CEONAME %>";
	var TENBYTEN_SOCADDR = "<%= TENBYTEN_SOCADDR %>";
	var TENBYTEN_SOCSTATUS = "<%= TENBYTEN_SOCSTATUS %>";
	var TENBYTEN_SOCEVENT = "<%= TENBYTEN_SOCEVENT %>";
	var TENBYTEN_MANAGERNAME = "<%= TENBYTEN_MANAGERNAME %>";
	var TENBYTEN_MANAGERPHONE = "<%= TENBYTEN_MANAGERPHONE %>";
	var TENBYTEN_MANAGERMAIL = "<%= TENBYTEN_MANAGERMAIL %>";

	var SUPPLY_SOCNO = "<%= oTax.FOneItem.FsupplybusiNo %>";
	var SUPPLY_SOCNAME = "<%= oTax.FOneItem.FsupplybusiName %>";
	var SUPPLY_CEONAME = "<%= oTax.FOneItem.FsupplybusiCEOName %>";
	var SUPPLY_SOCADDR = "<%= oTax.FOneItem.FsupplybusiAddr %>";
	var SUPPLY_SOCSTATUS = "<%= oTax.FOneItem.FsupplybusiType %>";
	var SUPPLY_SOCEVENT = "<%= oTax.FOneItem.FsupplybusiItem %>";
	var SUPPLY_MANAGERNAME = "<%= oTax.FOneItem.FsupplyrepName %>";
	var SUPPLY_MANAGERPHONE = "<%= oTax.FOneItem.FsupplyrepTel %>";
	var SUPPLY_MANAGERMAIL = "<%= oTax.FOneItem.FsupplyrepEmail %>";

	// 01, 02, 03, 51, 11, 99
	if ((frm.billdiv.value == "01") || (frm.billdiv.value == "02") || (frm.billdiv.value == "03") || (frm.billdiv.value == "51")) {
		setSupplyCompanyInfo(TENBYTEN_SOCNO, "", TENBYTEN_SOCNAME, TENBYTEN_CEONAME, TENBYTEN_SOCADDR, TENBYTEN_SOCSTATUS, TENBYTEN_SOCEVENT, TENBYTEN_MANAGERNAME, TENBYTEN_MANAGERPHONE, TENBYTEN_MANAGERMAIL);
	} else if (frm.billdiv.value == "11") {
		setSupplyCompanyInfo(SUPPLY_SOCNO, "", SUPPLY_SOCNAME, SUPPLY_CEONAME, SUPPLY_SOCADDR, SUPPLY_SOCSTATUS, SUPPLY_SOCEVENT, SUPPLY_MANAGERNAME, SUPPLY_MANAGERPHONE, SUPPLY_MANAGERMAIL);
	}
}

function chgHandTax(comp){
    var txbox = comp.form.totaltaxprice;

    if (comp.checked){
        txbox.readOnly = false;
        txbox.className = "writebox";
    }else{
        txbox.readOnly = true;
        txbox.className = "readonlybox";
    }
}

function CalcPrice(comp) {
	var frm = document.frm;

	if (frm.taxtype.value.length < 1) {
		alert('과세구분을 입력하세요.');
		return;
	}

	if ((comp.name == "taxtype") || (comp.name == "totalprice")) {
		if ((frm.totalprice.value == "") || (frm.totalprice.value*0 != 0)) {
			// alert("먼저 합계금액을 입력하세요.");
			return;
		}

		if (frm.taxtype.value == "Y") {
			// 세액은 공급가를 구하고 0.1 후 반올림 해주면 된다.
			frm.totaltaxprice.value = Math.round(1.0 * frm.totalprice.value / 1.1 / 10.0);
			frm.totaltaxprice2.value = frm.totaltaxprice.value;
		} else {
			frm.totaltaxprice.value = 0;
		}
		frm.totaltaxprice2.value = frm.totaltaxprice.value;

		frm.totalsupplyprice.value = frm.totalprice.value*1 - frm.totaltaxprice2.value*1;
		frm.totalsupplyprice2.value = frm.totalsupplyprice.value;
		frm.totalsupplyprice3.value = frm.totalsupplyprice.value;
	} else if (comp.name == "totaltaxprice") {
		if ((frm.totaltaxprice.value == "") || (frm.totaltaxprice.value*0 != 0)) {
			// alert("먼저 세액을 입력입력하세요.");
			return;
		}

		frm.totaltaxprice2.value = frm.totaltaxprice.value;

		frm.totalsupplyprice.value = frm.totalprice.value*1 - frm.totaltaxprice2.value*1;
		frm.totalsupplyprice2.value = frm.totalsupplyprice.value;
		frm.totalsupplyprice3.value = frm.totalsupplyprice.value;
	} else if (comp.name == "totalsupplyprice2") {
		if ((frm.totalsupplyprice2.value == "") || (frm.totalsupplyprice2.value*0 != 0)) {
			// alert("먼저 세액을 입력입력하세요.");
			return;
		}

		frm.totalsupplyprice.value = frm.totalsupplyprice2.value;
		frm.totalsupplyprice3.value = frm.totalsupplyprice2.value;

		if (frm.taxtype.value == "Y") {
			// 세액은 공급가를 구하고 0.1 후 반올림 해주면 된다.
			frm.totaltaxprice.value = parseInt(frm.totalsupplyprice2.value*0.1);
			frm.totaltaxprice2.value = frm.totaltaxprice.value;
		} else {
			frm.totaltaxprice.value = 0;
		}

		frm.totalprice.value = frm.totalsupplyprice.value*1 + frm.totaltaxprice.value*1;
	}

	frm.totalprice2.value = frm.totalprice.value;
	frm.totalprice3.value = frm.totalprice.value;
}

function doRegisterSheet(){

	if (errMSG != "") {
		alert(errMSG);
		return;
	}

	if (frm.issuetype.value != "") {
		if ((frm.issuetype.value == "orderserial") && (frm.billdiv.value != "01") && (frm.billdiv.value != "11")) {
			alert('소비자 매출만 작성할 수 있습니다.');
			frm.billdiv.focus();
			return;
		}

		if ((frm.issuetype.value == "etcmeachul") && (frm.billdiv.value != "51") && (frm.billdiv.value != "99")) {
			alert('기타매출만 작성할 수 있습니다.');
			frm.billdiv.focus();
			return;
		}

		if(frm.orderidx.value == "") {
			alert('주문번호 또는 기타매출 코드가 비고에 입력되어 있어야 합니다.');
			return;
		}
	}

	// 20036 => 4010005
	if ((frm.selltype.value == "4010005") && (frm.taxtype.value != "0")) {
		alert('계정과목이 영세인 경우 영세계산서만 작성가능합니다.');
		return;
	} else if ((frm.selltype.value != "4010005") && (frm.taxtype.value == "0")) {
		alert('계정과목이 영세가 아니면 영세계산서를 작성할 수 없습니다.');
		return;
	}

	if(frm.billdiv.value == "0") {
		alert('공급자를 선택하세요.');
		return;
	}

	if (frm.socname.value.length<1){
		alert('사업자 등록상의 회사명을 입력하세요.');
		frm.socname.focus();
		return;
	}

	if (frm.ceoname.value.length<1){
		alert('사업자 등록상의 대표자명을 입력하세요.');
		frm.ceoname.focus();
		return;
	}

	if (frm.socno.value.length<1){
		alert('사업자 등록 번호를 입력하세요.');
		frm.socno.focus();
		return;
	}

	if (frm.socaddr.value.length<1){
		alert('사업자 등록상의 주소를 입력하세요.');
		frm.socaddr.focus();
		return;
	}

	if (frm.socstatus.value.length<1){
		alert('사업자 등록상의 업태를 입력하세요.');
		frm.socstatus.focus();
		return;
	}

	if (frm.socevent.value.length<1){
		alert('사업자 등록상의 업종을 입력하세요.');
		frm.socevent.focus();
		return;
	}

	if (frm.managername.value.length<1){
		alert('담당자 성함을 입력하세요.');
		frm.managername.focus();
		return;
	}

	if (frm.managerphone.value.length<1){
		alert('담당자 전화번호를 입력하세요.');
		frm.managerphone.focus();
		return;
	}

	if (frm.managerphone.value.indexOf("-") == -1) {
		alert('담당자 연락처는 000-000-0000 형식이어야 합니다.');
		frm.managerphone.focus();
		return;
	}

	if (frm.managermail.value.length<1){
		alert('담당자 이메일주소를 입력하세요.');
		frm.managermail.focus();
		return;
	}

	if (frm.yyyymmdd.value.length<1){
		alert('작성일을 입력하세요.');
		return;
	}

	if (frm.itemname.value.length<1){
		alert('품목을 입력하세요.');
		return;
	}

	if (frm.totalprice.value.length<1){
		alert('합계금액을 입력하세요.');
		return;
	}

	if (frm.taxtype.value.length<1){
		alert('과세구분을 입력하세요.');
		return;
	}

	if ((frm.subsocno.value.length != 0) && (frm.subsocno.value.length != 4)) {
		alert('종사업장번호를 4자리로 입력하세요');
		return;
	}

	if ((frm.billdiv.value == "01") || (frm.billdiv.value == "11")) {
		if(frm.orderidx.value == "") {
			alert('비고에 주문번호를 입력하세요.');
			return;
		}
	} else if ((frm.orderidx.value != "") && (frm.billdiv.value != "03") && (frm.billdiv.value != "51") && (frm.billdiv.value != "99")) {
		alert('소비자매출/프로모션/기타매출에만 비고에 주문번호 또는 출고코드를 넣을 수 있습니다.');
		return;
	}

	if (frm.billdiv.value != "99") {
		if (frm.sellBizCd.value.length<1){
			alert('매출부서를 지정하세요.');
			return;
		}
	} else {
		if (frm.sellBizCd.value.length > 0){
			alert('3PL매출에는 부서를 지정할 수 없습니다.');
			return;
		}
	}

	if (frm.selltype.value.length<1){
		alert('매출계정을 지정하세요.');
		return;
	}

	if (frm.taxissuetype.value.length<1){
		alert('세부내역을 지정하세요.');
		return;
	}

	<% if (groupID <> "") then %>
	if (frm.billdiv.value == "11") {
		// 소비자(업체매출)
		if ((frm.consignYN.value == "N") && (frm.reg_socno.value != "211-87-00620")) {
			alert("공급자가 텐바이텐이 아닌 경우 위수탁구분 설정하세요.");
			return;
		}
		if ((frm.consignYN.value == "Y") && (frm.reg_socno.value == "211-87-00620")) {
			alert("공급자가 텐바이텐인 경우 위수탁구분을 [정상] 으로 설정하세요.");
			return;
		}
		if ((frm.consignYN.value == "Y") && (frm.issueMethod.value != "eSero")) {
			alert("위수탁 계산서는 이세로에서 수기발행만 발행가능합니다.");
			return;
		}
	}
	<% end if %>

	setRegisterInfo();

    if (confirm('세금계산서 발행신청을 하시겠습니까?')){
		document.frm.mode.value = "tax_register_new";

		if (frm.billdiv.value == "11") {
			<% if (groupID <> "") then %>
			// 2014년 이후 계산서 발행(업체별)
			document.frm.mode.value = "tax_register_new_2014_upche";
			<% else %>
			// 2014년 이후 요청(업체별)
			document.frm.mode.value = "tax_register_new_2014";
			<% end if %>
		}

        document.frm.submit();
    }
}

function popMeachulDetailList() {
	if (frm.taxissuetype.value != "E") {
		alert("기타매출인 경우에만 내역을 추가할 수 있습니다.");
		return;
	}

	var popwin = window.open('pop_etc_meachul_list.asp?idx=<%= oTax.FOneItem.ForderIdx %>','popMeachulDetailList','width=1000, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popListPreviousCustomerTaxSheet(userid){
    var popwin=window.open("/cscenter/taxsheet/popListPreviousCustomerTaxSheet.asp?userid=" + userid,"popListPreviousCustomerTaxSheet","width=700,height=400,scrollbars=yes,resizable=yes");
    popwin.focus();
}

function ReactMeachulDetailList(arrchk, tottaxsum) {
    var frm = document.frm;

    frm.totalprice.value = tottaxsum;
    frm.orderidx.value = arrchk;

    CalcPrice(frm.totalprice);
}

function SearchSocno() {
	if (frm.socno.value == "") {
		alert("사업자번호를 입력하세요.");
		return;
	}

	if (frm.socno.value.length != 12) {
		alert("사업자번호는 아래와 같은 형식으로 입력하세요.\n\n000-00-00000");
		return;
	}

	icheckframe.location.href="isearchframe.asp?socno=" + frm.socno.value;
}

function setCompanyInfo(subsocno, socname, ceoname, socaddr, socstatus, socevent, managername, managerphone, managermail)
{
	frm.subsocno.value = subsocno;
	frm.socname.value = socname;
	frm.ceoname.value = ceoname;
	frm.socaddr.value = socaddr;
	frm.socstatus.value = socstatus;
	frm.socevent.value = socevent;
	frm.managername.value = managername;
	frm.managerphone.value = managerphone;
	frm.managermail.value = managermail;
}

function setSupplyCompanyInfo(socno, subsocno, socname, ceoname, socaddr, socstatus, socevent, managername, managerphone, managermail)
{
	frm.reg_socno.value = socno;
	frm.reg_subsocno.value = subsocno;
	frm.reg_socname.value = socname;
	frm.reg_ceoname.value = ceoname;
	frm.reg_socaddr.value = socaddr;
	frm.reg_socstatus.value = socstatus;
	frm.reg_socevent.value = socevent;
	frm.reg_managername.value = managername;
	frm.reg_managerphone.value = managerphone;
	frm.reg_managermail.value = managermail;
}

// 요청서 수정
function GotoTaxModify(){
	if (frm.managername.value.length<1){
		alert('담당자 성함을 입력하세요.');
		frm.managername.focus();
		return;
	}

	if (frm.managerphone.value.length<1){
		alert('담당자 전화번호를 입력하세요.');
		frm.managerphone.focus();
		return;
	}

	if (frm.managermail.value.length<1){
		alert('담당자 이메일주소를 입력하세요.');
		frm.managermail.focus();
		return;
	}

	if (frm.itemname.value.length<1){
		alert('품목을 입력하세요.');
		return;
	}

	if ((frm.billdiv.value == "52") || (frm.billdiv.value == "55")) {
		alert('발행불가');
		return;
	}

	if (frm.totalsupplyprice2.value.length<1){
		alert('단가를 입력하세요.');
		return;
	}

	if (frm.totalprice.value.length<1){
		alert('합계금액을 입력하세요.');
		return;
	}

	if (frm.totalprice.value*1 != frm.orgtotalprice.value*1) {
		alert('금액을 수정할 수 없습니다.\n\n삭제 후 재작성하세요.');
		return;
	}

	if (frm.sellBizCd.value.length<1){
		alert('매출부서를 지정하세요.');
		return;
	}

	if (frm.selltype.value.length<1){
		alert('매출계정을 지정하세요.');
		return;
	}

	if (frm.taxissuetype.value.length<1){
		alert('세부내역을 지정하세요.');
		return;
	}

	if (frm.taxtype.value.length<1){
		alert('과세구분을 입력하세요.');
		return;
	}

	if (frm.managerphone.value.indexOf("-") == -1) {
		alert('담당자 연락처는 000-000-0000 형식이어야 합니다.');
		frm.managerphone.focus();
		return;
	}

	if (frm.billdiv.value == "11") {
		// 소비자(업체매출)
		if ((frm.consignYN.value == "N") && (frm.reg_socno.value != "211-87-00620")) {
			alert("공급자가 텐바이텐이 아닌 경우 위수탁구분 설정하세요.");
			return;
		}
		if ((frm.consignYN.value == "Y") && (frm.reg_socno.value == "211-87-00620")) {
			alert("공급자가 텐바이텐인 경우 위수탁구분을 [정상] 으로 설정하세요.");
			return;
		}
		if ((frm.consignYN.value == "Y") && (frm.issueMethod.value != "eSero")) {
			alert("위수탁 계산서는 이세로에서 수기발행만 발행가능합니다.");
			return;
		}
	}

	setRegisterInfo();

	if (confirm('요청서를 수정 하시겠습니까?')){
		document.frm.mode.value="tax_modify";
		document.frm.submit();
	}
}

//function PopCommonSampleTaxReg(){
//	var popCommonSamplewin = window.open("/cscenter/taxsheet/popCommonSampleWehagotaxregapi.asp?taxIdx=<%'=taxIdx %>&taxType=<%'= oTax.FOneItem.Fbilldiv %>","popCommonSampletaxreg","width=1200 height=768 scrollbars=yes resizable=yes");
//	popCommonSamplewin.focus();
//}

function PopCommonWehagoTaxReg(){
	<% if (IsIssusOK = False) then %>
		<% if (session("ssbctID") = "icommang") or (session("ssbctID") = "coolhas") or (session("ssbctID") = "tozzinet") then %>
			if (!confirm('<%= ErrMSG %>\n\n계속하시겠습니까?(관리자권한)')) return;
		<% else %>
			alert('<%= ErrMSG %>');
			return;
		<% end if %>
	<% end if %>

	<% if ((session("ssAdminLsn") = "1") or (session("ssAdminLsn") = "2") or (session("ssAdminPsn") = "8") or (session("ssAdminPsn") = "17") or (session("ssbctID") = "josin222")) then %>
		// 마스터이상 권한 또는 경영관리팀, 관계사(아이띵소)
		if (confirm('세금계산서를 발행하시겠습니까?')){
			var popCommonWehagowin = window.open('/cscenter/taxsheet/popCommonWehagotaxregapi.asp?taxIdx=<%=taxIdx %>&taxType=<%= oTax.FOneItem.Fbilldiv %>','popCommonWehagotaxreg','width=1200,height=768,scrollbars=yes,resizable=yes');
			popCommonWehagowin.focus()
		}
	<% else %>
		alert('권한이 없습니다.[2]');
	<% end if %>
}

/*
function TaxEvalBill36524api(){
	<% if (IsIssusOK = False) then %>
		<% if (session("ssbctID") = "icommang") or (session("ssbctID") = "coolhas") or (session("ssbctID") = "tozzinet") then %>
			if (!confirm('<%= ErrMSG %>\n\n계속하시겠습니까?(관리자권한)')) return;
		<% else %>
			alert('<%= ErrMSG %>');
			return;
		<% end if %>
	<% end if %>

	<% if ((session("ssAdminLsn") = "1") or (session("ssAdminLsn") = "2") or (session("ssAdminPsn") = "8") or (session("ssAdminPsn") = "17") or (session("ssbctID") = "josin222")) then %>
		// 마스터이상 권한 또는 경영관리팀, 관계사(아이띵소)
		if (confirm('세금계산서를 발행하시겠습니까?')){
			var popwin = window.open('/cscenter/taxsheet/evalTaxBill36524api.asp?taxIdx=<%=taxIdx %>&taxType=<%= oTax.FOneItem.Fbilldiv %>','evalTaxBill36524api','width=1024,height=768,scrollbars=yes,resizable=yes');
			popwin.focus()
		}
	<% else %>
		alert('권한이 없습니다.[2]');
	<% end if %>
}

    function TaxEvalBill36524(){
    	<% if (IsIssusOK = False) then %>
    	    <% if (session("ssbctID") = "icommang") or (session("ssbctID") = "coolhas") or (session("ssbctID") = "tozzinet") then %>
    	    if (!confirm('<%= ErrMSG %>\n\n계속하시겠습니까?(관리자권한)')) return;
    	    <% else %>
    		alert('<%= ErrMSG %>');
    		return;
    		<% end if %>
    	<% end if %>

<% if ((session("ssAdminLsn") = "1") or (session("ssAdminLsn") = "2") or (session("ssAdminPsn") = "8") or (session("ssAdminPsn") = "17") or (session("ssbctID") = "josin222")) then %>
		// 마스터이상 권한 또는 경영관리팀, 관계사(아이띵소)
        if (confirm('세금계산서를 발행하시겠습니까?')){
            var popwin = window.open('evalTaxBill36524.asp?taxIdx=<%=taxIdx %>&taxType=<%= oTax.FOneItem.Fbilldiv %>','evalTaxBill36524','width=400,height=300,scrollbars=yes,resizable=yes');
            popwin.focus()
        }
<% else %>
		alert('권한이 없습니다.[2]');
<% end if %>
    }
*/

    function TaxEvaleSero() {
    	<% if (IsIssusOK = False) then %>
    	    <% if (session("ssbctID") = "icommang") or (session("ssbctID") = "coolhas") or (session("ssbctID") = "tozzinet") then %>
    	    if (!confirm('<%= ErrMSG %>\n\n계속하시겠습니까?(관리자권한)')) return;
    	    <% else %>
    		alert('<%= ErrMSG %>');
    		return;
    		<% end if %>
    	<% end if %>

<% if ((session("ssAdminLsn") = "1") or (session("ssAdminLsn") = "2") or (session("ssAdminPsn") = "8") or (session("ssAdminPsn") = "17") or (session("ssbctID") = "josin222")) then %>
		// 마스터이상 권한 또는 경영관리팀, 관계사(아이띵소)
        if (confirm('수기발행 완료처리 하시겠습니까?')){
			document.frm.mode.value="finishSheetByESero";
			document.frm.submit();
        }
<% else %>
		alert('권한이 없습니다.[2]');
<% end if %>
    }

	// 요청서 삭제
	function GotoTaxDel(isIssued){
		if (confirm('요청서를 삭제 하시겠습니까?\n\n계산서가 발행된 경우 발행이 취소된후 삭제하시기 바랍니다.')){
		    if (isIssued == "Y") {
    		    <% if C_ADMIN_AUTH or C_MngPowerUser then %>
    		    alert('관리자 권한 실행. 반드시 수정세금계산서 확인.');
    			document.frm.mode.value="sheetDel";
    			document.frm.submit();
    			<% else %>
    			alert('권한이 없습니다. 관리자 문의 요망[1]');
    			<% end if %>
    		}else{
    		    document.frm.mode.value="sheetDel";
    			document.frm.submit();
    		}
		}
	}

	// 세금계산서 보기
	function goView(tax_no, b_biz_no, s_biz_no)
	{
		<% if (application("Svr_Info")="Dev") then %>
			// 테스트
			window.open("http://ifs.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no="+tax_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
		<% else %>
			// 실서버
			window.open("http://web1.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no=" + tax_no + "&cur_biz_no="+b_biz_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
		<% end if %>
	}

	function goView2(tax_no, b_biz_no, s_biz_no){
		<% if (application("Svr_Info")="Dev") then %>
			// 테스트
			window.open("http://ifs.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no="+tax_no+"&cur_biz_no="+s_biz_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
		<% else %>
			// 실서버
			window.open("http://web1.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no="+tax_no+"&cur_biz_no="+s_biz_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
		<% end if %>
	}

	function goView_Bill36524(tax_no, b_biz_no)
	{
			window.open("http://www.bill36524.com/popupBillTax.jsp?NO_TAX=" + tax_no + "&NO_BIZ_NO="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
	}

    function GotoTaxMapHand(){
        var popwin = window.open('popTaxMapHand.asp?taxIdx=<%=taxIdx%>','popTaxMapHand','scrollbars=yes,resizable=yes,width=400,height=300');
        popwin.focus();
    }

	// 사업자검색
	function PopUpcheSelectBySocno(frmname){
		var socno = eval(frmname+'.socno').value;
		if (socno==''){
			alert('사업자 등록번호를 입력하세요');
			eval(frmname+'.socno').focus();
			return;
		}

		var popwin = window.open("/admin/member/popupcheselect.asp?mode=tax&frmname=" + frmname + "&rectsocno=" + socno,"popupcheselectbysocno","width=1280 height=960 scrollbars=yes resizable=yes");
		popwin.focus();
	}

	function jsGetCust(frmname){
		var socno = eval(frmname+'.socno').value;
		if (socno==''){
			alert('사업자 등록번호를 입력하세요');
			eval(frmname+'.socno').focus();
			return;
		}

		var Strparm = "";
		Strparm = "?selSTp=5&sSTx="+ socno;
		Strparm = Strparm + "&opnType=eTaxdetail";
		var winC = window.open("/admin/linkedERP/cust/popGetCust.asp"+Strparm,"popC","width=1280, height=960,resizable=yes, scrollbars=yes");
		winC.focus();
	}

	//거래처 선택
	function jsSetCust(custcd, custnm, ceonm, custno, addr, bscd, intp, email, telno){
		frm.socname.value = custnm;		// 상호
		frm.ceoname.value = ceonm;		// 대표자
		//frm.socno.value = custno;		// 사업자번호
		frm.socaddr.value = addr;		// 사업장주소
		frm.socstatus.value = bscd;		// 업태
		frm.socevent.value = intp;		// 종목
		frm.managerphone.value = telno;		// 연락처
		frm.managermail.value = email;		// 이메일
	}

</script>

<style type="text/css">
.Readonlybox { border:0px; }
.writebox { border:10px; background:#E6E6E6; }
</style>

<table width="800" border="0" class="a" cellpadding="0" cellspacing="0">
<tr>
	<td>

		<!-- 세금계산 요청서 정보 시작 -->
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
			<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td colspan="4" align="left">
					<b>세금계산서 발행정보</b>
				</td>
			</tr>
<% if (oTax.FOneItem.Fbilldiv = "01") or (oTax.FOneItem.Fbilldiv = "11") then %>
			<tr height="25">
				<td align="center" width="120" bgcolor="#F0F0FD">요청자</td>
				<td bgcolor="#FFFFFF" colspan="3"><%= oTax.FOneItem.Fuserid %></td>
			</tr>
			<tr height="25">
				<td align="center" bgcolor="#F0F0FD">입금확인일</td>
				<td bgcolor="#FFFFFF" width="120">
					<% if IsNULL(oTax.FOneItem.Fipkumdate) then %>

				 	<% else %>
						<% if oTax.FOneItem.Fipkumdate <> "" then %>
							<%=FormatDate(oTax.FOneItem.Fipkumdate,"0000-00-00")%>
						<% end if %>
			    	<% end if %>
				</td>
				<td align="center" bgcolor="#F0F0FD" width="120">등록일</td>
				<td bgcolor="#FFFFFF"><%= oTax.FOneItem.Fregdate %></td>
			</tr>
<% end if %>
<%	if oTax.FOneItem.FisueYn="Y" then %>
			<tr>
				<td align="center" bgcolor="#F0F0FD">발급 여부</td>
				<td bgcolor="#F8F8FF" colspan="3">
					<font color=darkblue>발급</font>
					&nbsp;
					<% if (Left(oTax.FOneItem.FneoTaxNo,2)="TX") or (Left(oTax.FOneItem.FneoTaxNo,2)="FX") then %>
					<input type="button" class="button" value="공급받는자 보관용" onClick="goView_Bill36524('<%=oTax.FOneItem.FneoTaxNo%>', '<%=Replace(oTax.FOneItem.FbusiNo,"-", "")%>')" style="cursor:pointer" align="absmiddle">
					<% else %>
					<input type="button" class="button" value="공급받는자 보관용" onClick="goView(<%=oTax.FOneItem.FneoTaxNo%>, '<%=Replace(oTax.FOneItem.FbusiNo,"-", "")%>', '2118700620')" style="cursor:pointer" align="absmiddle">
					<% end if %>
					<% if (Left(oTax.FOneItem.FneoTaxNo,2)="TX") or (Left(oTax.FOneItem.FneoTaxNo,2)="FX") then %>
					<input type="button" class="button" value="공급자 보관용" onClick="goView_Bill36524('<%=oTax.FOneItem.FneoTaxNo%>', '2118700620')" style="cursor:pointer" align="absmiddle">

					<% else %>
					<input type="button" class="button" value="공급자 보관용" onClick="goView2(<%=oTax.FOneItem.FneoTaxNo%>, '<%=Replace(oTax.FOneItem.FbusiNo,"-", "")%>', '2118700620')" style="cursor:pointer" align="absmiddle">
				    <% end if %>
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="#F0F0FD">발급자 아이디</td>
				<td bgcolor="#FFFFFF" colspan="3"><%=oTax.FOneItem.FcurUserId%></td>
			</tr>
			<tr>
				<td align="center" bgcolor="#F0F0FD">계산서 작성일자</td>
				<td bgcolor="#FFFFFF" colspan="3"><b><%=oTax.FOneItem.FisueDate%></b></td>
			</tr>
			<tr>
				<td align="center" bgcolor="#F0F0FD">발급일시</td>
				<td bgcolor="#FFFFFF" colspan="3"><%=oTax.FOneItem.Fprintdate%></td>
			</tr>
<%	else %>
	<% if (orderserial <> "") and (ordercancelyn <> "") then %>
			<tr>
				<td align="center" bgcolor="#F0F0FD">주문번호</td>
				<td bgcolor="#FFFFFF" colspan="3">
					<%= orderserial %>
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="#F0F0FD">주문상태</td>
				<td bgcolor="#FFFFFF" colspan="3">
				    <%= iIpkumDivName %>
				    &nbsp;/&nbsp;
					<% if (ordercancelyn <> "N") then %>
						<font color=red>취소</font>
					<% else %>
						정상
					<% end if %>
					&nbsp;
					<% IF (ordercancelyn<>"N") then %>
					<strong>[취소된 주문건은 발행 불가 합니다.]</strong>
					<% elseIF (IpkumDiv<8) then %>
					<strong>[출고전 부분취소등으로 금액 변동이 생길 수 있으니 가능한 상품출고 이후에 발행 하세요.]</strong>
					<% end if %>
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="#F0F0FD">결제총액</td>
				<td bgcolor="#FFFFFF" colspan="3">
					<%= FormatNumber(subtotalprice, 0) %>
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="#F0F0FD">실결제액</td>
				<td bgcolor="#FFFFFF" colspan="3">
					<%= FormatNumber((subtotalprice - sumPaymentEtc), 0) %>
					&nbsp;
					<% if (accountdiv = "400") then %>(휴대폰결제)<% end if %>
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="#F0F0FD">보조결제</td>
				<td bgcolor="#FFFFFF" colspan="3">
					<%= FormatNumber(sumPaymentEtc, 0) %>
				</td>
			</tr>
	<% end if %>
	<% if (mode <> "new") then %>
			<tr>
				<td align="center" bgcolor="#F0F0FD">발급 여부</td>
				<td bgcolor="#FFFFFF" colspan="3">
					<font color=darkred>미발급</font>
					&nbsp;
					<% if (mode = "view") and (oTax.FOneItem.Fbilldiv = "11") and (oTax.FOneItem.FissueMethod = "eSero") then %>
					<input type="button" class="button" value="수기발급완료(eSero)" onClick="TaxEvaleSero()">
					<% else %>
						<% '<input type="button" class="button" value="발행(Bill36524 플래시)" onClick="TaxEvalBill36524()"> %>
						<% '<input type="button" class="button" value="발행(Bill36524 API)" onClick="TaxEvalBill36524api()"> %>
						<input type="button" class="button" value="발행(위하고)" onClick="PopCommonWehagoTaxReg()">
        				<% '<input type="button" class="button" value="샘플" onclick="PopCommonSampleTaxReg(); return false;"> %>
					<% end if %>
					&nbsp;
				</td>
			</tr>
	<%	end if %>
<% end if %>
		</table>
	</td>
</tr>
<tr height="20">
	<td>
	</td>
</tr>
<tr height="20">
	<td>
		<br>*<font color=red>금액이 상이</font>할 경우 금액을 수정할 수 없고, 삭제후 재작성해야 합니다.
		<br>*연락처의 경우 xx-xxx-xxxx(정상) xx.xxx.xxxx(오류) 로 입력해 주세요
	</td>
</tr>
<tr>
	<td>
		<br>
    	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="POST" action="doTaxOrder.asp" style="margin:0px;">
			<input type="hidden" name="taxIdx" value="<%= taxIdx %>">
			<input type=hidden name=issuetype value="<%= issuetype %>">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="tplcompanyid" value="<%= tplcompanyid %>">
			<input type="hidden" name="mode" value="">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    			<td height="25" width="15%"><b>부서</b></td>
    			<td align="left" bgcolor="#FFFFFF" width="35%">
					<select class="select" name="sellBizCd">
					<option value="">--선택--</option>
					<% For i = 0 To UBound(arrBizList,2)	%>
						<option value="<%=arrBizList(0,i)%>" <%IF Cstr(oTax.FOneItem.FsellBizCd) = Cstr(arrBizList(0,i)) THEN%> selected <%END IF%>><%=arrBizList(1,i)%></option>
					<% Next %>
					</select>
    				<%'= fndrawSaleBizSecCombo(true,"sellBizCd", oTax.FOneItem.FsellBizCd,"") %>
    			</td>
    			<td height="25" width="15%"><b>계정</b></td>
    			<td align="left" bgcolor="#FFFFFF">
    				<% drawPartnerCommCodeBox true,"sellacccd","selltype", oTax.FOneItem.Fselltype,"" %>
    			</td>
    		</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    			<td height="25"><b>세부내역</b></td>
    			<td align="left" bgcolor="#FFFFFF">
    				<select class="select" name="taxissuetype">
    					<option value="">-선택-</option>
    					<option value="C" <% if (oTax.FOneItem.Ftaxissuetype = "C") then %>selected<% end if %> >온라인주문</option>
						<option value="F" <% if (oTax.FOneItem.Ftaxissuetype = "F") then %>selected<% end if %> >오프라인주문</option>
						<option value="E" <% if (oTax.FOneItem.Ftaxissuetype = "E") then %>selected<% end if %> >기타매출</option>
						<!-- <option value="S" <% if (oTax.FOneItem.Ftaxissuetype = "S") then %>selected<% end if %> >출고리스트</option> -->
    					<option value="X" <% if (oTax.FOneItem.Ftaxissuetype = "X") then %>selected<% end if %> >내역없음</option>
    				</select>
    			</td>
    			<td height="25"></td>
    			<td align="left" bgcolor="#FFFFFF">
    			</td>
    		</tr>
    	</table>

<Br>

    	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    			<td height="25" width="15%"><b>위수탁구분</b></td>
    			<td align="left" bgcolor="#FFFFFF" width="35%">
    				<select class="select" name="consignYN">
    					<option value="">-선택-</option>
    					<option value="N" <% if (oTax.FOneItem.FconsignYN = "N") then %>selected<% end if %> >정상</option>
						<option value="Y" <% if (oTax.FOneItem.FconsignYN = "Y") then %>selected<% end if %> >위수탁(업체소비자매출)</option>
    				</select>
    			</td>
    			<td height="25" width="15%"><b>계산서발행</b></td>
    			<td align="left" bgcolor="#FFFFFF">
    				<select class="select" name="issueMethod">
    					<option value="">-선택-</option>
    					<!--<option value="bill36524" <% 'if (oTax.FOneItem.FissueMethod = "bill36524") then %>selected<% 'end if %> >BILL36524</option>-->
						<option value="WEHAGO" <% if (oTax.FOneItem.FissueMethod = "WEHAGO") then %>selected<% end if %> >위하고</option>
						<option value="eSero" <% if (oTax.FOneItem.FissueMethod = "eSero") then %>selected<% end if %> >이세로 수기</option>
    				</select>
    			</td>
    		</tr>
    	</table>

<Br>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr valign="top">
        <td width="49%">
        	<!-- 공급자정보 시작 -->
        	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        			<td colspan="4" height="25"><b>공급자 정보</b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">등록번호</td>
        			<td colspan="3">
        				<input type=text name="reg_socno" size=12 value="<%= oTax.FOneItem.FsupplyBusiNo %>" class="readonlybox" readonly>
        				<select class="select" name="billdiv" onchange="setRegisterInfo()">
							<option value="">공급자선택</option>
							<option value="">-----------</option>
							<% if (oTax.FOneItem.Fbilldiv <> "99") then  %>
							<option value="11" <% if (oTax.FOneItem.Fbilldiv = "11") then %>selected<%end if %>>소비자(업체매출)</option>
							<option value="01" <% if (oTax.FOneItem.Fbilldiv = "01") then %>selected<%end if %>>소비자(2013년출고)</option><!-- customer -->
							<option value="">-----------</option>
        					<option value="02" <% if (oTax.FOneItem.Fbilldiv = "02") then %>selected<%end if %>>가맹점(accounts)</option>
        					<option value="03" <% if (oTax.FOneItem.Fbilldiv = "03") then %>selected<%end if %>>프로모션(promotion)</option>
        					<option value="51" <% if (oTax.FOneItem.Fbilldiv = "51") then %>selected<%end if %>>기타매출(accounts)</option>
							<% if (oTax.FOneItem.Fbilldiv = "52") then %>
        						<option value="52" selected>유아러걸(youareagirl)</option>
							<%end if %>
							<% if (oTax.FOneItem.Fbilldiv = "53") then %>
        						<option value="53" selected>아이씽소(ithinkso)</option>
							<%end if %>
							<% if (oTax.FOneItem.Fbilldiv = "54") then %>
        						<option value="54" selected>텐텐리빙(living1010)</option>
							<%end if %>
							<% if (oTax.FOneItem.Fbilldiv = "55") then %>
        						<option value="55" selected>에이플러스비(aplusb)</option>
							<%end if %>
							<% else %>
							<option value="99" <% if (oTax.FOneItem.Fbilldiv = "99") then %>selected<%end if %>><%= tplcompanyname %>(<%= tplbillUserID %>)</option>
							<% end if%>
        				</select>
        			</td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td width="70" bgcolor="#F0F0FD" height="25">상호</td>
        			<td><input type=text name="reg_socname" size=14 value="<%= oTax.FOneItem.FsupplyBusiName %>" border=0 class="readonlybox" readonly></td>
        			<td width="70" bgcolor="#F0F0FD">대표자</td>
        			<td><input type=text name="reg_ceoname" size=8 value="<%= oTax.FOneItem.FsupplyBusiCEOName %>" class="readonlybox" readonly></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">사업장주소</td>
        			<td colspan="3"><input type=text name="reg_socaddr" size=40 value="<%= oTax.FOneItem.FsupplyBusiAddr %>" class="readonlybox" readonly></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">업태</td>
        			<td colspan=2><input type=text name="reg_socstatus" size=20 value="<%= oTax.FOneItem.FsupplyBusiType %>" class="readonlybox" readonly></td>
        			<td bgcolor="#F0F0FD">종사업장번호</td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">종목</td>
        			<td colspan=2><input type=text name="reg_socevent" size=20 value="<%= oTax.FOneItem.FsupplyBusiItem %>" class="readonlybox" readonly></td>
        			<td><input type=text name="reg_subsocno" size=4 value="<%= oTax.FOneItem.FsupplyBusiSubNo %>" class="readonlybox" readonly></td>
        		</tr>
        		<tr><td height="1" colspan="4" bgcolor="#FFFFFF"></td></tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">담당자</td>
        			<td><input type=text name="reg_managername" size=14 value="<%= oTax.FOneItem.FsupplyRepName %>" class="readonlybox" readonly></td>
        			<td bgcolor="#F0F0FD">연락처</td>
        			<td><input type=text name="reg_managerphone" size=14 value="<%= oTax.FOneItem.FsupplyRepTel %>" class="readonlybox" readonly></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">이메일</td>
        			<td colspan=3><input type=text name="reg_managermail" size=20 value="<%= oTax.FOneItem.FsupplyRepEmail %>" class="readonlybox" readonly></td>
        		</tr>
        	</table>
        	<!-- 공급자정보 끝 -->
        </td>
        <td>&nbsp;</td>
        <td width="49%">
			<!-- 공급받는자정보 시작 -->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td colspan="4" height="25" align="right">
						<b>공급받는자 정보</b>
						&nbsp;
						&nbsp;
						&nbsp;
						<input type="button" class="button" value="[거래처검색]" onClick="jsGetCust('frm');">
						<input type="button" class="button" value="[사업자검색]" onClick="PopUpcheSelectBySocno('frm');">
						<% '<a href="http://www.nts.go.kr/cal/cal_check_02.asp" target="_blank">[사업자조회]</a> %>
					</td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF">
					<td bgcolor="#F0F0FD" height="25">등록번호</td>
					<td colspan="3">
						<input type=text name="socno" size=13 value="<%= oTax.FOneItem.FbusiNo %>" class="writebox">
						<% if (mode = "new") and (issuetype <> "etcmeachul") then %>
						<input type="button" class="button_s" value="검 색" onClick="SearchSocno()">
						<% end if %>
						<% if (userid <> "") then %>
							<input type="button" class="button_s" value="기존(<%= previssuecount %>)" onClick="popListPreviousCustomerTaxSheet('<%= userid %>')" <% if (previssuecount < 1) then %>disabled<% end if %>>
						<% end if %>
					</td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF">
					<td width="70" bgcolor="#F0F0FD" height="25">상호</td>
					<td align="left"><input type=text name="socname" size=14 value="<%= oTax.FOneItem.FbusiName %>" border=0 class="writebox"></td>
					<td width="70" bgcolor="#F0F0FD">대표자</td>
					<td align="left"><input type=text name="ceoname" size=14 value="<%= oTax.FOneItem.FbusiCEOName %>" class="writebox"></td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF">
					<td bgcolor="#F0F0FD" height="25">사업장주소</td>
					<td align="left" colspan="3"><input type=text name="socaddr" size=40 value="<%= oTax.FOneItem.FbusiAddr %>" class="writebox"></td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF">
					<td bgcolor="#F0F0FD" height="25">업태</td>
					<td colspan=2><input type=text name="socstatus" size=20 value="<%= oTax.FOneItem.FbusiType %>" class="writebox"></td>
					<td bgcolor="#F0F0FD">종사업장번호</td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF">
					<td bgcolor="#F0F0FD" height="25">종목</td>
					<td colspan=2><input type=text name="socevent" size=20 value="<%= oTax.FOneItem.FbusiItem %>" class="writebox"></td>
					<td><input type=text name="subsocno" size=4 value="" class="writebox"></td>
				</tr>
				<tr><td height="1" colspan="4" bgcolor="#FFFFFF"></td></tr>
				<tr align="center" bgcolor="#FFFFFF">
					<td bgcolor="#F0F0FD" height="25">담당자</td>
					<td align="left"><input type=text name="managername" size=14 value="<%= oTax.FOneItem.FrepName %>" class="writebox"></td>
					<td bgcolor="#F0F0FD">연락처</td>
					<td align="left"><input type=text name="managerphone" size=14 value="<%= oTax.FOneItem.FrepTel %>" class="writebox"></td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF">
					<td bgcolor="#F0F0FD" height="25">이메일</td>
					<td align="left" colspan="3"><input type=text name="managermail" size=40 value="<%= oTax.FOneItem.FrepEmail %>" class="writebox"></td>
				</tr>
			</table>
			<!-- 공급받는자정보 끝 -->
        </td>
	</tr>
</table>

	</td>
</tr>
<tr height="5">
	<td>
	</td>
</tr>
<tr>
	<td>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="#F0F0FD">
		<td width="120" height="25">발행일</td>
		<td width="100">공급가액</td>
		<td width="100">과세구분</td>
		<td width="100">세액</td>
		<td width="100">합계금액</td>
		<td>비고</td>
	</tr>
    <tr align="center" bgcolor="#FFFFFF">
		<td height="25">
			<input type="text" size="10" name="yyyymmdd" value="<%= oTax.FOneItem.FisueDate %>" onClick="jsPopCal('frm','yyyymmdd');" style="cursor:hand;" class="writebox">
		</td>
		<td><input type=text name="totalsupplyprice" size=10 value="" class="readonlybox" readonly></td>
		<td>
			<select name=taxtype class="select" onChange="CalcPrice(this)">
				<option value="">====</option>
				<option value="Y" <% if (oTax.FOneItem.FTaxType = "Y") then %>selected<% end if %>>과세</option>
				<option value="N" <% if (oTax.FOneItem.FTaxType = "N") then %>selected<% end if %>>면세</option>
				<option value="0" <% if (oTax.FOneItem.FTaxType = "0") then %>selected<% end if %>>영세</option>
			</select>
		</td>
		<td><input type=text name="totaltaxprice" size=10 value="<%= (oTax.FOneItem.FtotalTax) %>" class="readonlybox" readonly  onkeyup="CalcPrice(this)"></td>
		<td><input type=text name="totalprice" size=10 value="<%= (oTax.FOneItem.FtotalPrice) %>" class="writebox" onkeyup="CalcPrice(this)"></td>
		<% if (mode <> "new") then %>
		<input type="hidden" name="orgtotalprice" value="<%= oTax.FOneItem.FtotalPrice %>">
		<% end if %>
		<td>
			<%
			if (oTax.FOneItem.FtaxIdx = "") then
				'// 작성
				%>
				<input type=text name="orderidx" size=20 value="<%= oTax.FOneItem.Forderidx %>" class="writebox">
				<% if (mode = "new") and (issuetype = "etcmeachul") and (etcmeachulIdx <> "") then %>
				<input type=button class=button name="btnCombine" value="추가" onClick="popMeachulDetailList()">
				<% end if %>
				<%
			else
				'// 수정
				%>
				<% if (Trim(oTax.FOneItem.Forderserial) <> "") then %>
				주문번호/출고코드 : <%= oTax.FOneItem.Forderserial %>
				<% elseif (CStr(oTax.FOneItem.Forderidx) <> "0") and (CStr(oTax.FOneItem.Forderidx) <> "") then %>
				인덱스코드 : <%=oTax.FOneItem.Forderidx %>
				<% elseif Not IsNull(oTax.FOneItem.GetMultiOrderIdxList) then  %>
				인덱스코드 : <%= oTax.FOneItem.GetMultiOrderIdxList %>
				<% end if %>
				<%
			end if
			%>
		</td>
	</tr>
</table>

	</td>
</tr>
<tr height="5">
	<td>
	</td>
</tr>
<tr>
	<td>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="#F0F0FD">
		<td width="30" height="25">월</td>
		<td width="30">일</td>
		<td>품목</td>
		<td width="50">규격</td>
		<td width="50">수량</td>
		<td width="100">단가</td>
		<td width="100">공급가액</td>
		<td width="100">세액</td>
	</tr>
    <tr align="center" bgcolor="#FFFFFF">
		<td height="25">
			<%= mid(oTax.FOneItem.FisueDate,6,2) %>
		</td>
		<td><%= mid(oTax.FOneItem.FisueDate,9,2) %></td>
		<td><input type=text name="itemname" size=40 value="<%=db2html(oTax.FOneItem.Fitemname)%>" class="writebox"></td>
		<td></td>
		<td><%= CHKIIF(oTax.FOneItem.Fbilldiv = "01","","1") %></td>

		<td><input type=text name="totalsupplyprice2" size=10 value="" class="writebox" onkeyup="CalcPrice(this)"></td>
		<td><input type=text name="totalsupplyprice3" size=10 value="" class="readonlybox" readonly></td>
		<td><input type=text name="totaltaxprice2" size=10 value="<%= (oTax.FOneItem.FtotalTax) %>" class="readonlybox" readonly></td>
	</tr>
</table>
	</td>
</tr>
<tr height="5">
	<td>
	</td>
</tr>
<tr>
	<td>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="#F0F0FD">
		<td height="25"><b>합계금액</b></td>
		<td width="100">현금</td>
		<td width="100">수표</td>
		<td width="100">어음</td>
		<td width="100">외상미수금</td>
	</tr>
    <tr align="center" bgcolor="#FFFFFF">
		<td height="25"><input type=text name="totalprice2" size=10 value="<%= (oTax.FOneItem.FtotalPrice) %>" class="readonlybox" readonly></td>
		<td>
<% if (oTax.FOneItem.Fbilldiv = "01") then %>
			<input type=text name="totalprice3" size=10 value="<%= (oTax.FOneItem.FtotalPrice) %>" class="readonlybox" readonly>
<% end if %>
		</td>
		<td></td>
		<td></td>
		<td>
<% if (oTax.FOneItem.Fbilldiv <> "01") then %>
			<input type=text name="totalprice3" size=10 value="<%= (oTax.FOneItem.FtotalPrice) %>" class="readonlybox" readonly>
<% end if %>
		</td>
	</tr>
</table>

	</td>
</tr>
<% if (mode = "new") then %>
<tr>
	<td align="right">
		<input type="checkbox" name="ckHand" onClick="chgHandTax(this)">세액 수기입력
	</td>
</tr>
<% end if %>
</form>
<tr>
	<td>

		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
		    <tr align="center">
				<td align="center" height="25">
					<% if (mode = "new") then %>
										<input type="button" class="button" value="작성" onClick="doRegisterSheet()">
					<% else %>
						<% if (oTax.FOneItem.FisueYn = "N") then %>
						<input type="button" class="button" value="수정" onClick="GotoTaxModify()">
						&nbsp;
						<% end if %>
						<input type="button" class="button" value="목록" onClick="self.location='Tax_list.asp?menupos=<%=menupos %>'">
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<% if (oTax.FOneItem.FisueYn = "Y") then %>
					        <% if (oTax.FOneItem.FdelYn = "Y") then %>
					        <input type="button" class="button" value="삭제(불가)" onClick="GotoTaxDel('<%= oTax.FOneItem.FdelYn %>')" disabled >
					        <font color="red">(발행후 삭제된 내역입니다.)</font>
					        <% else %>
					        	<% if C_ADMIN_AUTH or C_MngPart or C_PSMngPart then %>
					        		<input type="button" class="button" value="삭제(관리자)" onClick="GotoTaxDel('<%= oTax.FOneItem.FdelYn %>')" >
					        	<% else %>
					        		<input type="button" class="button" value="삭제(불가)" onClick="GotoTaxDel('<%= oTax.FOneItem.FdelYn %>')" disabled >
					        	<% end if %>
					        <% end if %>
						<% ELSE %>
					        <% if (oTax.FOneItem.FdelYn = "Y") then %>
							<input type="button" class="button" value="삭제(불가)" disabled >
							(발행전 삭제된 내역입니다.)
							<% else %>
							<input type="button" class="button" value="삭제" onClick="GotoTaxDel('<%= oTax.FOneItem.FisueYn %>')">
							<% end if %>
						<% end if %>
					<% end if %>
				</td>
			</tr>
		</table>

	</td>
</tr>
</table>

<br>

<iframe src="" name="icheckframe" width="0" height="0" frameborder="1" marginwidth="0" marginheight="0" topmargin="0" scrolling="no"></iframe>

<script type='text/javascript'>

// 페이지 시작시 작동하는 스크립트
function getOnload(){
	setRegisterInfo();
	CalcPrice(document.frm.taxtype);
}

window.onload = getOnload;

</script>

<%
function Is3PLShopid(shopid)
	dim sqlStr

	Is3PLShopid = False

	sqlStr = " select top 1 p.id as shopid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_partner p with (nolock)"
	sqlStr = sqlStr + " 	join db_partner.dbo.tbl_partner_tpl t with (nolock)"
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		p.tplcompanyid = t.tplcompanyid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and p.id = '" + CStr(shopid) + "' "
	sqlStr = sqlStr + " 	and IsNull(p.tplcompanyid, '') <> '' "

	'response.write sqlStr & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.EOF then
		Is3PLShopid = True
	end if
	rsget.close
end function

function Get3PLUpcheInfoByShopid(shopid, byRef tplcompanyid, byRef tplcompanyname, byRef tplgroupid, byRef tplbillUserID)
	dim sqlStr

	sqlStr = " select top 1 t.tplcompanyid, t.tplcompanyname, t.groupid as tplgroupid, billUserID as tplbillUserID "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_partner p with (nolock)"
	sqlStr = sqlStr + " 	join db_partner.dbo.tbl_partner_tpl t with (nolock)"
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		p.tplcompanyid = t.tplcompanyid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and p.id = '" + CStr(shopid) + "' "
	sqlStr = sqlStr + " 	and IsNull(p.tplcompanyid, '') <> '' "

	'response.write sqlStr & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.EOF then
		tplcompanyid = rsget("tplcompanyid")
		tplcompanyname = db2html(rsget("tplcompanyname"))
		tplgroupid = rsget("tplgroupid")
		tplbillUserID = rsget("tplbillUserID")
	end if
	rsget.close
end function

function Get3PLUpcheInfo(tplcompanyid, byRef tplcompanyname, byRef tplgroupid, byRef tplbillUserID)
	dim sqlStr

	sqlStr = " select top 1 t.tplcompanyid, t.tplcompanyname, t.groupid as tplgroupid, billUserID as tplbillUserID "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_partner p with (nolock)"
	sqlStr = sqlStr + " 	join db_partner.dbo.tbl_partner_tpl t with (nolock)"
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		p.tplcompanyid = t.tplcompanyid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and IsNull(p.tplcompanyid, '') = '" + CStr(tplcompanyid) + "' "

	'response.write sqlStr & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.EOF then
		tplcompanyname = db2html(rsget("tplcompanyname"))
		tplgroupid = rsget("tplgroupid")
		tplbillUserID = rsget("tplbillUserID")
	end if
	rsget.close
end function
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
