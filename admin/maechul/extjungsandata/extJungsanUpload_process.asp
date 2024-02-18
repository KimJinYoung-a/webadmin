<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

function getLtiMallRound(iorgprice)
	dim rudprc : rudprc = ROUND(iorgprice,1)
	dim multiple : multiple = 1
	if (iorgprice<0) then multiple=-1

	if RIGHT(CStr(rudprc),2)=".5" then
		if LEFT(RIGHT(CStr(rudprc),3),1)="0" then
			getLtiMallRound = ROUND(rudprc-0.5*multiple,0)
		elseif LEFT(RIGHT(CStr(rudprc),3),1)="9" then
			getLtiMallRound = ROUND(rudprc+0.5*multiple,0)
		else
			getLtiMallRound = ROUND(rudprc,0)
		end if
	else
		getLtiMallRound = ROUND(rudprc,0)
	end if
end function

''제휴몰 정산내역 추가방법
''ADD EXT SHOP 검색하여 추가한다.

''response.write "작업중.."
''response.end

dim extsellsite
dim errMSG
dim affectedRows


'==============================================================================
dim lastPageTime, pageElapsedTime
lastPageTime = Timer

function checkAndWriteElapsedTime(memo)
	pageElapsedTime = Timer - lastPageTime
	lastPageTime = Timer
	response.write "<!-- Page Execute Time Check : " & FormatNumber(pageElapsedTime, 4) & " : " & memo & " -->" & vbCrLf
end function

Call checkAndWriteElapsedTime("001")

''response.end
dim DefaultPath : DefaultPath =	server.mappath("/admin/maechul/extjungsandata/upFiles/")
dim FileMaxSize : FileMaxSize = 15


'// ============================================================================
'// 업로드 컨퍼넌트 선언 //
dim fso
dim uprequest, sqlStr
dim extMeachulMonth

dim objfso
set objfso = server.CreateObject("scripting.Filesystemobject")

if not objfso.FolderExists(DefaultPath) then
	objfso.CreateFolder(DefaultPath)
end if

set uprequest = Server.CreateObject("TABSUpload4.Upload")
uprequest.Start DefaultPath

extsellsite = requestCheckvar(uprequest("extsellsite"),32)
extMeachulMonth = requestCheckvar(uprequest("extMeachulMonth"),7)

if (extsellsite = "") then
	dbget.close()
    response.write "No site name..."
    response.end
end if

dim fullpath, filename

'// YYYYMMDDHHmmSS
sqlStr = " select Left(Replace(Replace(Replace(CONVERT(varchar, GETDATE(), 127), '-', ''), ':', ''), 'T', ''), 14) as filename" + VbCrlf
rsget.CursorLocation = adUseClient
rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	filename = rsget("filename")
rsget.close


fullpath = uprequest("sFile").saveas((DefaultPath &"\"& filename + "_" + extsellsite + "."+ uprequest("sFile").FileType), True)


'// ============================================================================
dim conXL, rsXL, sheetName

set conXL = Server.CreateObject("ADODB.Connection")
if (extsellsite="kakaogift") or (extsellsite="kakaostore") or (extsellsite="boribori1010") or (extsellsite="wconcept1010") or (extsellsite="goodwearmall10") or (extsellsite="goodwearmall10beasongpay") or (extsellsite="goodshop1010") or (extsellsite="coupang") or (extsellsite="ssg6006") or (extsellsite="ssg6007") or (extsellsite = "nvstorefarm") or (extsellsite = "nvstoremoonbangu") or (extsellsite = "Mylittlewhoopee") or (extsellsite = "nvstoregift") or (extsellsite = "nvstorefarmclass") or (extsellsite = "wadsmartstore") or (extsellsite = "lotteon") or (extsellsite = "yes24") or (extsellsite = "halfclubproduct") or (extsellsite = "halfclubbeasongpay") or (extsellsite = "gsshopproduct") or (extsellsite = "gsshopbeasongpay") or (extsellsite = "gsshopproductday") or (extsellsite = "WMP") or (extsellsite = "WMPbeasongpay") or (extsellsite = "wmpfashion") or (extsellsite = "wmpfashionbeasongpay") or (extsellsite = "ohou1010") or (extsellsite = "LFmall") or (extsellsite = "aboutpet") then
    conXL.Provider = "Microsoft.ACE.OLEDB.12.0"
    'conXL.Properties("ExtEnded Properties").Value = "Excel 12.0;IMEX=1"

	if (extsellsite="WMP") or (extsellsite="WMPbeasongpay") or (extsellsite="wmpfashion") or (extsellsite="wmpfashionbeasongpay") then
		conXL.Properties("ExtEnded Properties").Value = "Excel 12.0;HDR=NO;IMEX=1;TypeGuessRows=0;ImportMixedTypes=Text;"  '' 한컬럼에 text와 money가 혼재되어 있음.
	else
		conXL.Properties("ExtEnded Properties").Value = "Excel 12.0;IMEX=1"
	end if
else
    conXL.Provider = "Microsoft.Jet.oledb.4.0"
    conXL.Properties("ExtEnded Properties").Value = "Excel 8.0;IMEX=1"
end if


conXL.Open fullpath

if (ERR) then
	uprequest("sFile").delete
	set objfso = Nothing
	set uprequest = Nothing
	response.write "ERROR : 오류가 발생했습니다. 시스템팀 문의[0]"
	response.end
end if

set rsXL = conXL.OpenSchema(adSchemaTables)

if Not rsXL.Eof then
	sheetName = rsXL.Fields("table_name").Value
end if

set rsXL = Nothing

if (extsellsite="ezwel") then
	sheetName = "Sheet2$"
end if
'// ============================================================================
dim sellsite, extOrderserial, extOrderserSeq, extOrgOrderserial, extItemNo, extItemCost, extReducedPrice, extOwnCouponPrice, extTenCouponPrice, extJungsanType, extCommPrice, extTenMeachulPrice, extTenJungsanPrice
dim extMeachulDate, extJungsanDate, isextMeachulDate
dim extItemName, extItemOptionName, tmpItemname, tmpSellerAddSalePriceBy10x10, tmpSellerAddSalePriceBy11st1010
dim IsOrderData, IsValidInput, IsReturnOrder
dim extVatYN, extCommSupplyPrice, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice
dim tmpDate, tmpDate2, totExtCommSupplyPrice, totExtTenMeachulSupplyPrice, totExtCommSupplyPrice2, totExtTenMeachulSupplyPrice2
dim totExtTenMeachulPrice, totExtCommPrice, totExtTenMeachulPrice2, totExtCommPrice2
dim tenitemid,tenitemoption,extitemid,extitemoption
dim tmpStr, pMxMeachulMonth
dim i
dim diffPrice
dim plusMinus
dim validitemno, extReItemNo, extChgItemno
dim siteNo, dvlprice, lfmallEtcPrice
Dim lotteondvlprice1, lotteondvlprice2, lotteondvlPGCommprice, lotteonBanpoomDate, lotteondvlTotCommprice, dlvCommprice
dim p_extOrderserial, p_extOrderserSeq
dim dlvsheettype
Dim cjbeasongGubun
set rsXL = Server.CreateObject("ADODB.Recordset")

sqlStr = "select * from [" & sheetName & "]"

rsXL.Open sqlStr, conXL

if (ERR) then
	uprequest("sFile").delete
	set objfso = Nothing
	set uprequest = Nothing
	response.write "ERROR : 오류가 발생했습니다. 시스템팀 문의[1]"
	response.end
end if

Call checkAndWriteElapsedTime("002")

rw extsellsite
If Not rsXL.Eof Then

	''ADD EXT SHOP. 01. 사이트구분
	Select Case extsellsite
		Case "interparkproduct"
			'// 인터파크 상품정산 상세내역
			sellsite = "interpark"
		''\upload\linkweb\extjungsandata\extJungsanUpload_process.asp ' 샘플은 여기서 확인.
		Case "kakaogift"
		    sellsite = "kakaogift"
		Case "kakaostore"
		    sellsite = "kakaostore"
		Case "boriboriproduct", "boriboribeasongpay"
			sellsite = "boribori1010"
        Case "coupang"
		    sellsite = "coupang"
		Case "11st1010"
		    sellsite = "11st1010"
		Case "GS25"
		    sellsite = "GS25"
		Case "withnature1010"
		    sellsite = "withnature1010"
		Case "ssg6006","ssg6007"
		    sellsite = "ssg"
		Case "cjmallbeasongpay", "cjmallproduct"
			sellsite = "cjmall"
		Case "wconcept1010"
			sellsite = "wconcept1010"
		Case "goodshop1010"
			sellsite = "goodshop1010"
		Case "gmarket1010","gmarket1010beasongpay"
			sellsite = "gmarket1010"
		Case "auction1010","auction1010beasongpay"
			sellsite = "auction1010"
		Case "ezwel"
			sellsite = "ezwel"
		Case "nvstorefarm"
		    sellsite = "nvstorefarm"
		Case "nvstorefarmclass"
			sellsite = "nvstorefarmclass"
		Case "nvstoremoonbangu"
			sellsite = "nvstoremoonbangu"
		Case "Mylittlewhoopee"
			sellsite = "Mylittlewhoopee"
		Case "nvstoregift"
			sellsite = "nvstoregift"
		Case "wadsmartstore"
			sellsite = "wadsmartstore"
		Case "lotteCom", "lotteCombeasongpay"
			sellsite = "lotteCom"
		Case "halfclubproduct", "halfclubbeasongpay"
			sellsite = "halfclub"
		Case "gsshopproduct", "gsshopbeasongpay", "gsshopproductday"
			sellsite = "gseshop"
		Case "lotteimall"
			sellsite = "lotteimall"
		Case "hmallproduct", "hmallbeasongpay"
			sellsite = "hmall1010"
		Case "WMP","WMPbeasongpay"
			sellsite = "WMP"
		Case "wmpfashion","wmpfashionbeasongpay"
			sellsite = "wmpfashion"
		Case "LFmall"
			sellsite = "LFmall"
		Case "lotteon"
		    sellsite = "lotteon"
		Case "yes24"
		    sellsite = "yes24"
		Case "alphamallMaechul", "alphamallHuanBool"
		    sellsite = "alphamall"
		Case "ohou1010"
		    sellsite = "ohou1010"
		Case "casamia_good_com"
		    sellsite = "casamia_good_com"
		Case "aboutpet"
		    sellsite = "aboutpet"
		Case "shintvshopping", "shintvshoppingbeasongpay"
		    sellsite = "shintvshopping"
		Case "wetoo1300k", "wetoo1300kbeasongpay"
		    sellsite = "wetoo1300k"
		Case "skstoa", "skstoabeasongpay"
		    sellsite = "skstoa"
		Case "goodwearmall10", "goodwearmall10beasongpay"
		    sellsite = "goodwearmall10"
		Case Else
			sellsite = ""
	End Select

	if (sellsite = "") then
		uprequest("sFile").delete
		set objfso = Nothing
		set uprequest = Nothing
		response.write "ERROR : 오류가 발생했습니다. 시스템팀 문의[2]"
		response.end
	end if

	'if (sellsite<>"ssg") then
		sqlStr = " delete from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite = '" + CStr(sellsite) + "' "
		''response.write sqlStr
		dbget.execute sqlStr
	'end if

	Call checkAndWriteElapsedTime("003")

	errMSG = ""
	extJungsanDate = ""
	i = 0
	do Until rsXL.Eof

		IsOrderData = False
		IsValidInput = True
		''IsReturnOrder = False

		'//ADD EXT SHOP. 02.
		'// 데이타 검증 01 (주문내역인지)
        ''\upload\linkweb\extjungsandata\extJungsanUpload_process.asp ' 샘플은 여기서 확인.
		Select Case extsellsite
			Case "interparkproduct"
				if (IsNumeric(rsXL(0)) = True) then
					IsOrderData = True
				end if
			Case "interparkbeasongpay"
				if (IsNumeric(rsXL(0)) = True) then
					IsOrderData = True
				end if
			Case "interparkipoint"
				if (IsNumeric(rsXL(0)) = True) and (rsXL(0) <> "0") then
					IsOrderData = True
				end if

				if (i = 0) and (rsXL(0) <> "기간별수수료사용현황 - I-Point적립") then
					IsValidInput = False
					exit do
				end if
			Case "interparkpointmall"
				if (IsNumeric(rsXL(0)) = True) and (rsXL(0) <> "0") then
					IsOrderData = True
				end if

				if (i = 0) and (rsXL(0) <> "기간별수수료사용현황 - 포인트몰판매수수료") then
					IsValidInput = False
					exit do
				end if

            Case "kakaogift"
                if (IsNumeric(rsXL(0)) and (rsXL(0) <> "")) then
					IsOrderData = True
				end if
            Case "kakaostore"
                if (IsNumeric(rsXL(0)) and (rsXL(0) <> "")) then
					IsOrderData = True
				end if
            Case "boriboriproduct"
                If (IsNumeric(rsXL(7)) and (rsXL(7) <> "")) Then
					IsOrderData = True
				End If
            Case "wconcept1010"
                If (rsXL(5) <> "") Then
					IsOrderData = True
				End If
            Case "goodshop1010"
                If (rsXL(4) <> "") Then
					IsOrderData = True
				End If
            Case "boriboribeasongpay"
                If (IsNumeric(rsXL(7)) and (rsXL(7) <> "")) Then
					IsOrderData = True
				End If
			Case "coupang"
                if (IsNumeric(rsXL(0)) and (rsXL(0) <> "")) then
					IsOrderData = True
				end if
			Case "11st1010"
				if (rsXL(0) <> "") then
					if (IsNumeric(rsXL(1)) = True) then
						IsOrderData = True
					end if
				end if
			Case "GS25"
				if (rsXL(3) <> "") then
					if (Len(rsXL(3)) = 13) then
						IsOrderData = True
					end if
				end if
			Case "withnature1010"
				if (rsXL(2) <> "") then
					if (Len(rsXL(2)) = 12) then
						IsOrderData = True
					end if
				end if
			Case "ssg6006", "ssg6007"
				if (rsXL(0) <> "") then
					if (IsNumeric(rsXL(0)) = True) then
						IsOrderData = True
					end if
				end if
			Case "cjmallbeasongpay"
				If (uprequest("cjbeasongGubun") = "1") OR uprequest("cjbeasongGubun") = "2" Then	'1 : 대금지불현황(교환택배비) / 2 : 대금지불현황(반품택배비)
					if (rsXL(4) <> "") then
						if (IsNumeric(rsXL(4)) = True) then
							IsOrderData = True
						end if
					end if
				ElseIf uprequest("cjbeasongGubun") = "3" Then										'3 : 대금지불현황(저단가배송비)
					if (rsXL(5) <> "") then
						if (IsNumeric(rsXL(5)) = True) then
							IsOrderData = True
						end if
					end if
				End If
				' if (sheetName="'대금지급현황(저단가배송비)$'") or (sheetName="sheet1$") then  ''(sheetName="sheet1$") : 무료배송쿠폰
				' 	if (rsXL(4) <> "") then
				' 		if (IsNumeric(rsXL(4)) = True) then
				' 			IsOrderData = True
				' 		end if
				' 	end if
				' elseif (sheetName="'대금지불현황(반품택배비)$'") or (sheetName="'대금지불현황(교환택배비)$'") then
				' 	if (rsXL(3) <> "") then
				' 		if (IsNumeric(rsXL(3)) = True) then
				' 			IsOrderData = True
				' 		end if
				' 	end if
				' elseif (sheetName="'대금지급현황(AS 택배비)$'") then
				' 	if (rsXL(7) <> "") then
				' 		if (IsNumeric(rsXL(7)) = True) then ''상품코드
				' 			IsOrderData = True
				' 		end if
				' 	end if
				' end if
			Case "cjmallproduct"
				if (rsXL(23) <> "") then
					if (Len(rsXL(23)) = 26) then
						IsOrderData = True
					end if
				end if
			Case "gmarket1010","gmarket1010beasongpay"
				if (rsXL(0) <> "") then  ''20200110 번호필드추가됨 '' 20200301 rsXL(1) 없는경우 있음.
					if (IsNumeric(rsXL(0)) = True) then ''장바구니번호
						IsOrderData = True
					end if
				end if
			Case "auction1010","auction1010beasongpay"
				if (rsXL(1) <> "") then
					if (Len(rsXL(1)) = 10) then
						IsOrderData = True
					end if
				end if
			Case "ezwel"
				if (rsXL(0) <> "") then
					if rsXL(1) = "배송완료일" then
						IsOrderData = True
					end if
				end if
			Case "nvstorefarm", "nvstorefarmclass", "nvstoremoonbangu", "nvstoregift", "wadsmartstore", "Mylittlewhoopee"
				if (rsXL(0) <> "") then
					if (rsXL(13) = "일반정산") or (rsXL(13) = "정산후 취소") or (rsXL(13) = "빠른정산") or (rsXL(13) = "빠른정산 회수") then
						IsOrderData = True
					end if
				end if
			Case "lotteon"
				if (rsXL(7) <> "") then
					if (Len(rsXL(7)) = 16) then
						IsOrderData = True
					end if
				end if
			Case "yes24"
				if (rsXL(1) <> "") then
					if (Len(rsXL(1)) = 11) then
						IsOrderData = True
					end if
				end if
			Case "alphamallMaechul", "alphamallHuanBool"
				if (rsXL(2) <> "") then
					if (Len(rsXL(2)) = 21) then
						IsOrderData = True
					end if
				end if
			Case "ohou1010"
				if (rsXL(5) <> "") then
					if (Len(rsXL(5)) = 9) then
						IsOrderData = True
					end if
				end if
			Case "casamia_good_com"
				if (rsXL(2) <> "") then
					if (Len(rsXL(2)) = 14) then
						IsOrderData = True
					end if
				end if
			Case "aboutpet"
				if (rsXL(0) <> "") then
					if (Len(rsXL(0)) = 16) then
						IsOrderData = True
					end if
				end if
			Case "lotteCom"
				if (Len(rsXL(24)) = 18) and (IsNumeric(rsXL(10)) = True) then
					IsOrderData = True
				end if
			Case "lotteCombeasongpay"
				if (Len(rsXL(3)) = 18) and (IsNumeric(rsXL(0)) = True) then
					IsOrderData = True
				end if
			Case "halfclubproduct"
				if (Len(rsXL(3)) = 12) and (IsNumeric(rsXL(3)) = True) then
					IsOrderData = True
				end if
			Case "halfclubbeasongpay"
				if (Len(rsXL(1)) = 12) and (rsXL(1)<>"") and (IsNumeric(rsXL(1)) = True) then
					IsOrderData = True
				end if
			Case "gsshopproduct"
				if (rsXL(0) <> "") then
					if (IsNumeric(rsXL(0)) = True) then
						IsOrderData = True
					end if
				end if
			Case "gsshopbeasongpay"
				if (rsXL(1) <> "") then
					if (IsNumeric(rsXL(1)) = True) then
						IsOrderData = True
					end if
				end if
			Case "gsshopproductday"
				if (rsXL(0) <> "") then
					if (IsNumeric(rsXL(8)) = True) then
						IsOrderData = True
					end if
				end if
			Case "lotteimall"
				if (rsXL(0) <> "") then
					if (IsNumeric(rsXL(0)) = True) then
						IsOrderData = True
					end if
				end if
			Case "hmallproduct"
				if (rsXL(0) <> "") then
					if (Len(rsXL(2)) = 15) and (IsNumeric(rsXL(0)) = True) then
						IsOrderData = True
					end if
				end if
			Case "hmallbeasongpay"
				if (rsXL(0) <> "") then
					if (Len(rsXL(3)) = 15) and (IsNumeric(rsXL(0)) = True) then
						IsOrderData = True
					end if
				end if
			Case "WMP", "wmpfashion"
				if (rsXL(0) <> "") then
					if (Len(rsXL(1+1)) = 8 or Len(rsXL(1+1)) = 9) and (IsNumeric(rsXL(0)) = True) then
						IsOrderData = True
					end if
				end if
			Case "WMPbeasongpay", "wmpfashionbeasongpay"
				if (rsXL(0) <> "") then
					if (Len(rsXL(0)) = 8 or Len(rsXL(0)) = 9) and (IsNumeric(rsXL(0)) = True) then
						IsOrderData = True
					end if
				end if
			Case "shintvshopping", "shintvshoppingbeasongpay"
				if (rsXL(1) <> "") then
					if (Len(Replace(rsXL(1), "-", "")) = 14 or Len(Replace(rsXL(1), "-", "")) = 14) and (IsNumeric(Replace(rsXL(1), "-", "")) = True) then
						IsOrderData = True
					end if
				end if
			Case "skstoa", "skstoabeasongpay"
				if (rsXL(3) <> "") then
					if (Len(Replace(rsXL(3), "-", "")) = 14 or Len(Replace(rsXL(3), "-", "")) = 14) and (IsNumeric(Replace(rsXL(3), "-", "")) = True) then
						IsOrderData = True
					end if
				end if
			Case "goodwearmall10"
				if (rsXL(0) <> "") then
					if (Len(rsXL(0)) = 17) then
						IsOrderData = True
					end if
				end if
			Case "goodwearmall10beasongpay"
				if (rsXL(4) <> "") then
					if (Len(rsXL(4)) = 17) then
						IsOrderData = True
					end if
				end if
			Case "wetoo1300k", "wetoo1300kbeasongpay"
				if (rsXL(0) <> "") then
					if (Len(rsXL(0)) = 15) and (IsNumeric(rsXL(0)) = True) then
						IsOrderData = True
					end if
				end if
			Case "LFmall"
				if (rsXL(1) <> "") then
					if (Len(rsXL(0)) = 10) and (IsNumeric(rsXL(1)) = True) and rsXL(9) <> "추가정산" then
						IsOrderData = True
					end if
				end if
			Case Else
				IsOrderData = False
		End Select

		if (IsOrderData = True) then

			'//ADD EXT SHOP. 03. 처리
            ''\upload\linkweb\extjungsandata\extJungsanUpload_process.asp ' 샘플은 여기서 확인.
			dvlprice = 0
			lfmallEtcPrice = 0
			lotteondvlprice1 = 0
			lotteondvlprice2 = 0
			lotteondvlPGCommprice = 0
			lotteondvlTotCommprice = 0
			dlvCommprice = 0
			lotteonBanpoomDate = ""
			Select Case extsellsite
				Case "interparkproduct"

				Case "interparkbeasongpay"

				Case "interparkipoint"

				Case "interparkpointmall"

				Case "boriboriproduct"
					'// 보리보리 상품정산 상세내역
					extOrderserial = rsXL(7)	'주문번호
					If Len(extOrderserial) = 12 Then
						extJungsanDate = ""
						extMeachulDate = LEFT(rsXL(6), 4) & "-" & MID(rsXL(6), 5, 2) & "-" & MID(rsXL(6), 7, 2)		'출고/반입일
						extOrderserSeq = extMeachulDate & "-" & i  ''Seq에 정산일을 더한다.
						extItemNo = Trim(rsXL(15))		'출고수량
						extOrgOrderserial = ""
						If (extItemNo <= 0) Then
							extitemNo = -1
							IsReturnOrder = True
							extOrgOrderserial = extOrderserial
							extOrderserial = extOrderserial&"-1"
							extItemCost			= rsXL(18)	'반품금액
						Else
							extItemCost			= rsXL(16)	'출고금액
							IsReturnOrder = False
						End If

						' extTenCouponPrice	= CLNG((CLNG(rsXL(23)) - CLNG(rsXL(24))) / extitemNo)	'출고분 쿠폰 부담액(23) - 반품 쿠폰 부담액(24)
						' extOwnCouponPrice	= 0
						' extTenCouponPrice	= 0
						extTenCouponPrice	= CLNG((CLNG(rsXL(26)) - CLNG(rsXL(27))) / extitemNo)	'출고분 쿠폰 부담액 - 반품 쿠폰 부담액..2023-08-01 김진영 추가
						extOwnCouponPrice	= CLNG((CLNG(rsXL(23))) / extitemNo)	'당사부담 할인액(B)(23)
						extTenMeachulPrice	= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice		= CLng(extTenMeachulPrice)
						extJungsanType		= "C"
						extCommPrice		= CLNG((CLNG(rsXL(22)) - CLNG(rsXL(23))) / extitemNo)	'수수료금액(A)(22) - 당사부담 할인액(B)(23)
						extTenJungsanPrice	= CLNG(rsXL(28) / extItemNo * 100)/100		'실지급금액(28)
						extTenJungsanPrice	= extReducedPrice - extCommPrice
						extItemName			= html2db(rsXL(10))			'상품명
						extItemOptionName	= html2db(rsXL(11))			'옵션명
						extVatYN = "Y"
						If rsXL(0) = "정발행(과세)" OR rsXL(0) = "역발행(과세)" Then	'구분
							extVatYN = "Y"
						Else
							extVatYN = "N"
						End If
						extitemid = rsXL(13)	'업체상품코드
					Else
						IsValidInput = False
					End If
				Case "boriboribeasongpay"
					'// 보리보리 배송비 상세내역
					extOrderserial = rsXL(7)
					If Len(rsXL(7)) = 12 Then
						extMeachulDate = LEFT(rsXL(0), 4) & "-" & MID(rsXL(0), 5, 2) & "-" & MID(rsXL(0), 7, 2)
						extJungsanDate = ""
						extVatYN = "Y"
						extOrgOrderserial = ""
						extItemNo				= 1
						extItemName				= "배송비"
						extOrderserSeq			= "D-"&i
						extItemCost				= CLNG(rsXL(8)) + CLNG(rsXL(11))	'주문배송비(A) + 주문 도서산간 추가비
						If (rsXL(14) <> 0) Then		'반품배송비
							extItemNo = 1
							extItemName			= "반품배송비"
							extItemCost			= CLNG(rsXL(12)) + CLNG(rsXL(13)) + CLNG(rsXL(14)) + CLNG(rsXL(15))		'취소배송비 + 취소 도서산간 추가비 + 반품배송비 + 반품 도서산간 추가비
							extOrderserSeq		= "DD-"&i
						ElseIf (rsXL(12) <> 0) Then		'반품배송비
							extItemNo = 1
							extItemName			= "취소배송비"
							extItemCost			= CLNG(rsXL(12)) + CLNG(rsXL(13))		'취소배송비 + 취소 도서산간 추가비
							extOrderserSeq		= "DDD-"&i
						End If
						extOwnCouponPrice		= 0
						extTenCouponPrice		= rsXL(9)		'업체부담배송쿠폰(B)
						extJungsanType			= "D"
						extCommPrice			= 0
						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extTenJungsanPrice		= extReducedPrice - extCommPrice
						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0
						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0
						''금액이 없으면 넣지 않는다.
						If (extItemCost=0) Then
							extItemNo = 0
						End If
					Else
						IsValidInput = False
					End If
				Case "wetoo1300k"
					extOrderserial			= rsXL(0)

					if (Len(extOrderserial) = 15) then
						extMeachulDate = LEFT(rsXL(4), 4) & "-" & MID(rsXL(4), 5, 2) & "-" & MID(rsXL(4), 7, 2)
						extJungsanDate = ""
						extOrderserSeq			= extMeachulDate&"-"&i  ''Seq에 정산일을 더한다.
						extItemNo				= Trim(rsXL(8))

						if (extItemNo <= 0) then
							IsReturnOrder = True
						else
							IsReturnOrder = False
						end if

						extItemCost				= rsXL(7)	'상품단가
						extOwnCouponPrice		= CLNG(CLNG(rsXL(12)) / extItemNo) 'rsXL(12)	'쿠폰금액
						extTenCouponPrice		= CLNG(CLNG(rsXL(13)) / extItemNo) 'rsXL(13)	'쿠폰금액 업체부담

						extOwnCouponPrice = 0	'2022-05-03 김진영 수정..정산에 쿠폰금액이 있는 데, 매출액 / 수수료 모두 쿠폰금액 74,716 차감되어 있는 최종금액 입니다...

						extTenMeachulPrice		= extItemCost-extOwnCouponPrice-extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extJungsanType			= "C"

						extCommPrice			= CLNG(CLNG(rsXL(11)) / extItemNo)
						extTenJungsanPrice		= extReducedPrice - extCommPrice

						extItemName				= html2db(rsXL(6))
						If Instr(extItemName, "// {옵션}") > 0 Then
							extItemOptionName		= Split(extItemName, "// {옵션}")(1)
						Else
							extItemOptionName = ""
						End If
						extVatYN = "Y"
						extitemid = ""
					else
						IsValidInput = False
					end if
				Case "wetoo1300kbeasongpay"
					extOrderserial			= rsXL(0)
					if Len(rsXL(0)) = 15 then
						extMeachulDate = LEFT(rsXL(4), 4) & "-" & MID(rsXL(4), 5, 2) & "-" & MID(rsXL(4), 7, 2)
						extJungsanDate = ""
						extVatYN = "Y"
						extOrgOrderserial = ""

						extItemNo				= 1
						extItemName				= "배송비"
						extOrderserSeq			= "D-"&i ''rsXL(3)&"-"&rsXL(4)
						extItemCost				= rsXL(7)

						if (extItemCost<0) then
							extItemNo = -1
							extItemCost = extItemCost*-1
							extOrderserSeq		= "DD-"&i
						end if
						extReducedPrice			= extItemCost
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extJungsanType			= "D"
						extCommPrice			= 0
						extTenMeachulPrice		= extItemCost
						extTenJungsanPrice		= extItemCost

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

						''금액이 없으면 넣지 않는다.
						if (extItemCost=0) then
							extItemNo = 0
						end if
					else
						IsValidInput = False
					end if
				Case "shintvshopping"
					extOrderserial = Replace(rsXL(1), "-", "")						'주문번호
					if (Len(extOrderserial) = 14) then
						extJungsanDate = ""
						extMeachulDate = Replace(rsXL(3), "/", "-")					'매출일자
						extOrderserSeq = extMeachulDate&"-"&i  						'Seq에 정산일을 더한다.
						extItemNo = Trim(rsXL(24))									'주문수량

						If (extItemNo <= 0) Then
							extItemCost				= rsXL(18)*-1					'판매가
							IsReturnOrder = True
						Else
							extItemCost				= rsXL(18)						'판매가
							IsReturnOrder = False
						End If

						If (extItemNo = 0) Then
							extItemNo = 1
						End If

						extOwnCouponPrice		= CLNG(CLNG(rsXL(29)) / extItemNo) + CLNG(CLNG(rsXL(30)) / extItemNo) '신세계티비쇼핑 계 + 제휴 계
						extTenCouponPrice		= CLNG(CLNG(rsXL(31)) / extItemNo)	'업체 계
						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extJungsanType			= "C"
						extCommPrice			= CLNG(CLNG(rsXL(36)) / extItemNo)
						extTenJungsanPrice		= extReducedPrice - extCommPrice
						extItemName				= html2db(rsXL(14))
						extItemOptionName		= html2db(rsXL(16))
						If rsXL(17) = "과세" Then									'과세구분
							extVatYN = "Y"
						Else
							extVatYN = "N"
						End If
                        extitemid				= rsXL(13)							'상품코드
                        extitemoption			= rsXL(15)							'단품코드
					else
						IsValidInput = False
					End If
				Case "shintvshoppingbeasongpay"
					extOrderserial			= rsXL(1)				'주문번호

					if (Len(extOrderserial) = 14) then

						'// 신세계TV쇼핑 배송비의 매출일이 넘어오지 않음..주문번호로 잘라서 쓰기..그래도 매출일과 차이는 있다.
						extMeachulDate = LEFT(rsXL(1), 4) & "-" & MID(rsXL(1), 5, 2) & "-" & MID(rsXL(1), 7, 2)
						extJungsanDate = ""

						If (extMeachulMonth <> LEFT(extMeachulDate,7)) Then
							extMeachulDate = extMeachulMonth+"-01"
						End If

						extVatYN = "Y"

						extOrgOrderserial = ""

						extItemNo				= 1
						extItemName				= "배송비"
						extOrderserSeq			= "D-"&i 

						extItemCost				= rsXL(6)						'배송비
						if (extItemCost<0) then
							extItemNo = -1
							extOrderserSeq		= "DD-"&i
							extItemCost = extItemCost * -1
						end if

						extReducedPrice			= extItemCost
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extJungsanType			= "D"
						extCommPrice			= 0
						extTenMeachulPrice		= extItemCost
						extTenJungsanPrice		= extItemCost

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

						''금액이 없으면 넣지 않는다.
						if (extItemCost=0) then
							extItemNo = 0
						end if
					else
						IsValidInput = False
					end if
				Case "goodwearmall10"
					extOrderserial = Replace(rsXL(0), "-", "")						'주문번호
					If (Len(extOrderserial) = 17) Then

					Else
						IsValidInput = False
					End If
				Case "goodwearmall10beasongpay"
					extOrderserial = Replace(rsXL(4), "-", "")						'주문번호
					If (Len(extOrderserial) = 17) Then

					Else
						IsValidInput = False
					End If
				Case "skstoa"
					extOrderserial = Replace(rsXL(3), "-", "")						'주문번호
					if (Len(extOrderserial) = 14) then
						extJungsanDate = ""
						extMeachulDate = Replace(rsXL(11), "/", "-")				'완료일
						extOrderserSeq = extMeachulDate&"-"&i  						'Seq에 정산일을 더한다.
						extItemNo = Trim(rsXL(15))									'주문수량
						extItemCost				= rsXL(13)							'판매가
						If (extItemNo <= 0) Then
							IsReturnOrder = True
						Else
							IsReturnOrder = False
						End If

						If (extItemNo = 0) Then
							extItemNo = 1
						End If

						extOwnCouponPrice		= CLNG(CLNG(rsXL(30)) / extItemNo) 	'당사프로모션
						extTenCouponPrice		= CLNG(CLNG(rsXL(29)) / extItemNo)	'업체프로모션
						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extJungsanType			= "C"
						extCommPrice			= CLNG(CLNG(rsXL(35)) / extItemNo)	'위수탁수수료(합계)
						extTenJungsanPrice		= extReducedPrice - extCommPrice
						extItemName				= html2db(rsXL(7))
						extItemOptionName		= html2db(rsXL(8))
						If rsXL(10) = "과세" Then									'과세구분
							extVatYN = "Y"
						Else
							extVatYN = "N"
						End If
                        extitemid				= rsXL(5)							'상품코드
					else
						IsValidInput = False
					End If
				Case "skstoabeasongpay"
					extOrderserial			= rsXL(3)				'주문번호

					if (Len(extOrderserial) = 14) then

						'// skstoa 배송비의 매출일이 넘어오지 않음..주문번호로 잘라서 쓰기..그래도 매출일과 차이는 있다.
						extMeachulDate = LEFT(rsXL(3), 4) & "-" & MID(rsXL(3), 5, 2) & "-" & MID(rsXL(3), 7, 2)
						extJungsanDate = ""

						If (extMeachulMonth <> LEFT(extMeachulDate,7)) Then
							extMeachulDate = extMeachulMonth+"-01"
						End If

						extVatYN = "Y"

						extOrgOrderserial = ""

						extItemNo				= 1
						extItemName				= "배송비"
						extOrderserSeq			= "D-"&i 

						extItemCost				= rsXL(7)						'배송비
						if (extItemCost<0) then
							extItemNo = -1
							extOrderserSeq		= "DD-"&i
							extItemCost = extItemCost * -1
						end if

						extReducedPrice			= extItemCost
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extJungsanType			= "D"
						extCommPrice			= 0
						extTenMeachulPrice		= extItemCost
						extTenJungsanPrice		= extItemCost

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

						''금액이 없으면 넣지 않는다.
						if (extItemCost=0) then
							extItemNo = 0
						end if
					else
						IsValidInput = False
					end if
				Case "WMP", "wmpfashion"
					extOrderserial			= rsXL(1+1) ''배송번호(이게 주문번호)
					if (Len(extOrderserial) = 8 or Len(extOrderserial) = 9) and IsNumeric(rsXL(12+1)) then

						extJungsanDate = ""
						extMeachulDate			= rsXL(28) ''배송일
						extOrderserSeq			= rsXL(2+1)  ''주문번호(Seq로 쓰자.)
						'extOrderserSeq			= rsXL(5+1)  ''옵션주문번호.. 2020-10-28 김진영 / 옵션주문번호로 변경
						If rsXL(5+1) <> "0" then
							extOrderserSeq = extOrderserSeq & "-" & rsXL(5+1)
						End If

						if (rsXL(11+1) <> "배송") then  '' 환불
							'// 반품
						 	extOrgOrderserial		= extOrderserial
						 	extOrderserSeq			= extOrderserSeq & "-1"

							extMeachulDate			= rsXL(29) ''환불완료일

							''환불완료일이 없는케이스가 있다. (배송번호 : 19906380)
							if (extMeachulDate="") then
								extOrderserSeq			= extOrgOrderserial & "-2"
								extMeachulDate			= rsXL(28)
								if (extMeachulDate="") then				''배송완료일도 없으면.  또는 환불완료일/배송완료일이 없는경우 상품번호로 찾아서 정산일을 배송완료일에 넣자.(엑셀에는없음.)
									extMeachulDate			= rsXL(27)  ''일단 결제완료일로 넣자.
'									2020-09-01 김진영..배송번호 : 151652292 8월정산중인데, 위 처럼 결제완료일로 넣을 시 7월이라서 8월 매출에 안 잡힘..결
'									8월로 넣을 수 있는 어떤 껀덕지도 발견이 안 됨..WMP에서는 8월 18일로 매출일이 잡힘..난감 그 자체..엑셀에 강제로 8월18일 집어넣어서 해결
'									추후 이처럼 발생될 일이 농후할 듯 함..
								end if
							end if
						else
						 	'// 정상출고

						 	extOrgOrderserial		= ""

						end if

						if (LEN(extMeachulDate)=8) then
							extMeachulDate = LEFT(extMeachulDate,4)&"-"&MID(extMeachulDate,5,2)&"-"&MID(extMeachulDate,7,2)
						end if

						extVatYN = "Y"  ''구분없음.
						' if (rsXL(23+1) = "면세") then
						' 	extVatYN = "N"
						' end if

						extItemNo				= CLNG(replace(replace(rsXL(12+1),",","")," ",""))  ''1,260

						extItemCost = 0
						extReducedPrice = 0
						extOwnCouponPrice = 0
						extTenCouponPrice = 0
						extTenJungsanPrice = 0
						extCommPrice = 0
						extTenMeachulPrice = 0

						If (rsXL(11+1) <> "배송") Then
							If extItemNo = "0" Then
								If rsXL(18+1) > 0 Then
									extItemNo = 1
								Else
									extItemNo = -1
								End If
							End If
						End If

						If (rsXL(11+1) <> "배송") and (rsXL(12+1)) = "0" then
							extItemCost				= 0
							extOwnCouponPrice		= CLNG((CLNG(rsXL(18+1))+CLNG(rsXL(21+1)))/extItemNo*100)/100 '' 위메프부담 상품쿠폰 , 위메프부담 장바구니쿠폰 추가 2019/06/02
							extTenCouponPrice		= CLNG(CLNG(rsXL(20+1))/extItemNo*100)/100  '' 판매업체부담 상품쿠폰
							extJungsanType			= "C"

							extTenMeachulPrice		= extItemCost-extOwnCouponPrice-extTenCouponPrice
							extReducedPrice			= CLNG(extTenMeachulPrice)

							extCommPrice			= CLNG(CLNG(rsXL(15+1))/extItemNo*100)/100 - extOwnCouponPrice + CLNG(CLNG(rsXL(17+1))/extItemNo*100) / 100 + CLNG(CLNG(rsXL(16+1))/extItemNo*100) / 100 ''판매대행수수료-위메프부담쿠폰
							extTenJungsanPrice		= extTenMeachulPrice-extCommPrice

							extCommSupplyVatPrice	= 0
							extCommSupplyPrice		= 0

							extTenMeachulSupplyVatPrice	= 0
							extTenMeachulSupplyPrice	= 0

							extitemID = rsXL(3+1)
							tenitemid = rsXL(30)
						elseif (replace(replace(rsXL(12+1),",","")," ","") <> 0) then
							extItemCost				= CLNG((CLNG(rsXL(8+1)) + CLNG(rsXL(10))) *100) / 100  ''상품판매가 + 옵션가
							extOwnCouponPrice		= CLNG((CLNG(rsXL(18+1))+CLNG(rsXL(21+1)))/extItemNo*100)/100 '' 위메프부담 상품쿠폰 , 위메프부담 장바구니쿠폰 추가 2019/06/02
							extTenCouponPrice		= CLNG(CLNG(rsXL(20+1))/extItemNo*100)/100  '' 판매업체부담 상품쿠폰
							extJungsanType			= "C"

							extTenMeachulPrice		= extItemCost-extOwnCouponPrice-extTenCouponPrice
							extReducedPrice			= CLNG(extTenMeachulPrice)


							extCommPrice			= CLNG(CLNG(rsXL(15+1))/extItemNo*100)/100 - extOwnCouponPrice + CLNG(CLNG(rsXL(17+1))/extItemNo*100) / 100 + CLNG(CLNG(rsXL(16+1))/extItemNo*100) / 100 ''판매대행수수료-위메프부담쿠폰
							extTenJungsanPrice		= extTenMeachulPrice-extCommPrice

							extCommSupplyVatPrice	= 0
							extCommSupplyPrice		= 0



							extTenMeachulSupplyVatPrice	= 0
							extTenMeachulSupplyPrice	= 0

							extitemID = rsXL(3+1)
							tenitemid = rsXL(30)
						end if
					else
						IsValidInput = False
					end if
				Case "WMPbeasongpay", "wmpfashionbeasongpay"
					extOrderserial			= rsXL(0)

					if (Len(extOrderserial) = 8 or Len(extOrderserial) = 9) and IsNumeric(rsXL(0)) then

						extJungsanDate = ""
						extMeachulDate			= rsXL(16)

						extJungsanType			= "D"

						if (extMeachulDate="") then
							extMeachulDate = rsXL(17)
						end if

						if (LEN(extMeachulDate)=8) then
							extMeachulDate = LEFT(extMeachulDate,4)&"-"&MID(extMeachulDate,5,2)&"-"&MID(extMeachulDate,7,2)
						end if

						extItemNo				= 1
						if (rsXL(11+1)<0) then
							extItemNo = -1
						end if

						if (rsXL(3+1) = "배송") then
							'// 정상출고
							extOrderserSeq			= "D"
							extOrgOrderserial		= ""
							extItemName				= "배송비"
							extItemCost				= CLNG(rsXL(11+1)/extItemNo*100)/100
						elseif (rsXL(3+1) = "환불") then
							'// 반품
							extOrderserSeq			= "D-"&rsXL(1)
							extOrgOrderserial		= extOrderserial
							extItemName				= "반품배송비"
'위메프w패션은 아래 주석으로 extItemCost 이 값을 구하는 데, 위메프는 아닌듯/ 확인필요..2020-11-03 김진영
'							extItemCost				= CLNG(rsXL(10+1)/extItemNo*100)/100
							extItemCost				= CLNG(rsXL(11+1)/extItemNo*100)/100
						else
							'// ??
							extOrderserSeq			= "DD-"&rsXL(1)
							extOrgOrderserial		= extOrderserial
							extItemName				= rsXL(3+1)
							extItemCost				= CLNG(rsXL(11+1)/extItemNo*100)/100
						end if

						extVatYN = "Y"
						extReducedPrice			= extItemCost
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0

						extTenJungsanPrice		= extItemCost

						extCommPrice			= CLNG(extReducedPrice-extTenJungsanPrice)
						extCommSupplyVatPrice	= 0
						extCommSupplyPrice		= 0

						extTenMeachulPrice		= extReducedPrice
						extReducedPrice			= CLNG(extReducedPrice)
						extTenMeachulSupplyVatPrice	= 0
						extTenMeachulSupplyPrice	= 0

						if (extMeachulDate="") and (extReducedPrice=0) and (extTenJungsanPrice=0) then
							extItemNo = 0
						end if
					else
						IsValidInput = False
					end if
				Case "hmallproduct"
					'// --------------------------------------------------------
					'// hmall 정산 상세내역
					extOrderserial			= replace(rsXL(2),"-","") '' 대시 제거
					''2019/11/03 (19) ''카드사부담액. => 2020/03/03 에누리에포함하지 않음.
					if (Len(extOrderserial) = 14) and IsNumeric(rsXL(14)) then

						extJungsanDate = ""
						extMeachulDate			= rsXL(1)
						extOrderserSeq			= rsXL(3)

						'extOrderserSeq 	= extOrderserSeq & "-" & rsXL(5)
						sqlStr = ""
						sqlStr = sqlStr & " SELECT extOrderserSeq FROM db_jungsan.dbo.tbl_xSite_JungsanData "
						sqlStr = sqlStr & " WHERE extOrderserial = '"& extOrderserial &"' "
						sqlStr = sqlStr & " and extOrderserSeq = '"& extOrderserSeq &"' "
						sqlStr = sqlStr & " and sellsite = 'hmall1010' "
						rsget.CursorLocation = adUseClient
						rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
						if Not rsget.Eof then
							extOrderserSeq = extOrderserSeq & "-A"
						end if
						rsget.Close

						sqlStr = ""
						sqlStr = sqlStr & " SELECT extOrderserSeq FROM db_temp.dbo.tbl_xSite_Jungsantmp "
						sqlStr = sqlStr & " WHERE extOrderserial = '"& extOrderserial &"' "
						sqlStr = sqlStr & " and extOrderserSeq = '"& extOrderserSeq &"' "
						sqlStr = sqlStr & " and sellsite = 'hmall1010' "
						rsget.CursorLocation = adUseClient
						rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
						if Not rsget.Eof then
							extOrderserSeq = extOrderserSeq & "-AA"
						end if
						rsget.Close

						' if (Right(extOrderserial, 3) = "001") then
						' 	'// 정상출고
						' 	extOrderserSeq			= extOrderserial
						' 	extOrgOrderserial		= ""
						' 	extOrderserial			= Left(extOrderserial, 14)
						' else
						' 	'// 반품
						' 	extOrderserSeq			= extOrderserial
						' 	extOrgOrderserial		= Left(extOrderserial, 14)
						' 	extOrderserial			= Left(extOrderserial, 14) ''& "-" & i
						' end if

						extVatYN = "Y"
						if (rsXL(23+1) = "면세") then
							extVatYN = "N"
						end if

						extItemNo				= rsXL(14)  '' 0인 CASE가 있음. 반품인듯.

						extItemCost = 0
						extReducedPrice = 0
						extOwnCouponPrice = 0
						extTenCouponPrice = 0
						extTenJungsanPrice = 0
						extCommPrice = 0
						extTenMeachulPrice = 0

						if (extItemNo<>0) then
							extItemCost				= CLNG(rsXL(8)/extItemNo*100)/100  ''판매금액
							extOwnCouponPrice		= CLNG(rsXL(16)/extItemNo*100)/100  + CLNG(rsXL(18)/extItemNo*100)/100 ''// + CLNG(rsXL(19)/extItemNo*100)/100 ''현대부담. 제휴사부담. '카드사부담.
							extTenCouponPrice		= CLNG(rsXL(17)/extItemNo*100)/100  '' 협력사부담.
							extJungsanType			= "C"

							extTenMeachulPrice		= extItemCost-extOwnCouponPrice-extTenCouponPrice
							extReducedPrice			= CLNG(extTenMeachulPrice)


							extCommPrice			= CLNG(rsXL(21+1)/extItemNo*100)/100  ''최종수수료
							extTenJungsanPrice		= extTenMeachulPrice-extCommPrice

							extCommSupplyVatPrice	= 0
							extCommSupplyPrice		= 0



							extTenMeachulSupplyVatPrice	= 0
							extTenMeachulSupplyPrice	= 0
						end if
					else
						IsValidInput = False
					end if
				Case "hmallbeasongpay"
					extOrderserial			= replace(rsXL(3),"-","") '' 대시 제거

					if (Len(extOrderserial) = 14) and IsNumeric(rsXL(0)) then

						extJungsanDate = ""
						extMeachulDate			= CStr(dateadd("d",-1,dateadd("m",1,rsXL(1)+"-01")))

						extJungsanType			= "D"

						if (rsXL(2) = "주문") then
							'// 정상출고
							extOrderserSeq			= "D"
							extOrgOrderserial		= ""
							extItemName				= "배송비"
						elseif (rsXL(2) = "반품") then
							'// 반품
							extOrderserSeq			= "D-2"
							extOrgOrderserial		= ""
							extItemName				= "반품배송비"
						elseif (rsXL(2) = "취소(주문)") then
							'// 반품
							extOrderserSeq			= "D-3"
							extOrgOrderserial		= ""
							extItemName				= "배송비"
						elseif (rsXL(2) = "기타") then
							'// 반품
							extOrderserSeq			= "D-4"
							extOrgOrderserial		= ""
							extItemName				= "배송비"
						else
							'// 취소주문
							extOrderserSeq			= "D-1"
							extOrgOrderserial		= extOrderserial
							extItemName				= rsXL(2)
						end if

						extVatYN = "Y"

						extItemNo				= 1
						if (rsXL(5)<0) then
							extItemNo = -1
						end if

						extItemCost				= CLNG(rsXL(5)/extItemNo*100)/100
						extReducedPrice			= extItemCost
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0

						extTenJungsanPrice		= extItemCost

						extCommPrice			= CLNG(extReducedPrice-extTenJungsanPrice)
						extCommSupplyVatPrice	= 0
						extCommSupplyPrice		= 0

						extTenMeachulPrice		= extReducedPrice
						extReducedPrice			= CLNG(extReducedPrice)
						extTenMeachulSupplyVatPrice	= 0
						extTenMeachulSupplyPrice	= 0
					else
						IsValidInput = False
					end if
				Case "lotteimall"
					'// --------------------------------------------------------
					'// 롯데아이몰 정산 상세내역
					extOrderserial			= rsXL(1)

					if (Len(extOrderserial) = 14) and IsNumeric(rsXL(7)) then

						'// 정산내역에 매출일자 없다. 따라서 가장큰 주문번호에서 매출월을 만들어낸다.
						'// (ex, 20140220H10035 -> 2014-02-28)
						'extMeachulDate = "0000-00-00"
						extMeachulDate = requestCheckvar(uprequest("extMeachulDate"),10)

						extJungsanDate = ""

						extOrderserSeq			= CStr(rsXL(2)) & "-" & CStr(rsXL(3))

						extItemNo				= rsXL(7)
						''4017377 교환비
						'// 원주문-반품주문 구분법 : 상품 = 수량이 0 보다 작은 경우 반품!!, 배송비 = 반품비(4017388) or 이전 상품주문이 반품인 경우 다음 배송비
						if (rsXL(2) = 4017357) or (rsXL(2) = 4017388) or (rsXL(2) = 4017377) then
							if (rsXL(2) = 4017377) then
								extOrderserSeq = extOrderserSeq &"-"&CStr(rsXL(0)) ''2019/11/01 추가
							elseif (rsXL(2) = 4017388) or (extItemNo <= 0) then
								extOrderserial 		= rsXL(1) & "-1"
								extOrgOrderserial	= rsXL(1)

								sqlStr = ""
								sqlStr = sqlStr & " SELECT COUNT(*) as cnt FROM db_temp.dbo.tbl_xSite_JungsanTmp "
								sqlStr = sqlStr & " WHERE extOrderserial = '"& extOrderserial &"' "
								sqlStr = sqlStr & " and extOrderserSeq = '"& extOrderserSeq &"' "
								rsget.CursorLocation = adUseClient
								rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
								If rsget("cnt") > 0 Then
									extOrderserSeq 	= extOrderserSeq & "-" & rsXL(0)
								End If
								rsget.Close
							else
								extOrgOrderserial	= ""
							end if
						else
							if (extItemNo <= 0) then
								IsReturnOrder = True
								extOrderserial = rsXL(1) & "-1"
								extOrgOrderserial		= rsXL(1)
							else
								IsReturnOrder = False
								extOrgOrderserial		= ""
							end if
						end if

						extItemCost				= 0
						extReducedPrice			= rsXL(8)
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0

						if (rsXL(2) = 4017357) or (rsXL(2) = 4017388) or (rsXL(2) = 4017377) then
							extJungsanType			= "D"
						else
							extJungsanType			= "C"
						end if

						extCommPrice			= rsXL(10)
						extTenMeachulPrice		= rsXL(8)
						extTenJungsanPrice		= rsXL(9)

						extItemName				= html2db(rsXL(4))
						extItemOptionName		= html2db(rsXL(5))

						extVatYN = "Y"
						if (rsXL(6) = "면세") then
							extVatYN = "N"
						end if

						if (extItemNo<>0) then
							extCommSupplyPrice		= CLNG(extCommPrice/extItemNo*100)/100
							extCommSupplyVatPrice	= 0

							extTenMeachulSupplyPrice	= CLNG(extTenMeachulPrice/extItemNo*100)/100
							extTenMeachulSupplyVatPrice	= 0

							if (extVatYN="Y") then

								extTenMeachulPrice = CLNG(getLtiMallRound(extTenMeachulPrice*1.1)/extItemNo*100)/100
								'extCommPrice	   = CLNG(getLtiMallRound(extCommPrice*1.1)/extItemNo*100)/100
								'extTenJungsanPrice = extTenMeachulPrice-extCommPrice
								extTenJungsanPrice = CLNG(getLtiMallRound(extTenJungsanPrice*1.1)/extItemNo*100)/100
								extCommPrice       = extTenMeachulPrice-extTenJungsanPrice
								extReducedPrice	   = CLNG(extTenMeachulPrice)
							else
								extTenMeachulPrice = CLNG(extTenMeachulPrice/extItemNo*100)/100
								'extCommPrice	   = CLNG(extCommPrice/extItemNo*100)/100
								'extTenJungsanPrice = extTenMeachulPrice-extCommPrice
								extTenJungsanPrice = CLNG(extTenJungsanPrice/extItemNo*100)/100
								extCommPrice	   = extTenMeachulPrice-extTenJungsanPrice
								extReducedPrice	   = CLNG(extTenMeachulPrice)
							end if

							extItemCost = extTenMeachulPrice
						end if
						extitemid = rsXL(2)
					else
						IsValidInput = False
					end if
				Case "LFmall"
					'// --------------------------------------------------------
					'// LFmall 정산 상세내역
					''2020/03/10 LF부담 배송비금액 추가됨(15)
					extOrderserial			= rsXL(1)		'주문번호

					if (Len(extOrderserial) = 8) and IsNumeric(Trim((rsXL(11)))) then	'rsXL(11) : 판매수량

						'// 정산내역에 매출일자 없다. 따라서 가장큰 주문번호에서 매출월을 만들어낸다. like lotteimall
						'// (ex, 20140220H10035 -> 2014-02-28)
						'extMeachulDate = "0000-00-00"
						lfmallEtcPrice = CLNG(rsXL(27))		'정산조정금액
						extMeachulDate = requestCheckvar(uprequest("extMeachulDate"),10)
						extOrderserSeq			= extMeachulDate&"-"&i  ''Seq에 정산일을 더한다.
						extJungsanDate = ""
						extItemNo				= Trim(rsXL(11))	'판매수량

						If extItemNo = "0" AND rsXL(10) = "배송비" Then	'정산구분
							extItemNo = 1
						End If

						if (extItemNo <= 0) then
							IsReturnOrder = True
							'extOrderserial 		= rsXL(1) & "-1"
							'extOrgOrderserial	= rsXL(1)
						else
							IsReturnOrder = False
							'extOrgOrderserial		= ""
						end if

						extItemCost				= CLNG(CLNG(rsXL(13)) / extItemNo)		'rsXL(13) : 판매금액
						extOwnCouponPrice		= CLNG(CLNG((rsXL(18)) + CLNG(rsXL(20))) / extItemNo)		'LF부담금액(쿠폰) + LF부담금액(마일리지)
						extTenCouponPrice		= CLNG(CLNG((rsXL(16)) + CLNG(rsXL(17))) / extItemNo) 	'업체부담금액(쿠폰) + 업체부담금액(EGM)
						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= extTenMeachulPrice
						extJungsanType			= "C"
						extCommPrice			= (CLNG(CLNG(rsXL(14)) * -1 / extItemNo) - extOwnCouponPrice)	'rsXL(14) : 수수료율


						extTenJungsanPrice		= extReducedPrice - extCommPrice
						extItemName				= html2db(rsXL(7))	'상품명
						extItemOptionName		= html2db(rsXL(6))	'사이즈
						extVatYN = "Y"
						extitemid = rsXL(5)	'상품코드
						dvlprice = CLNG(rsXL(21)) + CLNG(rsXL(22)) + CLNG(rsXL(23)) + CLNG(rsXL(24))		'배송비금액 (기본 + 반품 + 교환) '2023-03-02 김진영 : 추가배송비(rsXl(23)) 추가
					else
						IsValidInput = False
					end if
				Case "auction1010beasongpay"
					'// --------------------------------------------------------
					'// 옥션 배송비 정산 상세내역
					extOrderserial			= rsXL(1) '' 결제번호

					dim chkTypeStr : chkTypeStr = rsXL(12) ''2019/08/13 수정 (9=>10)
					dim auctionxltype : auctionxltype=0
					if (chkTypeStr="환불금차감") or (chkTypeStr="종류") then
						auctionxltype=1
					end if

					if (auctionxltype=0) then
						extItemName = rsXL(13)
					else
						extItemName = rsXL(12)
					end if

					if (Len(extOrderserial) = 10) and ((extItemName = "최초배송비") or (extItemName = "반품배송비") or (extItemName = "교환배송비") or (extItemName = "환불금차감") ) then

						extItemNo = 1
						if (auctionxltype=1) then
							extMeachulDate = Left(rsXL(8), 10)  ''2019/08/13 분기
						else
							extMeachulDate = Left(rsXL(7), 10)  ''매출기준일 == 정산완료일
						end if
						if (extItemName = "환불금차감") then
							if (extMeachulDate>="2019-10-25") then
								extOrderserSeq = CStr(rsXL(2)) & "-DDDD"
							else
								extOrderserSeq = CStr(rsXL(1)) & "-DDDD"
							end if
							extItemCost	   = CLng(rsXL(9)/extItemNo)  ''2019/08/13 수정 (8=>9)
							extCommPrice			= CLng(rsXL(10)/extItemNo)
						elseif (extItemName = "교환배송비") then
							if (rsXL(6)="") or isEmpty(rsXL(6)) or isNULL(rsXL(6)) then
								extOrderserSeq = CStr(rsXL(1)) & "-DD"
							else
								extOrderserSeq = CStr(rsXL(1)) & "-DD2"
							end if

							if (extItemName <> "교환배송비") then
								extOrderserSeq = extOrderserSeq & "D"

								if (extItemName="반품배송비") then
									extOrderserSeq = extOrderserSeq & "D"
								end if
							end if

							extItemCost				= CLng(rsXL(10)/extItemNo)
							extCommPrice			= CLng(rsXL(11)/extItemNo)
						else
							''일반배송비
							if (rsXL(6)="") or isEmpty(rsXL(6)) or isNULL(rsXL(6)) then
								extOrderserSeq = CStr(rsXL(1)) & "-D"
							else
								extOrderserSeq = CStr(rsXL(1)) & "-D2"
							end if

							if (rsXL(9) <> "") then
								'// 반품
								extItemNo = -1
							end if

							extItemCost				= CLng(rsXL(10)/extItemNo)
							extCommPrice			= CLng(rsXL(11)/extItemNo)
						end if



						extJungsanDate = ""


						extReducedPrice			= extItemCost
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extJungsanType			= "D"

'						extCommPrice			= 0
						extTenMeachulPrice		= extReducedPrice
						extTenJungsanPrice		= extTenMeachulPrice-extCommPrice

						''extItemName				= replace(extItemName,"국내","")
						extItemOptionName		= ""

						extVatYN = "Y"

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0
					else
						IsValidInput = False
					end if

				Case "auction1010"
					'// --------------------------------------------------------
					'// 옥션 정산 상세내역
					extOrderserial			= rsXL(1) '' 결제번호

					if (Len(extOrderserial) = 10) and IsNumeric(rsXL(16)) then

						extMeachulDate = LEFT(rsXL(11),10) ''매출기준일

						extJungsanDate = ""

						extOrderserSeq = rsXL(2) '' 주문번호

						if rsXL(10) <> "" then  ''환불일
							extOrderserSeq			= extOrderserSeq & "-1"
						end if

						extItemNo = rsXL(16) ''주문수량

						''extItemCost				= ABS(rsXL(17)) + CLNG(CLNG(rsXL(18)) / extItemNo)    '' 주문수량*상품판매가+옥션상품판매가
						extItemCost				= CLNG(CLNG(rsXL(18)+rsXL(17)) / extItemNo)    '' 주문수량*상품판매가+옥션상품판매가 //2019/05/30 포멧 바뀌었음

						''extOwnCouponPrice		= CLng(CLNG(rsXL(19)+rsXL(20)-rsXL(32)-rsXL(33))/ extItemNo*100)/100  ''옥션 상품별 할인 + 옥션 구매자 쿠폰 할인
						''extOwnCouponPrice		= CLng(CLNG(rsXL(19)+rsXL(20)+rsXL(32)+rsXL(33))/ extItemNo*100)/100  '' //2019/05/30 포멧 바뀌었음
						extOwnCouponPrice		= CLng(CLNG(rsXL(19)+rsXL(20))/ extItemNo*100)/100  '' //2019/05/31 재수정
						extTenCouponPrice		= CLng(CLNG(rsXL(21))/ extItemNo*100)/100   '''판매자 할인  :: 판매자 펀딩 상품별 할인(?)
						extJungsanType			= "C"


						''extTenMeachulPrice		= CLng((CLNG(rsXL(22)+rsXL(32)+rsXL(33)) / extItemNo)*100)/100
						''extTenMeachulPrice		= CLng((CLNG(rsXL(22)) / extItemNo)*100)/100 '' //2019/05/30 포멧 바뀌었음
						extTenMeachulPrice		= CLng((CLNG(rsXL(22)+rsXL(32)+rsXL(33)) / extItemNo)*100)/100 '' //2019/05/31 재수정
						extReducedPrice			= CLng(extTenMeachulPrice)

						''extItemCost				= extTenMeachulPrice+extOwnCouponPrice+extTenCouponPrice
						extCommPrice			= CLng((CLNG(rsXL(29)) / extItemNo)*100)/100  ''서비스이용료

						extTenJungsanPrice		= extTenMeachulPrice-extCommPrice  ''CLng((CLng(rsXL(45)) / extItemNo)*100)/100
						''extTenJungsanPrice		= CLng((CLNG(rsXL(26)) / extItemNo)*100)/100


						extItemName				= html2db(rsXL(4))
						extItemOptionName		= ""
						extitemID				= rsXL(3)

						extVatYN = "Y"

						if (rsXL(40)<>"과세") then
							extVatYN = "N"
						end if

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0
					else
						IsValidInput = False
					end if
				Case "auctionbuga"
					'// --------------------------------------------------------
					'// 옥션 정산 상세내역
					extOrderserial			= rsXL(1) '' 결제번호

					if (Len(extOrderserial) = 10) and IsNumeric(rsXL(16)) then

						extMeachulDate = LEFT(rsXL(11),10) ''매출기준일

						extJungsanDate = ""

						extOrderserSeq = rsXL(2) '' 주문번호

						if rsXL(10) <> "" then  ''환불일
							extOrderserSeq			= extOrderserSeq & "-1"
						end if

						extItemNo = rsXL(16) ''주문수량

						''extItemCost				= ABS(rsXL(17)) + CLNG(CLNG(rsXL(18)) / extItemNo)    '' 주문수량*상품판매가+옥션상품판매가
						extItemCost				= CLNG(CLNG(rsXL(18)+rsXL(17)) / extItemNo)    '' 주문수량*상품판매가+옥션상품판매가 //2019/05/30 포멧 바뀌었음

						''extOwnCouponPrice		= CLng(CLNG(rsXL(19)+rsXL(20)-rsXL(32)-rsXL(33))/ extItemNo*100)/100  ''옥션 상품별 할인 + 옥션 구매자 쿠폰 할인
						''extOwnCouponPrice		= CLng(CLNG(rsXL(19)+rsXL(20)+rsXL(32)+rsXL(33))/ extItemNo*100)/100  '' //2019/05/30 포멧 바뀌었음
						extOwnCouponPrice		= CLng(CLNG(rsXL(19)+rsXL(20))/ extItemNo*100)/100  '' //2019/05/31 재수정
						extTenCouponPrice		= CLng(CLNG(rsXL(21))/ extItemNo*100)/100   '''판매자 할인  :: 판매자 펀딩 상품별 할인(?)
						extJungsanType			= "C"


						''extTenMeachulPrice		= CLng((CLNG(rsXL(22)+rsXL(32)+rsXL(33)) / extItemNo)*100)/100
						''extTenMeachulPrice		= CLng((CLNG(rsXL(22)) / extItemNo)*100)/100 '' //2019/05/30 포멧 바뀌었음
						extTenMeachulPrice		= CLng((CLNG(rsXL(22)+rsXL(32)+rsXL(33)) / extItemNo)*100)/100 '' //2019/05/31 재수정
						extReducedPrice			= CLng(extTenMeachulPrice)

						''extItemCost				= extTenMeachulPrice+extOwnCouponPrice+extTenCouponPrice
						extCommPrice			= CLng((CLNG(rsXL(29)) / extItemNo)*100)/100  ''서비스이용료

						extTenJungsanPrice		= extTenMeachulPrice-extCommPrice  ''CLng((CLng(rsXL(45)) / extItemNo)*100)/100
						''extTenJungsanPrice		= CLng((CLNG(rsXL(26)) / extItemNo)*100)/100


						extItemName				= html2db(rsXL(4))
						extItemOptionName		= ""
						extitemID				= rsXL(3)

						extVatYN = "Y"

						if (rsXL(40)<>"과세") then
							extVatYN = "N"
						end if

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0
					else
						IsValidInput = False
					end if
				Case "gmarket1010"
					'// --------------------------------------------------------
					'// 지마켓 정산 상세내역
					'// 20200110 번호필드(0) 추가, 공제/환급금(서비스이용료 미포함)(28) 추가

					extOrderserial			= rsXL(0+1)

					if (Len(extOrderserial) = 10) and IsNumeric(rsXL(2+1)) then

						if (rsXL(9+1) <> "") then
							'// 반품
							extMeachulDate = Left(rsXL(9+1), 10)
						else
							extMeachulDate = Left(rsXL(8+1), 10)  ''배송완료일 하자.
						end if

						extJungsanDate = ""

						if (rsXL(9+1) <> "") then
							extOrderserSeq			= CStr(rsXL(1+1)) & "-1"
						else
							extOrderserSeq			= CStr(rsXL(1+1))
						end if

						'// 원주문-반품주문 구분법 : 환불일자 있는지
						extItemNo = rsXL(14+1)

						extItemCost				= CLng((CLng(rsXL(15+1))+CLng(rsXL(16+1))+CLng(rsXL(17+1))+CLng(rsXL(18+1))) / extItemNo*100)/100  ''판매가격+필수선택상품금액+추가구성상품금액+옵션상품

						extOwnCouponPrice		= CLng((CLng(rsXL(19+1))+CLng(rsXL(20+1)))/ extItemNo*100)/100*-1
						extTenCouponPrice		= CLng((CLng(rsXL(21+1))+CLng(rsXL(36+2))*-1)/ extItemNo*100)/100*-1  ''36판매자 펀딩 구매쿠폰 할인

						extTenMeachulPrice		= CLng(rsXL(22+1)/ extItemNo*100)/100
						extReducedPrice			= CLng(extTenMeachulPrice)
						extJungsanType			= "C"

						extCommPrice			= CLng(rsXL(30+2)/ extItemNo*100)/100 ''서비스이용료

						extTenJungsanPrice		= CLng(rsXL(27+2)/ extItemNo*100)/100 ''판매자 최종정산금

						extItemName				= "" ''html2db(rsXL(3))
						extItemOptionName		= ""
						extitemID 				= (rsXL(2))
						extVatYN = "Y"
						if (rsXL(40+2) = "면세") then
							extVatYN = "N"
						end if

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0
					else
						IsValidInput = False
					end if
				Case "gmarket1010beasongpay"
					'// --------------------------------------------------------
					'// 지마켓 배송비 정산 상세내역
					extOrderserial			= rsXL(0)

					''and IsNumeric(rsXL(2))
					' extItemName = rsXL(19) ''현금,카드,복합결제,모바일

					' ''if NOT ((extItemName="반품배송비") or (extItemName="사용한 원배송비 환불차감") or (extItemName="종류") or (extItemName="무료반품배송비") or (extItemName="기타반품비")) then
					' if (extItemName="현금" or extItemName="카드" or extItemName="복합결제" or extItemName="모바일" or extItemName="알리페이" or extItemName="글로벌결제") then
					' 	extItemName = rsXL(22)
					' end if

					On Error Resume Next
					extItemName = rsXL(23)
					If ERR Then
						extItemName = rsXL(15)
					end if
					On Error Goto 0

					''if (Len(extOrderserial) = 10) and ((extItemName = "국내배송비") or (extItemName = "반품배송비") or (extItemName = "사용한 원배송비 환불차감") or (extItemName="무료반품배송비") or (extItemName="기타반품비")) then
					if (Len(extOrderserial) = 10 ) then
						extItemNo = 1
						'if (extItemName = "반품배송비") or (extItemName = "사용한 원배송비 환불차감") or (extItemName="무료반품배송비") or (extItemName="기타반품비") then
						''rw  extItemName
						if (extItemName<>"국내배송비") and (extItemName<>"추가배송국내") then
							if (isNULL(rsXL(1))) then
								extOrderserSeq = "-DD"
							else
								extOrderserSeq = CStr(rsXL(1)) & "-DD"
							end if


							if (extItemName <> "반품배송비") then
								extOrderserSeq = extOrderserSeq & "D"
							end if

							if (extItemName="무료반품배송비") then
								extOrderserSeq = extOrderserSeq & "F"
							end if

							if (extItemName="기타반품비") then
								extOrderserSeq = extOrderserSeq & "E"
							end if

							if (rsXL(9) <> "") then
								extMeachulDate = Left(rsXL(9), 10) ''배송완료일 하자.
							else
								extMeachulDate = Left(rsXL(8), 10)  ''배송일
							end if

							extItemCost				= CLng(rsXL(14)/extItemNo)
							extCommPrice			= 0
						else
							''일반배송비
							if (isNULL(rsXL(1))) then
								If rsXL(16) = "구매미결정" Then
									extOrderserSeq = rsXL(7) & "-D"
								Else
									extOrderserSeq = "-D"
								End If
							elseif extItemName = "추가배송국내" then
								extOrderserSeq = CStr(rsXL(1)) & "-DDD"
							else
								extOrderserSeq = CStr(rsXL(1)) & "-D"
							end if

							if (rsXL(10) <> "") then
								'// 반품
								extMeachulDate = Left(rsXL(10), 10)
								if (Left(rsXL(9), 10)>extMeachulDate) then extMeachulDate=Left(rsXL(9), 10)  ''환불일,배송완료일중 느린날짜인듯함
								extOrderserSeq	=extOrderserSeq & "-1"
								extItemNo = -1

								'' 입금일이 없으면 환불일 기준인듯함
								if (rsXL(6) = "") then
									extMeachulDate = Left(rsXL(10), 10)
									extOrderserSeq	=extOrderserSeq & "-1"
								end if

								if rsXL(16)="취소" then extOrderserSeq	=extOrderserSeq & "-R"
								if rsXL(14)>0 then extOrderserSeq	=extOrderserSeq & "-1"
							else
								extMeachulDate = Left(rsXL(9), 10)  ''배송완료일 하자.
							end if

							extItemCost				= CLng(rsXL(13)/extItemNo)
							extCommPrice			= CLng(rsXL(14)/extItemNo)
							'rw extOrderserial &":"&extOrderserSeq&":"
						end if



						extJungsanDate = ""


						extReducedPrice			= extItemCost
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extJungsanType			= "D"


						extTenMeachulPrice		= extReducedPrice
					If (extItemName<>"국내배송비") and (extItemName<>"추가배송국내") Then
						extTenJungsanPrice		= extReducedPrice
					Else
						extTenJungsanPrice 		= CLng(rsXL(15)/extItemNo)
					End If


						extItemName				= replace(extItemName,"국내","")
						extItemOptionName		= ""

						extVatYN = "Y"

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0
					else
						IsValidInput = False
					end if
				Case "cjmallproduct"
					'// --------------------------------------------------------
					'// CJ몰 상품정산 상세내역
					extOrderserial			= rsXL(23)		'주문번호

					if IsNumeric(rsXL(11)) and Len(extOrderserial) = 26 then

						extJungsanDate = ""
						extMeachulDate			= Replace(rsXL(1), "/", "-")	'매출일자
						if (Right(extOrderserial, 3) = "001") then
							'// 정상출고
							extOrderserSeq			= extOrderserial
							extOrgOrderserial		= ""
							extOrderserial			= Left(extOrderserial, 14)
						else
							'// 반품
							extOrderserSeq			= extOrderserial
							extOrgOrderserial		= Left(extOrderserial, 14)
							extOrderserial			= Left(extOrderserial, 14) ''& "-" & i
						end if

						extVatYN = "Y"
						if (rsXL(8) = "면세") then	'과세구분
							extVatYN = "N"
						end if

						extItemNo				= rsXL(9)	'매출량

						'// 우선 총액을 입력하고, 총액의 합계금액을 구한 후에 수량으로 나누고 차액을 보정해준다.
						extItemCost				= CLNG(rsXL(11)/extItemNo*100)/100	'원판매금액
						extReducedPrice			= CLNG(rsXL(15)/extItemNo*100)/100	'소비자판매가
						extOwnCouponPrice		= extItemCost - extReducedPrice
						extTenCouponPrice		= 0
						extJungsanType			= "C"

						extTenJungsanPrice		= CLNG(rsXL(20)/extItemNo*100)/100	'상품대금지급예정액(공제전)

						extCommPrice			= CLNG(rsXL(18)/extItemNo*100)/100	'cj합계금액
						extCommSupplyVatPrice	= CLNG(rsXL(17)/extItemNo*100)/100	'cj부가세
						extCommSupplyPrice		= CLNG(rsXL(16)/extItemNo*100)/100	'cj수수료

						extTenMeachulPrice		= extReducedPrice
						extReducedPrice			= CLNG(extReducedPrice)
						extTenMeachulSupplyVatPrice	= CLNG(rsXL(14)/extItemNo*100)/100	'소비자부가세
						extTenMeachulSupplyPrice	= CLNG(rsXL(13)/extItemNo*100)/100	'소비자공급가
					else
						IsValidInput = False
					end if
				Case "cjmallbeasongpay"
					' uprequest("cjbeasongGubun") = "1"		//대금지불현황(교환택배비)
					' uprequest("cjbeasongGubun") = "2"		//대금지불현황(반품택배비)
					' uprequest("cjbeasongGubun") = "3"		//대금지불현황(저단가배송비)
					Dim ccnt
					ccnt = 0
					If uprequest("cjbeasongGubun") = "1" OR uprequest("cjbeasongGubun") = "2" OR uprequest("cjbeasongGubun") = "3" Then
						extOrderserial			= rsXL(4)
						If uprequest("cjbeasongGubun") = "3" Then
							extOrderserial			= rsXL(5)
						End If

						If (Len(extOrderserial) = 14) Then
							If uprequest("cjbeasongGubun") = "1" OR uprequest("cjbeasongGubun") = "2" Then
								extMeachulDate = rsXL(2)  ''처리일자.
							Else
								extMeachulDate = LEFT(rsXL(3),10)  ''처리일자 같은달이 아니다..?
								extMeachulDate = LEFT(extMeachulDate,4) & "-" & MID(extMeachulDate,5,2) & "-" & MID(extMeachulDate,7,2)
							End If

							If (extMeachulMonth <> LEFT(extMeachulDate,7)) then
								extMeachulDate = extMeachulMonth+"-01"
							End If
							extJungsanDate = ""
							extVatYN = "Y"
							extOrderserSeq			= "1-"

							If uprequest("cjbeasongGubun") = "1" Then
								extOrderserSeq = extOrderserSeq + "DDD"
								extItemName = "교환택배비"
							End If

							If uprequest("cjbeasongGubun") = "2" Then
								extOrderserSeq = extOrderserSeq + "DD"
								extItemName = "반품택배비"
							End If

							If uprequest("cjbeasongGubun") = "3" Then
								extOrderserSeq = extOrderserSeq + "D"
								extItemName = "배송비"
								if LEN(rsXL(3)) > 8 then extOrderserSeq = extOrderserSeq + "E" ''저단가배송비 기타

								sqlStr = ""
								sqlStr = sqlStr & " SELECT COUNT(*) as cnt FROM db_temp.dbo.tbl_xSite_JungsanTmp "
								sqlStr = sqlStr & " WHERE extOrderserial = '"& rsXL(5) &"' "
								sqlStr = sqlStr & " and extOrderserSeq = '"& extOrderserSeq &"' "
								sqlStr = sqlStr & " and sellsite='cjmall' "
								rsget.CursorLocation = adUseClient
								rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
								If rsget("cnt") > 0 Then
									ccnt = "1"
									extOrderserSeq 	= extOrderserSeq & "-" & rsXL(0)
								End If
								rsget.Close

								If ccnt <> 1 Then 
									sqlStr = ""
									sqlStr = sqlStr & " SELECT COUNT(*) as cnt FROM db_jungsan.dbo.tbl_xSite_JungsanData "
									sqlStr = sqlStr & " WHERE extOrderserial = '"& rsXL(5) &"' "
									sqlStr = sqlStr & " and extOrderserSeq = '"& extOrderserSeq &"' "
									sqlStr = sqlStr & " and sellsite='cjmall' "
									rsget.CursorLocation = adUseClient
									rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
									If rsget("cnt") > 0 Then
										extOrderserSeq 	= extOrderserSeq & "-" & rsXL(0)
									End If
									rsget.Close
								End If
							End If

							extOrgOrderserial		= ""
							extItemNo				= 1	 ''판매량
							extOwnCouponPrice		= 0
							extTenCouponPrice		= 0
							extItemCost				= CLNG(rsXL(7)*-1)   ''공제액 이므로 -1을 곱하자  /반품택배비, 교환택배비
							If extItemCost < 0 then ''금액이-이면
								extOrderserSeq = extOrderserSeq + "-1"
							End If
							extReducedPrice			= extItemCost - extTenCouponPrice - extOwnCouponPrice
							extJungsanType			= "D"
							extCommPrice			= 0
							extCommSupplyPrice		= extCommPrice
							extCommSupplyVatPrice	= 0
							extTenMeachulPrice			= extReducedPrice
							extTenMeachulSupplyPrice	= extTenMeachulPrice
							extTenMeachulSupplyVatPrice	= 0
							extTenJungsanPrice		= extReducedPrice - extCommPrice  ''== ROUND(rsXL(32) / extItemNo,0)
							extitemid		= 0
							extitemoption	= "0000"

							extItemName = ""
							extItemOptionName = ""
						else
							IsValidInput = False
						end if
					Else

					End If

					' if (sheetName="'대금지급현황(저단가배송비)$'") or (sheetName="'대금지불현황(반품택배비)$'") or (sheetName="'대금지불현황(교환택배비)$'") or (sheetName="'대금지급현황(AS 택배비)$'") or (sheetName="sheet1$")  then
					' 	extOrderserial			= rsXL(4)
					' 	if (sheetName="'대금지불현황(반품택배비)$'") or (sheetName="'대금지불현황(교환택배비)$'") then
					' 		extOrderserial			= rsXL(3)
					' 	end if

					' 	if (sheetName="'대금지급현황(AS 택배비)$'") then
					' 		extOrderserial			= LEFT(rsXL(5),14)
					' 	end if

					' 	if (Len(extOrderserial) = 14)  then
					' 		if (sheetName="'대금지불현황(반품택배비)$'") or (sheetName="'대금지불현황(교환택배비)$'") then
					' 			extMeachulDate = rsXL(1)  ''처리일자.
					' 		else
					' 			extMeachulDate = LEFT(rsXL(2),10)  ''처리일자 같은달이 아니다..?
					' 		end if

					' 		if (extMeachulMonth<>LEFT(extMeachulDate,7)) then
					' 			extMeachulDate = extMeachulMonth+"-01"
					' 		end if

					' 		extJungsanDate = ""

					' 		extVatYN = "Y"


					' 		extOrderserSeq			= "1-"

					' 		if (sheetName="'대금지급현황(저단가배송비)$'") then
					' 			extOrderserSeq = extOrderserSeq + "D"
					' 			extItemName = "배송비"

					' 			if LEN(rsXL(2))>10 then extOrderserSeq = extOrderserSeq + "E" ''저단가배송비 기타
					' 		end if

					' 		if (sheetName="'대금지불현황(반품택배비)$'") then
					' 			extOrderserSeq = extOrderserSeq + "DD"
					' 			extItemName = "반품택배비"
					' 		end if

					' 		if (sheetName="'대금지불현황(교환택배비)$'") then
					' 			extOrderserSeq = extOrderserSeq + "DDD"
					' 			extItemName = "교환택배비"
					' 		end if

					' 		if (sheetName="'대금지급현황(AS 택배비)$'") then
					' 			extOrderserSeq = extOrderserSeq + "DDDD"
					' 			extItemName = "AS택배비"
					' 		end if

					' 		if (sheetName="sheet1$") then  ''[대금정산] 무료배송쿠폰 기타정산으로 올리는게 좋을듯..
					' 			extOrderserSeq = extOrderserSeq + "C"
					' 			extItemName = "배송비쿠폰"

					' 		end if

					' 		' if (p_extOrderserial = extOrderserial) and  (p_extOrderserSeq = extOrderserSeq) then
					' 		' 	extOrderserSeq = extOrderserSeq+"-1"
					' 		' end if

					' 		extOrgOrderserial		= ""

					' 		extItemNo				= 1	 ''판매량

					' 		extOwnCouponPrice		= 0
					' 		extTenCouponPrice		= 0

					' 		if (sheetName="'대금지급현황(AS 택배비)$'") then
					' 			extItemCost = 0
					' 			if rsXL(11)<>"" then
					' 				extItemCost				= CLNG(rsXL(11)*-1)
					' 			end if

					' 			if rsXL(10)<>"" then
					' 				extTenCouponPrice		= CLNG(rsXL(10))
					' 			end if
								
					' 			If rsXL(5) = "20230312052749-001-001-004" Then		'2023-05-02 김진영 추가
					' 				extOrderserSeq = extOrderserSeq + "D"
					' 			End If
					' 		else
					' 			extItemCost				= CLNG(rsXL(6)*-1)   ''공제액 이므로 -1을 곱하자  /반품택배비, 교환택배비
					' 		end if

					' 		if extItemCost<0 then ''금액이-이면
					' 			extOrderserSeq = extOrderserSeq + "-1"
					' 		end if


					' 		extReducedPrice			= extItemCost - extTenCouponPrice - extOwnCouponPrice

					' 		extJungsanType			= "D"

					' 		extCommPrice			= 0
					' 		extCommSupplyPrice		= extCommPrice
					' 		extCommSupplyVatPrice	= 0

					' 		extTenMeachulPrice			= extReducedPrice
					' 		extTenMeachulSupplyPrice	= extTenMeachulPrice
					' 		extTenMeachulSupplyVatPrice	= 0

					' 		extTenJungsanPrice		= extReducedPrice - extCommPrice  ''== ROUND(rsXL(32) / extItemNo,0)

					' 		if (sheetName="sheet1$") then  ''[대금정산] 무료배송쿠폰 기타정산으로 올리는게 좋을듯..
					' 			extJungsanType = "E"
					' 			extItemCost = 0
					' 			extTenCouponPrice = 0
					' 			extOwnCouponPrice = 0
					' 			extReducedPrice = 0
					' 			extCommPrice = 0
					' 			extCommSupplyPrice = 0
					' 			extTenMeachulPrice = 0
					' 			extTenMeachulSupplyPrice = 0


					' 			extTenJungsanPrice 	= CLNG(rsXL(6)*-1)
					' 		end if

					' 		extitemid		= 0
					' 		extitemoption	= "0000"

					' 		extItemName = ""
					' 		extItemOptionName = ""
					' 	else
					' 		IsValidInput = False
					' 	end if
					' else

					' end if
				Case "ssg6006", "ssg6007"
					'// --------------------------------------------------------
					'// SSG
					if extsellsite="ssg6006" then siteNo="6006"
					if extsellsite="ssg6007" then siteNo="6007"

					extOrderserial			= rsXL(7+1)  ''주문ID
					extOrgOrderserial		= rsXL(6+1)  ''원주문ID
					if (extOrderserial=extOrgOrderserial) then extOrgOrderserial=""

					if (Len(extOrderserial) = 14)  then

                        extMeachulDate = replace(rsXL(1),"/","-")  ''정산일
						extJungsanDate = ""

						extOrderserSeq			= CStr(rsXL(8+1)) '주문순번


						extItemNo				= rsXL(15+2)


						if (extItemNo < 0) then  ' 반품
							IsReturnOrder = True
							extOrderserial = extOrderserial & "-" & extOrderserSeq
							extOrgOrderserial = rsXL(6+1)   ''원주문ID
						' elseif (extItemNo = 0) then  '' 수량이 0이면 배송비 또는 보정금액
						'  	IsReturnOrder = False
						'  	''순판매금액이 0이 아닌 CASE가 있음.
						' 	extItemNo = 1
						else
							IsReturnOrder = False
						end if



						extItemCost				= 0
						extReducedPrice			= 0
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0

                        extItemName				= ""     ''상품명
						extItemOptionName		= rsXL(14+2)     ''옵션명(단품)
						if isNULL(extItemOptionName) then extItemOptionName=""

						' if (rsXL(25)<>0) then          ''배송비 (VAT포함)
						' 	extJungsanType			= "D"
						' 	extItemNo = 1
						' 	if (IsReturnOrder) then extItemNo=extItemNo*-1
						' 	extOrderserSeq = extOrderserSeq + "-" + extJungsanType
						' 	extItemName = "배송비"

						' 	if (rsXL(27)<>0) then
						' 		extOrderserSeq = extOrderserSeq + extJungsanType  ''-DD
						' 		extItemName = "반품배송비"
						' 	end if

						' 	if (rsXL(25)<>0) then
						' 		extOrderserSeq = extOrderserSeq + "R"  ''-DDR ,  DR
						' 		extItemName = "반품배송비"
						' 	end if
						' else
						' 	extJungsanType			= "C"
						' end if
						extJungsanType = "C"
						dvlprice = rsXL(25+3)   ''배송비

						''기타 보정작업인듯함..
						if (extItemNo=0) then
							extOwnCouponPrice		= CLNG(rsXL(19+3))
							extTenCouponPrice		= CLNG(rsXL(18+3))
							extCommPrice			= CLNG(rsXL(21+3)-rsXL(16+2))  ''순판매액-정산금액
							extItemCost				= CLNG(rsXL(17+2))
							extTenJungsanPrice		= CLNG(rsXL(16+2))  ''정산금액

							if (dvlprice=0) and (rsXL(21+3)<>0 or rsXL(16+2)<>0) then  ''배송비가 아니고 순판매액이나. 정산액이 <>0 이면 수량을 1로 하고 넣자. 보정금액?
								extItemNo = 1
								extOrderserSeq = extOrderserSeq & "-" & replace(extMeachulDate,"-","")
								''중복되는 케이스가 있다..
								'' 순차적이지는 않다..
								' rw p_extOrderserial&":"&extOrderserial&"::"&p_extOrderserSeq&":"&extOrderserSeq
								' if (p_extOrderserial = extOrderserial) and (p_extOrderserSeq = extOrderserSeq) then
								' 	extOrderserSeq = extOrderserSeq + "-" & extOrderserSeq
								' end if
							end if
						else
							extOwnCouponPrice		= CLNG(rsXL(19+3)/extItemNo*100)/100   ''SSG할인
							extTenCouponPrice		= CLNG(rsXL(18+3)/extItemNo*100)/100   ''협력사할인부담금.

							extCommPrice			= CLNG((rsXL(21+3)-rsXL(16+2))/extItemNo*100)/100  ''수수료 :: 순판매액-정산액
							extItemCost				= CLNG((rsXL(17+2)-dvlprice)/extItemNo*100)/100   ''판매금액합계
							extTenJungsanPrice		= CLNG((rsXL(16+2)-dvlprice)/extItemNo*100)/100   ''정산금액합계
						end if

						extReducedPrice			= CLNG(extItemCost-(extTenCouponPrice+extOwnCouponPrice))   ''소수점 재낌.
						extTenMeachulPrice      = (extItemCost-(extTenCouponPrice+extOwnCouponPrice)) 		''소수점 1자리

						extVatYN = "Y"

						if (rsXL(6) = "면세") then
							extVatYN = "N"
						end if

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

						extitemid = rsXL(12+2)
						if (extItemOptionName="") then
							extitemoption="00000"
						else
							extitemoption=""
						end if
					else
						IsValidInput = False
					end if
				Case "11st1010"
					'// --------------------------------------------------------
					'// 11번가 정산 상세내역
					extOrderserial			= rsXL(1)

					if (Len(extOrderserial) = 15) or (Len(extOrderserial) = 17) then  ''and IsNumeric(rsXL(7)

                        extMeachulDate = replace(rsXL(11),"/","-")  ''정산확정일
						extJungsanDate = ""

						extOrderserSeq			= CStr(rsXL(2))

						extItemNo				= rsXL(16)

						'// 원주문-반품주문 구분법 : 상품 = 수량이 0 보다 작은 경우 반품!!, 배송비 = 반품비(4017388) or 이전 상품주문이 반품인 경우 다음 배송비
''						if (rsXL(2) = 4017357) or (rsXL(2) = 4017388) then
''							if (IsReturnOrder = True or rsXL(2) = 4017388) or (extItemNo <= 0) then
''								extOrderserial 		= rsXL(1) & "-" & rsXL(0)
''								extOrgOrderserial	= rsXL(1)
''							else
''								extOrgOrderserial	= ""
''							end if
''						else


						if (extItemNo < 0) then  ''수량이 0이면 배송비.
							IsReturnOrder = True
							extOrderserial = rsXL(1) & "-" & extOrderserSeq
							extOrgOrderserial		= rsXL(1)
						elseif (extItemNo = 0)	and (rsXL(18)<0) then  ''배송비고 반품인 CASE
							IsReturnOrder = True
							extOrderserial = rsXL(1) & "-" & extOrderserSeq
							extOrgOrderserial		= rsXL(1)
						else
							IsReturnOrder = False
							extOrgOrderserial		= ""
						end if

						extItemCost				= 0
						extReducedPrice			= 0
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						tmpSellerAddSalePriceBy10x10	= 0
						tmpSellerAddSalePriceBy11st1010	= 0

                        extItemName				= html2db(rsXL(14))     ''상품명
						extItemOptionName		= html2db(rsXL(15))     ''옵션명
						'2018-01-18 김진영 하단으로 수정
						extItemOptionName		= requestCheckvar(extItemOptionName, 128)

						if (extItemNo=0) then          ''수량이 0이면 배송비.
							extJungsanType			= "D"
							extItemNo = 1
							if (IsReturnOrder) then extItemNo=extItemNo*-1
							extOrderserSeq = extOrderserSeq + "-" + extJungsanType
							extItemName = "배송비"

							if (rsXL(27+1)<>0) then
								extOrderserSeq = extOrderserSeq + extJungsanType  ''-DD
								extItemName = "반품배송비"
							end if

							if (rsXL(25+1)<>0) then
								extOrderserSeq = extOrderserSeq + "R"  ''-DDR ,  DR
								extItemName = "반품배송비"
							end if

							extItemOptionName = ""
						else
							extJungsanType			= "C"
						end if

						'' 5월1일부터 먼가 바뀌었음.
						' extOwnCouponPrice		= CLNG(rsXL(44)/extItemNo*1000)/1000   ''11번가할인
						' extTenCouponPrice		= CLNG(rsXL(41)/extItemNo*1000)/1000   ''할인쿠폰이용료
						'' 2022-12-15 김진영 위에서 아래로 수정
						' extOwnCouponPrice		= CLNG((CLNG(rsXL(43)) + CLNG(rsXL(44))) /extItemNo*1000)/1000   ''11번가할인
						' extTenCouponPrice		= CLNG(rsXL(42)/extItemNo*1000)/1000   ''할인쿠폰이용료
						'' 2022-12-15 김진영 위에서 아래로 수정
						tmpSellerAddSalePriceBy10x10	= CLNG(CLNG(rsXL(43)) * 0.15)				'판매자추가할인의 15%는 텐바이텐 부담이라함
						tmpSellerAddSalePriceBy11st1010	= CLNG(CLNG(rsXL(43)) * 0.85)				'판매자추가할인의 85%는 11st 부담이라함

						extOwnCouponPrice		= CLNG((tmpSellerAddSalePriceBy11st1010 + CLNG(rsXL(44))) /extItemNo*1000)/1000   ''11번가할인
						extTenCouponPrice		= CLNG((tmpSellerAddSalePriceBy10x10 + CLNG(rsXL(42))) / extItemNo*1000)/1000   ''할인쿠폰이용료

						extCommPrice			= CLNG((CLNG(rsXL(39)) + CLNG(rsXL(40)) +CLNG(rsXL(51))- CLNG(rsXL(44)))/extItemNo*1000)/1000  ''서비스이용료(상품)+ 서비스이용료(선결제배송비) + 후불광고비-11번가할인
						''extCommPrice			= CLNG((rsXL(20+1)-rsXL(39))*100/extItemNo)/100 - CLNG(rsXL(44)/extItemNo*100)/100  ''공제금액합계 :: 할인쿠폰이용료(38) 도 합쳐져 있다.
						extItemCost				= CLNG(rsXL(18)/extItemNo*100)/100   ''판매금액합계
						extTenJungsanPrice		= CLNG(rsXL(17)/extItemNo*100)/100   ''정산금액

						''extCommPrice = extItemCost-extOwnCouponPrice-extTenCouponPrice-extTenJungsanPrice

						extReducedPrice			= CLNG(extItemCost-(extTenCouponPrice+extOwnCouponPrice))   ''소수점 재낌.
						extTenMeachulPrice      = (extItemCost-(extTenCouponPrice+extOwnCouponPrice)) 		''소수점 1자리

						extVatYN = "Y"
						''if (rsXL(6) = "면세") then
						''	extVatYN = "N"
						''end if

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0
					else
						IsValidInput = False
					end if
				Case "gsshopproductday"
					'// --------------------------------------------------------
					'// GS샵 상품정산 상세내역
					extOrderserial			= rsXL(8)

					''if IsNumeric(extOrderserial) and Len(rsXL(5)) = 10 then
					if IsNumeric(extOrderserial) and (Len(extOrderserial) = 9 or Len(extOrderserial) = 10) then

						extMeachulDate = rsXL(32)  ''매출완료일
						extJungsanDate = ""

						if (extMeachulDate = "--") or IsNull(extMeachulDate) then
							extMeachulDate = rsXL(30)  ''배송완료일

							if (extMeachulDate = "--") or IsNull(extMeachulDate) then
								extMeachulDate = rsXL(29)  ''출고완료일
							end if
						end if

						' if (extMeachulDate)="--" or IsNull(extMeachulDate) then
						' 	extMeachulDate = replace(replace(RIGHT(sheetName,9),"$",""),"_","")
						' 	extMeachulDate = LEFT(extMeachulDate,4)&"-"&MID(extMeachulDate,5,2)&"-"&MID(extMeachulDate,7,2)
						' end if

						extVatYN = "Y"

						extOrderserSeq			= rsXL(9)
						extOrgOrderserial		= rsXL(78)

						if IsNull(extOrgOrderserial) then
							extOrgOrderserial = ""
						end if

						extItemNo				= rsXL(43)

						extItemCost				= CLNG(rsXL(46)) ''단가이다. 정산과다름.
						extTenMeachulPrice		= CLNG(rsXL(47) / extItemNo*100)/100
						extOwnCouponPrice		= extItemCost - extTenMeachulPrice
						extTenCouponPrice		= 0
						extJungsanType			= "C"

						extTenJungsanPrice		= CLNG(rsXL(48) / extItemNo*100)/100

						extCommPrice			= (extTenMeachulPrice - extTenJungsanPrice)
						extCommSupplyPrice		= 0
						extCommSupplyVatPrice	= 0

						extReducedPrice			= CLNG(extTenMeachulPrice)
						extTenMeachulSupplyPrice	= 0
						extTenMeachulSupplyVatPrice	= 0

						extitemID = rsXL(37)
						extItemOptionName = html2db(rsXL(39))
					else
					    response.write "["&rsXL(5)&"]"
						IsValidInput = False
					end if
				Case "gsshopproduct"
					'// --------------------------------------------------------
					'// GS샵 상품정산 상세내역
					extOrderserial			= rsXL(0)

					''if IsNumeric(extOrderserial) and Len(rsXL(5)) = 10 then
					if IsNumeric(extOrderserial) and (Len(extOrderserial) = 9 or Len(extOrderserial) = 10) then

						extMeachulDate = rsXL(19+1)  ''매출완료일
						extJungsanDate = ""

						if (extMeachulDate = "--") or IsNull(extMeachulDate) then
							extMeachulDate = rsXL(18+1)  ''배송완료일

							if (extMeachulDate = "--") or IsNull(extMeachulDate) then
								extMeachulDate = rsXL(17+1)  ''출고완료일
							end if
						end if

						if (extMeachulDate)="--" or IsNull(extMeachulDate) then
							extMeachulDate = replace(replace(RIGHT(sheetName,9),"$",""),"_","")
							extMeachulDate = LEFT(extMeachulDate,4)&"-"&MID(extMeachulDate,5,2)&"-"&MID(extMeachulDate,7,2)

							rw extMeachulDate
							if NOT isDate(extMeachulDate) then
								extMeachulDate = RIGHT(sheetName,10)
								rw extMeachulDate
								extMeachulDate = LEFT(extMeachulDate,4)&"-"&MID(extMeachulDate,5,2)&"-"&MID(extMeachulDate,7,2)
								rw extMeachulDate
							end if
						end if

						extVatYN = "Y"

						extOrderserSeq			= rsXL(2+1)
						extOrgOrderserial		= rsXL(6+1)

						if IsNull(extOrgOrderserial) then
							extOrgOrderserial = ""
						end if

						extItemNo				= rsXL(16+1)

						extItemCost				= CLNG(rsXL(20+1) / extItemNo*100)/100
						extTenMeachulPrice		= CLNG(rsXL(21+1) / extItemNo*100)/100
						extOwnCouponPrice		= extItemCost - extTenMeachulPrice
						extTenCouponPrice		= 0
						extJungsanType			= "C"

						extTenJungsanPrice		= CLNG(rsXL(22+1) / extItemNo*100)/100

						extCommPrice			= (extTenMeachulPrice - extTenJungsanPrice)
						extCommSupplyPrice		= 0
						extCommSupplyVatPrice	= 0

						extReducedPrice			= CLNG(extTenMeachulPrice)
						extTenMeachulSupplyPrice	= 0
						extTenMeachulSupplyVatPrice	= 0

						extitemID = rsXL(9+1)
						extItemOptionName = html2db(rsXL(11+1))
					else
					    response.write "["&rsXL(5+1)&"]"&rsXL(0)
						IsValidInput = False
					end if
				Case "gsshopbeasongpay"
					'// --------------------------------------------------------
					'// GS샵 배송비정산 상세내역
					extOrderserial			= rsXL(1)

					if (Len(extOrderserial) = 9 or Len(extOrderserial) = 10) and IsNumeric(rsXL(4)) then

						'// 우선은 원주문일자 넣고 아래에서 출고일자 입력한다.
						extMeachulDate = rsXL(6)
						extJungsanDate = ""

						extVatYN = "Y"

						tmpStr = rsXL(2)
						if IsNull(tmpStr) then
							tmpStr = ""
						end if

						extOrgOrderserial = ""
						if (tmpStr <> "") then
							extOrderserial = rsXL(2)
							extOrgOrderserial = rsXL(1)
						end if

						extItemNo				= 1
						extItemName				= rsXL(0)
						extOrderserSeq			= rsXL(3)&"-"&rsXL(4)

						if isNULL(extItemName) then extItemName="기타"
						if (extOrgOrderserial = "") then


							extItemCost				= rsXL(10)
							extReducedPrice			= extItemCost
							extOwnCouponPrice		= 0
							extTenCouponPrice		= 0
							extJungsanType			= "D"
							extCommPrice			= 0
							extTenMeachulPrice		= extItemCost
							extTenJungsanPrice		= extItemCost
						else
							extItemCost				= rsXL(10)
							if (rsXL(0) = "반품배송비(환불)") then
								extItemNo = -1
								extItemCost = extItemCost * -1
							end if

							extReducedPrice			= extItemCost
							extOwnCouponPrice		= 0
							extTenCouponPrice		= 0
							extJungsanType			= "D"
							extCommPrice			= 0
							extTenMeachulPrice		= extItemCost
							extTenJungsanPrice		= extItemCost
						end if

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

						if (pMxMeachulMonth="") then
							pMxMeachulMonth = LEFT(extMeachulDate,7)
						end if

						if (pMxMeachulMonth<LEFT(extMeachulDate,7)) then
							extMeachulDate = (dateAdd("d",-1,LEFT(extMeachulDate,7)+"-01"))
						end if
					else
						IsValidInput = False
					end if
				Case "dnshopproduct"
					'// --------------------------------------------------------
					'// 디앤샵 상품정산 상세내역

				Case "cjmallproduct"
					'// --------------------------------------------------------
					'// CJ몰 상품정산 상세내역

				Case "wizwidproduct"
					'// --------------------------------------------------------
					'// 위즈위드 상품정산 상세내역

				Case "gabangpopproduct"
					'// --------------------------------------------------------
					'// 패션팝(가방팝) 상품정산 상세내역
				Case "goodshop1010"
					extOrderserial = rsXL(4)	'주문번호
					If Len(extOrderserial) = 19 Then
						extJungsanDate = ""
						extOrgOrderserial = ""
						extMeachulDate =  LEFT(Trim(rsXL(6)), 10)		'주문일자
						extOrderserSeq = extMeachulDate & "-" & i  ''Seq에 정산일을 더한다.
						extItemNo = Trim(rsXL(16))		'주문수량
						If (extItemNo <= 0) Then
							IsReturnOrder = True
							extOrgOrderserial = extOrderserial
							extOrderserial = extOrderserial&"-1"
						Else
							IsReturnOrder = False
						End If
						extJungsanType			= "C"
						extItemName			= html2db(rsXL(14))			'상품명
						extItemOptionName	= html2db(rsXL(15))			'옵션명
						extItemCost	= CLNG(rsXL(18) / extitemNo)		'판매가
						extTenCouponPrice	= 0
						extOwnCouponPrice	= 0

						extTenMeachulPrice	= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice		= CLng(extTenMeachulPrice)
						extCommPrice		= CLNG(CLNG(rsXL(24)) / extitemNo)		'굿샵정산금액
						extTenJungsanPrice	= extReducedPrice - extCommPrice
						
						If rsXL(26) = "과세" Then
							extVatYN = "Y"
						Else
							extVatYN = "N"
						End If
						extitemid = rsXL(5)		'상품번호
					Else
						IsValidInput = False
					End If
				Case "wconcept1010"
					extOrderserial = rsXL(5)	'주문번호
					If Len(extOrderserial) = 9 Then
						extJungsanDate = ""
						extOrgOrderserial = ""
						extMeachulDate =  Trim(rsXL(4))		'정산확정일(처리완료일)
						extOrderserSeq = Trim(rsXL(37))
						extItemNo = Trim(rsXL(17))		'수량
						If (extItemNo <= 0) Then
							IsReturnOrder = True
							extOrgOrderserial = extOrderserial
							extOrderserial = extOrderserial&"-1"
						Else
							IsReturnOrder = False
						End If

						If rsXL(38) = "배송료" then          '구분
							extJungsanType	= "D"
							extItemNo = 1
							extOrderserSeq = extMeachulDate & "-" & i & "-" & extJungsanType
							extItemName = "배송비"
							extItemOptionName = ""
							extItemCost	= CLNG(rsXL(27) / extitemNo)		'배송료
						Else
							extJungsanType			= "C"
							extItemName			= html2db(rsXL(13))			'상품명
							extItemOptionName	= html2db(rsXL(15))			'옵션1
							extItemCost	= CLNG(rsXL(31) / extitemNo)		'순매출금액
						End If
						' extTenCouponPrice	= CLNG(CLNG(rsXL(21)) / extitemNo)	'업체쿠폰
						' extOwnCouponPrice	= CLNG(CLNG(rsXL(22)) / extitemNo)	'본사쿠폰
						extTenCouponPrice	= 0
						extOwnCouponPrice	= 0

						extTenMeachulPrice	= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice		= CLng(extTenMeachulPrice)
						extCommPrice		= CLNG(CLNG(rsXL(32)) / extitemNo) + CLNG(CLNG(rsXL(33)) / extitemNo)   	'세금계산서 + 해외판매 수수료
						extTenJungsanPrice	= CLNG(rsXL(34) / extItemNo * 100)/100	'업체지급액
						extTenJungsanPrice	= extReducedPrice - extCommPrice
						extVatYN = "Y"
						extitemid = ""
					Else
						IsValidInput = False
					End If
				Case "withnature1010"
					extOrderserial = rsXL(2)	'주문번호
					If Len(extOrderserial) = 12 Then
						extJungsanDate = ""
						extOrgOrderserial = ""
						extOrderserSeq = ""
						extMeachulDate =  Trim(rsXL(3))		'입고일자
						
						extItemNo = Trim(rsXL(7))			'수량
						If (extItemNo <= 0) Then
							IsReturnOrder = True
							extOrgOrderserial = extOrderserial
							extOrderserial = extOrderserial&"-1"
						Else
							IsReturnOrder = False
						End If
						extJungsanType			= "C"
						extItemName			= html2db(rsXL(6))			'상품명
						extItemOptionName	= ""

						extTenCouponPrice	= 0	'업체쿠폰
						extOwnCouponPrice	= 0	'본사쿠폰

						sqlStr = ""
						sqlStr = sqlStr & " SELECT TOP 1 SellPrice, matchItemID, matchitemoption, OrgDetailKey "
						sqlStr = sqlStr & " FROM db_temp.dbo.tbl_xSite_TMPOrder "
						If extItemNo < 0 Then
							sqlStr = sqlStr & " WHERE OutMallOrderSerial= '"&rsXL(2)&"' "
						Else
							sqlStr = sqlStr & " WHERE OutMallOrderSerial= '"&extOrderserial&"' "
						End If
						sqlStr = sqlStr & " and outMallGoodsNo = '"&rsXL(5)&"' "
						sqlStr = sqlStr & " and SellSite = 'withnature1010' "
						If extItemNo < 0 Then
							sqlStr = sqlStr & " and ItemOrderCount = '"&extItemNo * -1 &"' "
						Else
							sqlStr = sqlStr & " and ItemOrderCount = '"&extItemNo&"' "
						End If
						sqlStr = sqlStr & " and OrderSerial is not null "
						rsget.CursorLocation = adUseClient
						rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
						if Not rsget.Eof then
							extItemCost		= rsget("SellPrice")		'매출금액이 안 넘어온다. 주문내역에서 가져와야한다.
							extitemid		= rsget("matchItemID")
							extitemoption	= rsget("matchitemoption")
							extOrderserSeq	= rsget("OrgDetailKey")
						End If
						rsget.Close

						If extItemNo < 0 Then
							extOrderserSeq = extOrderserSeq & "-1"
						End If

						If extOrderserSeq = "" Then
							sqlStr = ""
							sqlStr = sqlStr & " SELECT TOP 1 SellPrice, matchItemID, matchitemoption, OrgDetailKey "
							sqlStr = sqlStr & " FROM db_temp.dbo.tbl_xSite_TMPOrder "
							sqlStr = sqlStr & " WHERE ref_OutMallOrderSerial= '"&extOrderserial&"' "
							sqlStr = sqlStr & " and outMallGoodsNo = '"&rsXL(5)&"' "
							sqlStr = sqlStr & " and SellSite = 'withnature1010' "
							If extItemNo < 0 Then
								sqlStr = sqlStr & " and ItemOrderCount = '"&extItemNo * -1 &"' "
							Else
								sqlStr = sqlStr & " and ItemOrderCount = '"&extItemNo&"' "
							End If
							sqlStr = sqlStr & " and OrderSerial is not null "
							rsget.CursorLocation = adUseClient
							rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
							if Not rsget.Eof then
								extItemCost		= rsget("SellPrice")		'매출금액이 안 넘어온다. 주문내역에서 가져와야한다.
								extitemid		= rsget("matchItemID")
								extitemoption	= rsget("matchitemoption")
								extOrderserSeq	= rsget("OrgDetailKey")
							End If
							rsget.Close
						End If

						If extOrderserSeq = "" Then
							extOrderserSeq = Trim(rsXL(0))		'No. | 주문디테일키 안 넘어옴 임시로 이걸로..
						End If

						extTenMeachulPrice	= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice		= CLng(extTenMeachulPrice)
						extCommPrice		= extItemCost - (CLNG(rsXL(11) / extItemNo * 100) / 100)
						extTenJungsanPrice	= CLNG(rsXL(11) / extItemNo * 100) / 100
						extVatYN = "Y"
						dvlprice = 0			'자연이랑은 배송비 무배
					Else
						IsValidInput = False
					End If
				Case "GS25"
					extOrderserial = rsXL(3)	'주문번호
					If Len(extOrderserial) = 13 Then
						extJungsanDate = ""
						extOrgOrderserial = ""
						extMeachulDate = Trim(rsXL(2))		'배송완료일
						extMeachulDate = Left(extMeachulDate, 4) & "-" & Right(Left(extMeachulDate, 6), 2) & "-" & Right(extMeachulDate, 2)
						extOrderserSeq = Trim(rsXL(4))
						extItemNo = Trim(rsXL(7))		'수량
						If (extItemNo <= 0) Then
							IsReturnOrder = True
							extOrgOrderserial = extOrderserial
							extOrderserial = extOrderserial&"-1"
						Else
							IsReturnOrder = False
						End If
						extJungsanType		= "C"
						extItemName			= html2db(rsXL(6))			'상품명
						extItemOptionName	= ""
						extItemCost			= CLNG(rsXL(8))				'순매출금액
						extTenCouponPrice	= 0
						extOwnCouponPrice	= 0

						extTenMeachulPrice	= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice		= CLng(extTenMeachulPrice)
						extCommPrice		= extTenMeachulPrice - (CLNG(rsXL(12) / extItemNo))
						extTenJungsanPrice	= CLNG(rsXL(12) / extItemNo)
						extVatYN = "Y"
						extitemid = ""
						Select Case rsXL(5)
							Case "2800100203602"	extitemid = "3313868"
							Case "2800100204449"	extitemid = "3471382"
							Case "2800100204456"	extitemid = "4524679"
							Case "2800100204463"	extitemid = "4890940"
							Case "2800100204487"	extitemid = "4509495"
							Case "2800100204494"	extitemid = "4509498"
							Case "2800100204500"	extitemid = "3504305"
							Case "2800100204517"	extitemid = "4728736"
						End Select
						extitemoption = "0000"
					Else
						IsValidInput = False
					End If
				Case "priviaproduct"
					'// --------------------------------------------------------
					'// 현대프리비아 상품정산 상세내역

				Case "playerproduct"
					'// --------------------------------------------------------
					'// 플레이어 상품정산 상세내역

				Case "lotteComM"  ''201605 수정
					'// --------------------------------------------------------
					'// 롯데닷컴(직매입) 상품정산

				Case "lotteComM_201604"
					'// --------------------------------------------------------
					'// 롯데닷컴(직매입) 상품정산
				Case "lotteCombeasongpay"
					'// --------------------------------------------------------
					'// 롯데닷컴 배송비

					if (rsXL(11) = "정산마감") then
						extOrderserial		= Replace(rsXL(3),"-","")
						extMeachulDate = rsXL(2)
						extItemCost = rsXL(10)
						extItemNo = 1


						'extOrderserSeq = replace(extMeachulDate,"-","")&"-D"&"-"&i

						' if (rsXL(5)="일반") then

						' elseif (rsXL(5)="고객클레임") then
						' 	extOrderserSeq = extOrderserSeq&"C"
						' elseif (rsXL(5)="플래티넘+무료배송") then
						' 	extOrderserSeq = extOrderserSeq&"P"
						' elseif (rsXL(5)="무료배송권") then
						' 	extOrderserSeq = extOrderserSeq&"F"
						' elseif (rsXL(5)="롯데지정-고객") then
						' 	extOrderserSeq = extOrderserSeq&"T"
						' elseif (rsXL(5)="롯데지정-업체") then
						' 	extOrderserSeq = extOrderserSeq&"U"
						' elseif (rsXL(5)="업체수거-고객") then
						' 	extOrderserSeq = extOrderserSeq&"S"
						' else
						' 	extOrderserSeq = extOrderserSeq&"E"
						' end if

						if (extItemCost<0) then
							extOrderserSeq = replace(extMeachulDate,"-","")&"-D"
						else
							extOrderserSeq = replace(extMeachulDate,"-","")&"D"
						end if

						if (rsXL(7)<>0) then
							extOrderserSeq = extOrderserSeq&"D"
						end if

						if (rsXL(8)<>0) then
							extOrderserSeq = extOrderserSeq&"C"
						end if

						if (rsXL(9)<>0) then
							extOrderserSeq = extOrderserSeq&"U"
						end if

						extOrderserSeq = extOrderserSeq &"-"&i

						if (extItemCost=0) then
							extItemNo = 0
						end if



						extJungsanDate = ""


						extReducedPrice			= extItemCost
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extJungsanType			= "D"

						extCommPrice			= 0
						extTenMeachulPrice		= extReducedPrice
						extTenJungsanPrice		= extReducedPrice

						extItemName				= "배송비"
						extItemOptionName		= ""

						extVatYN = "Y"

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0
					else
						IsValidInput = False
					end if
				Case "lotteCom"  ''2016/05/01 주문상세번호 필드 생긴듯.
					'// --------------------------------------------------------
					'// 롯데닷컴 위탁정산 상세내역
					if (rsXL(28) = "텐바이텐") then  ''24+1 =>24+2 2016/05/01  // 27 2019/07/12 (27) (등록상품명), 거래처분담금

						extMeachulDate = rsXL(0)
						extJungsanDate = ""

						extItemNo				= rsXL(10)
						plusMinus				= extItemNo / Abs(extItemNo)
						if (extItemNo >= 0) then
							'// 정상출고
							extOrderserial 			= Replace(rsXL(24), "-", "")
							extOrderserSeq			= TRim(rsXL(25)) ''=>  2016/05/01 //
							extOrgOrderserial		= ""
						else
							'// 반품
							extOrderserial 			= Replace(rsXL(24), "-", "") & "-" & i
							extOrderserSeq			= TRim(rsXL(25)) ''=>  2016/05/01 //
							extOrgOrderserial		= Replace(rsXL(24), "-", "")
						end if

						extVatYN = "Y"
						if (rsXL(14) = 0) then
							extVatYN = "N"
						end if

						extItemCost				= CLNG(rsXL(12)  / extItemNo * 100)/100
						extTenCouponPrice		= CLNG((rsXL(15) - rsXL(20)) / extItemNo * 100)/100  ''CLNG(rsXL(16)  / extItemNo * 100)/100  ''
						extTenMeachulPrice		= CLNG(rsXL(17)  / extItemNo * 100)/100
						extOwnCouponPrice		= extItemCost - extTenMeachulPrice - extTenCouponPrice

						extJungsanType			= "C"

						extCommPrice			= CLNG((rsXL(19) - rsXL(20))  / extItemNo * 100)/100
						extReducedPrice			= CLNG(extTenMeachulPrice)
						extTenJungsanPrice		= CLNG(rsXL(23) / extItemNo * 100)/100

						extItemName				= html2db(rsXL(7))			'// 외부몰 상품명이 바뀐다. 상품명 대신 외부몰 상품코드로 매칭
						extItemOptionName		= html2db(rsXL(8))						'// 정산내역에 옵션정보가 없다. ==>있음.

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0
					else
						IsValidInput = False
					end if
				Case "halfclubproduct"
					'// --------------------------------------------------------
					'// 하프클럽 상품 상세내역
					if (LEN(rsXL(3)) = 12) then

						extOrderserial = rsXL(3)
						extMeachulDate = rsXL(1)
						extJungsanDate = ""
						extMeachulDate = Left(extMeachulDate, 4) & "-" & Right(Left(extMeachulDate, 6), 2) & "-" & Right(extMeachulDate, 2)

						extOrgOrderserial		= CSTR(rsXL(2))
						extOrgOrderserial = ""

						extOrderserSeq			= rsXL(4)
						extItemNo				= rsXL(17)   ''실출고수량


						if (extItemNo <0) then
							'// 반품
							extOrgOrderserial = extOrderserial
							extOrderserial = extOrderserial&"-1"
						end if

						extVatYN = "Y"

						extItemCost				= CLNG(rsXL(16)  / extItemNo * 100)/100
						extTenCouponPrice		= CLNG((rsXL(22) + rsXL(23)*-1) / extItemNo * 100)/100
						extOwnCouponPrice		= CLNG(rsXL(19)  / extItemNo * 100)/100
						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice


						extJungsanType			= "C"

						extCommPrice			= CLNG((rsXL(20))  / extItemNo * 100)/100
						extReducedPrice			= CLNG(extTenMeachulPrice)
						extTenJungsanPrice		= CLNG(rsXL(24) / extItemNo * 100)/100

						extItemName				= ""
						extItemOptionName		= ""

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

						extitemID = rsXL(7)
						extItemOptionName = rsXL(8)
					else
						IsValidInput = False
					end if
				Case "halfclubbeasongpay"
					'// --------------------------------------------------------
					'// 하프클럽 배송비 상세내역
					if (LEN(rsXL(1)) = 12) then
						extOrderserial		= Replace(rsXL(1),"-","")
						extMeachulDate = rsXL(0)
						extMeachulDate = Left(extMeachulDate, 4) & "-" & Right(Left(extMeachulDate, 6), 2) & "-" & Right(extMeachulDate, 2)

						extItemCost = rsXL(7)
						extItemNo = 1



						if (extItemCost<0) then
							extOrderserSeq = replace(extMeachulDate,"-","")&"-D"
						else
							extOrderserSeq = replace(extMeachulDate,"-","")&"D"
						end if

						if (rsXL(4)<>0) then
							extOrderserSeq = extOrderserSeq&"C"
						end if

						if (rsXL(5)<>0) then
							extOrderserSeq = extOrderserSeq&"R"
						end if

						if (rsXL(6)<>0) then
							extOrderserSeq = extOrderserSeq&"F"
						end if

						''extOrderserSeq = extOrderserSeq &"-"&i

						if (extItemCost=0) then
							extItemNo = 0
						end if



						extJungsanDate = ""


						extReducedPrice			= extItemCost
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extJungsanType			= "D"

						extCommPrice			= 0
						extTenMeachulPrice		= extReducedPrice
						extTenJungsanPrice		= extReducedPrice

						extItemName				= "배송비"
						extItemOptionName		= ""

						extVatYN = "Y"

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0
					else
						IsValidInput = False
					end if
				Case "nvstorefarm","nvstorefarmclass", "nvstoremoonbangu", "nvstoregift", "wadsmartstore", "Mylittlewhoopee"
					'// --------------------------------------------------------
					'// 스토어팜 상세내역
					if (Len(rsXL(0)) = 16) then
						isextMeachulDate = ""
						extMeachulDate = rsXL(5)
						extJungsanDate = "" 'rsXL(6)
						extMeachulDate = Left(extMeachulDate, 4) & "-" & Right(Left(extMeachulDate, 6), 2) & "-" & Right(extMeachulDate, 2)
						extJungsanDate = "" 'Left(extJungsanDate, 4) & "-" & Right(Left(extJungsanDate, 6), 2) & "-" & Right(extJungsanDate, 2)

						extItemNo				= 1					'// 수량없음 **



						extOrderserial 			= rsXL(0)
						extOrderserSeq 			= rsXL(1)
						If (rsXL(13)="빠른정산 회수") Then
							extOrderserSeq 	= extOrderserSeq & "-1"
						ElseIf (rsXL(13)="빠른정산") Then
							'extOrderserSeq 	= extOrderserSeq & "-" & rsXL(5)
							sqlStr = ""
							sqlStr = sqlStr & " SELECT COUNT(*) as cnt FROM db_temp.dbo.tbl_xSite_JungsanTmp "
							sqlStr = sqlStr & " WHERE extOrderserial = '"& rsXL(0) &"' "
							sqlStr = sqlStr & " and extOrderserSeq = '"& rsXL(1) &"' "
							rsget.CursorLocation = adUseClient
							rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
							If rsget("cnt") > 0 Then
								extOrderserSeq 	= extOrderserSeq & "-" & rsXL(5)
							End If
							rsget.Close
						End If

'전달 데이터 땜에 정산등록이 안 되는 현상..
'						If Date() <= "2021-08-02" Then
							'extOrderserSeq 	= extOrderserSeq & "-" & rsXL(5)
							sqlStr = ""
							sqlStr = sqlStr & " SELECT COUNT(*) as cnt FROM db_temp.dbo.tbl_xSite_JungsanTmp "
							sqlStr = sqlStr & " WHERE extOrderserial = '"& rsXL(0) &"' "
							sqlStr = sqlStr & " and extOrderserSeq = '"& extOrderserSeq &"' "
							rsget.CursorLocation = adUseClient
							rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
							If rsget("cnt") > 0 Then
								extOrderserSeq 	= extOrderserSeq & "-" & rsXL(5)
							End If
							rsget.Close

							'1. seq로 정산내역 있는 지 확인, 이때 매출일도 가져온다
							sqlStr = ""
							sqlStr = sqlStr & " SELECT extMeachulDate FROM db_jungsan.dbo.tbl_xSite_JungsanData "
							sqlStr = sqlStr & " WHERE extOrderserial = '"& rsXL(0) &"' "
							sqlStr = sqlStr & " and extOrderserSeq = '"& extOrderserSeq &"' "
							rsget.CursorLocation = adUseClient
							rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
							if Not rsget.Eof then
								isextMeachulDate = rsget("extMeachulDate")
							End If
							rsget.Close

							'2. 만약 정산내역이 있으면 가져온 매출일과 등록할 매출일을 비교한다.
							If isextMeachulDate <> "" Then
								If isextMeachulDate <> extMeachulDate Then
							'3. 비교시 날짜가 다르면 -A를 붙인다
									extOrderserSeq 	= extOrderserSeq & "-A"
								End If
							End If
'						End If

						extOrgOrderserial		= ""
						extVatYN = "Y"

						if rsXL(2) = "배송비" then
							extJungsanType			= "D"
							extItemNo				= 1
						else
							extJungsanType			= "C"
						end if

						if (rsXL(13)="정산후 취소") then extItemNo=-1

						extItemCost				= rsXL(9)			'// 판매가 없음
						if (extItemCost=0 and extJungsanType="D") then extItemNo=0

						extTenCouponPrice		= 0
						extReducedPrice			= rsXL(9)
						extOwnCouponPrice		= 0

						extCommPrice			= (rsXL(10) + rsXL(11) + rsXL(12)) * -1
						extTenMeachulPrice		= extReducedPrice
						extTenJungsanPrice		= extReducedPrice - extCommPrice

						extItemName				= html2db(rsXL(3))
						extItemOptionName		= ""

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

						if (extItemNo<0) then
							extOrderserSeq=extOrderserSeq&"-1"
							extItemCost = extItemCost*-1
							extReducedPrice = extReducedPrice*-1
							extCommPrice	= extCommPrice*-1
							extTenMeachulPrice = extTenMeachulPrice*-1
							extTenJungsanPrice = extTenJungsanPrice*-1
							extCommSupplyPrice = extCommSupplyPrice*-1
							extTenMeachulSupplyPrice = extTenMeachulSupplyPrice*-1
						end if
					else
						IsValidInput = False
					end if
				Case "lotteon"
					'// --------------------------------------------------------
					'// 롯데온 상세내역
					If (Len(rsXL(7)) = 16) then
						extOrderserial 	= rsXL(7)   											''주문번호
						extOrgOrderserial = ""
						extOrderserSeq = ""
						extCommPrice = ""
						extMeachulDate		= rsXL(20)  										''구매확정일
						lotteonBanpoomDate	= rsXL(21)  										''반품완료일

						If Trim(lotteonBanpoomDate) <> "" Then
							extMeachulDate = lotteonBanpoomDate
						End If

						If extMeachulDate = "" Then
							extMeachulDate = rsXL(23)											'매출대상기간
						End If

						If extMeachulDate = "" AND rsXL(22) <> "" Then							'rsXL(19) : 정산예정일 / 구매확정일, 배송완료일, 반품완료일 전부 데이터가 없고 예정일만 있음
							extMeachulDate = dateadd("d", -1, rsXL(22))
						End If

						extJungsanDate	= ""
						extItemNo		= rsXL(24)  											''판매수량
						extOrderserSeq	= extOrderserial &"-"& rsXL(0)
						extOrderserSeq	= extOrderserSeq & "-" & replace(extMeachulDate,"-","")

						If rsXL(23) <> "" AND extItemNo < 1 Then
							extItemNo = 1
						Else
							If (extItemNo <= 0) Then
								IsReturnOrder = True
								extOrgOrderserial = extOrderserial
								extOrderserSeq	= extOrderserSeq & "-1"
							Else
								IsReturnOrder = False
							End If
						End If

						'엑셀의 셀러==텐바이텐, 당사==롯데
						extItemCost				= rsXL(25)										'판매단가
						extOwnCouponPrice		= CLNG(rsXL(28) / Chkiif(extItemNo="0", "1", extItemNo) )				'rsXL(25) : 상품할인당사부담금액
						extTenCouponPrice		= CLNG(CLNG(rsXL(27) + rsXL(29)) / Chkiif(extItemNo="0", "1", extItemNo))	'rsXL(24) : 셀러즉시할인금액, rsXL(26) : 상품할인셀러부담금액
						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extJungsanType			= "C"

                    If extItemNo < 0 AND CLNG(rsXL(28)) < 0 AND CLNG(rsXL(43)) = 0 Then
                        'extCommPrice            = CLNG(CLNG(CLNG(rsXL(35)) + CLNG(rsXL(37)) + CLNG(rsXL(44))) / Chkiif(extItemNo="0", "1", extItemNo))
						'maybe									기본수수료 + PCS수수료금액 + PG수수료 총합계
						'extCommPrice            = CLNG(CLNG(CLNG(rsXL(36)) + CLNG(rsXL(39)) + CLNG(rsXL(46))) / Chkiif(extItemNo="0", "1", extItemNo))
						'maybe									기본수수료 + PCS수수료금액 + PG수수료 총합계
						extCommPrice            = CLNG(CLNG(CLNG(rsXL(36)) + CLNG(rsXL(40)) + CLNG(rsXL(47))) / Chkiif(extItemNo="0", "1", extItemNo))
                    Else
                        'extCommPrice            = CLNG(CLNG(CLNG(rsXL(35)) + CLNG(rsXL(37)) + CLNG(rsXL(44)) - rsXL(28)) / Chkiif(extItemNo="0", "1", extItemNo))
						'maybe									기본수수료 + PCS수수료금액 + PG수수료 총합계 - 상품할인당사부담금액
						extCommPrice            = CLNG(CLNG(CLNG(rsXL(36)) + CLNG(rsXL(40)) + CLNG(rsXL(47)) - CLNG(rsXL(28))) / Chkiif(extItemNo="0", "1", extItemNo))
                    End If
						extTenJungsanPrice		= extReducedPrice - extCommPrice

						lotteondvlprice1		=  CLNG(rsXL(30))	'배송비정산대상금액-> 총배송비 
						lotteondvlprice2		=  CLNG(rsXL(32))	'배송비할인(당사부담)
						lotteondvlPGCommprice	=  CLNG(rsXL(47))	'PG수수료 총합계
						lotteondvlTotCommprice	=  CLNG(rsXL(48))	'총 수수료 합계
						dlvCommprice			= CLNG(rsXL(38))	'배송비 수수료

						'dvlprice				= lotteondvlprice1 - lotteondvlprice2			'//2020-05-26 김진영..이렇게 빼야 실 배송비 금액이 나오는 듯?
						dvlprice				= lotteondvlprice1 								'배송비의 경우 배송비할인(당사부담) : 롯데부담분은 차감하지 않는것이 맞는듯함. by)eastone

						extItemName				= LeftB(html2db(rsXL(12)), 80)					'rsXL(9) : 상품명(옵션명포함) / LeftB처리..2021-01-18..2021010913167292
						extItemOptionName		= ""											'위처럼 상품명과 옵션명이 같이 옴..빈값처리
						extitemoption			= html2db(rsXL(11))								'rsXL(8) : 단품번호

						If rsXL(14) = "과세" Then												'과세구분은 : 과세유형 필드를 볼것..by)eastone
							extVatYN = "Y"
						Else
							extVatYN = "N"
						End If
						extitemid = rsXL(10)														'rsXL(7) : 판매자상품번호

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

						IsValidInput = True
					Else
						IsValidInput = False
					End If
				Case "yes24"
					'// --------------------------------------------------------
					'// yes24 상세내역
					If (Len(rsXL(1)) = 11) then
						extOrderserial 	= rsXL(1)   				''주문번호
						extOrgOrderserial = ""
						extOrderserSeq = ""
						extCommPrice = ""
						extMeachulDate		= rsXL(0)  				''일자

						extJungsanDate	= ""
						extItemNo		= 1  						''수량 / 수량 엑셀에 없음..강제 1처리
						extOrderserSeq	= extOrderserial &"-"& rsXL(0)

						extOwnCouponPrice = 0						'쿠폰금액 알수없음
						extTenCouponPrice = 0						'쿠폰금액 알수없음

						If (rsXL(4) <> 0) Then						''반품금액이 0이 아니면
							IsReturnOrder = True
							extOrgOrderserial = extOrderserial
							extOrderserSeq	= extOrderserSeq & "-" & rsXL(2)
							extItemCost	= rsXL(4)					'반품금액
							extItemNo		= -1
							extCommPrice			= CLNG(rsXL(4)) - CLNG(rsXL(6))	'수수료 = 반품금액 - 반품원가
						Else
							IsReturnOrder = False
							extItemCost	= rsXL(3)					'주문금액
							extCommPrice			= CLNG(rsXL(3)) - CLNG(rsXL(5))	'수수료 = 주문금액 - 주문원가
						End If

						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extJungsanType			= "C"

						extTenJungsanPrice		= extReducedPrice - extCommPrice
						dvlprice				= rsXL(7) 			'배송비

						extItemName				= ""				'상품명 알수없음
						extItemOptionName		= ""				'옵션명 알수없음
						extitemoption			= ""				'옵션 알수없음

						extVatYN = "Y"								'과세여부 알수없음
						extitemid = ""								'상품번호 알수없음

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

						IsValidInput = True
					Else
						IsValidInput = False
					End If
				Case "ohou1010"
					'// --------------------------------------------------------
					If (Len(rsXL(5)) = 9) then
						extOrderserial 	= rsXL(5)   				''주문번호
						extOrgOrderserial = ""
						extOrderserSeq = ""
						extCommPrice = ""
						extMeachulDate		= LEFT(rsXL(18),10)  	''정산기준일

						extJungsanDate	= ""
						extItemNo		= rsXL(8)  					''수량

'						extOrderserSeq	= extOrderserial &"-"& rsXL(1)
						extOrderserSeq	= extOrderserial & "-" & rsXL(6) & "-" & i	''주문옵션번호rsXL(6)

						extOwnCouponPrice = 0						'쿠폰금액 알수없음
						extTenCouponPrice = 0						'쿠폰금액 알수없음

						If (extItemNo < 0) Then						''수량 0이 아니면
							IsReturnOrder = True
							extOrgOrderserial = extOrderserial
							extOrderserSeq	= extOrderserSeq & "-" & rsXL(6)
							extItemCost		= CLNG(rsXL(10) / extItemNo)		'상품판매가A
							extItemNo		= -1
							extCommPrice	= CLNG(rsXL(13) / extItemNo )		'수수료D
						Else
							IsReturnOrder = False
							extItemCost		= CLNG(rsXL(10) / extItemNo)		'상품판매가A
							extCommPrice	= CLNG(rsXL(13) / extItemNo )		'수수료D
						End If

						extJungsanType			= "C"
						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extTenJungsanPrice		= extReducedPrice - extCommPrice
						dvlprice				= rsXL(11) 			'배송비

						extItemName				= rsXL(7)			'상품명
						extItemOptionName		= ""
						extitemoption			= ""				'옵션 알수없음

						extVatYN = "Y"								'과세여부 알수없음
						extitemid = rsXL(6)							'상품번호 알수없음

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

						IsValidInput = True
					Else
						IsValidInput = False
					End If
				Case "casamia_good_com"		'2021-02-02 김진영..수량이 1인 주문만 있다. 2이상일 때 나눠야 될 지도..
					'// --------------------------------------------------------
					If (Len(rsXL(2)) = 14) then
						extOrderserial 	= rsXL(2)   				''주문번호
						extOrgOrderserial = ""
						extOrderserSeq = ""
						extCommPrice = 0
						extOwnCouponPrice = 0
						extTenCouponPrice = 0

						extMeachulDate		= LEFT(rsXL(1),10)  	''구매확정일

						extJungsanDate	= ""
						extItemNo		= rsXL(11)  					''수량
						extOrderserSeq	= extOrderserial & "-" & rsXL(3)

						If (extItemNo < 0) Then						''수량 0이 아니면
							IsReturnOrder = True
							extOrgOrderserial = extOrderserial
							extOrderserSeq	= extOrderserSeq & "-" & rsXL(3)
							extItemCost		= CLNG(rsXL(12)/extItemNo*100)/100
							extItemNo		= -1
'							extCommPrice	= CLNG(rsXL(42)/extItemNo*100)/100	'판매대행수수료
							'2021-04-02 하단으로 수정
							extCommPrice	= (CLNG(rsXL(14)/extItemNo*100)/100) - (CLNG(rsXL(25)/extItemNo*100)/100)	'판매대행수수료
						Else
							IsReturnOrder = False
							extItemCost		= CLNG(rsXL(12)/extItemNo*100)/100
							'2021-04-02 하단으로 수정
'							extCommPrice	= CLNG(rsXL(14)/extItemNo*100)/100	'판매대행수수료
							extCommPrice	= (CLNG(rsXL(14)/extItemNo*100)/100) - (CLNG(rsXL(25)/extItemNo*100)/100)	'판매대행수수료
						End If

						if (rsXL(15) > 0) then
							extJungsanType			= "D"
							extItemNo				= 1
							extItemCost		= CLNG((rsXL(15)))
							extOrderserSeq = extOrderserSeq+"-D"
						else
							extJungsanType			= "C"
						end if

						extOwnCouponPrice = CLNG(rsXL(31)/extItemNo)
						extTenCouponPrice = CLNG(rsXL(32)/extItemNo)

						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extTenJungsanPrice		= extReducedPrice - extCommPrice
						dvlprice				= rsXL(15) 			'배송비

						extItemName				= rsXL(7)			'상품명
						extItemOptionName		= rsXL(8)			'옵션명
						extitemoption			= ""				'옵션 알수없음

						extVatYN = "Y"								'과세여부 알수없음
						extitemid = rsXL(9)

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

						IsValidInput = True
					Else
						IsValidInput = False
					End If
				Case "aboutpet"
					'// --------------------------------------------------------
					If (Len(rsXL(0)) = 16) then
						extOrderserial 	= rsXL(0)   							'주문번호
						extOrgOrderserial = ""
						extOrderserSeq = ""
						extCommPrice = 0
						extOwnCouponPrice = 0
						extTenCouponPrice = 0

'						extMeachulDate = LEFT(Replace(rsXL(0), "C", ""), 10)	'주문일 안 넘어옴..주문번호로 개조
'						extMeachulDate = LEFT(extMeachulDate,4)&"-"&MID(extMeachulDate,5,2)&"-"&MID(extMeachulDate,7,2)
						extMeachulDate = LEFT(rsXL(86), 10)						'배송완료일자

						extJungsanDate	= ""
						extItemNo		= rsXL(28)  							'수량
						extOrderserSeq	= extOrderserial & "-" & rsXL(1)		'주문번호 "-" 주문상세순번

						If (rsXL(28) < 0) Then									'수량 0이 아니면..케이스 아직 미발견
 							IsReturnOrder = True
 							extOrgOrderserial = extOrderserial
' 							extOrderserSeq	= extOrderserSeq & "-" & rsXL(1) & "-1" 
 							extItemNo		= -1
							extItemCost		= rsXL(22)								'판매가
							extCommPrice	= (rsXL(22) - rsXL(20)) 				'판매가 - 매입가
						Else
							IsReturnOrder = False
							extItemCost		= rsXL(22)								'판매가
							extCommPrice	= (rsXL(22) - rsXL(20)) 				'판매가 - 매입가
						End If

						extJungsanType			= "C"
						extOwnCouponPrice = CLNG((rsXL(35) / extItemNo * 100) / 100) 		'총 할인금액
						extCommPrice = extCommPrice - extOwnCouponPrice
						extTenCouponPrice = 0

						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extTenJungsanPrice		= extReducedPrice - extCommPrice
						dvlprice				= rsXL(44) 						'실배송비

						tmpItemname = ""
						tmpItemname = rsXL(18)

						If Instr(tmpItemname, "_") > 0 Then
							extItemName 		= Split(tmpItemname, "_")(0)
							extItemOptionName 	= Split(tmpItemname, "_")(1)
						Else
							extItemName = tmpItemname
							extItemOptionName = ""
						End If
						extitemoption			= ""				'옵션 알수없음

						extVatYN = "Y"								'과세여부 알수없음
						extitemid = ""

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

						IsValidInput = True
					Else
						IsValidInput = False
					End If
				Case "alphamallMaechul"
					'// --------------------------------------------------------
					'// alphamall 상세내역
					If (Len(rsXL(2)) = 21) then
						extOrderserial 	= rsXL(2)   				''주문번호
						extOrgOrderserial = ""
						extOrderserSeq = ""
						extCommPrice = ""
						extMeachulDate		= LEFT(rsXL(1),10)  				''일자

						extJungsanDate	= ""
						extItemNo		= rsXL(7)  					''수량
						extOrderserSeq	= extOrderserial &"-"& rsXL(0)

						extOwnCouponPrice = 0						'쿠폰금액 알수없음
						extTenCouponPrice = 0						'쿠폰금액 알수없음

						IsReturnOrder = False
						extItemCost	= rsXL(10)					'주문금액
						extCommPrice			= CLNG(rsXL(10)) - CLNG(rsXL(8))	'수수료 = 주문금액 - 주문원가

						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extJungsanType			= "C"

						extTenJungsanPrice		= extReducedPrice - extCommPrice
						dvlprice				= rsXL(12) 			'배송비

						extItemName				= rsXL(6)			'상품명
						extItemOptionName		= ""				'옵션명 알수없음
						extitemoption			= ""				'옵션 알수없음

						extVatYN = "Y"								'과세여부 알수없음
						extitemid = rsXL(5)							'상품번호

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

						IsValidInput = True
					Else
						IsValidInput = False
					End If
				Case "alphamallHuanBool"
					'// --------------------------------------------------------
					'// alphamall 상세내역
					If (Len(rsXL(2)) = 21) then
						extOrderserial 	= rsXL(2)   				''주문번호
						extOrgOrderserial = ""
						extOrderserSeq = ""
						extCommPrice = ""
						extMeachulDate		= LEFT(rsXL(1),10)  				''일자

						extJungsanDate	= ""
						extItemNo		= rsXL(7) * -1 					''수량
						extOrderserSeq	= extOrderserial &"-"& rsXL(0)

						extOwnCouponPrice = 0						'쿠폰금액 알수없음
						extTenCouponPrice = 0						'쿠폰금액 알수없음

						IsReturnOrder = True
						extOrgOrderserial = extOrderserial
						extOrderserSeq	= extOrderserSeq & "-" & rsXL(0)
						extItemCost	= rsXL(10)					'반품금액
						extCommPrice			= CLNG(rsXL(10)) - CLNG(rsXL(8))	'수수료 = 반품금액 - 반품원가

						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extJungsanType			= "C"

						extTenJungsanPrice		= extReducedPrice - extCommPrice
						dvlprice				= rsXL(12) 			'배송비

						extItemName				= rsXL(6)			'상품명
						extItemOptionName		= ""				'옵션명 알수없음
						extitemoption			= ""				'옵션 알수없음

						extVatYN = "Y"								'과세여부 알수없음
						extitemid = rsXL(5)							'상품번호

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

						IsValidInput = True
					Else
						IsValidInput = False
					End If
				Case "ezwel"
					'// --------------------------------------------------------
					'// 이지웰페어 상세내역
					'// 2019/11/01 순번이 생겼음 (6)
					if (Len(rsXL(4)) = 10) then
						extMeachulDate = rsXL(23+1)  ''배송완료일


						extJungsanDate = ""

						extItemNo				= rsXL(17+1)  ''수량

						if (rsXL(22+1) <> "") and (extItemNo<0) then ''취소일 수량이 마이너스인거만
							extMeachulDate = rsXL(22+1)
						end if

						extOrderserial 			= rsXL(4)  ''복지매장주문번호
						extOrderserSeq 			= rsXL(5) ''rsXL(8+1) & "-" & rsXL(0)			'// extOrderserSeq 이 없고, 상품코드가 있다, 상품코드 => extOrderserSeq 변환 필요 프로시져에서 일괄변경하자.

						if (rsXL(22+1) <> "") then ''취소일
							extOrderserSeq = extOrderserSeq&"-1"
						end if

						''일괄로 변경하자.
						' sqlStr = "select IsNull(min(OrgDetailKey), -1) as OrgDetailKey "
						' sqlStr = sqlStr + " from "
 						' sqlStr = sqlStr + " 	db_temp.dbo.tbl_xSite_TMPOrder o "
 						' sqlStr = sqlStr + " 	left join db_temp.dbo.tbl_xSite_JungsanTmp T "
 						' sqlStr = sqlStr + " 	on "
 						' sqlStr = sqlStr + " 		1 = 1 "
 						' sqlStr = sqlStr + " 		and T.sellsite = o.sellsite "
 						' sqlStr = sqlStr + " 		and T.extOrderserial = o.OutMallOrderSerial "
 						' sqlStr = sqlStr + " 		and T.extOrderserSeq = o.OrgDetailKey "
 						' sqlStr = sqlStr + " where o.sellsite = '" & sellsite & "' and OutMallOrderSerial = '" & extOrderserial & "' and outMallGoodsNo = '" & rsXL(8) & "' and T.sellsite is NULL"

						' ''response.write sqlstr & "<Br>"
						' rsget.Open sqlStr,dbget,1
						' 	extOrderserSeq = rsget("OrgDetailKey")
						' rsget.Close


						extOrgOrderserial		= ""
						extVatYN = "Y"

						extJungsanType			= "C"

						extItemName				= html2db(rsXL(9+1)) ''html2db(rsXL(3))
						dvlprice				= CLNG(rsXL(14+1))	'// 배송비를 상품금액과 같은 라인에 준다. 차후에 배송비내역 생성필요.

						extItemCost				= CLNG((rsXL(11+1) - dvlprice) / extItemNo*100)/100
						extTenMeachulPrice		= extItemCost ''쿠폰금액을 넣지 말자. 않맞음 2019/05/08
						''extTenMeachulPrice		= CLNG((rsXL(12) - dvlprice) / extItemNo*100)/100

						extOwnCouponPrice		= extItemCost - CLNG((rsXL(12+1) - dvlprice) / extItemNo*100)/100
						extTenCouponPrice		= extOwnCouponPrice*-1

						extTenJungsanPrice		= CLNG((rsXL(19+1) - dvlprice) / extItemNo*100)/100

						extReducedPrice			= CLNG(extTenMeachulPrice)
						extCommPrice			= extTenMeachulPrice - extTenJungsanPrice


						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

						extitemid = rsXL(8+1)
					else
						IsValidInput = False
					end if
				Case "homeplus"
					'// --------------------------------------------------------
					'// 홈플러스 정산 상세내역

                Case "kakaogift"
					'// --------------------------------------------------------
					'// kakaogift 정산 상세내역
					'// 톡딜 할인(26) 추가됨. 2019/05/22
					'// 판매자할인쿠폰(27) 추가됨 2019/12/19
					extOrderserial			= rsXL(2)

					if (Len(extOrderserial) = 9 or Len(extOrderserial) = 8 or Len(extOrderserial) = 10) and IsNumeric(extOrderserial) then

						extMeachulDate = rsXL(7)  ''정산기준일
						extJungsanDate = ""

						extOrderserSeq			= "" ''rsXL(5) ''없다... kakaoGIFT는 단품구매만 있다.

                        tenitemid =""
                        tenitemoption =""
                        extitemid = rsXL(15)  ''상품번호(kakaogift)
                        extitemoption = ""


                        extOrgOrderserial	= ""
						if (rsXL(13) <> "") then				'취소/환불일
							extOrderserSeq = "-1"
						end if
						' 	'// 반품
						' 	extOrgOrderserial	= ""  '' 어딘지 모름.
						' else
						' 	extOrgOrderserial	= ""
						' end if

                        '' 배송비가 따로 없다/ 상품대에 포함된다..;; => 이걸 어떻게 처리해야할까.
                        extJungsanType			= "C"
                        extItemNo				= rsXL(32)			'수량

                        extItemCost				= CLNG(rsXL(29) / extItemNo*100)/100 ''정산기준금액(28)  ''동일하다.
						extTenMeachulPrice		= extItemCost
						extReducedPrice			= CLNG(extTenMeachulPrice)

						extOwnCouponPrice		= 0                             '' 없음
						extTenCouponPrice       = 0                             '' 없음


						if rsXL(41) = 0 then		'수수료합계
							extCommPrice			= 0
						else
							extCommPrice			= CLNG(rsXL(41) / extItemNo*100)/100       ''수수료합(41)
						end if


                        extTenJungsanPrice      = CLNG(rsXL(42) / extItemNo*100)/100           ''판매정산금액
						extItemName				= html2db(rsXL(16))
						extItemOptionName		= html2db(rsXL(18))

						extVatYN = "Y"
						' if (rsXL(20) = rsXL(21)) and (rsXL(20) <> 0) then             ''없음.
						' 	extVatYN = "N"
						' end if

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

                        IsValidInput = True
					else
						IsValidInput = False
					end if
                Case "kakaostore"
					extOrderserial			= rsXL(0)			'결제번호

					if (Len(extOrderserial) = 9 or Len(extOrderserial) = 8 or Len(extOrderserial) = 10) and IsNumeric(extOrderserial) then

						extMeachulDate = rsXL(7)  ''정산기준일
						extJungsanDate = ""

						extOrderserSeq			= rsXL(2)		'주문번호

                        tenitemid =""
                        tenitemoption =""
                        extitemid = rsXL(15)  ''상품번호(kakaogift)
                        extitemoption = ""
						extOrgOrderserial	= ""
						if (rsXL(13) <> "") then
							extOrderserSeq = extOrderserSeq & "-1"
						end if
                        '' 배송비가 따로 없다/ 상품대에 포함된다..;; => 이걸 어떻게 처리해야할까.

						If (rsXL(3) = "배송비") Then
							extJungsanType			= "D"
							extItemNo				= 1
							extItemCost				= CLNG((rsXL(30))) + CLNG((rsXL(31)))		'선불배송비(30), 반품배송비(31)
							extOrderserSeq = "D-"&i
							extItemName				= rsXL(3)
						Else
                        	extJungsanType			= "C"
							extItemNo				= rsXL(32)					'수량
							extItemCost				= CLNG(rsXL(29) / extItemNo*100)/100 ''정산기준금액(29)  ''동일하다.
							extItemName				= html2db(rsXL(16))	'상품명
						End If

						extTenMeachulPrice		= extItemCost
						extReducedPrice			= CLNG(extTenMeachulPrice)

						extOwnCouponPrice		= 0                             '' 없음
						extTenCouponPrice       = 0                             '' 없음

						if rsXL(41) = 0 then		'수수료합계
							extCommPrice			= 0
						else
							extCommPrice			= CLNG(rsXL(41) / extItemNo*100)/100       ''수수료합(41)
						end if

                        extTenJungsanPrice      = CLNG(rsXL(42) / extItemNo*100)/100           ''판매정산금액
						extItemOptionName		= html2db(rsXL(18))	'옵션명

						extVatYN = "Y"
						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0
						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0
                        IsValidInput = True
					else
						IsValidInput = False
					end if
				Case "coupang"
					'// --------------------------------------------------------
					'// coupang 정산 상세내역
					'' 쿠팡스토어수수료할인추가됨 2019/10/28 (17)
					'' 우대수수료 필드추가됨 4 (19,20,21,22)
					extOrderserial			= rsXL(0)

					if (Len(extOrderserial) >= 13) and IsNumeric(extOrderserial) then

						extMeachulDate = rsXL(24+3+1+4)  ''구매확정일
						if (rsXL(24+3+1+1+4)<>"") then ''취소완료일
							extMeachulDate = rsXL(24+3+1+1+4)
						end if
						extJungsanDate = ""

						extOrderserSeq			= "" ''rsXL(5) ''없다...

                        tenitemid =""
                        tenitemoption =""
                        extitemid = rsXL(2)  ''Product ID
                        extitemoption = rsXL(6)	 ''Option ID

                        ''
                        extJungsanType			= "C"
						extItemNo = 0
						extReItemNo = 0
						extChgItemno = 0
						validitemno = 0
                        extItemNo				= rsXL(9)
						extReItemNo				= rsXL(10)
						extChgItemno			= rsXL(27+3+1+4)
						validitemno				= extItemNo-extReItemNo

						'' 판매수량(9) 과 환불수량(10) 이 있음  둘다 + 값임 ,교환수량도 있음, 무시.
						'' 교환적용일 (26)
                        extOrgOrderserial	= ""
						if (extReItemNo <> "0") then
							'// 반품
							extOrgOrderserial	= extOrderserial
						else
							extOrgOrderserial	= ""
						end if

						extOrderserSeq = extitemoption '' 단품ID를 넣자.

						''배송비부분.
						'' 17000021040462 반품의 케이스이나 기본배송료로 들어오는케이스가 있음. 9/22->9/13
						if (extitemoption = "<기본배송료>") or (extitemoption = "<추가배송료>") then
							extJungsanType			= "D"

							if (rsXL(11)<>0) or (rsXL(20+3+1+4)<>0) then   ''판매액(A) or 정산액이 있으면.
								validitemno = 1
								extItemNo   = 1
							else
								validitemno = 0
								extItemNo   = 0
							end if

							if (rsXL(11)<0) then  ''판매액이 마이너스이다..
								validitemno = -1
								extItemNo   = -1
							end if

							if (extitemoption = "<추가배송료>") then
								extOrderserSeq			= "D1"
								extItemName				= "배송비"
								extItemOptionName		= "추가배송료"
							else
								extOrderserSeq			= "D"
								extItemName				= "배송비"
								extItemOptionName		= "기본배송료"
							end if
						end if

						extOwnCouponPrice		= rsXL(12)                      '' 쿠팡지원할인(B) : 현재 없는듯.
						extTenCouponPrice       = rsXL(16)                      '' 판매자할인쿠폰 (몇개 있으나 계산방식이 좀이상.)
						if (extOwnCouponPrice="") then extOwnCouponPrice=0
						if (extTenCouponPrice="") then extTenCouponPrice=0
						extItemCost				= rsXL(8)						'' 판매가(8) 판매액(A)-11. :판매수량-환불수량=0 이면 판매액이 0 이다.
						if (extJungsanType<>"C") then '' 배송비는 판매액(A)으로 하자.
							extItemCost				= rsXL(11)
							if (validitemno<>0) then
								extItemCost = extItemCost/validitemno
							end if
						end if

						extItemCost = (extItemCost)
						extOwnCouponPrice = (extOwnCouponPrice)
						extTenCouponPrice = (extTenCouponPrice)

						extCommPrice			= rsXL(17+1)  ''rsXL(17)						''서비스이용료 10% 부가세별도.  쿠팡스토어수수료할인추가됨 2019/10/28

						if (validitemno<>0) then
							extOwnCouponPrice = CLNG(extOwnCouponPrice/validitemno*100)/100
							extTenCouponPrice = CLNG(extTenCouponPrice/validitemno*100)/100
							extCommPrice	  = CLNG(extCommPrice/validitemno*100)/100

							' if (extItemNo=0) and (extJungsanType="C") then
							' 	extItemNo = validitemno
							' end if

							if (extJungsanType="C") then
							 	extItemNo = validitemno
							end if
						end if

						extTenMeachulPrice		= extItemCost-extTenCouponPrice-extOwnCouponPrice
						extReducedPrice			= CLNG(extTenMeachulPrice) ''정산기준금액(26) ''매출금액(D=A-B) 판매자 할인쿠폰은?

                        extTenJungsanPrice      = extTenMeachulPrice-extCommPrice 			 ''rsXL(20)   ''판매정산금액 수량으로 나누어야..

						''구매확정일과 취수완료일이 다른것은 일자별로 다시 넣어야 한다..
						if (validitemno=0) and (extJungsanType="C") and (rsXL(13)=0) and (rsXL(17+1)=0) then
						 	extItemNo = 0
						end if

						if (validitemno<0)   then ''반품
							extOrderserSeq = extOrderserSeq&"-1"
						end if

						if (rsXL(26+3+1+4)<>"") then '' 교환적용일이 있다..
							if (extChgItemno<>0)   then ''교환 수량이 있으면
								extOrderserSeq = extOrderserSeq&"-2"
							else
								extOrderserSeq = extOrderserSeq&"-3"
							end if
						end if


						if (extJungsanType="C") then
							extItemName				= ""
							extItemOptionName		= ""
						end if

						extVatYN = "Y"
						if (extJungsanType="C") and (NOT rsXL(1)="TAX") then
							extVatYN = "N"
						end if

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

                        IsValidInput = True
			'rw extJungsanType&"|"&extItemNo&"|"&extCommPrice&"|"&extCommPrice&"|"&extTenMeachulPrice&"|"&extReducedPrice&"|"&extTenJungsanPrice&"|"&extMeachulDate
					else
						IsValidInput = False

						rw IsValidInput
					end if
				Case Else
					IsValidInput = False
			End Select

			if (IsValidInput = False) then
				Exit Do
			end if

			''response.write extOrgOrderserial & "---<br>"

			if (sellsite="ssg") or (sellsite="cjmall")  then
				if (((p_extOrderserial = extOrderserial) and (LEFT(p_extOrderserSeq,LEN(extOrderserSeq)) = extOrderserSeq)) or (sheetName="'대금지불현황(반품택배비)$'")) and (extItemNo<>0) then ''이정도의 조건으로 제한하자.
					'' SSG  20180719769649 중복케이스.
					''sqlStr = " select count(*) CNT from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite='"&sellsite&"' and extOrderserial='"&extOrderserial&"' and extOrderserSeq='" &extOrderserSeq&"'" &vbCRLF
					sqlStr = " select count(*) CNT from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite='"&sellsite&"' and extOrderserial='"&extOrderserial&"' and LEFT(extOrderserSeq,LEN('"&extOrderserSeq&"'))='" &extOrderserSeq&"'" &vbCRLF

					rsget.CursorLocation = adUseClient
					rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
					if Not rsget.Eof then
						if (rsget("CNT")>0) then
							extOrderserSeq = extOrderserSeq + "-"&rsget("CNT")
						end if
					end if
					rsget.Close
				end if
			end if

			''배송비 중복 케이스 있음. (20190412-296463)
			if (sellsite="hmall1010") and (extsellsite="hmallbeasongpay") then
				if ((p_extOrderserial = extOrderserial) and (LEFT(p_extOrderserSeq,LEN(extOrderserSeq)) = extOrderserSeq)) then
					sqlStr = " select count(*) CNT from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite='"&sellsite&"' and extOrderserial='"&extOrderserial&"' and LEFT(extOrderserSeq,LEN('"&extOrderserSeq&"'))='" &extOrderserSeq&"'" &vbCRLF

					rsget.CursorLocation = adUseClient
					rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
					if Not rsget.Eof then
						if (rsget("CNT")>0) then
							extOrderserSeq = extOrderserSeq + "-"&rsget("CNT")
						end if
					end if
					rsget.Close
				end if

			end if

			if (extsellsite="auction1010beasongpay") or (extsellsite="cjmallbeasongpay") then
				if ((p_extOrderserial = extOrderserial) and (LEFT(p_extOrderserSeq,LEN(extOrderserSeq)) = extOrderserSeq)) then
					sqlStr = " select count(*) CNT from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite='"&sellsite&"' and extOrderserial='"&extOrderserial&"' and LEFT(extOrderserSeq,LEN('"&extOrderserSeq&"'))='" &extOrderserSeq&"'" &vbCRLF

					rsget.CursorLocation = adUseClient
					rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
					if Not rsget.Eof then
						if (rsget("CNT")>0) then
							extOrderserSeq = extOrderserSeq + "-"&rsget("CNT")
						end if
					end if
					rsget.Close
				end if
			end if


			' if (extsellsite="lotteCombeasongpay") then
			' 	if ((p_extOrderserial = extOrderserial) and (extItemNo<>0)) then '' and (p_extOrderserSeq = extOrderserSeq)

			' 		'sqlStr = " select count(*) CNT from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite='"&sellsite&"' and extOrderserial='"&extOrderserial&"' and extOrderserSeq='" &extOrderserSeq&"'" &vbCRLF
			' 		sqlStr = " select count(*) CNT from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite='"&sellsite&"' and extOrderserial='"&extOrderserial&"' and LEFT(extOrderserSeq,LEN('"&extOrderserSeq&"'))='" &extOrderserSeq&"'" &vbCRLF

			' 		rsget.CursorLocation = adUseClient
			' 		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			' 		if Not rsget.Eof then
			' 			if (rsget("CNT")>0) then
			' 				extOrderserSeq = extOrderserSeq + "-"&rsget("CNT")
			' 			end if
			' 		end if
			' 		rsget.Close
			' 	end if
			' end if

            sqlStr = " insert into db_temp.dbo.tbl_xSite_JungsanTmp"
            sqlStr = sqlStr + " (sellsite, extOrderserial, extOrderserSeq"
            sqlStr = sqlStr + ", extOrgOrderserial, extItemNo, extItemCost"
            sqlStr = sqlStr + ", extReducedPrice, extOwnCouponPrice, extTenCouponPrice"
            sqlStr = sqlStr + ", extJungsanType, extCommPrice, extTenMeachulPrice"
            sqlStr = sqlStr + ", extTenJungsanPrice, extMeachulDate, extJungsanDate"
            sqlStr = sqlStr + ", extItemName, extItemOptionName, extVatYN"
            sqlStr = sqlStr + ", extCommSupplyPrice, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice"
            sqlStr = sqlStr + ", itemid, itemoption"
            sqlStr = sqlStr + ", extitemid, extitemoption,siteNo"
            sqlStr = sqlStr + " ) "
            sqlStr = sqlStr + " values('" + CStr(sellsite) + "', '" + CStr(extOrderserial) + "', '" + CStr(extOrderserSeq) + "'"
            sqlStr = sqlStr + ", '" + CStr(extOrgOrderserial) + "', '" + CStr(extItemNo) + "', '" + CStr(extItemCost) + "'"
            sqlStr = sqlStr + ", '" + CStr(extReducedPrice) + "', '" + CStr(extOwnCouponPrice) + "', '" + CStr(extTenCouponPrice) + "'"
            sqlStr = sqlStr + ", '" + CStr(extJungsanType) + "', '" + CStr(extCommPrice) + "', '" + CStr(extTenMeachulPrice) + "'"
            sqlStr = sqlStr + ", '" + CStr(extTenJungsanPrice) + "', '" + CStr(extMeachulDate) + "', '" + CStr(extJungsanDate) + "'"
            sqlStr = sqlStr + ", '" + CStr(extItemName) + "', convert(varchar(128),'" & CStr(extItemOptionName) & "'), '" + CStr(extVatYN) + "'"
            sqlStr = sqlStr + ", '" + CStr(extCommSupplyPrice) + "', '" + CStr(extCommSupplyVatPrice) + "', '" + CStr(extTenMeachulSupplyPrice) + "', '" + CStr(extTenMeachulSupplyVatPrice) + "'"
            if (tenitemid<>"") then
                sqlStr = sqlStr + ", '" + CStr(tenitemid) + "'"
            else
                sqlStr = sqlStr + ", NULL"
            end if
            if (tenitemoption<>"") then
                sqlStr = sqlStr + ", '" + CStr(tenitemoption) + "'"
            else
                sqlStr = sqlStr + ", NULL"
            end if
            if (extitemid<>"") then
                sqlStr = sqlStr + ", '" + CStr(extitemid) + "'"
            else
                sqlStr = sqlStr + ", NULL"
            end if
            if (extitemoption<>"") then
                sqlStr = sqlStr + ", '" + CStr(extitemoption) + "'"
            else
                sqlStr = sqlStr + ", NULL"
            end if
			if (siteNo<>"") then
                sqlStr = sqlStr + ", '" + CStr(siteNo) + "'"
            else
                sqlStr = sqlStr + ", NULL"
            end if
            sqlStr = sqlStr + ") "

			if (extItemNo<>0) then
				p_extOrderserial = extOrderserial
				p_extOrderserSeq = extOrderserSeq

				on Error Resume Next
            	dbget.execute sqlStr
				if Err then
					rw sqlStr
					on Error Goto 0
					response.end
				end if


			end if

			''if (sellsite="coupang") or (sellsite="ssg") then
			if (sellsite="ssg") or (sellsite="ezwel") or (sellsite="LFmall") or (sellsite="lotteon") or (sellsite="yes24") or (sellsite="ohou1010") or (extsellsite="alphamallMaechul") or (extsellsite="alphamallHuanBool") or (extsellsite="aboutpet") then
				''validitemno
				' if (sellsite="coupang") then ''반품이 별도필드로 있다.
				' 	if (CStr(extReItemNo)<>"0") then
				' 		extItemNo=extReItemNo*-1
				' 		extOrderserial = extOrderserial&"-1"
				' 		''extOrderserSeq = extOrderserSeq&"-1"

				' 		''rw  extOrderserial
				' 	end if
				' end if
				If (sellsite="yes24") Then
					if (dvlprice<>0) then  '배송비가 한줄로 있다.
						extJungsanType = "D"
						extOrderserSeq = extOrderserSeq+"-D"
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extCommPrice			= 0

						If (rsXL(4) <> 0) Then						''반품금액이 0보다 크다면
							extItemNo = -1
						Else
							extItemNo = 1
							' extCommPrice = rsXL(4) - rsXL(6)
							' extTenJungsanPrice	= dvlprice/extItemNo - extCommPrice
						End If

						extItemCost				= dvlprice/extItemNo
						extTenJungsanPrice		= dvlprice/extItemNo


						extReducedPrice			= dvlprice/extItemNo-extOwnCouponPrice
						extTenMeachulPrice      = dvlprice/extItemNo-extOwnCouponPrice

						extCommSupplyPrice		= 0
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= 0
						extTenMeachulSupplyVatPrice	= 0
					else
						extItemNo = 0
					end if
				End If

				If (sellsite="ohou1010") Then
					if (dvlprice<>0) then  '배송비가 한줄로 있다.
						extJungsanType = "D"
						extOrderserSeq = extOrderserSeq+"-D"
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extCommPrice			= 0
						
						If (extItemNo < 0) Then						''반품금액이 0보다 크다면
							extItemNo = -1
						Else
							extItemNo = 1
							' extCommPrice = rsXL(4) - rsXL(6)
							' extTenJungsanPrice	= dvlprice/extItemNo - extCommPrice
						End If

						extCommPrice			= CLNG(rsXL(13) / extItemNo )		'수수료D
						extItemCost				= dvlprice/extItemNo
						extTenJungsanPrice		= (dvlprice - extCommPrice) / extItemNo

						extReducedPrice			= dvlprice/extItemNo-extOwnCouponPrice
						extTenMeachulPrice      = dvlprice/extItemNo-extOwnCouponPrice

						extCommSupplyPrice		= 0
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= 0
						extTenMeachulSupplyVatPrice	= 0
					else
						extItemNo = 0
					end if
				End If

				If (sellsite="aboutpet") Then
					if (dvlprice<>0) then  '배송비가 한줄로 있다.
						extJungsanType = "D"
						extOrderserSeq = extOrderserSeq+"-D"
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extCommPrice			= 0

						If (extItemNo < 0) Then						''반품금액이 0보다 크다면
							extItemNo = -1
						Else
							extItemNo = 1
							' extCommPrice = rsXL(4) - rsXL(6)
							' extTenJungsanPrice	= dvlprice/extItemNo - extCommPrice
						End If

						extItemCost				= dvlprice/extItemNo
						extTenJungsanPrice		= dvlprice/extItemNo


						extReducedPrice			= dvlprice/extItemNo-extOwnCouponPrice
						extTenMeachulPrice      = dvlprice/extItemNo-extOwnCouponPrice

						extCommSupplyPrice		= 0
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= 0
						extTenMeachulSupplyVatPrice	= 0
					else
						extItemNo = 0
					end if
				End If

				If (extsellsite="alphamallMaechul") Then
					if (dvlprice<>0) then  '배송비가 한줄로 있다.
						extJungsanType = "D"
						extOrderserSeq = extOrderserSeq+"-D"
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extCommPrice			= 0

						extItemNo = 1

						extItemCost				= dvlprice/extItemNo
						extTenJungsanPrice		= dvlprice/extItemNo


						extReducedPrice			= dvlprice/extItemNo-extOwnCouponPrice
						extTenMeachulPrice      = dvlprice/extItemNo-extOwnCouponPrice

						extCommSupplyPrice		= 0
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= 0
						extTenMeachulSupplyVatPrice	= 0
					else
						extItemNo = 0
					end if
				End If

				If (extsellsite="alphamallHuanBool") Then
					if (dvlprice<>0) then  '배송비가 한줄로 있다.
						extJungsanType = "D"
						extOrderserSeq = extOrderserSeq+"-D"
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extCommPrice			= 0

						extItemNo = -1

						extItemCost				= dvlprice/1
						extTenJungsanPrice		= dvlprice/1


						extReducedPrice			= dvlprice/1-  extOwnCouponPrice
						extTenMeachulPrice      = dvlprice/1 - extOwnCouponPrice

						extCommSupplyPrice		= 0
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= 0
						extTenMeachulSupplyVatPrice	= 0
					else
						extItemNo = 0
					end if
				End If

				If (sellsite="lotteon") Then
					If (dvlprice <> 0) Then  '배송비가 한줄로 있다.
						extJungsanType = "D"
						extOrderserSeq = extOrderserSeq+"-D"
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extCommPrice			= 0

						If extItemNo = 0 AND dvlprice > 0 Then
							extItemNo = 1
							extCommPrice		= lotteondvlTotCommprice
							extTenJungsanPrice	= dvlprice/extItemNo - lotteondvlPGCommprice
						Else
							If (dvlprice < 0) Then
								extItemNo = -1
								extCommPrice		= dlvCommprice
								If extOrderserial = "2022112916918870" or extOrderserial = "2022122010117929" Then
									extCommPrice		= dlvCommprice * -1
								End If
							Else
								extItemNo = 1
								'extCommPrice		= (lotteondvlprice2) * -1
								extCommPrice		= dlvCommprice
							End If
							extTenJungsanPrice		= (dvlprice - dlvCommprice) / extItemNo
						End If

						extItemCost				= (dvlprice - lotteondvlprice2) /extItemNo
						extReducedPrice			= (dvlprice - lotteondvlprice2)/extItemNo-extOwnCouponPrice
						extTenMeachulPrice      = (dvlprice - lotteondvlprice2)/extItemNo-extOwnCouponPrice

						extCommSupplyPrice		= 0
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= 0
						extTenMeachulSupplyVatPrice	= 0
					Else
						extItemNo = 0
					End If
				End If

				if (sellsite="ssg") then
					if (dvlprice<>0) then  '배송비가 한줄로 있다.
						extJungsanType = "D"
						extOrderserSeq = extOrderserSeq+"-D"

						if (extItemNo<0) then
							extItemNo = -1
							extOwnCouponPrice		= 0
							extTenCouponPrice		= 0
							extCommPrice			= 0
						elseif (extItemNo=0) then
							extItemNo = 1
							extOrderserSeq = extOrderserSeq+"D"
							if (dvlprice<0) then extOrderserSeq = extOrderserSeq+"D"

							'extOwnCouponPrice		= 0   ''20180807467429 CASE
							'extTenCouponPrice		= 0
							extCommPrice			= extOwnCouponPrice*-1   ''20180807467429 CASE
						else
							extItemNo = 1
							extOwnCouponPrice		= 0
							extTenCouponPrice		= 0
							extCommPrice			= 0
						end if


						extItemCost				= dvlprice/extItemNo
						extTenJungsanPrice		= dvlprice/extItemNo


						extReducedPrice			= dvlprice/extItemNo-extOwnCouponPrice
						extTenMeachulPrice      = dvlprice/extItemNo-extOwnCouponPrice

						''extVatYN = "Y" 비교를 쉽게하기위해 따라가자
						''if (rsXL(6) = "면세") then
						''	extVatYN = "N"
						''end if

						extCommSupplyPrice		= 0
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= 0
						extTenMeachulSupplyVatPrice	= 0
					else
						extItemNo = 0
					end if
				end if

				if (sellsite="ezwel") then
					if (dvlprice<>0) then  '배송비가 한줄로 있다.
						extJungsanType = "D"
						extOrderserSeq = extOrderserSeq+"-D"
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extCommPrice			= 0

						if (dvlprice<0) then
							extItemNo = -1
						else
							extItemNo = 1
						end if

						extItemCost				= dvlprice/extItemNo
						extTenJungsanPrice		= dvlprice/extItemNo


						extReducedPrice			= dvlprice/extItemNo-extOwnCouponPrice
						extTenMeachulPrice      = dvlprice/extItemNo-extOwnCouponPrice

						extCommSupplyPrice		= 0
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= 0
						extTenMeachulSupplyVatPrice	= 0
					else
						extItemNo = 0
					end if
				end if

				if (sellsite="LFmall") then
					if (dvlprice<>0) then  '배송비가 한줄로 있다.
						extJungsanType = "D"
						extOrderserSeq = extOrderserSeq+"-D"
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extCommPrice			= 0

						if (dvlprice<0) then
							extItemNo = -1
						else
							extItemNo = 1
						end if

						extItemCost				= dvlprice/extItemNo
						extTenJungsanPrice		= dvlprice/extItemNo


						extReducedPrice			= dvlprice/extItemNo-extOwnCouponPrice
						extTenMeachulPrice      = dvlprice/extItemNo-extOwnCouponPrice

						extCommSupplyPrice		= 0
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= 0
						extTenMeachulSupplyVatPrice	= 0
					else
						extItemNo = 0
					end if
				end if

				if (extItemNo<>0) then

					'' SSG는 배송비도 안분하는듯 동일한 내역이 있으면 더하자.
					if (sellsite="ssg") and (RIGHT(extOrderserSeq,1)="D") then
						sqlStr = " select count(*) CNT from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite='ssg' and extOrderserial='"&extOrderserial&"' and extOrderserSeq='" &extOrderserSeq&"'" &vbCRLF
						rsget.CursorLocation = adUseClient
						rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
						if Not rsget.Eof then
							if (rsget("CNT")>0) then
								extOrderserSeq = extOrderserSeq + "-"&rsget("CNT")
							end if
						end if
						rsget.Close
					end if


					sqlStr = " insert into db_temp.dbo.tbl_xSite_JungsanTmp"
					sqlStr = sqlStr + " (sellsite, extOrderserial, extOrderserSeq"
					sqlStr = sqlStr + ", extOrgOrderserial, extItemNo, extItemCost"
					sqlStr = sqlStr + ", extReducedPrice, extOwnCouponPrice, extTenCouponPrice"
					sqlStr = sqlStr + ", extJungsanType, extCommPrice, extTenMeachulPrice"
					sqlStr = sqlStr + ", extTenJungsanPrice, extMeachulDate, extJungsanDate"
					sqlStr = sqlStr + ", extItemName, extItemOptionName, extVatYN"
					sqlStr = sqlStr + ", extCommSupplyPrice, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice"
					sqlStr = sqlStr + ", itemid, itemoption"
					sqlStr = sqlStr + ", extitemid, extitemoption,siteNo"
					sqlStr = sqlStr + " ) "
					sqlStr = sqlStr + " values('" + CStr(sellsite) + "', '" + CStr(extOrderserial) + "', '" + CStr(extOrderserSeq) + "'"
					sqlStr = sqlStr + ", '" + CStr(extOrgOrderserial) + "', '" + CStr(extItemNo) + "', '" + CStr(extItemCost) + "'"
					sqlStr = sqlStr + ", '" + CStr(extReducedPrice) + "', '" + CStr(extOwnCouponPrice) + "', '" + CStr(extTenCouponPrice) + "'"
					sqlStr = sqlStr + ", '" + CStr(extJungsanType) + "', '" + CStr(extCommPrice) + "', '" + CStr(extTenMeachulPrice) + "'"
					sqlStr = sqlStr + ", '" + CStr(extTenJungsanPrice) + "', '" + CStr(extMeachulDate) + "', '" + CStr(extJungsanDate) + "'"
					sqlStr = sqlStr + ", '" + CStr(extItemName) + "', convert(varchar(128),'" & CStr(extItemOptionName) & "'), '" + CStr(extVatYN) + "'"
					sqlStr = sqlStr + ", '" + CStr(extCommSupplyPrice) + "', '" + CStr(extCommSupplyVatPrice) + "', '" + CStr(extTenMeachulSupplyPrice) + "', '" + CStr(extTenMeachulSupplyVatPrice) + "'"
					if (tenitemid<>"") then
						sqlStr = sqlStr + ", '" + CStr(tenitemid) + "'"
					else
						sqlStr = sqlStr + ", NULL"
					end if
					if (tenitemoption<>"") then
						sqlStr = sqlStr + ", '" + CStr(tenitemoption) + "'"
					else
						sqlStr = sqlStr + ", NULL"
					end if
					if (extitemid<>"") then
						sqlStr = sqlStr + ", '" + CStr(extitemid) + "'"
					else
						sqlStr = sqlStr + ", NULL"
					end if
					if (extitemoption<>"") then
						sqlStr = sqlStr + ", '" + CStr(extitemoption) + "'"
					else
						sqlStr = sqlStr + ", NULL"
					end if
					if (siteNo<>"") then
						sqlStr = sqlStr + ", '" + CStr(siteNo) + "'"
					else
						sqlStr = sqlStr + ", NULL"
					end if
					sqlStr = sqlStr + ") "
					dbget.execute sqlStr
				end if
			end if

		end if
		i = i + 1
		rsXL.MoveNext
	loop
	Call checkAndWriteElapsedTime("004")
end if
rsXL.Close
set rsXL = Nothing
set conXL = Nothing

if (IsValidInput = False) then
	uprequest("sFile").delete
	set objfso = Nothing
	set uprequest = Nothing
	response.write "ERROR : 오류가 발생했습니다. 시스템팀 문의[3]" & errMSG & extsellsite
	dbget.close
	response.end
end if


uprequest("sFile").delete
set objfso = Nothing
set uprequest = Nothing


Dim AssignedRow
if (sellsite="kakaogift") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_kakaogift]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="kakaostore") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_kakaostore]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="coupang") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_coupang]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="11st1010") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_11st1010]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="ssg") then
	'' XL로 업로드 할때 기존내역을 업어칠지 결정하자..
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_ssg] 1"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="cjmall") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_cjmall]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="gmarket1010") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_gmarket]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="auction1010") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_auction]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="ezwel") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_ezwel]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="nvstorefarm") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_nvstorefarm]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="nvstorefarmclass") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_nvstorefarmclass]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="nvstoremoonbangu") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_nvstoremoonbangu]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="Mylittlewhoopee") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_Mylittlewhoopee]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="nvstoregift") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_nvstoregift]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="wadsmartstore") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_wadsmartstore]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="lotteCom") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_lotteCom]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="halfclub") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_halfclub]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="gseshop") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_gseshop]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="lotteimall") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_lotteimall]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="hmall1010") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_hmall]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="WMP") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_WMP]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="wmpfashion") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_wmpfashion]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="LFmall") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_LFmall]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="lotteon") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_lotteon]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="yes24") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_yes24]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="ohou1010") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_ohou1010]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="alphamall") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_alphamall]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="casamia_good_com") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_casamia]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="aboutpet") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_aboutpet]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="shintvshopping") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_shintvshopping]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="skstoa") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_skstoa]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="wetoo1300k") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_wetoo1300k]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="boribori1010") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_boribori]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="wconcept1010") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_wconcept]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="GS25") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_GS25]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="withnature1010") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_withnature1010]"
	dbget.Execute sqlStr, AssignedRow
elseif (sellsite="goodshop1010") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_goodshop1010]"
	dbget.Execute sqlStr, AssignedRow
else
	rw "TT"
end if

%>
<script>
alert("저장되었습니다. ");
location.href = "<%= manageUrl %>/common/popReloadOpener.asp";
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
