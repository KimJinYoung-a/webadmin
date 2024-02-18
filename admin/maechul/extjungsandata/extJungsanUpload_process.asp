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

''���޸� ���곻�� �߰����
''ADD EXT SHOP �˻��Ͽ� �߰��Ѵ�.

''response.write "�۾���.."
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
'// ���ε� ���۳�Ʈ ���� //
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
		conXL.Properties("ExtEnded Properties").Value = "Excel 12.0;HDR=NO;IMEX=1;TypeGuessRows=0;ImportMixedTypes=Text;"  '' ���÷��� text�� money�� ȥ��Ǿ� ����.
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
	response.write "ERROR : ������ �߻��߽��ϴ�. �ý����� ����[0]"
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
	response.write "ERROR : ������ �߻��߽��ϴ�. �ý����� ����[1]"
	response.end
end if

Call checkAndWriteElapsedTime("002")

rw extsellsite
If Not rsXL.Eof Then

	''ADD EXT SHOP. 01. ����Ʈ����
	Select Case extsellsite
		Case "interparkproduct"
			'// ������ũ ��ǰ���� �󼼳���
			sellsite = "interpark"
		''\upload\linkweb\extjungsandata\extJungsanUpload_process.asp ' ������ ���⼭ Ȯ��.
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
		response.write "ERROR : ������ �߻��߽��ϴ�. �ý����� ����[2]"
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
		'// ����Ÿ ���� 01 (�ֹ���������)
        ''\upload\linkweb\extjungsandata\extJungsanUpload_process.asp ' ������ ���⼭ Ȯ��.
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

				if (i = 0) and (rsXL(0) <> "�Ⱓ������������Ȳ - I-Point����") then
					IsValidInput = False
					exit do
				end if
			Case "interparkpointmall"
				if (IsNumeric(rsXL(0)) = True) and (rsXL(0) <> "0") then
					IsOrderData = True
				end if

				if (i = 0) and (rsXL(0) <> "�Ⱓ������������Ȳ - ����Ʈ���Ǹż�����") then
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
				If (uprequest("cjbeasongGubun") = "1") OR uprequest("cjbeasongGubun") = "2" Then	'1 : ���������Ȳ(��ȯ�ù��) / 2 : ���������Ȳ(��ǰ�ù��)
					if (rsXL(4) <> "") then
						if (IsNumeric(rsXL(4)) = True) then
							IsOrderData = True
						end if
					end if
				ElseIf uprequest("cjbeasongGubun") = "3" Then										'3 : ���������Ȳ(���ܰ���ۺ�)
					if (rsXL(5) <> "") then
						if (IsNumeric(rsXL(5)) = True) then
							IsOrderData = True
						end if
					end if
				End If
				' if (sheetName="'���������Ȳ(���ܰ���ۺ�)$'") or (sheetName="sheet1$") then  ''(sheetName="sheet1$") : ����������
				' 	if (rsXL(4) <> "") then
				' 		if (IsNumeric(rsXL(4)) = True) then
				' 			IsOrderData = True
				' 		end if
				' 	end if
				' elseif (sheetName="'���������Ȳ(��ǰ�ù��)$'") or (sheetName="'���������Ȳ(��ȯ�ù��)$'") then
				' 	if (rsXL(3) <> "") then
				' 		if (IsNumeric(rsXL(3)) = True) then
				' 			IsOrderData = True
				' 		end if
				' 	end if
				' elseif (sheetName="'���������Ȳ(AS �ù��)$'") then
				' 	if (rsXL(7) <> "") then
				' 		if (IsNumeric(rsXL(7)) = True) then ''��ǰ�ڵ�
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
				if (rsXL(0) <> "") then  ''20200110 ��ȣ�ʵ��߰��� '' 20200301 rsXL(1) ���°�� ����.
					if (IsNumeric(rsXL(0)) = True) then ''��ٱ��Ϲ�ȣ
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
					if rsXL(1) = "��ۿϷ���" then
						IsOrderData = True
					end if
				end if
			Case "nvstorefarm", "nvstorefarmclass", "nvstoremoonbangu", "nvstoregift", "wadsmartstore", "Mylittlewhoopee"
				if (rsXL(0) <> "") then
					if (rsXL(13) = "�Ϲ�����") or (rsXL(13) = "������ ���") or (rsXL(13) = "��������") or (rsXL(13) = "�������� ȸ��") then
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
					if (Len(rsXL(0)) = 10) and (IsNumeric(rsXL(1)) = True) and rsXL(9) <> "�߰�����" then
						IsOrderData = True
					end if
				end if
			Case Else
				IsOrderData = False
		End Select

		if (IsOrderData = True) then

			'//ADD EXT SHOP. 03. ó��
            ''\upload\linkweb\extjungsandata\extJungsanUpload_process.asp ' ������ ���⼭ Ȯ��.
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
					'// �������� ��ǰ���� �󼼳���
					extOrderserial = rsXL(7)	'�ֹ���ȣ
					If Len(extOrderserial) = 12 Then
						extJungsanDate = ""
						extMeachulDate = LEFT(rsXL(6), 4) & "-" & MID(rsXL(6), 5, 2) & "-" & MID(rsXL(6), 7, 2)		'���/������
						extOrderserSeq = extMeachulDate & "-" & i  ''Seq�� �������� ���Ѵ�.
						extItemNo = Trim(rsXL(15))		'������
						extOrgOrderserial = ""
						If (extItemNo <= 0) Then
							extitemNo = -1
							IsReturnOrder = True
							extOrgOrderserial = extOrderserial
							extOrderserial = extOrderserial&"-1"
							extItemCost			= rsXL(18)	'��ǰ�ݾ�
						Else
							extItemCost			= rsXL(16)	'���ݾ�
							IsReturnOrder = False
						End If

						' extTenCouponPrice	= CLNG((CLNG(rsXL(23)) - CLNG(rsXL(24))) / extitemNo)	'���� ���� �δ��(23) - ��ǰ ���� �δ��(24)
						' extOwnCouponPrice	= 0
						' extTenCouponPrice	= 0
						extTenCouponPrice	= CLNG((CLNG(rsXL(26)) - CLNG(rsXL(27))) / extitemNo)	'���� ���� �δ�� - ��ǰ ���� �δ��..2023-08-01 ������ �߰�
						extOwnCouponPrice	= CLNG((CLNG(rsXL(23))) / extitemNo)	'���δ� ���ξ�(B)(23)
						extTenMeachulPrice	= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice		= CLng(extTenMeachulPrice)
						extJungsanType		= "C"
						extCommPrice		= CLNG((CLNG(rsXL(22)) - CLNG(rsXL(23))) / extitemNo)	'������ݾ�(A)(22) - ���δ� ���ξ�(B)(23)
						extTenJungsanPrice	= CLNG(rsXL(28) / extItemNo * 100)/100		'�����ޱݾ�(28)
						extTenJungsanPrice	= extReducedPrice - extCommPrice
						extItemName			= html2db(rsXL(10))			'��ǰ��
						extItemOptionName	= html2db(rsXL(11))			'�ɼǸ�
						extVatYN = "Y"
						If rsXL(0) = "������(����)" OR rsXL(0) = "������(����)" Then	'����
							extVatYN = "Y"
						Else
							extVatYN = "N"
						End If
						extitemid = rsXL(13)	'��ü��ǰ�ڵ�
					Else
						IsValidInput = False
					End If
				Case "boriboribeasongpay"
					'// �������� ��ۺ� �󼼳���
					extOrderserial = rsXL(7)
					If Len(rsXL(7)) = 12 Then
						extMeachulDate = LEFT(rsXL(0), 4) & "-" & MID(rsXL(0), 5, 2) & "-" & MID(rsXL(0), 7, 2)
						extJungsanDate = ""
						extVatYN = "Y"
						extOrgOrderserial = ""
						extItemNo				= 1
						extItemName				= "��ۺ�"
						extOrderserSeq			= "D-"&i
						extItemCost				= CLNG(rsXL(8)) + CLNG(rsXL(11))	'�ֹ���ۺ�(A) + �ֹ� �����갣 �߰���
						If (rsXL(14) <> 0) Then		'��ǰ��ۺ�
							extItemNo = 1
							extItemName			= "��ǰ��ۺ�"
							extItemCost			= CLNG(rsXL(12)) + CLNG(rsXL(13)) + CLNG(rsXL(14)) + CLNG(rsXL(15))		'��ҹ�ۺ� + ��� �����갣 �߰��� + ��ǰ��ۺ� + ��ǰ �����갣 �߰���
							extOrderserSeq		= "DD-"&i
						ElseIf (rsXL(12) <> 0) Then		'��ǰ��ۺ�
							extItemNo = 1
							extItemName			= "��ҹ�ۺ�"
							extItemCost			= CLNG(rsXL(12)) + CLNG(rsXL(13))		'��ҹ�ۺ� + ��� �����갣 �߰���
							extOrderserSeq		= "DDD-"&i
						End If
						extOwnCouponPrice		= 0
						extTenCouponPrice		= rsXL(9)		'��ü�δ�������(B)
						extJungsanType			= "D"
						extCommPrice			= 0
						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extTenJungsanPrice		= extReducedPrice - extCommPrice
						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0
						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0
						''�ݾ��� ������ ���� �ʴ´�.
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
						extOrderserSeq			= extMeachulDate&"-"&i  ''Seq�� �������� ���Ѵ�.
						extItemNo				= Trim(rsXL(8))

						if (extItemNo <= 0) then
							IsReturnOrder = True
						else
							IsReturnOrder = False
						end if

						extItemCost				= rsXL(7)	'��ǰ�ܰ�
						extOwnCouponPrice		= CLNG(CLNG(rsXL(12)) / extItemNo) 'rsXL(12)	'�����ݾ�
						extTenCouponPrice		= CLNG(CLNG(rsXL(13)) / extItemNo) 'rsXL(13)	'�����ݾ� ��ü�δ�

						extOwnCouponPrice = 0	'2022-05-03 ������ ����..���꿡 �����ݾ��� �ִ� ��, ����� / ������ ��� �����ݾ� 74,716 �����Ǿ� �ִ� �����ݾ� �Դϴ�...

						extTenMeachulPrice		= extItemCost-extOwnCouponPrice-extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extJungsanType			= "C"

						extCommPrice			= CLNG(CLNG(rsXL(11)) / extItemNo)
						extTenJungsanPrice		= extReducedPrice - extCommPrice

						extItemName				= html2db(rsXL(6))
						If Instr(extItemName, "// {�ɼ�}") > 0 Then
							extItemOptionName		= Split(extItemName, "// {�ɼ�}")(1)
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
						extItemName				= "��ۺ�"
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

						''�ݾ��� ������ ���� �ʴ´�.
						if (extItemCost=0) then
							extItemNo = 0
						end if
					else
						IsValidInput = False
					end if
				Case "shintvshopping"
					extOrderserial = Replace(rsXL(1), "-", "")						'�ֹ���ȣ
					if (Len(extOrderserial) = 14) then
						extJungsanDate = ""
						extMeachulDate = Replace(rsXL(3), "/", "-")					'��������
						extOrderserSeq = extMeachulDate&"-"&i  						'Seq�� �������� ���Ѵ�.
						extItemNo = Trim(rsXL(24))									'�ֹ�����

						If (extItemNo <= 0) Then
							extItemCost				= rsXL(18)*-1					'�ǸŰ�
							IsReturnOrder = True
						Else
							extItemCost				= rsXL(18)						'�ǸŰ�
							IsReturnOrder = False
						End If

						If (extItemNo = 0) Then
							extItemNo = 1
						End If

						extOwnCouponPrice		= CLNG(CLNG(rsXL(29)) / extItemNo) + CLNG(CLNG(rsXL(30)) / extItemNo) '�ż���Ƽ����� �� + ���� ��
						extTenCouponPrice		= CLNG(CLNG(rsXL(31)) / extItemNo)	'��ü ��
						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extJungsanType			= "C"
						extCommPrice			= CLNG(CLNG(rsXL(36)) / extItemNo)
						extTenJungsanPrice		= extReducedPrice - extCommPrice
						extItemName				= html2db(rsXL(14))
						extItemOptionName		= html2db(rsXL(16))
						If rsXL(17) = "����" Then									'��������
							extVatYN = "Y"
						Else
							extVatYN = "N"
						End If
                        extitemid				= rsXL(13)							'��ǰ�ڵ�
                        extitemoption			= rsXL(15)							'��ǰ�ڵ�
					else
						IsValidInput = False
					End If
				Case "shintvshoppingbeasongpay"
					extOrderserial			= rsXL(1)				'�ֹ���ȣ

					if (Len(extOrderserial) = 14) then

						'// �ż���TV���� ��ۺ��� �������� �Ѿ���� ����..�ֹ���ȣ�� �߶� ����..�׷��� �����ϰ� ���̴� �ִ�.
						extMeachulDate = LEFT(rsXL(1), 4) & "-" & MID(rsXL(1), 5, 2) & "-" & MID(rsXL(1), 7, 2)
						extJungsanDate = ""

						If (extMeachulMonth <> LEFT(extMeachulDate,7)) Then
							extMeachulDate = extMeachulMonth+"-01"
						End If

						extVatYN = "Y"

						extOrgOrderserial = ""

						extItemNo				= 1
						extItemName				= "��ۺ�"
						extOrderserSeq			= "D-"&i 

						extItemCost				= rsXL(6)						'��ۺ�
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

						''�ݾ��� ������ ���� �ʴ´�.
						if (extItemCost=0) then
							extItemNo = 0
						end if
					else
						IsValidInput = False
					end if
				Case "goodwearmall10"
					extOrderserial = Replace(rsXL(0), "-", "")						'�ֹ���ȣ
					If (Len(extOrderserial) = 17) Then

					Else
						IsValidInput = False
					End If
				Case "goodwearmall10beasongpay"
					extOrderserial = Replace(rsXL(4), "-", "")						'�ֹ���ȣ
					If (Len(extOrderserial) = 17) Then

					Else
						IsValidInput = False
					End If
				Case "skstoa"
					extOrderserial = Replace(rsXL(3), "-", "")						'�ֹ���ȣ
					if (Len(extOrderserial) = 14) then
						extJungsanDate = ""
						extMeachulDate = Replace(rsXL(11), "/", "-")				'�Ϸ���
						extOrderserSeq = extMeachulDate&"-"&i  						'Seq�� �������� ���Ѵ�.
						extItemNo = Trim(rsXL(15))									'�ֹ�����
						extItemCost				= rsXL(13)							'�ǸŰ�
						If (extItemNo <= 0) Then
							IsReturnOrder = True
						Else
							IsReturnOrder = False
						End If

						If (extItemNo = 0) Then
							extItemNo = 1
						End If

						extOwnCouponPrice		= CLNG(CLNG(rsXL(30)) / extItemNo) 	'������θ��
						extTenCouponPrice		= CLNG(CLNG(rsXL(29)) / extItemNo)	'��ü���θ��
						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extJungsanType			= "C"
						extCommPrice			= CLNG(CLNG(rsXL(35)) / extItemNo)	'����Ź������(�հ�)
						extTenJungsanPrice		= extReducedPrice - extCommPrice
						extItemName				= html2db(rsXL(7))
						extItemOptionName		= html2db(rsXL(8))
						If rsXL(10) = "����" Then									'��������
							extVatYN = "Y"
						Else
							extVatYN = "N"
						End If
                        extitemid				= rsXL(5)							'��ǰ�ڵ�
					else
						IsValidInput = False
					End If
				Case "skstoabeasongpay"
					extOrderserial			= rsXL(3)				'�ֹ���ȣ

					if (Len(extOrderserial) = 14) then

						'// skstoa ��ۺ��� �������� �Ѿ���� ����..�ֹ���ȣ�� �߶� ����..�׷��� �����ϰ� ���̴� �ִ�.
						extMeachulDate = LEFT(rsXL(3), 4) & "-" & MID(rsXL(3), 5, 2) & "-" & MID(rsXL(3), 7, 2)
						extJungsanDate = ""

						If (extMeachulMonth <> LEFT(extMeachulDate,7)) Then
							extMeachulDate = extMeachulMonth+"-01"
						End If

						extVatYN = "Y"

						extOrgOrderserial = ""

						extItemNo				= 1
						extItemName				= "��ۺ�"
						extOrderserSeq			= "D-"&i 

						extItemCost				= rsXL(7)						'��ۺ�
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

						''�ݾ��� ������ ���� �ʴ´�.
						if (extItemCost=0) then
							extItemNo = 0
						end if
					else
						IsValidInput = False
					end if
				Case "WMP", "wmpfashion"
					extOrderserial			= rsXL(1+1) ''��۹�ȣ(�̰� �ֹ���ȣ)
					if (Len(extOrderserial) = 8 or Len(extOrderserial) = 9) and IsNumeric(rsXL(12+1)) then

						extJungsanDate = ""
						extMeachulDate			= rsXL(28) ''�����
						extOrderserSeq			= rsXL(2+1)  ''�ֹ���ȣ(Seq�� ����.)
						'extOrderserSeq			= rsXL(5+1)  ''�ɼ��ֹ���ȣ.. 2020-10-28 ������ / �ɼ��ֹ���ȣ�� ����
						If rsXL(5+1) <> "0" then
							extOrderserSeq = extOrderserSeq & "-" & rsXL(5+1)
						End If

						if (rsXL(11+1) <> "���") then  '' ȯ��
							'// ��ǰ
						 	extOrgOrderserial		= extOrderserial
						 	extOrderserSeq			= extOrderserSeq & "-1"

							extMeachulDate			= rsXL(29) ''ȯ�ҿϷ���

							''ȯ�ҿϷ����� �������̽��� �ִ�. (��۹�ȣ : 19906380)
							if (extMeachulDate="") then
								extOrderserSeq			= extOrgOrderserial & "-2"
								extMeachulDate			= rsXL(28)
								if (extMeachulDate="") then				''��ۿϷ��ϵ� ������.  �Ǵ� ȯ�ҿϷ���/��ۿϷ����� ���°�� ��ǰ��ȣ�� ã�Ƽ� �������� ��ۿϷ��Ͽ� ����.(�������¾���.)
									extMeachulDate			= rsXL(27)  ''�ϴ� �����Ϸ��Ϸ� ����.
'									2020-09-01 ������..��۹�ȣ : 151652292 8���������ε�, �� ó�� �����Ϸ��Ϸ� ���� �� 7���̶� 8�� ���⿡ �� ����..��
'									8���� ���� �� �ִ� � �������� �߰��� �� ��..WMP������ 8�� 18�Ϸ� �������� ����..���� �� ��ü..������ ������ 8��18�� ����־ �ذ�
'									���� ��ó�� �߻��� ���� ������ �� ��..
								end if
							end if
						else
						 	'// �������

						 	extOrgOrderserial		= ""

						end if

						if (LEN(extMeachulDate)=8) then
							extMeachulDate = LEFT(extMeachulDate,4)&"-"&MID(extMeachulDate,5,2)&"-"&MID(extMeachulDate,7,2)
						end if

						extVatYN = "Y"  ''���о���.
						' if (rsXL(23+1) = "�鼼") then
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

						If (rsXL(11+1) <> "���") Then
							If extItemNo = "0" Then
								If rsXL(18+1) > 0 Then
									extItemNo = 1
								Else
									extItemNo = -1
								End If
							End If
						End If

						If (rsXL(11+1) <> "���") and (rsXL(12+1)) = "0" then
							extItemCost				= 0
							extOwnCouponPrice		= CLNG((CLNG(rsXL(18+1))+CLNG(rsXL(21+1)))/extItemNo*100)/100 '' �������δ� ��ǰ���� , �������δ� ��ٱ������� �߰� 2019/06/02
							extTenCouponPrice		= CLNG(CLNG(rsXL(20+1))/extItemNo*100)/100  '' �Ǹž�ü�δ� ��ǰ����
							extJungsanType			= "C"

							extTenMeachulPrice		= extItemCost-extOwnCouponPrice-extTenCouponPrice
							extReducedPrice			= CLNG(extTenMeachulPrice)

							extCommPrice			= CLNG(CLNG(rsXL(15+1))/extItemNo*100)/100 - extOwnCouponPrice + CLNG(CLNG(rsXL(17+1))/extItemNo*100) / 100 + CLNG(CLNG(rsXL(16+1))/extItemNo*100) / 100 ''�ǸŴ��������-�������δ�����
							extTenJungsanPrice		= extTenMeachulPrice-extCommPrice

							extCommSupplyVatPrice	= 0
							extCommSupplyPrice		= 0

							extTenMeachulSupplyVatPrice	= 0
							extTenMeachulSupplyPrice	= 0

							extitemID = rsXL(3+1)
							tenitemid = rsXL(30)
						elseif (replace(replace(rsXL(12+1),",","")," ","") <> 0) then
							extItemCost				= CLNG((CLNG(rsXL(8+1)) + CLNG(rsXL(10))) *100) / 100  ''��ǰ�ǸŰ� + �ɼǰ�
							extOwnCouponPrice		= CLNG((CLNG(rsXL(18+1))+CLNG(rsXL(21+1)))/extItemNo*100)/100 '' �������δ� ��ǰ���� , �������δ� ��ٱ������� �߰� 2019/06/02
							extTenCouponPrice		= CLNG(CLNG(rsXL(20+1))/extItemNo*100)/100  '' �Ǹž�ü�δ� ��ǰ����
							extJungsanType			= "C"

							extTenMeachulPrice		= extItemCost-extOwnCouponPrice-extTenCouponPrice
							extReducedPrice			= CLNG(extTenMeachulPrice)


							extCommPrice			= CLNG(CLNG(rsXL(15+1))/extItemNo*100)/100 - extOwnCouponPrice + CLNG(CLNG(rsXL(17+1))/extItemNo*100) / 100 + CLNG(CLNG(rsXL(16+1))/extItemNo*100) / 100 ''�ǸŴ��������-�������δ�����
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

						if (rsXL(3+1) = "���") then
							'// �������
							extOrderserSeq			= "D"
							extOrgOrderserial		= ""
							extItemName				= "��ۺ�"
							extItemCost				= CLNG(rsXL(11+1)/extItemNo*100)/100
						elseif (rsXL(3+1) = "ȯ��") then
							'// ��ǰ
							extOrderserSeq			= "D-"&rsXL(1)
							extOrgOrderserial		= extOrderserial
							extItemName				= "��ǰ��ۺ�"
'������w�м��� �Ʒ� �ּ����� extItemCost �� ���� ���ϴ� ��, �������� �ƴѵ�/ Ȯ���ʿ�..2020-11-03 ������
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
					'// hmall ���� �󼼳���
					extOrderserial			= replace(rsXL(2),"-","") '' ��� ����
					''2019/11/03 (19) ''ī���δ��. => 2020/03/03 ���������������� ����.
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
						' 	'// �������
						' 	extOrderserSeq			= extOrderserial
						' 	extOrgOrderserial		= ""
						' 	extOrderserial			= Left(extOrderserial, 14)
						' else
						' 	'// ��ǰ
						' 	extOrderserSeq			= extOrderserial
						' 	extOrgOrderserial		= Left(extOrderserial, 14)
						' 	extOrderserial			= Left(extOrderserial, 14) ''& "-" & i
						' end if

						extVatYN = "Y"
						if (rsXL(23+1) = "�鼼") then
							extVatYN = "N"
						end if

						extItemNo				= rsXL(14)  '' 0�� CASE�� ����. ��ǰ�ε�.

						extItemCost = 0
						extReducedPrice = 0
						extOwnCouponPrice = 0
						extTenCouponPrice = 0
						extTenJungsanPrice = 0
						extCommPrice = 0
						extTenMeachulPrice = 0

						if (extItemNo<>0) then
							extItemCost				= CLNG(rsXL(8)/extItemNo*100)/100  ''�Ǹűݾ�
							extOwnCouponPrice		= CLNG(rsXL(16)/extItemNo*100)/100  + CLNG(rsXL(18)/extItemNo*100)/100 ''// + CLNG(rsXL(19)/extItemNo*100)/100 ''����δ�. ���޻�δ�. 'ī���δ�.
							extTenCouponPrice		= CLNG(rsXL(17)/extItemNo*100)/100  '' ���»�δ�.
							extJungsanType			= "C"

							extTenMeachulPrice		= extItemCost-extOwnCouponPrice-extTenCouponPrice
							extReducedPrice			= CLNG(extTenMeachulPrice)


							extCommPrice			= CLNG(rsXL(21+1)/extItemNo*100)/100  ''����������
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
					extOrderserial			= replace(rsXL(3),"-","") '' ��� ����

					if (Len(extOrderserial) = 14) and IsNumeric(rsXL(0)) then

						extJungsanDate = ""
						extMeachulDate			= CStr(dateadd("d",-1,dateadd("m",1,rsXL(1)+"-01")))

						extJungsanType			= "D"

						if (rsXL(2) = "�ֹ�") then
							'// �������
							extOrderserSeq			= "D"
							extOrgOrderserial		= ""
							extItemName				= "��ۺ�"
						elseif (rsXL(2) = "��ǰ") then
							'// ��ǰ
							extOrderserSeq			= "D-2"
							extOrgOrderserial		= ""
							extItemName				= "��ǰ��ۺ�"
						elseif (rsXL(2) = "���(�ֹ�)") then
							'// ��ǰ
							extOrderserSeq			= "D-3"
							extOrgOrderserial		= ""
							extItemName				= "��ۺ�"
						elseif (rsXL(2) = "��Ÿ") then
							'// ��ǰ
							extOrderserSeq			= "D-4"
							extOrgOrderserial		= ""
							extItemName				= "��ۺ�"
						else
							'// ����ֹ�
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
					'// �Ե����̸� ���� �󼼳���
					extOrderserial			= rsXL(1)

					if (Len(extOrderserial) = 14) and IsNumeric(rsXL(7)) then

						'// ���곻���� �������� ����. ���� ����ū �ֹ���ȣ���� ������� ������.
						'// (ex, 20140220H10035 -> 2014-02-28)
						'extMeachulDate = "0000-00-00"
						extMeachulDate = requestCheckvar(uprequest("extMeachulDate"),10)

						extJungsanDate = ""

						extOrderserSeq			= CStr(rsXL(2)) & "-" & CStr(rsXL(3))

						extItemNo				= rsXL(7)
						''4017377 ��ȯ��
						'// ���ֹ�-��ǰ�ֹ� ���й� : ��ǰ = ������ 0 ���� ���� ��� ��ǰ!!, ��ۺ� = ��ǰ��(4017388) or ���� ��ǰ�ֹ��� ��ǰ�� ��� ���� ��ۺ�
						if (rsXL(2) = 4017357) or (rsXL(2) = 4017388) or (rsXL(2) = 4017377) then
							if (rsXL(2) = 4017377) then
								extOrderserSeq = extOrderserSeq &"-"&CStr(rsXL(0)) ''2019/11/01 �߰�
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
						if (rsXL(6) = "�鼼") then
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
					'// LFmall ���� �󼼳���
					''2020/03/10 LF�δ� ��ۺ�ݾ� �߰���(15)
					extOrderserial			= rsXL(1)		'�ֹ���ȣ

					if (Len(extOrderserial) = 8) and IsNumeric(Trim((rsXL(11)))) then	'rsXL(11) : �Ǹż���

						'// ���곻���� �������� ����. ���� ����ū �ֹ���ȣ���� ������� ������. like lotteimall
						'// (ex, 20140220H10035 -> 2014-02-28)
						'extMeachulDate = "0000-00-00"
						lfmallEtcPrice = CLNG(rsXL(27))		'���������ݾ�
						extMeachulDate = requestCheckvar(uprequest("extMeachulDate"),10)
						extOrderserSeq			= extMeachulDate&"-"&i  ''Seq�� �������� ���Ѵ�.
						extJungsanDate = ""
						extItemNo				= Trim(rsXL(11))	'�Ǹż���

						If extItemNo = "0" AND rsXL(10) = "��ۺ�" Then	'���걸��
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

						extItemCost				= CLNG(CLNG(rsXL(13)) / extItemNo)		'rsXL(13) : �Ǹűݾ�
						extOwnCouponPrice		= CLNG(CLNG((rsXL(18)) + CLNG(rsXL(20))) / extItemNo)		'LF�δ�ݾ�(����) + LF�δ�ݾ�(���ϸ���)
						extTenCouponPrice		= CLNG(CLNG((rsXL(16)) + CLNG(rsXL(17))) / extItemNo) 	'��ü�δ�ݾ�(����) + ��ü�δ�ݾ�(EGM)
						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= extTenMeachulPrice
						extJungsanType			= "C"
						extCommPrice			= (CLNG(CLNG(rsXL(14)) * -1 / extItemNo) - extOwnCouponPrice)	'rsXL(14) : ��������


						extTenJungsanPrice		= extReducedPrice - extCommPrice
						extItemName				= html2db(rsXL(7))	'��ǰ��
						extItemOptionName		= html2db(rsXL(6))	'������
						extVatYN = "Y"
						extitemid = rsXL(5)	'��ǰ�ڵ�
						dvlprice = CLNG(rsXL(21)) + CLNG(rsXL(22)) + CLNG(rsXL(23)) + CLNG(rsXL(24))		'��ۺ�ݾ� (�⺻ + ��ǰ + ��ȯ) '2023-03-02 ������ : �߰���ۺ�(rsXl(23)) �߰�
					else
						IsValidInput = False
					end if
				Case "auction1010beasongpay"
					'// --------------------------------------------------------
					'// ���� ��ۺ� ���� �󼼳���
					extOrderserial			= rsXL(1) '' ������ȣ

					dim chkTypeStr : chkTypeStr = rsXL(12) ''2019/08/13 ���� (9=>10)
					dim auctionxltype : auctionxltype=0
					if (chkTypeStr="ȯ�ұ�����") or (chkTypeStr="����") then
						auctionxltype=1
					end if

					if (auctionxltype=0) then
						extItemName = rsXL(13)
					else
						extItemName = rsXL(12)
					end if

					if (Len(extOrderserial) = 10) and ((extItemName = "���ʹ�ۺ�") or (extItemName = "��ǰ��ۺ�") or (extItemName = "��ȯ��ۺ�") or (extItemName = "ȯ�ұ�����") ) then

						extItemNo = 1
						if (auctionxltype=1) then
							extMeachulDate = Left(rsXL(8), 10)  ''2019/08/13 �б�
						else
							extMeachulDate = Left(rsXL(7), 10)  ''��������� == ����Ϸ���
						end if
						if (extItemName = "ȯ�ұ�����") then
							if (extMeachulDate>="2019-10-25") then
								extOrderserSeq = CStr(rsXL(2)) & "-DDDD"
							else
								extOrderserSeq = CStr(rsXL(1)) & "-DDDD"
							end if
							extItemCost	   = CLng(rsXL(9)/extItemNo)  ''2019/08/13 ���� (8=>9)
							extCommPrice			= CLng(rsXL(10)/extItemNo)
						elseif (extItemName = "��ȯ��ۺ�") then
							if (rsXL(6)="") or isEmpty(rsXL(6)) or isNULL(rsXL(6)) then
								extOrderserSeq = CStr(rsXL(1)) & "-DD"
							else
								extOrderserSeq = CStr(rsXL(1)) & "-DD2"
							end if

							if (extItemName <> "��ȯ��ۺ�") then
								extOrderserSeq = extOrderserSeq & "D"

								if (extItemName="��ǰ��ۺ�") then
									extOrderserSeq = extOrderserSeq & "D"
								end if
							end if

							extItemCost				= CLng(rsXL(10)/extItemNo)
							extCommPrice			= CLng(rsXL(11)/extItemNo)
						else
							''�Ϲݹ�ۺ�
							if (rsXL(6)="") or isEmpty(rsXL(6)) or isNULL(rsXL(6)) then
								extOrderserSeq = CStr(rsXL(1)) & "-D"
							else
								extOrderserSeq = CStr(rsXL(1)) & "-D2"
							end if

							if (rsXL(9) <> "") then
								'// ��ǰ
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

						''extItemName				= replace(extItemName,"����","")
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
					'// ���� ���� �󼼳���
					extOrderserial			= rsXL(1) '' ������ȣ

					if (Len(extOrderserial) = 10) and IsNumeric(rsXL(16)) then

						extMeachulDate = LEFT(rsXL(11),10) ''���������

						extJungsanDate = ""

						extOrderserSeq = rsXL(2) '' �ֹ���ȣ

						if rsXL(10) <> "" then  ''ȯ����
							extOrderserSeq			= extOrderserSeq & "-1"
						end if

						extItemNo = rsXL(16) ''�ֹ�����

						''extItemCost				= ABS(rsXL(17)) + CLNG(CLNG(rsXL(18)) / extItemNo)    '' �ֹ�����*��ǰ�ǸŰ�+���ǻ�ǰ�ǸŰ�
						extItemCost				= CLNG(CLNG(rsXL(18)+rsXL(17)) / extItemNo)    '' �ֹ�����*��ǰ�ǸŰ�+���ǻ�ǰ�ǸŰ� //2019/05/30 ���� �ٲ����

						''extOwnCouponPrice		= CLng(CLNG(rsXL(19)+rsXL(20)-rsXL(32)-rsXL(33))/ extItemNo*100)/100  ''���� ��ǰ�� ���� + ���� ������ ���� ����
						''extOwnCouponPrice		= CLng(CLNG(rsXL(19)+rsXL(20)+rsXL(32)+rsXL(33))/ extItemNo*100)/100  '' //2019/05/30 ���� �ٲ����
						extOwnCouponPrice		= CLng(CLNG(rsXL(19)+rsXL(20))/ extItemNo*100)/100  '' //2019/05/31 �����
						extTenCouponPrice		= CLng(CLNG(rsXL(21))/ extItemNo*100)/100   '''�Ǹ��� ����  :: �Ǹ��� �ݵ� ��ǰ�� ����(?)
						extJungsanType			= "C"


						''extTenMeachulPrice		= CLng((CLNG(rsXL(22)+rsXL(32)+rsXL(33)) / extItemNo)*100)/100
						''extTenMeachulPrice		= CLng((CLNG(rsXL(22)) / extItemNo)*100)/100 '' //2019/05/30 ���� �ٲ����
						extTenMeachulPrice		= CLng((CLNG(rsXL(22)+rsXL(32)+rsXL(33)) / extItemNo)*100)/100 '' //2019/05/31 �����
						extReducedPrice			= CLng(extTenMeachulPrice)

						''extItemCost				= extTenMeachulPrice+extOwnCouponPrice+extTenCouponPrice
						extCommPrice			= CLng((CLNG(rsXL(29)) / extItemNo)*100)/100  ''�����̿��

						extTenJungsanPrice		= extTenMeachulPrice-extCommPrice  ''CLng((CLng(rsXL(45)) / extItemNo)*100)/100
						''extTenJungsanPrice		= CLng((CLNG(rsXL(26)) / extItemNo)*100)/100


						extItemName				= html2db(rsXL(4))
						extItemOptionName		= ""
						extitemID				= rsXL(3)

						extVatYN = "Y"

						if (rsXL(40)<>"����") then
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
					'// ���� ���� �󼼳���
					extOrderserial			= rsXL(1) '' ������ȣ

					if (Len(extOrderserial) = 10) and IsNumeric(rsXL(16)) then

						extMeachulDate = LEFT(rsXL(11),10) ''���������

						extJungsanDate = ""

						extOrderserSeq = rsXL(2) '' �ֹ���ȣ

						if rsXL(10) <> "" then  ''ȯ����
							extOrderserSeq			= extOrderserSeq & "-1"
						end if

						extItemNo = rsXL(16) ''�ֹ�����

						''extItemCost				= ABS(rsXL(17)) + CLNG(CLNG(rsXL(18)) / extItemNo)    '' �ֹ�����*��ǰ�ǸŰ�+���ǻ�ǰ�ǸŰ�
						extItemCost				= CLNG(CLNG(rsXL(18)+rsXL(17)) / extItemNo)    '' �ֹ�����*��ǰ�ǸŰ�+���ǻ�ǰ�ǸŰ� //2019/05/30 ���� �ٲ����

						''extOwnCouponPrice		= CLng(CLNG(rsXL(19)+rsXL(20)-rsXL(32)-rsXL(33))/ extItemNo*100)/100  ''���� ��ǰ�� ���� + ���� ������ ���� ����
						''extOwnCouponPrice		= CLng(CLNG(rsXL(19)+rsXL(20)+rsXL(32)+rsXL(33))/ extItemNo*100)/100  '' //2019/05/30 ���� �ٲ����
						extOwnCouponPrice		= CLng(CLNG(rsXL(19)+rsXL(20))/ extItemNo*100)/100  '' //2019/05/31 �����
						extTenCouponPrice		= CLng(CLNG(rsXL(21))/ extItemNo*100)/100   '''�Ǹ��� ����  :: �Ǹ��� �ݵ� ��ǰ�� ����(?)
						extJungsanType			= "C"


						''extTenMeachulPrice		= CLng((CLNG(rsXL(22)+rsXL(32)+rsXL(33)) / extItemNo)*100)/100
						''extTenMeachulPrice		= CLng((CLNG(rsXL(22)) / extItemNo)*100)/100 '' //2019/05/30 ���� �ٲ����
						extTenMeachulPrice		= CLng((CLNG(rsXL(22)+rsXL(32)+rsXL(33)) / extItemNo)*100)/100 '' //2019/05/31 �����
						extReducedPrice			= CLng(extTenMeachulPrice)

						''extItemCost				= extTenMeachulPrice+extOwnCouponPrice+extTenCouponPrice
						extCommPrice			= CLng((CLNG(rsXL(29)) / extItemNo)*100)/100  ''�����̿��

						extTenJungsanPrice		= extTenMeachulPrice-extCommPrice  ''CLng((CLng(rsXL(45)) / extItemNo)*100)/100
						''extTenJungsanPrice		= CLng((CLNG(rsXL(26)) / extItemNo)*100)/100


						extItemName				= html2db(rsXL(4))
						extItemOptionName		= ""
						extitemID				= rsXL(3)

						extVatYN = "Y"

						if (rsXL(40)<>"����") then
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
					'// ������ ���� �󼼳���
					'// 20200110 ��ȣ�ʵ�(0) �߰�, ����/ȯ�ޱ�(�����̿�� ������)(28) �߰�

					extOrderserial			= rsXL(0+1)

					if (Len(extOrderserial) = 10) and IsNumeric(rsXL(2+1)) then

						if (rsXL(9+1) <> "") then
							'// ��ǰ
							extMeachulDate = Left(rsXL(9+1), 10)
						else
							extMeachulDate = Left(rsXL(8+1), 10)  ''��ۿϷ��� ����.
						end if

						extJungsanDate = ""

						if (rsXL(9+1) <> "") then
							extOrderserSeq			= CStr(rsXL(1+1)) & "-1"
						else
							extOrderserSeq			= CStr(rsXL(1+1))
						end if

						'// ���ֹ�-��ǰ�ֹ� ���й� : ȯ������ �ִ���
						extItemNo = rsXL(14+1)

						extItemCost				= CLng((CLng(rsXL(15+1))+CLng(rsXL(16+1))+CLng(rsXL(17+1))+CLng(rsXL(18+1))) / extItemNo*100)/100  ''�ǸŰ���+�ʼ����û�ǰ�ݾ�+�߰�������ǰ�ݾ�+�ɼǻ�ǰ

						extOwnCouponPrice		= CLng((CLng(rsXL(19+1))+CLng(rsXL(20+1)))/ extItemNo*100)/100*-1
						extTenCouponPrice		= CLng((CLng(rsXL(21+1))+CLng(rsXL(36+2))*-1)/ extItemNo*100)/100*-1  ''36�Ǹ��� �ݵ� �������� ����

						extTenMeachulPrice		= CLng(rsXL(22+1)/ extItemNo*100)/100
						extReducedPrice			= CLng(extTenMeachulPrice)
						extJungsanType			= "C"

						extCommPrice			= CLng(rsXL(30+2)/ extItemNo*100)/100 ''�����̿��

						extTenJungsanPrice		= CLng(rsXL(27+2)/ extItemNo*100)/100 ''�Ǹ��� ���������

						extItemName				= "" ''html2db(rsXL(3))
						extItemOptionName		= ""
						extitemID 				= (rsXL(2))
						extVatYN = "Y"
						if (rsXL(40+2) = "�鼼") then
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
					'// ������ ��ۺ� ���� �󼼳���
					extOrderserial			= rsXL(0)

					''and IsNumeric(rsXL(2))
					' extItemName = rsXL(19) ''����,ī��,���հ���,�����

					' ''if NOT ((extItemName="��ǰ��ۺ�") or (extItemName="����� ����ۺ� ȯ������") or (extItemName="����") or (extItemName="�����ǰ��ۺ�") or (extItemName="��Ÿ��ǰ��")) then
					' if (extItemName="����" or extItemName="ī��" or extItemName="���հ���" or extItemName="�����" or extItemName="�˸�����" or extItemName="�۷ι�����") then
					' 	extItemName = rsXL(22)
					' end if

					On Error Resume Next
					extItemName = rsXL(23)
					If ERR Then
						extItemName = rsXL(15)
					end if
					On Error Goto 0

					''if (Len(extOrderserial) = 10) and ((extItemName = "������ۺ�") or (extItemName = "��ǰ��ۺ�") or (extItemName = "����� ����ۺ� ȯ������") or (extItemName="�����ǰ��ۺ�") or (extItemName="��Ÿ��ǰ��")) then
					if (Len(extOrderserial) = 10 ) then
						extItemNo = 1
						'if (extItemName = "��ǰ��ۺ�") or (extItemName = "����� ����ۺ� ȯ������") or (extItemName="�����ǰ��ۺ�") or (extItemName="��Ÿ��ǰ��") then
						''rw  extItemName
						if (extItemName<>"������ۺ�") and (extItemName<>"�߰���۱���") then
							if (isNULL(rsXL(1))) then
								extOrderserSeq = "-DD"
							else
								extOrderserSeq = CStr(rsXL(1)) & "-DD"
							end if


							if (extItemName <> "��ǰ��ۺ�") then
								extOrderserSeq = extOrderserSeq & "D"
							end if

							if (extItemName="�����ǰ��ۺ�") then
								extOrderserSeq = extOrderserSeq & "F"
							end if

							if (extItemName="��Ÿ��ǰ��") then
								extOrderserSeq = extOrderserSeq & "E"
							end if

							if (rsXL(9) <> "") then
								extMeachulDate = Left(rsXL(9), 10) ''��ۿϷ��� ����.
							else
								extMeachulDate = Left(rsXL(8), 10)  ''�����
							end if

							extItemCost				= CLng(rsXL(14)/extItemNo)
							extCommPrice			= 0
						else
							''�Ϲݹ�ۺ�
							if (isNULL(rsXL(1))) then
								If rsXL(16) = "���Ź̰���" Then
									extOrderserSeq = rsXL(7) & "-D"
								Else
									extOrderserSeq = "-D"
								End If
							elseif extItemName = "�߰���۱���" then
								extOrderserSeq = CStr(rsXL(1)) & "-DDD"
							else
								extOrderserSeq = CStr(rsXL(1)) & "-D"
							end if

							if (rsXL(10) <> "") then
								'// ��ǰ
								extMeachulDate = Left(rsXL(10), 10)
								if (Left(rsXL(9), 10)>extMeachulDate) then extMeachulDate=Left(rsXL(9), 10)  ''ȯ����,��ۿϷ����� ������¥�ε���
								extOrderserSeq	=extOrderserSeq & "-1"
								extItemNo = -1

								'' �Ա����� ������ ȯ���� �����ε���
								if (rsXL(6) = "") then
									extMeachulDate = Left(rsXL(10), 10)
									extOrderserSeq	=extOrderserSeq & "-1"
								end if

								if rsXL(16)="���" then extOrderserSeq	=extOrderserSeq & "-R"
								if rsXL(14)>0 then extOrderserSeq	=extOrderserSeq & "-1"
							else
								extMeachulDate = Left(rsXL(9), 10)  ''��ۿϷ��� ����.
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
					If (extItemName<>"������ۺ�") and (extItemName<>"�߰���۱���") Then
						extTenJungsanPrice		= extReducedPrice
					Else
						extTenJungsanPrice 		= CLng(rsXL(15)/extItemNo)
					End If


						extItemName				= replace(extItemName,"����","")
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
					'// CJ�� ��ǰ���� �󼼳���
					extOrderserial			= rsXL(23)		'�ֹ���ȣ

					if IsNumeric(rsXL(11)) and Len(extOrderserial) = 26 then

						extJungsanDate = ""
						extMeachulDate			= Replace(rsXL(1), "/", "-")	'��������
						if (Right(extOrderserial, 3) = "001") then
							'// �������
							extOrderserSeq			= extOrderserial
							extOrgOrderserial		= ""
							extOrderserial			= Left(extOrderserial, 14)
						else
							'// ��ǰ
							extOrderserSeq			= extOrderserial
							extOrgOrderserial		= Left(extOrderserial, 14)
							extOrderserial			= Left(extOrderserial, 14) ''& "-" & i
						end if

						extVatYN = "Y"
						if (rsXL(8) = "�鼼") then	'��������
							extVatYN = "N"
						end if

						extItemNo				= rsXL(9)	'���ⷮ

						'// �켱 �Ѿ��� �Է��ϰ�, �Ѿ��� �հ�ݾ��� ���� �Ŀ� �������� ������ ������ �������ش�.
						extItemCost				= CLNG(rsXL(11)/extItemNo*100)/100	'���Ǹűݾ�
						extReducedPrice			= CLNG(rsXL(15)/extItemNo*100)/100	'�Һ����ǸŰ�
						extOwnCouponPrice		= extItemCost - extReducedPrice
						extTenCouponPrice		= 0
						extJungsanType			= "C"

						extTenJungsanPrice		= CLNG(rsXL(20)/extItemNo*100)/100	'��ǰ������޿�����(������)

						extCommPrice			= CLNG(rsXL(18)/extItemNo*100)/100	'cj�հ�ݾ�
						extCommSupplyVatPrice	= CLNG(rsXL(17)/extItemNo*100)/100	'cj�ΰ���
						extCommSupplyPrice		= CLNG(rsXL(16)/extItemNo*100)/100	'cj������

						extTenMeachulPrice		= extReducedPrice
						extReducedPrice			= CLNG(extReducedPrice)
						extTenMeachulSupplyVatPrice	= CLNG(rsXL(14)/extItemNo*100)/100	'�Һ��ںΰ���
						extTenMeachulSupplyPrice	= CLNG(rsXL(13)/extItemNo*100)/100	'�Һ��ڰ��ް�
					else
						IsValidInput = False
					end if
				Case "cjmallbeasongpay"
					' uprequest("cjbeasongGubun") = "1"		//���������Ȳ(��ȯ�ù��)
					' uprequest("cjbeasongGubun") = "2"		//���������Ȳ(��ǰ�ù��)
					' uprequest("cjbeasongGubun") = "3"		//���������Ȳ(���ܰ���ۺ�)
					Dim ccnt
					ccnt = 0
					If uprequest("cjbeasongGubun") = "1" OR uprequest("cjbeasongGubun") = "2" OR uprequest("cjbeasongGubun") = "3" Then
						extOrderserial			= rsXL(4)
						If uprequest("cjbeasongGubun") = "3" Then
							extOrderserial			= rsXL(5)
						End If

						If (Len(extOrderserial) = 14) Then
							If uprequest("cjbeasongGubun") = "1" OR uprequest("cjbeasongGubun") = "2" Then
								extMeachulDate = rsXL(2)  ''ó������.
							Else
								extMeachulDate = LEFT(rsXL(3),10)  ''ó������ �������� �ƴϴ�..?
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
								extItemName = "��ȯ�ù��"
							End If

							If uprequest("cjbeasongGubun") = "2" Then
								extOrderserSeq = extOrderserSeq + "DD"
								extItemName = "��ǰ�ù��"
							End If

							If uprequest("cjbeasongGubun") = "3" Then
								extOrderserSeq = extOrderserSeq + "D"
								extItemName = "��ۺ�"
								if LEN(rsXL(3)) > 8 then extOrderserSeq = extOrderserSeq + "E" ''���ܰ���ۺ� ��Ÿ

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
							extItemNo				= 1	 ''�Ǹŷ�
							extOwnCouponPrice		= 0
							extTenCouponPrice		= 0
							extItemCost				= CLNG(rsXL(7)*-1)   ''������ �̹Ƿ� -1�� ������  /��ǰ�ù��, ��ȯ�ù��
							If extItemCost < 0 then ''�ݾ���-�̸�
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

					' if (sheetName="'���������Ȳ(���ܰ���ۺ�)$'") or (sheetName="'���������Ȳ(��ǰ�ù��)$'") or (sheetName="'���������Ȳ(��ȯ�ù��)$'") or (sheetName="'���������Ȳ(AS �ù��)$'") or (sheetName="sheet1$")  then
					' 	extOrderserial			= rsXL(4)
					' 	if (sheetName="'���������Ȳ(��ǰ�ù��)$'") or (sheetName="'���������Ȳ(��ȯ�ù��)$'") then
					' 		extOrderserial			= rsXL(3)
					' 	end if

					' 	if (sheetName="'���������Ȳ(AS �ù��)$'") then
					' 		extOrderserial			= LEFT(rsXL(5),14)
					' 	end if

					' 	if (Len(extOrderserial) = 14)  then
					' 		if (sheetName="'���������Ȳ(��ǰ�ù��)$'") or (sheetName="'���������Ȳ(��ȯ�ù��)$'") then
					' 			extMeachulDate = rsXL(1)  ''ó������.
					' 		else
					' 			extMeachulDate = LEFT(rsXL(2),10)  ''ó������ �������� �ƴϴ�..?
					' 		end if

					' 		if (extMeachulMonth<>LEFT(extMeachulDate,7)) then
					' 			extMeachulDate = extMeachulMonth+"-01"
					' 		end if

					' 		extJungsanDate = ""

					' 		extVatYN = "Y"


					' 		extOrderserSeq			= "1-"

					' 		if (sheetName="'���������Ȳ(���ܰ���ۺ�)$'") then
					' 			extOrderserSeq = extOrderserSeq + "D"
					' 			extItemName = "��ۺ�"

					' 			if LEN(rsXL(2))>10 then extOrderserSeq = extOrderserSeq + "E" ''���ܰ���ۺ� ��Ÿ
					' 		end if

					' 		if (sheetName="'���������Ȳ(��ǰ�ù��)$'") then
					' 			extOrderserSeq = extOrderserSeq + "DD"
					' 			extItemName = "��ǰ�ù��"
					' 		end if

					' 		if (sheetName="'���������Ȳ(��ȯ�ù��)$'") then
					' 			extOrderserSeq = extOrderserSeq + "DDD"
					' 			extItemName = "��ȯ�ù��"
					' 		end if

					' 		if (sheetName="'���������Ȳ(AS �ù��)$'") then
					' 			extOrderserSeq = extOrderserSeq + "DDDD"
					' 			extItemName = "AS�ù��"
					' 		end if

					' 		if (sheetName="sheet1$") then  ''[�������] ���������� ��Ÿ�������� �ø��°� ������..
					' 			extOrderserSeq = extOrderserSeq + "C"
					' 			extItemName = "��ۺ�����"

					' 		end if

					' 		' if (p_extOrderserial = extOrderserial) and  (p_extOrderserSeq = extOrderserSeq) then
					' 		' 	extOrderserSeq = extOrderserSeq+"-1"
					' 		' end if

					' 		extOrgOrderserial		= ""

					' 		extItemNo				= 1	 ''�Ǹŷ�

					' 		extOwnCouponPrice		= 0
					' 		extTenCouponPrice		= 0

					' 		if (sheetName="'���������Ȳ(AS �ù��)$'") then
					' 			extItemCost = 0
					' 			if rsXL(11)<>"" then
					' 				extItemCost				= CLNG(rsXL(11)*-1)
					' 			end if

					' 			if rsXL(10)<>"" then
					' 				extTenCouponPrice		= CLNG(rsXL(10))
					' 			end if
								
					' 			If rsXL(5) = "20230312052749-001-001-004" Then		'2023-05-02 ������ �߰�
					' 				extOrderserSeq = extOrderserSeq + "D"
					' 			End If
					' 		else
					' 			extItemCost				= CLNG(rsXL(6)*-1)   ''������ �̹Ƿ� -1�� ������  /��ǰ�ù��, ��ȯ�ù��
					' 		end if

					' 		if extItemCost<0 then ''�ݾ���-�̸�
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

					' 		if (sheetName="sheet1$") then  ''[�������] ���������� ��Ÿ�������� �ø��°� ������..
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

					extOrderserial			= rsXL(7+1)  ''�ֹ�ID
					extOrgOrderserial		= rsXL(6+1)  ''���ֹ�ID
					if (extOrderserial=extOrgOrderserial) then extOrgOrderserial=""

					if (Len(extOrderserial) = 14)  then

                        extMeachulDate = replace(rsXL(1),"/","-")  ''������
						extJungsanDate = ""

						extOrderserSeq			= CStr(rsXL(8+1)) '�ֹ�����


						extItemNo				= rsXL(15+2)


						if (extItemNo < 0) then  ' ��ǰ
							IsReturnOrder = True
							extOrderserial = extOrderserial & "-" & extOrderserSeq
							extOrgOrderserial = rsXL(6+1)   ''���ֹ�ID
						' elseif (extItemNo = 0) then  '' ������ 0�̸� ��ۺ� �Ǵ� �����ݾ�
						'  	IsReturnOrder = False
						'  	''���Ǹűݾ��� 0�� �ƴ� CASE�� ����.
						' 	extItemNo = 1
						else
							IsReturnOrder = False
						end if



						extItemCost				= 0
						extReducedPrice			= 0
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0

                        extItemName				= ""     ''��ǰ��
						extItemOptionName		= rsXL(14+2)     ''�ɼǸ�(��ǰ)
						if isNULL(extItemOptionName) then extItemOptionName=""

						' if (rsXL(25)<>0) then          ''��ۺ� (VAT����)
						' 	extJungsanType			= "D"
						' 	extItemNo = 1
						' 	if (IsReturnOrder) then extItemNo=extItemNo*-1
						' 	extOrderserSeq = extOrderserSeq + "-" + extJungsanType
						' 	extItemName = "��ۺ�"

						' 	if (rsXL(27)<>0) then
						' 		extOrderserSeq = extOrderserSeq + extJungsanType  ''-DD
						' 		extItemName = "��ǰ��ۺ�"
						' 	end if

						' 	if (rsXL(25)<>0) then
						' 		extOrderserSeq = extOrderserSeq + "R"  ''-DDR ,  DR
						' 		extItemName = "��ǰ��ۺ�"
						' 	end if
						' else
						' 	extJungsanType			= "C"
						' end if
						extJungsanType = "C"
						dvlprice = rsXL(25+3)   ''��ۺ�

						''��Ÿ �����۾��ε���..
						if (extItemNo=0) then
							extOwnCouponPrice		= CLNG(rsXL(19+3))
							extTenCouponPrice		= CLNG(rsXL(18+3))
							extCommPrice			= CLNG(rsXL(21+3)-rsXL(16+2))  ''���Ǹž�-����ݾ�
							extItemCost				= CLNG(rsXL(17+2))
							extTenJungsanPrice		= CLNG(rsXL(16+2))  ''����ݾ�

							if (dvlprice=0) and (rsXL(21+3)<>0 or rsXL(16+2)<>0) then  ''��ۺ� �ƴϰ� ���Ǹž��̳�. ������� <>0 �̸� ������ 1�� �ϰ� ����. �����ݾ�?
								extItemNo = 1
								extOrderserSeq = extOrderserSeq & "-" & replace(extMeachulDate,"-","")
								''�ߺ��Ǵ� ���̽��� �ִ�..
								'' ������������ �ʴ�..
								' rw p_extOrderserial&":"&extOrderserial&"::"&p_extOrderserSeq&":"&extOrderserSeq
								' if (p_extOrderserial = extOrderserial) and (p_extOrderserSeq = extOrderserSeq) then
								' 	extOrderserSeq = extOrderserSeq + "-" & extOrderserSeq
								' end if
							end if
						else
							extOwnCouponPrice		= CLNG(rsXL(19+3)/extItemNo*100)/100   ''SSG����
							extTenCouponPrice		= CLNG(rsXL(18+3)/extItemNo*100)/100   ''���»����κδ��.

							extCommPrice			= CLNG((rsXL(21+3)-rsXL(16+2))/extItemNo*100)/100  ''������ :: ���Ǹž�-�����
							extItemCost				= CLNG((rsXL(17+2)-dvlprice)/extItemNo*100)/100   ''�Ǹűݾ��հ�
							extTenJungsanPrice		= CLNG((rsXL(16+2)-dvlprice)/extItemNo*100)/100   ''����ݾ��հ�
						end if

						extReducedPrice			= CLNG(extItemCost-(extTenCouponPrice+extOwnCouponPrice))   ''�Ҽ��� �糦.
						extTenMeachulPrice      = (extItemCost-(extTenCouponPrice+extOwnCouponPrice)) 		''�Ҽ��� 1�ڸ�

						extVatYN = "Y"

						if (rsXL(6) = "�鼼") then
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
					'// 11���� ���� �󼼳���
					extOrderserial			= rsXL(1)

					if (Len(extOrderserial) = 15) or (Len(extOrderserial) = 17) then  ''and IsNumeric(rsXL(7)

                        extMeachulDate = replace(rsXL(11),"/","-")  ''����Ȯ����
						extJungsanDate = ""

						extOrderserSeq			= CStr(rsXL(2))

						extItemNo				= rsXL(16)

						'// ���ֹ�-��ǰ�ֹ� ���й� : ��ǰ = ������ 0 ���� ���� ��� ��ǰ!!, ��ۺ� = ��ǰ��(4017388) or ���� ��ǰ�ֹ��� ��ǰ�� ��� ���� ��ۺ�
''						if (rsXL(2) = 4017357) or (rsXL(2) = 4017388) then
''							if (IsReturnOrder = True or rsXL(2) = 4017388) or (extItemNo <= 0) then
''								extOrderserial 		= rsXL(1) & "-" & rsXL(0)
''								extOrgOrderserial	= rsXL(1)
''							else
''								extOrgOrderserial	= ""
''							end if
''						else


						if (extItemNo < 0) then  ''������ 0�̸� ��ۺ�.
							IsReturnOrder = True
							extOrderserial = rsXL(1) & "-" & extOrderserSeq
							extOrgOrderserial		= rsXL(1)
						elseif (extItemNo = 0)	and (rsXL(18)<0) then  ''��ۺ�� ��ǰ�� CASE
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

                        extItemName				= html2db(rsXL(14))     ''��ǰ��
						extItemOptionName		= html2db(rsXL(15))     ''�ɼǸ�
						'2018-01-18 ������ �ϴ����� ����
						extItemOptionName		= requestCheckvar(extItemOptionName, 128)

						if (extItemNo=0) then          ''������ 0�̸� ��ۺ�.
							extJungsanType			= "D"
							extItemNo = 1
							if (IsReturnOrder) then extItemNo=extItemNo*-1
							extOrderserSeq = extOrderserSeq + "-" + extJungsanType
							extItemName = "��ۺ�"

							if (rsXL(27+1)<>0) then
								extOrderserSeq = extOrderserSeq + extJungsanType  ''-DD
								extItemName = "��ǰ��ۺ�"
							end if

							if (rsXL(25+1)<>0) then
								extOrderserSeq = extOrderserSeq + "R"  ''-DDR ,  DR
								extItemName = "��ǰ��ۺ�"
							end if

							extItemOptionName = ""
						else
							extJungsanType			= "C"
						end if

						'' 5��1�Ϻ��� �հ� �ٲ����.
						' extOwnCouponPrice		= CLNG(rsXL(44)/extItemNo*1000)/1000   ''11��������
						' extTenCouponPrice		= CLNG(rsXL(41)/extItemNo*1000)/1000   ''���������̿��
						'' 2022-12-15 ������ ������ �Ʒ��� ����
						' extOwnCouponPrice		= CLNG((CLNG(rsXL(43)) + CLNG(rsXL(44))) /extItemNo*1000)/1000   ''11��������
						' extTenCouponPrice		= CLNG(rsXL(42)/extItemNo*1000)/1000   ''���������̿��
						'' 2022-12-15 ������ ������ �Ʒ��� ����
						tmpSellerAddSalePriceBy10x10	= CLNG(CLNG(rsXL(43)) * 0.15)				'�Ǹ����߰������� 15%�� �ٹ����� �δ��̶���
						tmpSellerAddSalePriceBy11st1010	= CLNG(CLNG(rsXL(43)) * 0.85)				'�Ǹ����߰������� 85%�� 11st �δ��̶���

						extOwnCouponPrice		= CLNG((tmpSellerAddSalePriceBy11st1010 + CLNG(rsXL(44))) /extItemNo*1000)/1000   ''11��������
						extTenCouponPrice		= CLNG((tmpSellerAddSalePriceBy10x10 + CLNG(rsXL(42))) / extItemNo*1000)/1000   ''���������̿��

						extCommPrice			= CLNG((CLNG(rsXL(39)) + CLNG(rsXL(40)) +CLNG(rsXL(51))- CLNG(rsXL(44)))/extItemNo*1000)/1000  ''�����̿��(��ǰ)+ �����̿��(��������ۺ�) + �ĺұ����-11��������
						''extCommPrice			= CLNG((rsXL(20+1)-rsXL(39))*100/extItemNo)/100 - CLNG(rsXL(44)/extItemNo*100)/100  ''�����ݾ��հ� :: ���������̿��(38) �� ������ �ִ�.
						extItemCost				= CLNG(rsXL(18)/extItemNo*100)/100   ''�Ǹűݾ��հ�
						extTenJungsanPrice		= CLNG(rsXL(17)/extItemNo*100)/100   ''����ݾ�

						''extCommPrice = extItemCost-extOwnCouponPrice-extTenCouponPrice-extTenJungsanPrice

						extReducedPrice			= CLNG(extItemCost-(extTenCouponPrice+extOwnCouponPrice))   ''�Ҽ��� �糦.
						extTenMeachulPrice      = (extItemCost-(extTenCouponPrice+extOwnCouponPrice)) 		''�Ҽ��� 1�ڸ�

						extVatYN = "Y"
						''if (rsXL(6) = "�鼼") then
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
					'// GS�� ��ǰ���� �󼼳���
					extOrderserial			= rsXL(8)

					''if IsNumeric(extOrderserial) and Len(rsXL(5)) = 10 then
					if IsNumeric(extOrderserial) and (Len(extOrderserial) = 9 or Len(extOrderserial) = 10) then

						extMeachulDate = rsXL(32)  ''����Ϸ���
						extJungsanDate = ""

						if (extMeachulDate = "--") or IsNull(extMeachulDate) then
							extMeachulDate = rsXL(30)  ''��ۿϷ���

							if (extMeachulDate = "--") or IsNull(extMeachulDate) then
								extMeachulDate = rsXL(29)  ''���Ϸ���
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

						extItemCost				= CLNG(rsXL(46)) ''�ܰ��̴�. ������ٸ�.
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
					'// GS�� ��ǰ���� �󼼳���
					extOrderserial			= rsXL(0)

					''if IsNumeric(extOrderserial) and Len(rsXL(5)) = 10 then
					if IsNumeric(extOrderserial) and (Len(extOrderserial) = 9 or Len(extOrderserial) = 10) then

						extMeachulDate = rsXL(19+1)  ''����Ϸ���
						extJungsanDate = ""

						if (extMeachulDate = "--") or IsNull(extMeachulDate) then
							extMeachulDate = rsXL(18+1)  ''��ۿϷ���

							if (extMeachulDate = "--") or IsNull(extMeachulDate) then
								extMeachulDate = rsXL(17+1)  ''���Ϸ���
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
					'// GS�� ��ۺ����� �󼼳���
					extOrderserial			= rsXL(1)

					if (Len(extOrderserial) = 9 or Len(extOrderserial) = 10) and IsNumeric(rsXL(4)) then

						'// �켱�� ���ֹ����� �ְ� �Ʒ����� ������� �Է��Ѵ�.
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

						if isNULL(extItemName) then extItemName="��Ÿ"
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
							if (rsXL(0) = "��ǰ��ۺ�(ȯ��)") then
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
					'// ��ؼ� ��ǰ���� �󼼳���

				Case "cjmallproduct"
					'// --------------------------------------------------------
					'// CJ�� ��ǰ���� �󼼳���

				Case "wizwidproduct"
					'// --------------------------------------------------------
					'// �������� ��ǰ���� �󼼳���

				Case "gabangpopproduct"
					'// --------------------------------------------------------
					'// �м���(������) ��ǰ���� �󼼳���
				Case "goodshop1010"
					extOrderserial = rsXL(4)	'�ֹ���ȣ
					If Len(extOrderserial) = 19 Then
						extJungsanDate = ""
						extOrgOrderserial = ""
						extMeachulDate =  LEFT(Trim(rsXL(6)), 10)		'�ֹ�����
						extOrderserSeq = extMeachulDate & "-" & i  ''Seq�� �������� ���Ѵ�.
						extItemNo = Trim(rsXL(16))		'�ֹ�����
						If (extItemNo <= 0) Then
							IsReturnOrder = True
							extOrgOrderserial = extOrderserial
							extOrderserial = extOrderserial&"-1"
						Else
							IsReturnOrder = False
						End If
						extJungsanType			= "C"
						extItemName			= html2db(rsXL(14))			'��ǰ��
						extItemOptionName	= html2db(rsXL(15))			'�ɼǸ�
						extItemCost	= CLNG(rsXL(18) / extitemNo)		'�ǸŰ�
						extTenCouponPrice	= 0
						extOwnCouponPrice	= 0

						extTenMeachulPrice	= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice		= CLng(extTenMeachulPrice)
						extCommPrice		= CLNG(CLNG(rsXL(24)) / extitemNo)		'�¼�����ݾ�
						extTenJungsanPrice	= extReducedPrice - extCommPrice
						
						If rsXL(26) = "����" Then
							extVatYN = "Y"
						Else
							extVatYN = "N"
						End If
						extitemid = rsXL(5)		'��ǰ��ȣ
					Else
						IsValidInput = False
					End If
				Case "wconcept1010"
					extOrderserial = rsXL(5)	'�ֹ���ȣ
					If Len(extOrderserial) = 9 Then
						extJungsanDate = ""
						extOrgOrderserial = ""
						extMeachulDate =  Trim(rsXL(4))		'����Ȯ����(ó���Ϸ���)
						extOrderserSeq = Trim(rsXL(37))
						extItemNo = Trim(rsXL(17))		'����
						If (extItemNo <= 0) Then
							IsReturnOrder = True
							extOrgOrderserial = extOrderserial
							extOrderserial = extOrderserial&"-1"
						Else
							IsReturnOrder = False
						End If

						If rsXL(38) = "��۷�" then          '����
							extJungsanType	= "D"
							extItemNo = 1
							extOrderserSeq = extMeachulDate & "-" & i & "-" & extJungsanType
							extItemName = "��ۺ�"
							extItemOptionName = ""
							extItemCost	= CLNG(rsXL(27) / extitemNo)		'��۷�
						Else
							extJungsanType			= "C"
							extItemName			= html2db(rsXL(13))			'��ǰ��
							extItemOptionName	= html2db(rsXL(15))			'�ɼ�1
							extItemCost	= CLNG(rsXL(31) / extitemNo)		'������ݾ�
						End If
						' extTenCouponPrice	= CLNG(CLNG(rsXL(21)) / extitemNo)	'��ü����
						' extOwnCouponPrice	= CLNG(CLNG(rsXL(22)) / extitemNo)	'��������
						extTenCouponPrice	= 0
						extOwnCouponPrice	= 0

						extTenMeachulPrice	= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice		= CLng(extTenMeachulPrice)
						extCommPrice		= CLNG(CLNG(rsXL(32)) / extitemNo) + CLNG(CLNG(rsXL(33)) / extitemNo)   	'���ݰ�꼭 + �ؿ��Ǹ� ������
						extTenJungsanPrice	= CLNG(rsXL(34) / extItemNo * 100)/100	'��ü���޾�
						extTenJungsanPrice	= extReducedPrice - extCommPrice
						extVatYN = "Y"
						extitemid = ""
					Else
						IsValidInput = False
					End If
				Case "withnature1010"
					extOrderserial = rsXL(2)	'�ֹ���ȣ
					If Len(extOrderserial) = 12 Then
						extJungsanDate = ""
						extOrgOrderserial = ""
						extOrderserSeq = ""
						extMeachulDate =  Trim(rsXL(3))		'�԰�����
						
						extItemNo = Trim(rsXL(7))			'����
						If (extItemNo <= 0) Then
							IsReturnOrder = True
							extOrgOrderserial = extOrderserial
							extOrderserial = extOrderserial&"-1"
						Else
							IsReturnOrder = False
						End If
						extJungsanType			= "C"
						extItemName			= html2db(rsXL(6))			'��ǰ��
						extItemOptionName	= ""

						extTenCouponPrice	= 0	'��ü����
						extOwnCouponPrice	= 0	'��������

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
							extItemCost		= rsget("SellPrice")		'����ݾ��� �� �Ѿ�´�. �ֹ��������� �����;��Ѵ�.
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
								extItemCost		= rsget("SellPrice")		'����ݾ��� �� �Ѿ�´�. �ֹ��������� �����;��Ѵ�.
								extitemid		= rsget("matchItemID")
								extitemoption	= rsget("matchitemoption")
								extOrderserSeq	= rsget("OrgDetailKey")
							End If
							rsget.Close
						End If

						If extOrderserSeq = "" Then
							extOrderserSeq = Trim(rsXL(0))		'No. | �ֹ�������Ű �� �Ѿ�� �ӽ÷� �̰ɷ�..
						End If

						extTenMeachulPrice	= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice		= CLng(extTenMeachulPrice)
						extCommPrice		= extItemCost - (CLNG(rsXL(11) / extItemNo * 100) / 100)
						extTenJungsanPrice	= CLNG(rsXL(11) / extItemNo * 100) / 100
						extVatYN = "Y"
						dvlprice = 0			'�ڿ��̶��� ��ۺ� ����
					Else
						IsValidInput = False
					End If
				Case "GS25"
					extOrderserial = rsXL(3)	'�ֹ���ȣ
					If Len(extOrderserial) = 13 Then
						extJungsanDate = ""
						extOrgOrderserial = ""
						extMeachulDate = Trim(rsXL(2))		'��ۿϷ���
						extMeachulDate = Left(extMeachulDate, 4) & "-" & Right(Left(extMeachulDate, 6), 2) & "-" & Right(extMeachulDate, 2)
						extOrderserSeq = Trim(rsXL(4))
						extItemNo = Trim(rsXL(7))		'����
						If (extItemNo <= 0) Then
							IsReturnOrder = True
							extOrgOrderserial = extOrderserial
							extOrderserial = extOrderserial&"-1"
						Else
							IsReturnOrder = False
						End If
						extJungsanType		= "C"
						extItemName			= html2db(rsXL(6))			'��ǰ��
						extItemOptionName	= ""
						extItemCost			= CLNG(rsXL(8))				'������ݾ�
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
					'// ����������� ��ǰ���� �󼼳���

				Case "playerproduct"
					'// --------------------------------------------------------
					'// �÷��̾� ��ǰ���� �󼼳���

				Case "lotteComM"  ''201605 ����
					'// --------------------------------------------------------
					'// �Ե�����(������) ��ǰ����

				Case "lotteComM_201604"
					'// --------------------------------------------------------
					'// �Ե�����(������) ��ǰ����
				Case "lotteCombeasongpay"
					'// --------------------------------------------------------
					'// �Ե����� ��ۺ�

					if (rsXL(11) = "���긶��") then
						extOrderserial		= Replace(rsXL(3),"-","")
						extMeachulDate = rsXL(2)
						extItemCost = rsXL(10)
						extItemNo = 1


						'extOrderserSeq = replace(extMeachulDate,"-","")&"-D"&"-"&i

						' if (rsXL(5)="�Ϲ�") then

						' elseif (rsXL(5)="��Ŭ����") then
						' 	extOrderserSeq = extOrderserSeq&"C"
						' elseif (rsXL(5)="�÷�Ƽ��+������") then
						' 	extOrderserSeq = extOrderserSeq&"P"
						' elseif (rsXL(5)="�����۱�") then
						' 	extOrderserSeq = extOrderserSeq&"F"
						' elseif (rsXL(5)="�Ե�����-��") then
						' 	extOrderserSeq = extOrderserSeq&"T"
						' elseif (rsXL(5)="�Ե�����-��ü") then
						' 	extOrderserSeq = extOrderserSeq&"U"
						' elseif (rsXL(5)="��ü����-��") then
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

						extItemName				= "��ۺ�"
						extItemOptionName		= ""

						extVatYN = "Y"

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0
					else
						IsValidInput = False
					end if
				Case "lotteCom"  ''2016/05/01 �ֹ��󼼹�ȣ �ʵ� �����.
					'// --------------------------------------------------------
					'// �Ե����� ��Ź���� �󼼳���
					if (rsXL(28) = "�ٹ�����") then  ''24+1 =>24+2 2016/05/01  // 27 2019/07/12 (27) (��ϻ�ǰ��), �ŷ�ó�д��

						extMeachulDate = rsXL(0)
						extJungsanDate = ""

						extItemNo				= rsXL(10)
						plusMinus				= extItemNo / Abs(extItemNo)
						if (extItemNo >= 0) then
							'// �������
							extOrderserial 			= Replace(rsXL(24), "-", "")
							extOrderserSeq			= TRim(rsXL(25)) ''=>  2016/05/01 //
							extOrgOrderserial		= ""
						else
							'// ��ǰ
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

						extItemName				= html2db(rsXL(7))			'// �ܺθ� ��ǰ���� �ٲ��. ��ǰ�� ��� �ܺθ� ��ǰ�ڵ�� ��Ī
						extItemOptionName		= html2db(rsXL(8))						'// ���곻���� �ɼ������� ����. ==>����.

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0
					else
						IsValidInput = False
					end if
				Case "halfclubproduct"
					'// --------------------------------------------------------
					'// ����Ŭ�� ��ǰ �󼼳���
					if (LEN(rsXL(3)) = 12) then

						extOrderserial = rsXL(3)
						extMeachulDate = rsXL(1)
						extJungsanDate = ""
						extMeachulDate = Left(extMeachulDate, 4) & "-" & Right(Left(extMeachulDate, 6), 2) & "-" & Right(extMeachulDate, 2)

						extOrgOrderserial		= CSTR(rsXL(2))
						extOrgOrderserial = ""

						extOrderserSeq			= rsXL(4)
						extItemNo				= rsXL(17)   ''��������


						if (extItemNo <0) then
							'// ��ǰ
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
					'// ����Ŭ�� ��ۺ� �󼼳���
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

						extItemName				= "��ۺ�"
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
					'// ������� �󼼳���
					if (Len(rsXL(0)) = 16) then
						isextMeachulDate = ""
						extMeachulDate = rsXL(5)
						extJungsanDate = "" 'rsXL(6)
						extMeachulDate = Left(extMeachulDate, 4) & "-" & Right(Left(extMeachulDate, 6), 2) & "-" & Right(extMeachulDate, 2)
						extJungsanDate = "" 'Left(extJungsanDate, 4) & "-" & Right(Left(extJungsanDate, 6), 2) & "-" & Right(extJungsanDate, 2)

						extItemNo				= 1					'// �������� **



						extOrderserial 			= rsXL(0)
						extOrderserSeq 			= rsXL(1)
						If (rsXL(13)="�������� ȸ��") Then
							extOrderserSeq 	= extOrderserSeq & "-1"
						ElseIf (rsXL(13)="��������") Then
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

'���� ������ ���� �������� �� �Ǵ� ����..
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

							'1. seq�� ���곻�� �ִ� �� Ȯ��, �̶� �����ϵ� �����´�
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

							'2. ���� ���곻���� ������ ������ �����ϰ� ����� �������� ���Ѵ�.
							If isextMeachulDate <> "" Then
								If isextMeachulDate <> extMeachulDate Then
							'3. �񱳽� ��¥�� �ٸ��� -A�� ���δ�
									extOrderserSeq 	= extOrderserSeq & "-A"
								End If
							End If
'						End If

						extOrgOrderserial		= ""
						extVatYN = "Y"

						if rsXL(2) = "��ۺ�" then
							extJungsanType			= "D"
							extItemNo				= 1
						else
							extJungsanType			= "C"
						end if

						if (rsXL(13)="������ ���") then extItemNo=-1

						extItemCost				= rsXL(9)			'// �ǸŰ� ����
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
					'// �Ե��� �󼼳���
					If (Len(rsXL(7)) = 16) then
						extOrderserial 	= rsXL(7)   											''�ֹ���ȣ
						extOrgOrderserial = ""
						extOrderserSeq = ""
						extCommPrice = ""
						extMeachulDate		= rsXL(20)  										''����Ȯ����
						lotteonBanpoomDate	= rsXL(21)  										''��ǰ�Ϸ���

						If Trim(lotteonBanpoomDate) <> "" Then
							extMeachulDate = lotteonBanpoomDate
						End If

						If extMeachulDate = "" Then
							extMeachulDate = rsXL(23)											'������Ⱓ
						End If

						If extMeachulDate = "" AND rsXL(22) <> "" Then							'rsXL(19) : ���꿹���� / ����Ȯ����, ��ۿϷ���, ��ǰ�Ϸ��� ���� �����Ͱ� ���� �����ϸ� ����
							extMeachulDate = dateadd("d", -1, rsXL(22))
						End If

						extJungsanDate	= ""
						extItemNo		= rsXL(24)  											''�Ǹż���
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

						'������ ����==�ٹ�����, ���==�Ե�
						extItemCost				= rsXL(25)										'�ǸŴܰ�
						extOwnCouponPrice		= CLNG(rsXL(28) / Chkiif(extItemNo="0", "1", extItemNo) )				'rsXL(25) : ��ǰ���δ��δ�ݾ�
						extTenCouponPrice		= CLNG(CLNG(rsXL(27) + rsXL(29)) / Chkiif(extItemNo="0", "1", extItemNo))	'rsXL(24) : ����������αݾ�, rsXL(26) : ��ǰ���μ����δ�ݾ�
						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extJungsanType			= "C"

                    If extItemNo < 0 AND CLNG(rsXL(28)) < 0 AND CLNG(rsXL(43)) = 0 Then
                        'extCommPrice            = CLNG(CLNG(CLNG(rsXL(35)) + CLNG(rsXL(37)) + CLNG(rsXL(44))) / Chkiif(extItemNo="0", "1", extItemNo))
						'maybe									�⺻������ + PCS������ݾ� + PG������ ���հ�
						'extCommPrice            = CLNG(CLNG(CLNG(rsXL(36)) + CLNG(rsXL(39)) + CLNG(rsXL(46))) / Chkiif(extItemNo="0", "1", extItemNo))
						'maybe									�⺻������ + PCS������ݾ� + PG������ ���հ�
						extCommPrice            = CLNG(CLNG(CLNG(rsXL(36)) + CLNG(rsXL(40)) + CLNG(rsXL(47))) / Chkiif(extItemNo="0", "1", extItemNo))
                    Else
                        'extCommPrice            = CLNG(CLNG(CLNG(rsXL(35)) + CLNG(rsXL(37)) + CLNG(rsXL(44)) - rsXL(28)) / Chkiif(extItemNo="0", "1", extItemNo))
						'maybe									�⺻������ + PCS������ݾ� + PG������ ���հ� - ��ǰ���δ��δ�ݾ�
						extCommPrice            = CLNG(CLNG(CLNG(rsXL(36)) + CLNG(rsXL(40)) + CLNG(rsXL(47)) - CLNG(rsXL(28))) / Chkiif(extItemNo="0", "1", extItemNo))
                    End If
						extTenJungsanPrice		= extReducedPrice - extCommPrice

						lotteondvlprice1		=  CLNG(rsXL(30))	'��ۺ�������ݾ�-> �ѹ�ۺ� 
						lotteondvlprice2		=  CLNG(rsXL(32))	'��ۺ�����(���δ�)
						lotteondvlPGCommprice	=  CLNG(rsXL(47))	'PG������ ���հ�
						lotteondvlTotCommprice	=  CLNG(rsXL(48))	'�� ������ �հ�
						dlvCommprice			= CLNG(rsXL(38))	'��ۺ� ������

						'dvlprice				= lotteondvlprice1 - lotteondvlprice2			'//2020-05-26 ������..�̷��� ���� �� ��ۺ� �ݾ��� ������ ��?
						dvlprice				= lotteondvlprice1 								'��ۺ��� ��� ��ۺ�����(���δ�) : �Ե��δ���� �������� �ʴ°��� �´µ���. by)eastone

						extItemName				= LeftB(html2db(rsXL(12)), 80)					'rsXL(9) : ��ǰ��(�ɼǸ�����) / LeftBó��..2021-01-18..2021010913167292
						extItemOptionName		= ""											'��ó�� ��ǰ��� �ɼǸ��� ���� ��..��ó��
						extitemoption			= html2db(rsXL(11))								'rsXL(8) : ��ǰ��ȣ

						If rsXL(14) = "����" Then												'���������� : �������� �ʵ带 ����..by)eastone
							extVatYN = "Y"
						Else
							extVatYN = "N"
						End If
						extitemid = rsXL(10)														'rsXL(7) : �Ǹ��ڻ�ǰ��ȣ

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
					'// yes24 �󼼳���
					If (Len(rsXL(1)) = 11) then
						extOrderserial 	= rsXL(1)   				''�ֹ���ȣ
						extOrgOrderserial = ""
						extOrderserSeq = ""
						extCommPrice = ""
						extMeachulDate		= rsXL(0)  				''����

						extJungsanDate	= ""
						extItemNo		= 1  						''���� / ���� ������ ����..���� 1ó��
						extOrderserSeq	= extOrderserial &"-"& rsXL(0)

						extOwnCouponPrice = 0						'�����ݾ� �˼�����
						extTenCouponPrice = 0						'�����ݾ� �˼�����

						If (rsXL(4) <> 0) Then						''��ǰ�ݾ��� 0�� �ƴϸ�
							IsReturnOrder = True
							extOrgOrderserial = extOrderserial
							extOrderserSeq	= extOrderserSeq & "-" & rsXL(2)
							extItemCost	= rsXL(4)					'��ǰ�ݾ�
							extItemNo		= -1
							extCommPrice			= CLNG(rsXL(4)) - CLNG(rsXL(6))	'������ = ��ǰ�ݾ� - ��ǰ����
						Else
							IsReturnOrder = False
							extItemCost	= rsXL(3)					'�ֹ��ݾ�
							extCommPrice			= CLNG(rsXL(3)) - CLNG(rsXL(5))	'������ = �ֹ��ݾ� - �ֹ�����
						End If

						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extJungsanType			= "C"

						extTenJungsanPrice		= extReducedPrice - extCommPrice
						dvlprice				= rsXL(7) 			'��ۺ�

						extItemName				= ""				'��ǰ�� �˼�����
						extItemOptionName		= ""				'�ɼǸ� �˼�����
						extitemoption			= ""				'�ɼ� �˼�����

						extVatYN = "Y"								'�������� �˼�����
						extitemid = ""								'��ǰ��ȣ �˼�����

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
						extOrderserial 	= rsXL(5)   				''�ֹ���ȣ
						extOrgOrderserial = ""
						extOrderserSeq = ""
						extCommPrice = ""
						extMeachulDate		= LEFT(rsXL(18),10)  	''���������

						extJungsanDate	= ""
						extItemNo		= rsXL(8)  					''����

'						extOrderserSeq	= extOrderserial &"-"& rsXL(1)
						extOrderserSeq	= extOrderserial & "-" & rsXL(6) & "-" & i	''�ֹ��ɼǹ�ȣrsXL(6)

						extOwnCouponPrice = 0						'�����ݾ� �˼�����
						extTenCouponPrice = 0						'�����ݾ� �˼�����

						If (extItemNo < 0) Then						''���� 0�� �ƴϸ�
							IsReturnOrder = True
							extOrgOrderserial = extOrderserial
							extOrderserSeq	= extOrderserSeq & "-" & rsXL(6)
							extItemCost		= CLNG(rsXL(10) / extItemNo)		'��ǰ�ǸŰ�A
							extItemNo		= -1
							extCommPrice	= CLNG(rsXL(13) / extItemNo )		'������D
						Else
							IsReturnOrder = False
							extItemCost		= CLNG(rsXL(10) / extItemNo)		'��ǰ�ǸŰ�A
							extCommPrice	= CLNG(rsXL(13) / extItemNo )		'������D
						End If

						extJungsanType			= "C"
						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extTenJungsanPrice		= extReducedPrice - extCommPrice
						dvlprice				= rsXL(11) 			'��ۺ�

						extItemName				= rsXL(7)			'��ǰ��
						extItemOptionName		= ""
						extitemoption			= ""				'�ɼ� �˼�����

						extVatYN = "Y"								'�������� �˼�����
						extitemid = rsXL(6)							'��ǰ��ȣ �˼�����

						extCommSupplyPrice		= extCommPrice
						extCommSupplyVatPrice	= 0

						extTenMeachulSupplyPrice	= extTenMeachulPrice
						extTenMeachulSupplyVatPrice	= 0

						IsValidInput = True
					Else
						IsValidInput = False
					End If
				Case "casamia_good_com"		'2021-02-02 ������..������ 1�� �ֹ��� �ִ�. 2�̻��� �� ������ �� ����..
					'// --------------------------------------------------------
					If (Len(rsXL(2)) = 14) then
						extOrderserial 	= rsXL(2)   				''�ֹ���ȣ
						extOrgOrderserial = ""
						extOrderserSeq = ""
						extCommPrice = 0
						extOwnCouponPrice = 0
						extTenCouponPrice = 0

						extMeachulDate		= LEFT(rsXL(1),10)  	''����Ȯ����

						extJungsanDate	= ""
						extItemNo		= rsXL(11)  					''����
						extOrderserSeq	= extOrderserial & "-" & rsXL(3)

						If (extItemNo < 0) Then						''���� 0�� �ƴϸ�
							IsReturnOrder = True
							extOrgOrderserial = extOrderserial
							extOrderserSeq	= extOrderserSeq & "-" & rsXL(3)
							extItemCost		= CLNG(rsXL(12)/extItemNo*100)/100
							extItemNo		= -1
'							extCommPrice	= CLNG(rsXL(42)/extItemNo*100)/100	'�ǸŴ��������
							'2021-04-02 �ϴ����� ����
							extCommPrice	= (CLNG(rsXL(14)/extItemNo*100)/100) - (CLNG(rsXL(25)/extItemNo*100)/100)	'�ǸŴ��������
						Else
							IsReturnOrder = False
							extItemCost		= CLNG(rsXL(12)/extItemNo*100)/100
							'2021-04-02 �ϴ����� ����
'							extCommPrice	= CLNG(rsXL(14)/extItemNo*100)/100	'�ǸŴ��������
							extCommPrice	= (CLNG(rsXL(14)/extItemNo*100)/100) - (CLNG(rsXL(25)/extItemNo*100)/100)	'�ǸŴ��������
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
						dvlprice				= rsXL(15) 			'��ۺ�

						extItemName				= rsXL(7)			'��ǰ��
						extItemOptionName		= rsXL(8)			'�ɼǸ�
						extitemoption			= ""				'�ɼ� �˼�����

						extVatYN = "Y"								'�������� �˼�����
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
						extOrderserial 	= rsXL(0)   							'�ֹ���ȣ
						extOrgOrderserial = ""
						extOrderserSeq = ""
						extCommPrice = 0
						extOwnCouponPrice = 0
						extTenCouponPrice = 0

'						extMeachulDate = LEFT(Replace(rsXL(0), "C", ""), 10)	'�ֹ��� �� �Ѿ��..�ֹ���ȣ�� ����
'						extMeachulDate = LEFT(extMeachulDate,4)&"-"&MID(extMeachulDate,5,2)&"-"&MID(extMeachulDate,7,2)
						extMeachulDate = LEFT(rsXL(86), 10)						'��ۿϷ�����

						extJungsanDate	= ""
						extItemNo		= rsXL(28)  							'����
						extOrderserSeq	= extOrderserial & "-" & rsXL(1)		'�ֹ���ȣ "-" �ֹ��󼼼���

						If (rsXL(28) < 0) Then									'���� 0�� �ƴϸ�..���̽� ���� �̹߰�
 							IsReturnOrder = True
 							extOrgOrderserial = extOrderserial
' 							extOrderserSeq	= extOrderserSeq & "-" & rsXL(1) & "-1" 
 							extItemNo		= -1
							extItemCost		= rsXL(22)								'�ǸŰ�
							extCommPrice	= (rsXL(22) - rsXL(20)) 				'�ǸŰ� - ���԰�
						Else
							IsReturnOrder = False
							extItemCost		= rsXL(22)								'�ǸŰ�
							extCommPrice	= (rsXL(22) - rsXL(20)) 				'�ǸŰ� - ���԰�
						End If

						extJungsanType			= "C"
						extOwnCouponPrice = CLNG((rsXL(35) / extItemNo * 100) / 100) 		'�� ���αݾ�
						extCommPrice = extCommPrice - extOwnCouponPrice
						extTenCouponPrice = 0

						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extTenJungsanPrice		= extReducedPrice - extCommPrice
						dvlprice				= rsXL(44) 						'�ǹ�ۺ�

						tmpItemname = ""
						tmpItemname = rsXL(18)

						If Instr(tmpItemname, "_") > 0 Then
							extItemName 		= Split(tmpItemname, "_")(0)
							extItemOptionName 	= Split(tmpItemname, "_")(1)
						Else
							extItemName = tmpItemname
							extItemOptionName = ""
						End If
						extitemoption			= ""				'�ɼ� �˼�����

						extVatYN = "Y"								'�������� �˼�����
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
					'// alphamall �󼼳���
					If (Len(rsXL(2)) = 21) then
						extOrderserial 	= rsXL(2)   				''�ֹ���ȣ
						extOrgOrderserial = ""
						extOrderserSeq = ""
						extCommPrice = ""
						extMeachulDate		= LEFT(rsXL(1),10)  				''����

						extJungsanDate	= ""
						extItemNo		= rsXL(7)  					''����
						extOrderserSeq	= extOrderserial &"-"& rsXL(0)

						extOwnCouponPrice = 0						'�����ݾ� �˼�����
						extTenCouponPrice = 0						'�����ݾ� �˼�����

						IsReturnOrder = False
						extItemCost	= rsXL(10)					'�ֹ��ݾ�
						extCommPrice			= CLNG(rsXL(10)) - CLNG(rsXL(8))	'������ = �ֹ��ݾ� - �ֹ�����

						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extJungsanType			= "C"

						extTenJungsanPrice		= extReducedPrice - extCommPrice
						dvlprice				= rsXL(12) 			'��ۺ�

						extItemName				= rsXL(6)			'��ǰ��
						extItemOptionName		= ""				'�ɼǸ� �˼�����
						extitemoption			= ""				'�ɼ� �˼�����

						extVatYN = "Y"								'�������� �˼�����
						extitemid = rsXL(5)							'��ǰ��ȣ

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
					'// alphamall �󼼳���
					If (Len(rsXL(2)) = 21) then
						extOrderserial 	= rsXL(2)   				''�ֹ���ȣ
						extOrgOrderserial = ""
						extOrderserSeq = ""
						extCommPrice = ""
						extMeachulDate		= LEFT(rsXL(1),10)  				''����

						extJungsanDate	= ""
						extItemNo		= rsXL(7) * -1 					''����
						extOrderserSeq	= extOrderserial &"-"& rsXL(0)

						extOwnCouponPrice = 0						'�����ݾ� �˼�����
						extTenCouponPrice = 0						'�����ݾ� �˼�����

						IsReturnOrder = True
						extOrgOrderserial = extOrderserial
						extOrderserSeq	= extOrderserSeq & "-" & rsXL(0)
						extItemCost	= rsXL(10)					'��ǰ�ݾ�
						extCommPrice			= CLNG(rsXL(10)) - CLNG(rsXL(8))	'������ = ��ǰ�ݾ� - ��ǰ����

						extTenMeachulPrice		= extItemCost - extOwnCouponPrice - extTenCouponPrice
						extReducedPrice			= CLng(extTenMeachulPrice)
						extJungsanType			= "C"

						extTenJungsanPrice		= extReducedPrice - extCommPrice
						dvlprice				= rsXL(12) 			'��ۺ�

						extItemName				= rsXL(6)			'��ǰ��
						extItemOptionName		= ""				'�ɼǸ� �˼�����
						extitemoption			= ""				'�ɼ� �˼�����

						extVatYN = "Y"								'�������� �˼�����
						extitemid = rsXL(5)							'��ǰ��ȣ

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
					'// ��������� �󼼳���
					'// 2019/11/01 ������ ������ (6)
					if (Len(rsXL(4)) = 10) then
						extMeachulDate = rsXL(23+1)  ''��ۿϷ���


						extJungsanDate = ""

						extItemNo				= rsXL(17+1)  ''����

						if (rsXL(22+1) <> "") and (extItemNo<0) then ''����� ������ ���̳ʽ��ΰŸ�
							extMeachulDate = rsXL(22+1)
						end if

						extOrderserial 			= rsXL(4)  ''���������ֹ���ȣ
						extOrderserSeq 			= rsXL(5) ''rsXL(8+1) & "-" & rsXL(0)			'// extOrderserSeq �� ����, ��ǰ�ڵ尡 �ִ�, ��ǰ�ڵ� => extOrderserSeq ��ȯ �ʿ� ���ν������� �ϰ���������.

						if (rsXL(22+1) <> "") then ''�����
							extOrderserSeq = extOrderserSeq&"-1"
						end if

						''�ϰ��� ��������.
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
						dvlprice				= CLNG(rsXL(14+1))	'// ��ۺ� ��ǰ�ݾװ� ���� ���ο� �ش�. ���Ŀ� ��ۺ񳻿� �����ʿ�.

						extItemCost				= CLNG((rsXL(11+1) - dvlprice) / extItemNo*100)/100
						extTenMeachulPrice		= extItemCost ''�����ݾ��� ���� ����. �ʸ��� 2019/05/08
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
					'// Ȩ�÷��� ���� �󼼳���

                Case "kakaogift"
					'// --------------------------------------------------------
					'// kakaogift ���� �󼼳���
					'// ��� ����(26) �߰���. 2019/05/22
					'// �Ǹ�����������(27) �߰��� 2019/12/19
					extOrderserial			= rsXL(2)

					if (Len(extOrderserial) = 9 or Len(extOrderserial) = 8 or Len(extOrderserial) = 10) and IsNumeric(extOrderserial) then

						extMeachulDate = rsXL(7)  ''���������
						extJungsanDate = ""

						extOrderserSeq			= "" ''rsXL(5) ''����... kakaoGIFT�� ��ǰ���Ÿ� �ִ�.

                        tenitemid =""
                        tenitemoption =""
                        extitemid = rsXL(15)  ''��ǰ��ȣ(kakaogift)
                        extitemoption = ""


                        extOrgOrderserial	= ""
						if (rsXL(13) <> "") then				'���/ȯ����
							extOrderserSeq = "-1"
						end if
						' 	'// ��ǰ
						' 	extOrgOrderserial	= ""  '' ����� ��.
						' else
						' 	extOrgOrderserial	= ""
						' end if

                        '' ��ۺ� ���� ����/ ��ǰ�뿡 ���Եȴ�..;; => �̰� ��� ó���ؾ��ұ�.
                        extJungsanType			= "C"
                        extItemNo				= rsXL(32)			'����

                        extItemCost				= CLNG(rsXL(29) / extItemNo*100)/100 ''������رݾ�(28)  ''�����ϴ�.
						extTenMeachulPrice		= extItemCost
						extReducedPrice			= CLNG(extTenMeachulPrice)

						extOwnCouponPrice		= 0                             '' ����
						extTenCouponPrice       = 0                             '' ����


						if rsXL(41) = 0 then		'�������հ�
							extCommPrice			= 0
						else
							extCommPrice			= CLNG(rsXL(41) / extItemNo*100)/100       ''��������(41)
						end if


                        extTenJungsanPrice      = CLNG(rsXL(42) / extItemNo*100)/100           ''�Ǹ�����ݾ�
						extItemName				= html2db(rsXL(16))
						extItemOptionName		= html2db(rsXL(18))

						extVatYN = "Y"
						' if (rsXL(20) = rsXL(21)) and (rsXL(20) <> 0) then             ''����.
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
					extOrderserial			= rsXL(0)			'������ȣ

					if (Len(extOrderserial) = 9 or Len(extOrderserial) = 8 or Len(extOrderserial) = 10) and IsNumeric(extOrderserial) then

						extMeachulDate = rsXL(7)  ''���������
						extJungsanDate = ""

						extOrderserSeq			= rsXL(2)		'�ֹ���ȣ

                        tenitemid =""
                        tenitemoption =""
                        extitemid = rsXL(15)  ''��ǰ��ȣ(kakaogift)
                        extitemoption = ""
						extOrgOrderserial	= ""
						if (rsXL(13) <> "") then
							extOrderserSeq = extOrderserSeq & "-1"
						end if
                        '' ��ۺ� ���� ����/ ��ǰ�뿡 ���Եȴ�..;; => �̰� ��� ó���ؾ��ұ�.

						If (rsXL(3) = "��ۺ�") Then
							extJungsanType			= "D"
							extItemNo				= 1
							extItemCost				= CLNG((rsXL(30))) + CLNG((rsXL(31)))		'���ҹ�ۺ�(30), ��ǰ��ۺ�(31)
							extOrderserSeq = "D-"&i
							extItemName				= rsXL(3)
						Else
                        	extJungsanType			= "C"
							extItemNo				= rsXL(32)					'����
							extItemCost				= CLNG(rsXL(29) / extItemNo*100)/100 ''������رݾ�(29)  ''�����ϴ�.
							extItemName				= html2db(rsXL(16))	'��ǰ��
						End If

						extTenMeachulPrice		= extItemCost
						extReducedPrice			= CLNG(extTenMeachulPrice)

						extOwnCouponPrice		= 0                             '' ����
						extTenCouponPrice       = 0                             '' ����

						if rsXL(41) = 0 then		'�������հ�
							extCommPrice			= 0
						else
							extCommPrice			= CLNG(rsXL(41) / extItemNo*100)/100       ''��������(41)
						end if

                        extTenJungsanPrice      = CLNG(rsXL(42) / extItemNo*100)/100           ''�Ǹ�����ݾ�
						extItemOptionName		= html2db(rsXL(18))	'�ɼǸ�

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
					'// coupang ���� �󼼳���
					'' ���ν��������������߰��� 2019/10/28 (17)
					'' �������� �ʵ��߰��� 4 (19,20,21,22)
					extOrderserial			= rsXL(0)

					if (Len(extOrderserial) >= 13) and IsNumeric(extOrderserial) then

						extMeachulDate = rsXL(24+3+1+4)  ''����Ȯ����
						if (rsXL(24+3+1+1+4)<>"") then ''��ҿϷ���
							extMeachulDate = rsXL(24+3+1+1+4)
						end if
						extJungsanDate = ""

						extOrderserSeq			= "" ''rsXL(5) ''����...

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

						'' �Ǹż���(9) �� ȯ�Ҽ���(10) �� ����  �Ѵ� + ���� ,��ȯ������ ����, ����.
						'' ��ȯ������ (26)
                        extOrgOrderserial	= ""
						if (extReItemNo <> "0") then
							'// ��ǰ
							extOrgOrderserial	= extOrderserial
						else
							extOrgOrderserial	= ""
						end if

						extOrderserSeq = extitemoption '' ��ǰID�� ����.

						''��ۺ�κ�.
						'' 17000021040462 ��ǰ�� ���̽��̳� �⺻��۷�� ���������̽��� ����. 9/22->9/13
						if (extitemoption = "<�⺻��۷�>") or (extitemoption = "<�߰���۷�>") then
							extJungsanType			= "D"

							if (rsXL(11)<>0) or (rsXL(20+3+1+4)<>0) then   ''�Ǹž�(A) or ������� ������.
								validitemno = 1
								extItemNo   = 1
							else
								validitemno = 0
								extItemNo   = 0
							end if

							if (rsXL(11)<0) then  ''�Ǹž��� ���̳ʽ��̴�..
								validitemno = -1
								extItemNo   = -1
							end if

							if (extitemoption = "<�߰���۷�>") then
								extOrderserSeq			= "D1"
								extItemName				= "��ۺ�"
								extItemOptionName		= "�߰���۷�"
							else
								extOrderserSeq			= "D"
								extItemName				= "��ۺ�"
								extItemOptionName		= "�⺻��۷�"
							end if
						end if

						extOwnCouponPrice		= rsXL(12)                      '' ������������(B) : ���� ���µ�.
						extTenCouponPrice       = rsXL(16)                      '' �Ǹ����������� (� ������ ������� ���̻�.)
						if (extOwnCouponPrice="") then extOwnCouponPrice=0
						if (extTenCouponPrice="") then extTenCouponPrice=0
						extItemCost				= rsXL(8)						'' �ǸŰ�(8) �Ǹž�(A)-11. :�Ǹż���-ȯ�Ҽ���=0 �̸� �Ǹž��� 0 �̴�.
						if (extJungsanType<>"C") then '' ��ۺ�� �Ǹž�(A)���� ����.
							extItemCost				= rsXL(11)
							if (validitemno<>0) then
								extItemCost = extItemCost/validitemno
							end if
						end if

						extItemCost = (extItemCost)
						extOwnCouponPrice = (extOwnCouponPrice)
						extTenCouponPrice = (extTenCouponPrice)

						extCommPrice			= rsXL(17+1)  ''rsXL(17)						''�����̿�� 10% �ΰ�������.  ���ν��������������߰��� 2019/10/28

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
						extReducedPrice			= CLNG(extTenMeachulPrice) ''������رݾ�(26) ''����ݾ�(D=A-B) �Ǹ��� ����������?

                        extTenJungsanPrice      = extTenMeachulPrice-extCommPrice 			 ''rsXL(20)   ''�Ǹ�����ݾ� �������� �������..

						''����Ȯ���ϰ� ����Ϸ����� �ٸ����� ���ں��� �ٽ� �־�� �Ѵ�..
						if (validitemno=0) and (extJungsanType="C") and (rsXL(13)=0) and (rsXL(17+1)=0) then
						 	extItemNo = 0
						end if

						if (validitemno<0)   then ''��ǰ
							extOrderserSeq = extOrderserSeq&"-1"
						end if

						if (rsXL(26+3+1+4)<>"") then '' ��ȯ�������� �ִ�..
							if (extChgItemno<>0)   then ''��ȯ ������ ������
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
				if (((p_extOrderserial = extOrderserial) and (LEFT(p_extOrderserSeq,LEN(extOrderserSeq)) = extOrderserSeq)) or (sheetName="'���������Ȳ(��ǰ�ù��)$'")) and (extItemNo<>0) then ''�������� �������� ��������.
					'' SSG  20180719769649 �ߺ����̽�.
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

			''��ۺ� �ߺ� ���̽� ����. (20190412-296463)
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
				' if (sellsite="coupang") then ''��ǰ�� �����ʵ�� �ִ�.
				' 	if (CStr(extReItemNo)<>"0") then
				' 		extItemNo=extReItemNo*-1
				' 		extOrderserial = extOrderserial&"-1"
				' 		''extOrderserSeq = extOrderserSeq&"-1"

				' 		''rw  extOrderserial
				' 	end if
				' end if
				If (sellsite="yes24") Then
					if (dvlprice<>0) then  '��ۺ� ���ٷ� �ִ�.
						extJungsanType = "D"
						extOrderserSeq = extOrderserSeq+"-D"
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extCommPrice			= 0

						If (rsXL(4) <> 0) Then						''��ǰ�ݾ��� 0���� ũ�ٸ�
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
					if (dvlprice<>0) then  '��ۺ� ���ٷ� �ִ�.
						extJungsanType = "D"
						extOrderserSeq = extOrderserSeq+"-D"
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extCommPrice			= 0
						
						If (extItemNo < 0) Then						''��ǰ�ݾ��� 0���� ũ�ٸ�
							extItemNo = -1
						Else
							extItemNo = 1
							' extCommPrice = rsXL(4) - rsXL(6)
							' extTenJungsanPrice	= dvlprice/extItemNo - extCommPrice
						End If

						extCommPrice			= CLNG(rsXL(13) / extItemNo )		'������D
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
					if (dvlprice<>0) then  '��ۺ� ���ٷ� �ִ�.
						extJungsanType = "D"
						extOrderserSeq = extOrderserSeq+"-D"
						extOwnCouponPrice		= 0
						extTenCouponPrice		= 0
						extCommPrice			= 0

						If (extItemNo < 0) Then						''��ǰ�ݾ��� 0���� ũ�ٸ�
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
					if (dvlprice<>0) then  '��ۺ� ���ٷ� �ִ�.
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
					if (dvlprice<>0) then  '��ۺ� ���ٷ� �ִ�.
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
					If (dvlprice <> 0) Then  '��ۺ� ���ٷ� �ִ�.
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
					if (dvlprice<>0) then  '��ۺ� ���ٷ� �ִ�.
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

						''extVatYN = "Y" �񱳸� �����ϱ����� ������
						''if (rsXL(6) = "�鼼") then
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
					if (dvlprice<>0) then  '��ۺ� ���ٷ� �ִ�.
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
					if (dvlprice<>0) then  '��ۺ� ���ٷ� �ִ�.
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

					'' SSG�� ��ۺ� �Ⱥ��ϴµ� ������ ������ ������ ������.
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
	response.write "ERROR : ������ �߻��߽��ϴ�. �ý����� ����[3]" & errMSG & extsellsite
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
	'' XL�� ���ε� �Ҷ� ���������� ����ĥ�� ��������..
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
alert("����Ǿ����ϴ�. ");
location.href = "<%= manageUrl %>/common/popReloadOpener.asp";
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
