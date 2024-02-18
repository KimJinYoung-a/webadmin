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
dim DefaultPath : DefaultPath =	server.mappath("/admin/etc/difforder/upFiles/") 
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
if (extsellsite="kakaogift")  then
    conXL.Provider = "Microsoft.ACE.OLEDB.12.0"
    conXL.Properties("ExtEnded Properties").Value = "Excel 12.0;IMEX=1"
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
dim sellsite
dim extItemName, extItemOptionName
dim IsItemData, IsValidInput

dim i
dim tenitemid,tenitemoption,extitemid,extitemoption
dim tmpStr



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


If Not rsXL.Eof Then

	''ADD EXT SHOP. 01. 사이트구분
	Select Case extsellsite
		Case "lotteCom"
			sellsite = "lotteCom"
		Case "lotteimall"
			sellsite = "lotteimall"
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

    'sqlStr = " delete from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite = '" + CStr(sellsite) + "' "
    ''response.write sqlStr
    'dbget.execute sqlStr
	'Call checkAndWriteElapsedTime("003")

	errMSG = ""
	i = 0
	do Until rsXL.Eof

		IsItemData = False
		IsValidInput = True
		''IsReturnOrder = False

		'// 데이타 검증 01 (주문내역인지)
		Select Case extsellsite
			Case "lotteCom"
				if (Len(rsXL(21+1)) = 18) and (IsNumeric(rsXL(8+1)) = True) then
					IsItemData = True
				end if
			
			Case "lotteimall"
				if (rsXL(0) <> "") then
					if (IsNumeric(rsXL(0)) = True) then
						IsItemData = True
					end if
				end if
			Case Else
				IsItemData = False
		End Select

		if (IsItemData = True) then

			'//ADD EXT SHOP. 03. 처리
			Select Case extsellsite
					
				
				Case "lotteCom"  ''2016/05/01 주문상세번호 필드 생긴듯.
					'// --------------------------------------------------------
					'// 롯데닷컴 위탁정산 상세내역
					if (rsXL(26) = "텐바이텐") then  ''24+1 =>24+2 2016/05/01

						extMeachulDate = rsXL(0)
						extJungsanDate = ""

						extItemNo				= rsXL(9)
						plusMinus				= extItemNo / Abs(extItemNo)
						if (extItemNo >= 0) then
							'// 정상출고
							extOrderserial 			= Replace(rsXL(22), "-", "")
							extOrderserSeq			= TRim(rsXL(23)) ''=>  2016/05/01 //
							extOrgOrderserial		= ""
						else
							'// 반품
							extOrderserial 			= Replace(rsXL(22), "-", "") & "-" & i
							extOrderserSeq			= TRim(rsXL(23)) ''=>  2016/05/01 //
							extOrgOrderserial		= Replace(rsXL(22), "-", "")
						end if

						extVatYN = "Y"
						if (rsXL(12) = 0) then
							extVatYN = "N"
						end if

						extItemCost				= CLNG(rsXL(11)  / extItemNo * 100)/100
						extTenCouponPrice		= CLNG((rsXL(14) - rsXL(18)) / extItemNo * 100)/100
						extTenMeachulPrice		= CLNG(rsXL(15)  / extItemNo * 100)/100
						extOwnCouponPrice		= extItemCost - extTenMeachulPrice - extTenCouponPrice

						extJungsanType			= "C"

						extCommPrice			= CLNG((rsXL(17) - rsXL(18))  / extItemNo * 100)/100
						extReducedPrice			= CLNG(extTenMeachulPrice)
						extTenJungsanPrice		= CLNG(rsXL(21) / extItemNo * 100)/100

						extItemName				= html2db(rsXL(5))			'// 외부몰 상품명이 바뀐다. 상품명 대신 외부몰 상품코드로 매칭
						extItemOptionName		= html2db(rsXL(7))						'// 정산내역에 옵션정보가 없다. ==>있음.

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

						
					else
						IsValidInput = False
					end if
				Case "halfclubbeasongpay"
					'// --------------------------------------------------------
					'// 하프클럽 배송비 상세내역
					if (LEN(rsXL(1)) = 12) then  
						
					else
						IsValidInput = False
					end if
				Case "nvstorefarm"
					'// --------------------------------------------------------
					'// 스토어팜 상세내역
					if (Len(rsXL(0)) = 16) then
						
					else
						IsValidInput = False
					end if
				Case "ezwel"
					'// --------------------------------------------------------
					'// 이지웰페어 상세내역
					if (Len(rsXL(4)) = 10) then
					
					else
						IsValidInput = False
					end if
				Case "homeplus"
					'// --------------------------------------------------------
					'// 홈플러스 정산 상세내역
					
                Case "kakaogift"
					'// --------------------------------------------------------
					'// kakaogift 정산 상세내역

					if (Len(extOrderserial) = 9) and IsNumeric(extOrderserial) then

					
                        IsValidInput = True
					else
						IsValidInput = False
					end if
				Case "coupang"
					'// --------------------------------------------------------

					if (Len(extOrderserial) >= 13) and IsNumeric(extOrderserial) then

						
                        IsValidInput = True
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

            	'dbget.execute sqlStr
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

''response.end

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
if (sellsite="lotteCom") then
	sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_lotteCom]"
'	dbget.Execute sqlStr, AssignedRow
else
	rw "TT"
end if


%>
<script>
alert("저장되었습니다. ");
location.href = "<%= manageUrl %>/common/popReloadOpener.asp";
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
