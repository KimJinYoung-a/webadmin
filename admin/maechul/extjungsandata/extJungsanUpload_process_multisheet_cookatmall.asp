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
dim i

'// ============================================================================
'// 업로드 컨퍼넌트 선언 //
dim fso
dim uprequest, sqlStr

dim objfso
set objfso = server.CreateObject("scripting.Filesystemobject")

if not objfso.FolderExists(DefaultPath) then
	objfso.CreateFolder(DefaultPath)
end if

set uprequest = Server.CreateObject("TABSUpload4.Upload")
uprequest.Start DefaultPath

extsellsite = requestCheckvar(uprequest("extsellsite"),32)

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


fullpath = uprequest("sFile").saveas((DefaultPath + filename + "_" + extsellsite + "."+ uprequest("sFile").FileType), True)


'// ============================================================================
dim conXL, rsXL, sheetNameArr

set conXL = Server.CreateObject("ADODB.Connection")
if (extsellsite="11st1010") then
    conXL.Provider = "Microsoft.Jet.oledb.4.0"
    conXL.Properties("ExtEnded Properties").Value = "Excel 8.0;IMEX=1"
else
	conXL.Provider = "Microsoft.ACE.OLEDB.12.0"
    conXL.Properties("ExtEnded Properties").Value = "Excel 12.0;IMEX=1"
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

if (extsellsite="cookatmall") then
	extsellsite = "cookatmall"
	redim sheetNameArr(1)
	sheetNameArr=Array("텐바이텐$","배송비$")
end if

if Not IsArray(sheetNameArr) then
	response.write "Sheet not Define"
	response.end
end if



''지우기 먼저.
sqlStr = " delete from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite = '" + CStr(extsellsite) + "' "
'response.write sqlStr
'response.end
dbget.execute sqlStr

Dim retVal, retErrStr
for i=LBOUND(sheetNameArr) to UBOUND(sheetNameArr)
	retVal = fnOneSheetUpload(extsellsite,sheetNameArr(i),conXL,i,retErrStr)
	rw sheetNameArr(i)&":"&retVal
	response.flush

	if (NOT retVal) then Exit for
next

set conXL = Nothing

if (NOT retVal) then
	uprequest("sFile").delete
	set objfso = Nothing
	set uprequest = Nothing

	rw retErrStr
	dbget.close()
	response.end
end if


uprequest("sFile").delete
set objfso = Nothing
set uprequest = Nothing

%>
<script>
alert("저장되었습니다. ");
location.href = "<%= manageUrl %>/common/popReloadOpener.asp";
</script>

<%





function fnOneSheetUpload(iextsellsite,isheetname,conXL,sheetN,retErrStr)
	'// ============================================================================
	fnOneSheetUpload = FALSE

	dim sellsite, extOrderserial, extOrderserSeq, extOrgOrderserial, extItemNo, extItemCost, extReducedPrice, extOwnCouponPrice, extTenCouponPrice, extJungsanType, extCommPrice, extTenMeachulPrice, extTenJungsanPrice
	dim extMeachulDate, extJungsanDate
	dim extItemName, extItemOptionName
	dim IsOrderData, IsValidInput, IsReturnOrder
	dim extVatYN, extCommSupplyPrice, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice
	dim tmpDate, tmpDate2, totExtCommSupplyPrice, totExtTenMeachulSupplyPrice, totExtCommSupplyPrice2, totExtTenMeachulSupplyPrice2
	dim totExtTenMeachulPrice, totExtCommPrice, totExtTenMeachulPrice2, totExtCommPrice2
	dim tenitemid,tenitemoption,extitemid,extitemoption
	dim tmpStr
	dim i
	dim diffPrice
	dim plusMinus
	dim validitemno, extReItemNo
	dim siteNo, dvlprice
	dim p_extOrderserial, p_extOrderserSeq
	Dim casesite, AssignedRow

	set rsXL = Server.CreateObject("ADODB.Recordset")

	sqlStr = "select * from [" & isheetname & "]"

	rsXL.Open sqlStr, conXL

	if (ERR) then
		retErrStr = "ERROR : 오류가 발생했습니다. 시스템팀 문의[1]"
		exit function
	end if

	sellsite = iextsellsite
	casesite = iextsellsite&isheetname

	If Not rsXL.Eof Then

		errMSG = ""
		extJungsanDate = ""
		i = 0
		do Until rsXL.Eof

			IsOrderData = False
			IsValidInput = True
			''IsReturnOrder = False

			Select Case casesite
				Case "cookatmall텐바이텐$"
					if (rsXL(1) <> "") then
						if (Len(rsXL(1)) = 17) then
							IsOrderData = True
						end if
					end if
				Case "cookatmall배송비$"
					if (rsXL(0) <> "") then
						if (Len(rsXL(0)) = 17)  then
							IsOrderData = True
						end if
					end if
				Case Else
					IsOrderData = False
			End Select

			if (IsOrderData = True) then
				dvlprice = 0
				Select Case casesite
					Case "cookatmall텐바이텐$"
						'// 쿠켓몰 상품정산 상세내역
						extOrderserial			= rsXL(1)

						if (Len(extOrderserial) = 17)  then
							extMeachulDate = Left(rsXL(1),4) &"-"& mid(rsXL(1),5,2) &"-"& mid(rsXL(1),7,2)
							'매출데이터 안 넘어옴..강제로 엑셀에 입력
							'2021-09-02 김진영..정산셀 변경됨으로 -1처리
							If rsXL(31-1) <> "" Then
								extMeachulDate = rsXL(31-1)
							End If

							extVatYN = "Y"
							extOrderserSeq			= rsXL(2)		'품목별 주문번호를 키로쓰자
							If rsXL(0) = "반품완료" Then
								extOrderserSeq = extOrderserSeq & "-1"
							End If
							extOrgOrderserial		= ""

							extItemNo				= CLNG(rsXL(9))	 '수량
							extItemCost				= ABS(CLNG(rsXL(11)))   ''판매가
							extOwnCouponPrice		= ""
							extTenCouponPrice		= ""

							extReducedPrice			= ABS(CLNG(rsXL(11)))	'실판매단가

							extJungsanType			= "C"
							extCommPrice			= ROUND( ( (CLNG(rsXL(13)) - CLNG(rsXL(12)) )/ extItemNo),0)
							extCommSupplyPrice		= extCommPrice
							extCommSupplyVatPrice	= 0

							extTenMeachulPrice			= extReducedPrice
							extTenMeachulSupplyPrice	= extTenMeachulPrice
							extTenMeachulSupplyVatPrice	= 0

							extTenJungsanPrice		= extReducedPrice - extCommPrice  ''== ROUND(rsXL(32) / extItemNo,0)
							extItemName				= rsXL(7)
							extItemOptionName		= rsXL(8)

							' Select Case rsXL(8)
							' 	Case "[피너츠] 스누피와 친구들 머그컵_스누피"
							' 		extitemid		= "2785591"
							' 		extitemoption	= "0011"
							' 	Case "[피너츠] 스누피와 친구들 머그컵_샐리"
							' 		extitemid		= "2785591"
							' 		extitemoption	= "0015"
							' 	Case "[피너츠] 스누피와 친구들 머그컵_우드스탁"
							' 		extitemid		= "2785591"
							' 		extitemoption	= "0012"
							' 	Case "[피너츠] 스누피와 친구들 머그컵_찰리브라운"
							' 		extitemid		= "2785591"
							' 		extitemoption	= "0013"
							' 	Case "[피너츠] 스누피와 친구들 머그컵_루시"
							' 		extitemid		= "2785591"
							' 		extitemoption	= "0014"
							' 	Case "[피너츠] 스누피와 친구들 머그컵_라이너스"
							' 		extitemid		= "2785591"
							' 		extitemoption	= "0016"
							' 	Case "[피너츠] 프렌즈 유리컵_스누피 (+2,000원)"
							' 		extitemid		= "3649588"
							' 		extitemoption	= "0011"
							' 	Case "[피너츠] 프렌즈 유리컵_찰리브라운 (+2,000원)"
							' 		extitemid		= "3649588"
							' 		extitemoption	= "0012"
							' End Select
						else
							IsValidInput = False
						end if
					Case "cookatmall배송비$"
						extOrderserial			= rsXL(0)		'주문번호

						if (Len(extOrderserial) = 17)  then

							extMeachulDate = Left(rsXL(0),4) &"-"& mid(rsXL(0),5,2) &"-"& mid(rsXL(0),7,2)
							'매출데이터 안 넘어옴..강제로 엑셀에 입력
							If rsXL(6) <> "" Then
								extMeachulDate = rsXL(6)
							End If

							extJungsanDate = ""
							extVatYN = "Y"
							extOrderserSeq			= "1-D"
							'주문번호가 두개 넘어옴..seq 잡아야하는 데, 방도가 없음..2021-08-03 김진영
							If extOrderserial= "20210613553464876" AND rsXL(1) = "14900" Then
								extOrderserSeq			= "1-DD"
							End If
							extOrgOrderserial		= ""

							extItemNo				= 1	 ''판매량
							extItemCost				= CLNG(rsXL(5))
							extOwnCouponPrice		= 0
							extTenCouponPrice		= 0
							extReducedPrice			= extItemCost

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
					Case Else
						IsValidInput = False
				End Select

				if (IsValidInput = False) then
					Exit Do
				end if

				''response.write extOrgOrderserial & "---<br>"

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
					dbget.execute sqlStr
				end if
			end if

			i = i + 1
			rsXL.MoveNext
		loop
	end if
	rsXL.Close
	set rsXL = Nothing

	if (IsValidInput = False) then
		retErrStr = "ERROR : 오류가 발생했습니다. 시스템팀 문의[3]" & errMSG & extsellsite
		exit function
	end if

	if (sellsite="cookatmall") and casesite = "cookatmall배송비$" then
		sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_cookatmall] "
		dbget.Execute sqlStr, AssignedRow
	else
		rw "TT"
	end if

	fnOneSheetUpload = TRUE
end function
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->