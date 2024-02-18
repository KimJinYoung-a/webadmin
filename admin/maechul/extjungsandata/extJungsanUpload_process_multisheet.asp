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
'// ���ε� ���۳�Ʈ ���� //
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
	response.write "ERROR : ������ �߻��߽��ϴ�. �ý����� ����[0]"
	response.end
end if

set rsXL = conXL.OpenSchema(adSchemaTables)

if (extsellsite="interpark") then
	redim sheetNameArr(2)
	sheetNameArr=Array("��ǰ����$","�߰���ۺ�$","������ũ��ۺ�����$")
end if

if Not IsArray(sheetNameArr) then
	response.write "Sheet not Define"
	response.end
end if



''����� ����.
sqlStr = " delete from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite = '" + CStr(extsellsite) + "' "
''response.write sqlStr
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
alert("����Ǿ����ϴ�. ");
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
		retErrStr = "ERROR : ������ �߻��߽��ϴ�. �ý����� ����[1]"
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
				Case "interpark��ǰ����$"
					if (rsXL(0) <> "") then
						if (IsNumeric(rsXL(0)) = True)  then
							IsOrderData = True
						end if
					end if
				Case "interpark�߰���ۺ�$"
					if (rsXL(0) <> "") then
						if (IsNumeric(rsXL(0)) = True)  then
							IsOrderData = True
						end if
					end if
				Case "interpark������ũ��ۺ�����$"
					if (rsXL(0) <> "") then
						if (IsNumeric(rsXL(0)) = True)  then
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
				Select Case casesite
					Case "interpark��ǰ����$"
						'// ������ũ ��ǰ���� �󼼳���
						extOrderserial			= rsXL(4)

						if (Len(extOrderserial) = 20)  then

							extMeachulDate = Left(rsXL(3), 4) & "-" & Right(Left(rsXL(3), 6), 2) & "-" & Right(rsXL(3), 2)
							extJungsanDate = ""

							extVatYN = "Y"
							if rsXL(13)="�鼼" then extVatYN = "N"

							extOrderserSeq			= rsXL(5)
							if (rsXL(6)<>"") then extOrderserSeq = extOrderserSeq + "-"&rsXL(6) ''Ŭ��������
							extOrgOrderserial		= ""

							extItemNo				= CLNG(rsXL(24))	 ''�Ǹŷ�
							extItemCost				= ABS(CLNG(rsXL(14)))   ''�ǸŴܰ�
							extOwnCouponPrice		= ABS(CLNG(rsXL(16))+CLNG(rsXL(17))+CLNG(rsXL(19))+CLNG(rsXL(20)))
							extTenCouponPrice		= ABS(CLNG(rsXL(18))+CLNG(rsXL(15)))
							extReducedPrice			= ABS(CLNG(rsXL(21)))

							extJungsanType			= "C"

							extCommPrice			= ROUND( ( (CLNG(rsXL(27)) + CLNG(rsXL(30)) + CLNG(rsXL(31)) )/ extItemNo),0) - extOwnCouponPrice
							extCommSupplyPrice		= extCommPrice
							extCommSupplyVatPrice	= 0

							extTenMeachulPrice			= extReducedPrice
							extTenMeachulSupplyPrice	= extTenMeachulPrice
							extTenMeachulSupplyVatPrice	= 0

							extTenJungsanPrice		= extReducedPrice - extCommPrice  ''== ROUND(rsXL(32) / extItemNo,0)

							extitemid		= rsXL(9)
							extitemoption	= rsXL(11)


						else
							IsValidInput = False
						end if
					Case "interpark�߰���ۺ�$"
						extOrderserial			= rsXL(4)

						if (Len(extOrderserial) = 20)  then

							extMeachulDate = Left(rsXL(3), 4) & "-" & Right(Left(rsXL(3), 6), 2) & "-" & Right(rsXL(3), 2)
							extJungsanDate = ""

							extVatYN = "Y"


							extOrderserSeq			= "1-D"
							if (rsXL(5)<>"") then extOrderserSeq = "1-"&rsXL(5)&"-D" ''Ŭ��������

							extOrgOrderserial		= ""

							extItemNo				= 1	 ''�Ǹŷ�
							extItemCost				= CLNG(rsXL(8))+CLNG(rsXL(9))   ''�߰���ۺ�+��ǰ��ۺ�  //���̳ʽ��� ������ �ִ�.
							extOwnCouponPrice		= 0
							extTenCouponPrice		= 0
							extReducedPrice			= extItemCost

							extJungsanType			= "D"
							if (CLNG(rsXL(9))<>0) then extOrderserSeq=extOrderserSeq+"D"

							extCommPrice			= CLNG(rsXL(10))				''��ۺ������ 2021-05-03 ������ �߰�
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

					Case "interpark������ũ��ۺ�����$"
						extOrderserial			= rsXL(4)

						if (Len(extOrderserial) = 20)  then

							extMeachulDate = Left(rsXL(3), 4) & "-" & Right(Left(rsXL(3), 6), 2) & "-" & Right(rsXL(3), 2)
							extJungsanDate = ""

							extVatYN = "Y"


							extOrderserSeq			= "1-DC"
							if (rsXL(5)<>"") then extOrderserSeq = "1-"&rsXL(5)&"-DC" ''Ŭ��������

							extOrgOrderserial		= ""

							extItemNo				= 1	 ''�Ǹŷ�
							extItemCost				= CLNG(rsXL(8))+CLNG(rsXL(9))   ''�߰���ۺ�+��ǰ��ۺ�  //���̳ʽ��� ������ �ִ�.
							extOwnCouponPrice		= 0 ''CLNG(rsXL(8))+CLNG(rsXL(9))
							extTenCouponPrice		= 0
							extReducedPrice			= extItemCost-extOwnCouponPrice

							extJungsanType			= "D"
							if (CLNG(rsXL(9))<>0) then extOrderserSeq=extOrderserSeq+"D"

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

				if (sellsite="ssg") then
					if (p_extOrderserial = extOrderserial) and (p_extOrderserSeq = extOrderserSeq) and (extItemNo<>0) then ''�������� �������� ��������.
						'' SSG  20180719769649 �ߺ����̽�.
						sqlStr = " select count(*) CNT from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite='"&sellsite&"' and extOrderserial='"&extOrderserial&"' and extOrderserSeq='" &extOrderserSeq&"'" &vbCRLF
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

				''��ۺ�.
				' if (sellsite="coupang") or (sellsite="ssg") then
				' 	''validitemno
				' 	if (sellsite="coupang") then ''��ǰ�� �����ʵ�� �ִ�.
				' 		if (CStr(extReItemNo)<>"0") then
				' 			extItemNo=extReItemNo*-1
				' 			extOrderserSeq = extOrderserSeq&"-1"
				' 		end if
				' 	end if

				' 	if (sellsite="ssg") and (dvlprice<>0) then  '��ۺ� ���ٷ� �ִ�.
				' 		extJungsanType = "D"
				' 		extOrderserSeq = extOrderserSeq+"-D"

				' 		if (extItemNo<0) then
				' 			extItemNo = -1
				' 			extOwnCouponPrice		= 0
				' 			extTenCouponPrice		= 0
				' 			extCommPrice			= 0
				' 		elseif (extItemNo=0) then
				' 			extItemNo = 1
				' 			extOrderserSeq = extOrderserSeq+"D"
				' 			if (dvlprice<0) then extOrderserSeq = extOrderserSeq+"D"

				' 			'extOwnCouponPrice		= 0   ''20180807467429 CASE
				' 			'extTenCouponPrice		= 0
				' 			extCommPrice			= extOwnCouponPrice*-1   ''20180807467429 CASE
				' 		else
				' 			extItemNo = 1
				' 			extOwnCouponPrice		= 0
				' 			extTenCouponPrice		= 0
				' 			extCommPrice			= 0
				' 		end if


				' 		extItemCost				= dvlprice/extItemNo
				' 		extTenJungsanPrice		= dvlprice/extItemNo


				' 		extReducedPrice			= dvlprice/extItemNo-extOwnCouponPrice
				' 		extTenMeachulPrice      = dvlprice/extItemNo-extOwnCouponPrice

				' 		''extVatYN = "Y" �񱳸� �����ϱ����� ������
				' 		''if (rsXL(6) = "�鼼") then
				' 		''	extVatYN = "N"
				' 		''end if

				' 		extCommSupplyPrice		= 0
				' 		extCommSupplyVatPrice	= 0

				' 		extTenMeachulSupplyPrice	= 0
				' 		extTenMeachulSupplyVatPrice	= 0
				' 	else
				' 		extItemNo = 0
				' 	end if

				' 	if (extItemNo<>0) then

				' 		'' SSG�� ��ۺ� �Ⱥ��ϴµ� ������ ������ ������ ������.
				' 		if (sellsite="ssg") and (RIGHT(extOrderserSeq,1)="D") then
				' 			sqlStr = " select count(*) CNT from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite='ssg' and extOrderserial='"&extOrderserial&"' and extOrderserSeq='" &extOrderserSeq&"'" &vbCRLF
				' 			rsget.CursorLocation = adUseClient
				' 			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				' 			if Not rsget.Eof then
				' 				if (rsget("CNT")>0) then
				' 					extOrderserSeq = extOrderserSeq + "-"&rsget("CNT")
				' 				end if
				' 			end if
				' 			rsget.Close
				' 		end if


				' 		sqlStr = " insert into db_temp.dbo.tbl_xSite_JungsanTmp"
				' 		sqlStr = sqlStr + " (sellsite, extOrderserial, extOrderserSeq"
				' 		sqlStr = sqlStr + ", extOrgOrderserial, extItemNo, extItemCost"
				' 		sqlStr = sqlStr + ", extReducedPrice, extOwnCouponPrice, extTenCouponPrice"
				' 		sqlStr = sqlStr + ", extJungsanType, extCommPrice, extTenMeachulPrice"
				' 		sqlStr = sqlStr + ", extTenJungsanPrice, extMeachulDate, extJungsanDate"
				' 		sqlStr = sqlStr + ", extItemName, extItemOptionName, extVatYN"
				' 		sqlStr = sqlStr + ", extCommSupplyPrice, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice"
				' 		sqlStr = sqlStr + ", itemid, itemoption"
				' 		sqlStr = sqlStr + ", extitemid, extitemoption,siteNo"
				' 		sqlStr = sqlStr + " ) "
				' 		sqlStr = sqlStr + " values('" + CStr(sellsite) + "', '" + CStr(extOrderserial) + "', '" + CStr(extOrderserSeq) + "'"
				' 		sqlStr = sqlStr + ", '" + CStr(extOrgOrderserial) + "', '" + CStr(extItemNo) + "', '" + CStr(extItemCost) + "'"
				' 		sqlStr = sqlStr + ", '" + CStr(extReducedPrice) + "', '" + CStr(extOwnCouponPrice) + "', '" + CStr(extTenCouponPrice) + "'"
				' 		sqlStr = sqlStr + ", '" + CStr(extJungsanType) + "', '" + CStr(extCommPrice) + "', '" + CStr(extTenMeachulPrice) + "'"
				' 		sqlStr = sqlStr + ", '" + CStr(extTenJungsanPrice) + "', '" + CStr(extMeachulDate) + "', '" + CStr(extJungsanDate) + "'"
				' 		sqlStr = sqlStr + ", '" + CStr(extItemName) + "', convert(varchar(128),'" & CStr(extItemOptionName) & "'), '" + CStr(extVatYN) + "'"
				' 		sqlStr = sqlStr + ", '" + CStr(extCommSupplyPrice) + "', '" + CStr(extCommSupplyVatPrice) + "', '" + CStr(extTenMeachulSupplyPrice) + "', '" + CStr(extTenMeachulSupplyVatPrice) + "'"
				' 		if (tenitemid<>"") then
				' 			sqlStr = sqlStr + ", '" + CStr(tenitemid) + "'"
				' 		else
				' 			sqlStr = sqlStr + ", NULL"
				' 		end if
				' 		if (tenitemoption<>"") then
				' 			sqlStr = sqlStr + ", '" + CStr(tenitemoption) + "'"
				' 		else
				' 			sqlStr = sqlStr + ", NULL"
				' 		end if
				' 		if (extitemid<>"") then
				' 			sqlStr = sqlStr + ", '" + CStr(extitemid) + "'"
				' 		else
				' 			sqlStr = sqlStr + ", NULL"
				' 		end if
				' 		if (extitemoption<>"") then
				' 			sqlStr = sqlStr + ", '" + CStr(extitemoption) + "'"
				' 		else
				' 			sqlStr = sqlStr + ", NULL"
				' 		end if
				' 		if (siteNo<>"") then
				' 			sqlStr = sqlStr + ", '" + CStr(siteNo) + "'"
				' 		else
				' 			sqlStr = sqlStr + ", NULL"
				' 		end if
				' 		sqlStr = sqlStr + ") "

				' 		dbget.execute sqlStr
				' 	end if
				' end if

			end if

			i = i + 1
			rsXL.MoveNext
		loop
	end if
	rsXL.Close
	set rsXL = Nothing

	if (IsValidInput = False) then
		retErrStr = "ERROR : ������ �߻��߽��ϴ�. �ý����� ����[3]" & errMSG & extsellsite
		exit function
	end if

	if (sellsite="interpark") then
		sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_interpark] " &sheetN
		dbget.Execute sqlStr, AssignedRow

	elseif (sellsite="ssg1111") then
		'' XL�� ���ε� �Ҷ� ���������� ����ĥ�� ��������..
		' sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_ssg] 1"
		' dbget.Execute sqlStr, AssignedRow

	else
		rw "TT"
	end if

	fnOneSheetUpload = TRUE
end function



%>


<!-- #include virtual="/lib/db/dbclose.asp" -->
