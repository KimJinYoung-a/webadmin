<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2010.09.10 이상구 생성
'			   2011.06.14 한용민 수정
'	Description : 주문 엑셀파일 수기 등록
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/db/dbTPLHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/orderInput/incUploadFunction.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

Dim paramInfo, retParamInfo, RetErr, retErrStr, sqlStr ,iLine, iResult ,POS1,POS2,POS3, okCNT
Dim xlPosArr, ArrayLen, skipString, afile, aSheetName ,i,j ,xlRowALL
Dim uploadform, objfile, sDefaultPath, sFolderPath
Dim iML, sFile, sFilePath, xltype, iMaxLen, sUploadPath, orgFileName, maybeSheetName
Dim monthFolder : monthFolder = Replace(Left(CStr(now()),7),"-","")
dim partnerItemName, partnerOptionName, barcode
dim makerid, itemgubun, itemid, itemoption, itemname, optionname, orgprice, mainimage, listimage, smallimage
dim itemoptionname
dim tplcompanyid

IF (application("Svr_Info")	= "Dev") then
    Set uploadform = Server.CreateObject("TABS.Upload")	   '' - TEST : TABS.Upload
ELSE
    Set uploadform = Server.CreateObject("TABSUpload4.Upload")	''REAL : TABSUpload4.Upload
END IF

Set objfile	   = Server.CreateObject("Scripting.FileSystemObject")
sDefaultPath   = Server.MapPath("/admin/etc/orderInput/upFiles/")

uploadform.Start sDefaultPath '업로드경로

iMaxLen 		= uploadform.Form("iML")	'이미지파일크기
xltype 			= uploadform.Form("xltype")
tplcompanyid	= uploadform.Form("tplcompanyid")

IF (fnChkFile(uploadform("sFile"), iMaxLen,"xls")) THEN	'파일체크

    '폴더 생성
    sFolderPath = sDefaultPath&"/tplorder/"
    IF NOT  objfile.FolderExists(sFolderPath) THEN
    	objfile.CreateFolder sFolderPath
    END IF

    sFolderPath = sDefaultPath&"/order/"&monthFolder&"/"
    IF NOT  objfile.FolderExists(sFolderPath) THEN
    	objfile.CreateFolder sFolderPath
    END IF

    '파일저장
	sFile = fnMakeFileName(uploadform("sFile"))
	sFilePath = sFolderPath&sFile
	sFilePath = uploadform("sFile").SaveAs(sFilePath, False)

	orgFileName = uploadform("sFile").FileName
	maybeSheetName = Replace(orgFileName,"."&uploadform("sFile").FileType,"")

END IF

Set objfile		= Nothing
Set uploadform = Nothing

''''			아래 순서대로 한다.
''''			시작은 0부터 시작한다.
''''			수령인addr1,수령인addr2 이 한개의 필드일때는 동일하게 지정해준다.

''''             주문번호, 주문일, 입금일, 지불수단, 주문자ID, 주문자, 주문자전화,주문자휴대전화,주문자이메일,
''''             수령인,수령인전화,수령인핸드폰,수령인Zip,수령인addr1,수령인addr2
''''             상품코드, 옵션코드, 수량, 판매가, 소비자가, 정산액, 상품명, 옵션명,
''''             업체상품코드,업체옵션코드,주문디테일키, 주문유의사항, 상품요구사항1, 상품요구사항2, 상품요구사항3
''''              옵션가격, 옵션공급가격. 범용코드, 쇼핑몰명, 상품명2, 옵션명2

''''			쇼핑몰 주문번호, 사방넷 주문번호, 쇼핑몰명, 수취인명, 전화번호1,
''''			전화번호2, 우편번호, 주소, 수량(소), 특기사항,
''''			운임구분, 운임, 상품명1, 상품명(확정), 옵션명,
''''			옵션명(확정), 옵션별칭(바코드), 판매가(수집), EA(상품)*수량

if (xltype = "sabangnet") then
	'' 사방넷
	xlPosArr = Array(0,-1,-1,-1,-1,-1,-1,-1,-1,    3,4,5,6,7,7,    -1,-1,8,17,17,-1,12,14,    -1,-1,1,9,-1,-1,-1    -1,-1, 16, 2, 13,15)
	ArrayLen = UBound(xlPosArr)
	skipString="쇼핑몰 주문번호"
	afile = sFilePath
	aSheetName = maybeSheetName
elseif (xltype = "default") then
	'' 기본포맷
	'' 주문번호	주문디테일키	쇼핑몰명	수령인	수령인전화	수령인핸드폰
	'' 수령인Zip	수령인addr1	수령인addr2	상품명	옵션명	수량
	'' 범용코드	소비자가	판매가	주문유의사항
	xlPosArr = Array(0,-1,-1,-1,-1,-1,-1,-1,-1,    3,4,5,6,7,7,    -1,-1,11,13,12,-1,9,10,    -1,-1,1,14,-1,-1,-1    -1,-1, 11, 2, -1,-1)
	ArrayLen = UBound(xlPosArr)
	skipString="주문번호"
	afile = sFilePath
	aSheetName = maybeSheetName
else
    response.write "<script>alert('등록되지 않은 포멧입니다. -"&xltype&"');</script>"
    response.end
end if


if (true) then
	'// skip
elseif (companyid = "toms") then

	if (SellSite="5") then

	    '' 탐스 - 후이즈
	    xlPosArr = Array(0,2,42,39,6,5,9,10,11,13,16,17,14,15,15,22,-1,28,29,31,-1,24,25,-1,-1,-1,12,19,20,21)
	    ArrayLen = UBound(xlPosArr)
	    skipString="주문번호"
	    afile = sFilePath
	    aSheetName = "TOMS 주문내역"  '' sheet name Maybe filename in

	elseif (SellSite="4") then

	    ''텐바이텐 - 인덱스추가형식
	    xlPosArr = Array(1,2,-1,-1,-1,3,4,5,6,7,8,9,10,11,12,15,-1,19,18,-1,-1,16,17,21,-1,0,13,20,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="일련번호"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in
	else
	    response.write "<script>alert('쇼핑몰 코드가 지정되지 않았습니다. -"&SellSite&"');</script>"
	    response.end
	end if

elseif (companyid = "ithinkso") then
    if (SellSite="6") then
        ''하나투어
        xlPosArr = Array(0,15,15,-1,-1,2,4,5,-1,3,4,5,16,16,16,1,-1,8,9,9,10,6,7,-1,-1,-1,17,-1,-1,-1,11,12)
	    ArrayLen = UBound(xlPosArr)
	    skipString="주문번호"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in
    elseif (SellSite="10") then
        ''(주)더블유컨셉코리아
        xlPosArr = Array(2,1,1,-1,-1,13,19,20,-1,14,17,18,15,16,16,5,-1,10,12,12,11,6,8,-1,-1,-1,21,-1,-1,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="주문번호"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in
    elseif (SellSite="11") then
        ''위즈위드
        xlPosArr = Array(4,2,2,-1,-1,		14,20,21,-1,15,		18,19,16,17,17,		6,-1,11,13,13,		12,7,8,-1,-1,		-1,23,-1,-1,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="주문번호"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in
    elseif (SellSite="7") then
		''''            주문번호, 주문일, 입금일, 지불수단, 주문자ID

		''''			주문자, 주문자전화,주문자휴대전화,주문자이메일,수령인

		''''			수령인전화,수령인핸드폰,수령인Zip,수령인addr1,수령인addr2

		''''            상품코드, 옵션코드, 수량, 판매가, 소비자가

		''''			정산액, 상품명, 옵션명,업체상품코드,업체옵션코드

		''''			주문디테일키, 주문유의사항, 상품요구사항1, 상품요구사항2, 상품요구사항3

		''''            옵션가격, 옵션공급가격.
    	'/신한카드올댓샵					   .             .                       .                    .
    	'xlPosArr = Array(1,23,23,-1,-1,2,3,4,5, 6,7,8,9,10,11 ,16,-1,19,20,20,-1,17,18 ,16,-1,-1,12,-1,-1,-1, -1,-1)	'2011.10.17 구 변경사항
    	xlPosArr = Array(1,23,23,-1,-1,     2,3,4,-1, 5,     6,7,8,9,10 ,     15,-1,18,19,19,     -1,16,17 ,14,-1,     -1,11,-1,-1,-1,      -1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="주문번호"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in
    elseif (SellSite="8") then

    	''민트샵
    	''1			2		3			4		5			6			7			8		9			10					11		12				13		14
    	''주문번호	주문자	받는사람	갯수	전화번호	비상전화	우편번호	주소	전달사항	아이띵소 상품코드	상품명	옵션(특징포함)	판매가	주문일
    	''xlPosArr = Array(1,14,14,-1,-1,2,5,6,-1, 3,5,6,7,8,15 ,10,-1,4,13,13,-1,11,12 ,10,-1,-1,9,-1,-1,-1, -1,-1)
    	xlPosArr = Array(0,13,13,-1,-1,1,4,5,-1, 2,4,5,6,7,7 ,9,-1,3,12,12,-1,10,11 ,9,-1,-1,8,-1,-1,-1, -1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="주문번호"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in
    elseif (SellSite="9") then
		''''            주문번호, 주문일, 입금일, 지불수단, 주문자ID

		''''			주문자, 주문자전화,주문자휴대전화,주문자이메일,수령인

		''''			수령인전화,수령인핸드폰,수령인Zip,수령인addr1,수령인addr2

		''''            상품코드, 옵션코드, 수량, 판매가, 소비자가

		''''			정산액, 상품명, 옵션명,업체상품코드,업체옵션코드

		''''			주문디테일키, 주문유의사항, 상품요구사항1, 상품요구사항2, 상품요구사항3

		''''            옵션가격, 옵션공급가격.
    	'/패션플러스					   .             .                       .                    .
    	'xlPosArr = Array(1,23,23,-1,-1,2,3,4,5, 6,7,8,9,10,11 ,16,-1,19,20,20,-1,17,18 ,16,-1,-1,12,-1,-1,-1, -1,-1)	'2011.10.17 구 변경사항
    	xlPosArr = Array(1,5,5,-1,-1,     27,26,25,-1, 2,     26,25,24,22,22 ,     10,-1,12,13,13,     -1,11,4 ,10,-1,     -1,28,-1,-1,-1,      -1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="주문번호"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in
    else
	    response.write "<script>alert('쇼핑몰 코드가 지정되지 않았습니다. -"&SellSite&"');</script>"
	    response.end
	end if
else
    response.write "<script>alert('등록되지 않은 업체입니다. -"&companyid&"');</script>"
    response.end
end if

ReDim xlRow(ArrayLen)
rw "ArrayLen="&ArrayLen

dim ret : ret = fnGetXLFileArray(xlRowALL, afile, aSheetName, ArrayLen)

if (Not ret) or (Not IsArray(xlRowALL)) then
    response.write "<script>alert('파일이 올바르지 않거나 내용이 없습니다. "&Replace(Err.Description,"'","")&"');</script>"

    if (Err.Description="외부 테이블 형식이 잘못되었습니다.") then
        response.write "<script>alert('엑셀에서 Save As Excel 97 -2003 통합문서 형태로 저장후 사용하세요.');</script>"
    end if
    response.write "<script>history.back();</script>"
    response.end
end if

response.write "OK"
response.end

''데이터 처리.
okCNT = 0

dbget.BeginTrans
    for i=0 to UBound(xlRowALL)

    if IsObject(xlRowALL(i)) then
        set iLine = xlRowALL(i)
        ''옵션가격 있는경우.

		if IsNumeric(iLine.FItemArray(1)) then
			'// 20120101 or 120101
			if (Len(CStr(iLine.FItemArray(1))) = 8) then
				iLine.FItemArray(1) = Left(iLine.FItemArray(1), 4) & "-" & Mid(iLine.FItemArray(1), 5, 2) & "-" & Right(iLine.FItemArray(1), 2)
			elseif (Len(CStr(iLine.FItemArray(1))) = 6) then
				iLine.FItemArray(1) = Left(iLine.FItemArray(1), 2) & "-" & Mid(iLine.FItemArray(1), 3, 2) & "-" & Right(iLine.FItemArray(1), 2)
			end if
		end if

		if IsNumeric(iLine.FItemArray(2)) then
			'// 20120101 or 120101
			if (Len(CStr(iLine.FItemArray(2))) = 8) then
				iLine.FItemArray(2) = Left(iLine.FItemArray(2), 4) & "-" & Mid(iLine.FItemArray(2), 5, 2) & "-" & Right(iLine.FItemArray(2), 2)
			elseif (Len(CStr(iLine.FItemArray(2))) = 6) then
				iLine.FItemArray(2) = Left(iLine.FItemArray(2), 2) & "-" & Mid(iLine.FItemArray(2), 3, 2) & "-" & Right(iLine.FItemArray(2), 2)
			end if
		end if

        iLine.FItemArray(17) = Replace(iLine.FItemArray(17),",","")
        iLine.FItemArray(18) = Replace(iLine.FItemArray(18),",","")
        iLine.FItemArray(19) = Replace(iLine.FItemArray(19),",","")

        IF (UBound(iLine.FItemArray)>30) and (iLine.FItemArray(18)<>"") then
            IF (iLine.FItemArray(30)="") then iLine.FItemArray(30)="0"
            IF (iLine.FItemArray(31)="") then iLine.FItemArray(31)="0"

            iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)) + CLNG(iLine.FItemArray(30))
        END IF

        iLine.FItemArray(3) = convPayTypeStr2Code(iLine.FItemArray(3))
        IF (iLine.FItemArray(3)="") then iLine.FItemArray(3)="50"                           ''PayType
        IF (iLine.FItemArray(2)="") then iLine.FItemArray(2)=iLine.FItemArray(1)            ''Paydate
        IF (iLine.FItemArray(19)="") then iLine.FItemArray(19)=iLine.FItemArray(18)         ''RealSellPrice

        '''우편번호와 주소1이 같은경우 [ ]로 우편번호 추출 = 하나투어방식.
        iLine.FItemArray(12) = TRIM(Replace(iLine.FItemArray(12),"  "," "))
        iLine.FItemArray(13) = TRIM(Replace(iLine.FItemArray(13),"  "," "))
        iLine.FItemArray(14) = TRIM(Replace(iLine.FItemArray(14),"  "," "))
        IF (iLine.FItemArray(12)=iLine.FItemArray(13)) then
            POS1 = InStr(iLine.FItemArray(12),"[")
            POS2 = InStr(iLine.FItemArray(12),"]")
            IF (POS1>0) and (POS2>0) then
                iLine.FItemArray(12) = Mid(iLine.FItemArray(12),POS1+1,POS2-POS1-1)
                iLine.FItemArray(12) = Trim (iLine.FItemArray(12))

                IF (iLine.FItemArray(13)=iLine.FItemArray(14)) THEN
                    iLine.FItemArray(13) = TRIM(Mid(iLine.FItemArray(13),POS2+1,512))
                    iLine.FItemArray(14) = iLine.FItemArray(13)
                ELSE
                    iLine.FItemArray(13) = TRIM(Mid(iLine.FItemArray(13),POS2+1,512))
                END IF
            END IF
        END IF

        '''주소와 상세주소가 같은경우 3번째 Blank에서 끊음.
        POS1 = 0
        POS2 = 0
        POS3 = 0

        IF (iLine.FItemArray(13)=iLine.FItemArray(14)) then
            POS1 = InStr(iLine.FItemArray(14)," ")
            ''rw "POS1="&POS1
            IF (POS1>0) then
                POS2 = InStr(MID(iLine.FItemArray(14),POS1+1,512)," ")
                ''rw "POS2="&POS2
                IF POS2>0 then
                    POS3 = InStr(MID(iLine.FItemArray(14),POS1+POS2+1,512)," ")
                    IF POS3>0 then
                        iLine.FItemArray(13)=LEFT(iLine.FItemArray(14),POS1+POS2+POS3-1)
                        iLine.FItemArray(14)=MID(iLine.FItemArray(14),POS1+POS2+POS3+1,512)

                        'rw iLine.FItemArray(13)
                        'rw iLine.FItemArray(14)
                    END IF
                END IF
            END IF
        END IF

		'주문번호-순번(17709513-1)
		''if (SellSite = 9) then
		''	if (Ubound(Split(iLine.FItemArray(0), "-")) > 0) then
		''		iLine.FItemArray(0) = Split(iLine.FItemArray(0), "-")(0)
		''	end if
		''end if

		if False then
			rw "@tplcompanyid="&tplcompanyid
			rw "@sellsitename="&iLine.FItemArray(32)
			rw "@OutMallOrderSerial="&iLine.FItemArray(0)
			rw "@SellDate="&iLine.FItemArray(1)
			rw "@PayType="&iLine.FItemArray(3)
			rw "@Paydate="&iLine.FItemArray(2)
			rw "@partnerItemID="&iLine.FItemArray(15)
			rw "@partnerItemName="&iLine.FItemArray(21)
			rw "@partnerOption="&iLine.FItemArray(16)
			rw "@partnerOptionName="&iLine.FItemArray(22)
			rw "@OrderUserID="&iLine.FItemArray(4)
			rw "@OrderName="&iLine.FItemArray(5)
			rw "@OrderEmail="&iLine.FItemArray(8)
			rw "@OrderTelNo="&iLine.FItemArray(6)
			rw "@OrderHpNo="&iLine.FItemArray(7)
			rw "@ReceiveName="&iLine.FItemArray(9)
			rw "@ReceiveTelNo="&iLine.FItemArray(10)
			rw "@ReceiveHpNo="&iLine.FItemArray(11)
			rw "@ReceiveZipCode="&iLine.FItemArray(12)
			rw "@ReceiveAddr1="&iLine.FItemArray(13)
			rw "@ReceiveAddr2="&iLine.FItemArray(14)
			rw "@SellPrice="&iLine.FItemArray(18)
			rw "@RealSellPrice="&iLine.FItemArray(19)
			rw "@ItemOrderCount="&iLine.FItemArray(17)
			rw "@OrgDetailKey="&iLine.FItemArray(25)
			rw "@deliverymemo="&iLine.FItemArray(26)
			rw "@requireDetail="&iLine.FItemArray(27)
			rw "barcode=" & iLine.FItemArray(31)
		end if

		IF (iLine.FItemArray(0)<>"") then
			partnerItemName = iLine.FItemArray(21)
			partnerOptionName = iLine.FItemArray(22)
			barcode = iLine.FItemArray(31)
			''partnerItemName = iLine.FItemArray(33)
			''partnerOptionName = iLine.FItemArray(34)


			itemgubun = ""
			itemid = ""
			itemoption = ""
			itemname = ""
			itemoptionname = ""

			if barcode <> "" then
				'// 범용코드 있는 경우
				'// 1. 범용코드로 상품코드 찾고
				sqlStr = " select top 1 "
				sqlStr = sqlStr & " 	b.itemgubun, b.itemid, b.itemoption "
				sqlStr = sqlStr & " 	, IsNull(i.itemname, si.shopitemname) as itemname, IsNull(IsNull(o.optionname, si.shopitemoptionname), '') as itemoptionname "
				''sqlStr = sqlStr & " 	, IsNull(i.orgprice, si.orgsellprice) as orgprice "
				''sqlStr = sqlStr & " 	, IsNull(i.mainimage, si.offimgmain) as mainimage, IsNull(i.listimage, si.offimglist) as listimage, IsNull(i.smallimage, si.offimgsmall) as smallimage "
				sqlStr = sqlStr & " from "
				sqlStr = sqlStr & " 	[db_item].[dbo].[tbl_item_option_stock] b "
				sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_item] i "
				sqlStr = sqlStr & " 	on "
				sqlStr = sqlStr & " 		1 = 1 "
				sqlStr = sqlStr & " 		and b.itemgubun = '10' "
				sqlStr = sqlStr & " 		and b.itemid = i.itemid "
				sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_item_option] o "
				sqlStr = sqlStr & " 	on "
				sqlStr = sqlStr & " 		i.itemid = o.itemid and b.itemoption = o.itemoption "
				sqlStr = sqlStr & " 	left join [db_shop].[dbo].[tbl_shop_item] si "
				sqlStr = sqlStr & " 	on "
				sqlStr = sqlStr & " 		1 = 1 "
				sqlStr = sqlStr & " 		and b.itemgubun <> '10' "
				sqlStr = sqlStr & " 		and b.itemgubun = si.itemgubun "
				sqlStr = sqlStr & " 		and b.itemid = si.shopitemid "
				sqlStr = sqlStr & " 		and b.itemoption = si.itemoption "
				sqlStr = sqlStr & " 	left join [db_partner].[dbo].[tbl_partner] ip on i.makerid = ip.id "
				sqlStr = sqlStr & " 	left join [db_partner].[dbo].[tbl_partner] sp on si.makerid = sp.id "
				sqlStr = sqlStr & " where b.barcode = '" & barcode & "' and IsNull(ip.id, sp.id) = '" & makerid & "' "
				''rw sqlStr
				rsTENget.CursorLocation = adUseClient
				rsTENget.Open sqlStr,dbTENget,adOpenForwardOnly, adLockReadOnly
				if  not rsTENget.EOF  then
					itemgubun         = rsTENget("itemgubun")
					itemid            = rsTENget("itemid")
					itemoption        = rsTENget("itemoption")
					itemname          = rsTENget("itemname")
					itemoptionname    = rsTENget("itemoptionname")
					''orgprice          = rsTENget("orgprice")
					''mainimage         = rsTENget("mainimage")
					''listimage         = rsTENget("listimage")
					''smallimage        = rsTENget("smallimage")
				end if
				rsTENget.Close
			end if

			if (itemgubun <> "") then
				''rw "Found : " & barcode
			else
				''rw "Not Found" & partnerItemName
			end if

			'// ================================================================
			'// 주의 : 갯수도 순서도 프로시져와 동일해야 한다.!!!
			'// ================================================================
            paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
                ,Array("@tplcompanyid" , adVarchar	, adParamInput,32, tplcompanyid)	_
                ,Array("@SellSite" , adInteger	, adParamInput,, "")	_
				,Array("@SellSiteName" , adVarchar	, adParamInput,32, iLine.FItemArray(32))	_
				,Array("@OutMallOrderSerial"	, adVarchar	, adParamInput,23, Left(CStr(iLine.FItemArray(0)), 22))	_
				,Array("@OrgDetailKey"	,adVarchar, adParamInput,32, iLine.FItemArray(25)) _
    			,Array("@SellDate"	,adDate, adParamInput,, iLine.FItemArray(1)) _
    			,Array("@PayType"	,adVarchar, adParamInput,32, iLine.FItemArray(3)) _
    			,Array("@Paydate"	,adDate, adParamInput,, iLine.FItemArray(2)) _
				,Array("@makerid" , adVarchar	, adParamInput,32, makerid)	_
				,Array("@itemgubun" , adVarchar	, adParamInput,2, itemgubun)	_
				,Array("@itemid" , adInteger	, adParamInput,, itemid)	_
				,Array("@itemoption" , adVarchar	, adParamInput,4, itemoption)	_
				,Array("@itemname" , adVarchar	, adParamInput,128, itemname)	_
				,Array("@itemoptionname" , adVarchar	, adParamInput,128, itemoptionname)	_
    			,Array("@orderItemID"	,adVarchar, adParamInput,32, iLine.FItemArray(15)) _
    			,Array("@orderItemName"	,adVarchar, adParamInput,128, iLine.FItemArray(21)) _
    			,Array("@orderItemOption"	,adVarchar, adParamInput,128, iLine.FItemArray(16)) _
    			,Array("@orderItemOptionName"	,adVarchar, adParamInput,128, iLine.FItemArray(22)) _
    			,Array("@barcode"	,adVarchar, adParamInput,32, barcode) _
    			,Array("@OrderName"	,adVarchar, adParamInput,32, iLine.FItemArray(5)) _
    			,Array("@OrderEmail"	,adVarchar, adParamInput,100, iLine.FItemArray(8)) _
    			,Array("@OrderTelNo"	,adVarchar, adParamInput,16, iLine.FItemArray(6)) _
    			,Array("@OrderHpNo"	,adVarchar, adParamInput,16, iLine.FItemArray(7)) _
    			,Array("@ReceiveName"	,adVarchar, adParamInput,32, iLine.FItemArray(9)) _
    			,Array("@ReceiveTelNo"	,adVarchar, adParamInput,16, iLine.FItemArray(10)) _
    			,Array("@ReceiveHpNo"	,adVarchar, adParamInput,16, iLine.FItemArray(11)) _
    			,Array("@ReceiveZipCode"	,adVarchar, adParamInput,7, iLine.FItemArray(12)) _
    			,Array("@ReceiveAddr1"	,adVarchar, adParamInput,128, iLine.FItemArray(13)) _
    			,Array("@ReceiveAddr2"	,adVarchar, adParamInput,512, iLine.FItemArray(14)) _
    			,Array("@SellPrice"	,adCurrency, adParamInput,, iLine.FItemArray(18)) _
    			,Array("@RealSellPrice"	,adCurrency, adParamInput,, iLine.FItemArray(19)) _
				,Array("@vatinclude"	,adVarchar, adParamInput,1, "Y") _
    			,Array("@ItemOrderCount"	,adInteger, adParamInput,, iLine.FItemArray(17)) _
    			,Array("@DeliveryType"	,adInteger, adParamInput,, 0) _
    			,Array("@deliveryprice"	,adCurrency, adParamInput,, 0) _
    			,Array("@deliverymemo"	,adVarchar, adParamInput,400, iLine.FItemArray(26)) _
				,Array("@countryCode"	,adVarchar, adParamInput,2, "KR") _
    			,Array("@requireDetail"	,adVarchar, adParamInput,400, iLine.FItemArray(27)) _
    			,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
    			)

''''             0주문번호, 주문일, 입금일, 지불수단, 주문자ID, 주문자, 주문자전화,주문자휴대전화,주문자이메일,
''''             9수령인,수령인전화,수령인핸드폰,수령인Zip,수령인addr1,수령인addr2
''''             15상품코드, 옵션코드, 수량, 판매가, 소비자가, 정산액, 상품명, 옵션명,
''''             23업체상품코드,업체옵션코드,주문디테일키,주문유의사항, 상품요구사항1, 상품요구사항2, 상품요구사항3

            sqlStr = "db_threepl.dbo.usp_OnlineTmpOrder_Insert"
            retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

            RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
            retErrStr  = GetValue(retParamInfo, "@retErrStr") ' 생성된 송장번호

            okCNT = okCNT +1

        END IF
            set iLine = Nothing

            IF (retErr)<>0 then
                dbget.rollbackTrans
                response.write "ERROR["&retErr&"]"& retErrStr
                response.write "<script>alert('"&Replace("ERROR["&retErr&"]"& retErrStr,"'","")&"');</script>"
                response.write "<script>history.back();</script>"
                response.end
            END IF
        end if
    Next
dbget.CommitTrans

response.write "<script>alert('"&okCNT&"건 입력되었습니다.')</script>"
response.write "<script>opener.location.reload();self.close();</script>"
'''====================================================================================

Class TXLRowObj
    public FItemArray

    public function setArrayLength(ln)
        Redim FItemArray(ln)
    end function

End Class

function convPayTypeStr2Code(oStr)
    SELECT CASE oStr
        CASE "신용카드" : convPayTypeStr2Code="100"
        CASE "신용" : convPayTypeStr2Code="100"
        CASE "무통장" : convPayTypeStr2Code="7"
        CASE "실시간이체" : convPayTypeStr2Code="20"
        CASE "핸드폰결제" : convPayTypeStr2Code="400"
        CASE "휴대폰결제" : convPayTypeStr2Code="400"
        CASE "핸드폰" : convPayTypeStr2Code="400"
        CASE "휴대폰" : convPayTypeStr2Code="400"
        CASE ELSE : convPayTypeStr2Code="50"
    END SELECT
end function

function IsSKipRow(ixlRow, skipCol0Str)

    if Not IsArray(ixlRow) then
        IsSKipRow = true
        Exit function
    end if

    if  LCASE(ixlRow(0))=LCASE(skipCol0Str) then
        IsSKipRow = true
        Exit function
    end if

    IsSKipRow = false
end function

Function fnGetXLFileArray(byref xlRowALL, sFilePath, aSheetName, iArrayLen)
    Dim conDB, Rs, strQry, iResult, i, J, iObj
    Dim irowObj, strTable
    '' on Error 구문 쓰면 안됨.. 서버 무한루프 도는듯.

    Set conDB = Server.CreateObject("ADODB.Connection")
    conDB.Provider = "Microsoft.Jet.oledb.4.0"
    conDB.Properties("ExtEnded Properties").Value = "Excel 8.0;"

    On Error Resume Next
        conDB.Open sFilePath

        IF (ERR) then
            fnGetXLFileArray=false
            exit function
        End if
    On Error Goto 0

    '' get First Sheet Name=============''시트가 여러개인경우 오류날 수 있음.
    Set Rs = conDB.OpenSchema(adSchemaTables)

    IF Not Rs.Eof Then
        aSheetName = Rs.Fields("table_name").Value
        ''rw "aSheetName="&aSheetName
    ENd IF
    Set Rs = Nothing
    ''==================================

    Set Rs = Server.CreateObject("ADODB.Recordset")

    ''strQry = "Select * From [sheet1$]"
    strQry = "Select * From ["&aSheetName&"]"

    ReDim xlRowALL(0)
    fnGetXLFileArray = true

On Error Resume Next
    Rs.Open strQry, conDB
    	IF (ERR) then
            fnGetXLFileArray=false
            Rs.Close
            Set Rs = Nothing
            Set conDB = Nothing
            exit function
        End if

        If Not Rs.Eof Then
            Do Until Rs.Eof
                IF (ERR) then
                    fnGetXLFileArray=false
                    Rs.Close
                    Set Rs = Nothing
                    Set conDB = Nothing
                    exit function
                End if

                set irowObj = new TXLRowObj
                irowObj.setArrayLength(iArrayLen)

                For i=0 to ArrayLen
                    if (xlPosArr(i)<0) then
                        irowObj.FItemArray(i) = ""
                    else
                        irowObj.FItemArray(i) = Replace(null2blank(Rs(xlPosArr(i))),"*","")
                    end if
                Next

                IF (Not IsSKipRow(irowObj.FItemArray,skipString)) then
                    ReDim Preserve xlRowALL(UBound(xlRowALL)+1)

                    set xlRowALL(UBound(xlRowALL)) =  irowObj
                    ''xlRowALL(UBound(xlRowALL)).arrayObj = xlRow

                END IF
                set irowObj = Nothing
                Rs.MoveNext
            Loop
       else
          fnGetXLFileArray=false
       end if

       ''''On Error Goto 0
        IF (ERR) then
            fnGetXLFileArray=false
        End if
    Rs.Close
On Error Goto 0

    Set Rs = Nothing
    Set conDB = Nothing

    if Ubound(xlRowALL)< 1 then fnGetXLFileArray=false

End Function

Function AddTmpDbOrderData(ixlRowALL)
    AddTmpDbOrderData = false

'    Rs.Open strQry, conDB
'        If Not Rs.Eof Or Rs.Bof Then
'            Do Until Rs.Eof
'            		iResult =""
'            		'Add1 =""
'            		'Add2=""
'            		outmallorderserial 	= replace(Rs(1),"*","")	'주문번호
'            		partnerItemId 		= Rs(3)	'상품코드
'            		partnerItemName 	= ReplaceRequestSpecialChar(Rs(4))	'상품명
'            		partnerOptionName	= replace(replace(Rs(5),"-","")," ","")	'옵션
'            		SellPrice			= Rs(6)	'주문금액
'            		ItemOrderCount 	= Rs(7)	'주문수량
'            		OrderName		= Rs(9)	'이름
'            		ZipCode			= left(Rs(11),3)&right(Rs(11),3)	'우편번호
'            		arrAdd			= split(Rs(12)," ")	'수령지주소
'            		for i = 0 to 2
'            		Add1			= Add1 &" "& arrAdd(i)  	'수령지주소1
'            		next
'            		for i = 3 to ubound(arrAdd)
'            		Add2			= Add2 &" "& arrAdd(i)  	'수령지주소1
'            		next
'            		IF Add2 = "" THEN Add2 = "."
'            		ReceiveName		= Rs(13)	'수령자
'            		ReceiveTelNo		= Rs(14)	'전화번호
'            		ReceiveHpNo		= Rs(15)	'핸드폰
'            		EtcAsk			= ReplaceRequestSpecialChar(Rs(16))	'주문요청사항
'            		PayDate			= Rs(18)	'결제(입금)일자
'            		RealSellPrice		= Rs(21)	'판매금액
'
'            		SellSite			="3"
'            		PartnerSeq		="58"
'            		OrderEmail 		=""
'
''iResult =  clsConnDB.fnMultiExecSPReturnValue("db_agirlOrder.dbo.[usp_Back_OutMallOrder_Insert]("&SellSite&","&PartnerSeq &" ,'"&OutMallOrderSerial&"','"&partnerItemID&"','"&partnerItemName&"','','"&partnerOptionName&"','"&OrderName&"'"&_
''    				",'"&OrderEmail&"','"&ReceiveTelNo&"','"&ReceiveHpNo&"','"&ReceiveName&"','"&ReceiveTelNo&"','"&ReceiveHpNo&"','"&ZipCode&"','"&Add1&"','"&Add2&"','"&EtcAsk&"','"&SellPrice&"','"&RealSellPrice&"','"&PayDate&"',1,'"&ItemOrderCount&"')")
'
'    		IF iResult  = 0 THEN
%>

<%
''clsConnDB.RollbackTrans
''Set clsConnDB = nothing
'			Set Rs = Nothing
'			Set conDB = Nothing
'		ELSEIF iResult = 2 THEN
'			if  errMsg1 ="" then
'			errMsg1 = OutMallOrderSerial
'			else
'			errMsg1 = errMsg1 &","&OutMallOrderSerial
'			end if
'		ELSEIF iResult = 3 THEN
'			if  errMsg2 ="" then
'			errMsg2 = partnerItemID
'			else
'			errMsg2 = errMsg2 &","&partnerItemID
'			end if
'    		END IF
'            Rs.MoveNext
'            Loop
'         end if
            'clsConnDB.CommitTrans
            'Set clsConnDB = nothing
'    Set Rs = Nothing
'    Set conDB = Nothing
end Function
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
