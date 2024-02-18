<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 판매 등록 관리
' Hieditor : 2011.04.22 이상구 생성
'			 2012.08.24 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/orderInput/incUploadFunction.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%
Dim uploadform, objfile, sDefaultPath, sFolderPath ,orderCsGbn ,monthFolder
Dim iML, sFile, sFilePath, SellSite, iMaxLen, sUploadPath, orgFileName, maybeSheetName
Dim overseasPrice, overseasDeliveryPrice, overseasRealPrice, reserve01, beasongNum11st, outmalloptionno
Dim outMallGoodsNo,shoplinkermallname,shoplinkerPrdCode,shoplinkerOrderID,shoplinkerMallID
Dim tmpOverseasRealprice, sItemname
Dim tmpStr1, tmpStr2, loops
dim tmpVal, tmpItem
Dim isValid

monthFolder = Replace(Left(CStr(now()),7),"-","")

IF (application("Svr_Info")	= "Dev") then
    'Set uploadform = Server.CreateObject("TABS.Upload")	   '' - TEST : TABS.Upload
	'2019-10-11 15:05 김진영 TABSUpload4.Upload로 수정
	Set uploadform = Server.CreateObject("TABSUpload4.Upload")	''REAL : TABSUpload4.Upload
ELSE
    Set uploadform = Server.CreateObject("TABSUpload4.Upload")	''REAL : TABSUpload4.Upload
END IF

Set objfile	   = Server.CreateObject("Scripting.FileSystemObject")
''sDefaultPath   = Server.MapPath("\admin\etc\orderInput\upFiles\")
sDefaultPath   = Server.MapPath("/admin/etc/orderInput/upFiles/")
uploadform.Start sDefaultPath '업로드경로

iMaxLen 		= uploadform.Form("iML")	'이미지파일크기
SellSite 	= uploadform.Form("sellsite")

IF (fnChkFile(uploadform("sFile"), iMaxLen,"xls")) THEN	'파일체크

    '폴더 생성
    sFolderPath = sDefaultPath&"/"&sellsite&"/"
    IF NOT  objfile.FolderExists(sFolderPath) THEN
    	objfile.CreateFolder sFolderPath
    END IF

    sFolderPath = sDefaultPath&"/"&sellsite&"/"&monthFolder&"/"
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

Dim xlPosArr, ArrayLen, skipString, afile, aSheetName ,i,j, k, m
''''				0				1				2				3				4
''''				주문번호, 		주문일, 		입금일, 		지불수단, 		주문자ID,
''''				5				6				7				8				9
''''				주문자, 		주문자전화,		주문자휴대전화,	주문자이메일, 	수령인,
''''				10				11				12				13				14
''''				수령인전화,		수령인핸드폰,	수령인Zip,		수령인addr1,	수령인addr2,
''''				15				16				17				18				19
''''				상품코드, 		옵션코드, 		수량, 			판매가, 		실판매가 XX소비자가,
''''				20				21				22				23				24
''''				정산액, 		상품명, 		옵션명, 		업체상품코드,	업체옵션코드,
''''				25				26				27				28				29
''''				주문디테일키, 	주문유의사항, 	상품요구사항1, 	배송비, 		ETC1컴,
''''				30				31				32				33				34
''''				ETC2(샵링커쇼핑몰),		ETC3(샵링커상품코드),			ETC4(샵링커주문번호),			국가코드,			ETC5(샵링커착불여부)
''''				35				36				37				38				39
''''				해외판매가,		해외배송비		해외실판매가	사이트별추가정보(해외)	미사용

'' 18 판매가 = 할인가
'' 19 실판매가  (쿠폰 반영 금액) 있는경우 꼭 넣어주실것. 2013/10/24

    ''샵링커
    ''2013-11-27 10:45 김진영 수정.. 11,-1,22,27,31 => 11,-1,22,27,27로 업무협조18446번 관련사항
    if (SellSite="shoplinker") then
'        xlPosArr = Array(3,2,2,-1,-1,   33,34,35,-1,37,   38,39,40,41,41,   11,-1,22,27,31,   26,15,20,10,-1,   -1,42,-1,13,-1,	  6,9,47,-1,12)
'        xlPosArr = Array(3,2,2,-1,-1,   33,34,35,-1,37,   38,39,40,41,41,   11,-1,22,27,27,   26,15,20,10,-1,   -1,42,-1,13,-1,	  6,9,47,7,12)		'2015-07-03까지의 번호
        xlPosArr = Array(3,2,2,-1,-1,   36,37,38,-1,40,   41,42,43,44,44,   11,-1,25,30,30,   29,15,20,10,-1,   -1,45,-1,13,-1,	  6,9,50,7,12)
        ArrayLen = UBound(xlPosArr)
	    skipString = "No."
	    afile = sFilePath
	    aSheetName = ""

	'/hmall
    elseif (SellSite="hmall1010") then
		xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "Sheet1" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

	'/디앤샵 - 무료배송
    elseif (SellSite="dnshop") then
'	    xlPosArr = Array(2,41,41,-1,-1,  18,20,21,-1,19,   20,21,22,23,24,   37,-1,9,27,27,   -1,5,6,4,-1,   3,15,17,-1,-1, -1)
	    xlPosArr = Array(2,38,38,-1,-1,  15,17,18,-1,16,   17,18,19,20,21,   34,-1,9,24,24,   -1,5,6,4,-1,   3,13,14,-1,-1, -1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

	'/해외중국사이트
	elseif (SellSite="cn10x10") then
	    'xlPosArr = Array(0,1,3,-1,-1,   37,38,38,39,37,   38,38,40,41,41,   9,7,15,16,16,   20,-1,-1,9,7,   -1,42,-1,28,-1,	  -1,-1,-1,36,-1)
	    xlPosArr = Array(0,1,3,-1,-1,   37,38,38,39,37,   38,38,40,41,41,   9,7,15,16,16,   20,-1,-1,9,7,   -1,-1,-1,28,-1,	  -1,-1,-1,36,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Worksheet"  '' sheet name Maybe filename in

	'/해외중국사이트
	elseif (SellSite="cnglob10x10") then
		'xlPosArr = Array(0,13,49,-1,-1,   14,19,19,24,22,   29,30,31,33,33,   92,-1,7,11,12,   -1,68,6,1,91,   -1,40,21,61,12,	  89,-1,-1,35,-1)
		xlPosArr = Array(0,14,50,-1,-1,   15,20,20,25,23,   30,31,32,34,34,   93,-1,7,11,12,   -1,69,6,1,92,   -1,41,22,62,13,	  90,-1,-1,36,-1,	44,62,12,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "combine_data"  '' sheet name Maybe filename in

	elseif (SellSite="cnhigo") then
		'xlPosArr = Array(0,14,50,-1,-1,   15,20,20,25,23,   30,31,32,34,34,   93,-1,7,11,12,   -1,69,6,1,92,   -1,41,22,62,13,	  90,-1,-1,36,-1)
		xlPosArr = Array(0,20,21,-1,-1,   11,12,12,-1,13,   14,15,19,17,17,   6,8,5,1,2,   -1,7,9,-1,-1,   -1,-1,-1,3,4,	  -1,-1,-1,18,-1,	1,3,2,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in

	elseif (SellSite="celectory") then
	    'xlPosArr = Array(5,32,-1,-1,-1,   21,24,25,-1,21,    24,25,22,23,23,   -1,-1,14,16,16,   -1,12,13,-1,-1,	 6,26,-1,18,-1)
	    xlPosArr = Array(3,2,-1,-1,-1,   11,12,12,-1,13,    14,14,15,16,16,   5,-1,10,9,9,   -1,7,8,4,-1,	 -1,17,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in


	elseif (SellSite="cnugoshop") then
		'xlPosArr = Array(0,14,50,-1,-1,   15,20,20,25,23,   30,31,32,34,34,   93,-1,7,11,12,   -1,69,6,1,92,   -1,41,22,62,13,	  90,-1,-1,36,-1)
		xlPosArr = Array(1,21,22,-1,-1,   12,13,13,-1,14,   15,16,20,18,18,   7,9,6,2,3,   -1,8,10,-1,-1,   -1,-1,-1,4,5,	  -1,-1,-1,19,-1,	2,4,3,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in

	'/11번가 말레이시아
	elseif (SellSite="11stmy") then
		'xlPosArr = Array(2,3,5,-1,-1,   26,13,13,-1,12,   13,13,15,14,14,   8,-1,10,20,46,   -1,9,11,7,-1,   4,17,-1,23,18,	  -1,-1,-1,16,-1,	20,23,20,6,-1)
		xlPosArr = Array(0,2,3,-1,-1,   15,16,16,-1,15,   16,16,17,19,19,   9,-1,6,25,32,   -1,4,5,8,-1,   1,20,-1,24,22,	  -1,-1,-1,18,-1,	25,24,25,13,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in
''''				0				1				2				3				4
''''				주문번호, 		주문일, 		입금일, 		지불수단, 		주문자ID,
''''				5				6				7				8				9
''''				주문자, 		주문자전화,		주문자휴대전화,	주문자이메일, 	수령인,
''''				10				11				12				13				14
''''				수령인전화,		수령인핸드폰,	수령인Zip,		수령인addr1,	수령인addr2,
''''				15				16				17				18				19
''''				상품코드, 		옵션코드, 		수량, 			판매가, 		실판매가 XX소비자가,
''''				20				21				22				23				24
''''				정산액, 		상품명, 		옵션명, 		업체상품코드,	업체옵션코드,
''''				25				26				27				28				29
''''				주문디테일키, 	주문유의사항, 	상품요구사항1, 	배송비, 		ETC1컴,
''''				30				31				32				33				34
''''				ETC2(샵링커쇼핑몰),		ETC3(샵링커상품코드),			ETC4(샵링커주문번호),			국가코드,			ETC5(샵링커착불여부)
''''				35				36				37				38				39
''''				해외판매가,		해외배송비		해외실판매가	사이트별추가정보(해외)	미사용

'' 18 판매가 = 할인가
'' 19 실판매가  (쿠폰 반영 금액) 있는경우 꼭 넣어주실것. 2013/10/24

	elseif (SellSite="zilingo") then
		'xlPosArr = Array(0,14,50,-1,-1,   15,20,20,25,23,   30,31,32,34,34,   93,-1,7,11,12,   -1,69,6,1,92,   -1,41,22,62,13,	  90,-1,-1,36,-1)
		xlPosArr = Array(1,3,3,-1,-1,   17,21,21,-1,17,   21,21,20,19,19,   -1,-1,8,11,11,   -1,7,-1,0,-1,   -1,-1,-1,-1,-1,	  -1,-1,-1,18,-1,	9,-1,9,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "Sheet0" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in


	'//해외 etsy
	elseif (SellSite="etsy") then
		'xlPosArr = Array(2,3,5,-1,-1,   26,13,13,-1,12,   13,13,15,14,14,   8,-1,10,20,46,   -1,9,11,7,-1,   4,17,-1,23,18,	  -1,-1,-1,16,-1,	20,23,20,6,-1)
		xlPosArr = Array(26,16,16,-1,-1,   2,-1,-1,-1,18,   -1,-1,23,19,21,   -1,-1,3,4,7,   -1,1,-1,15,-1,   14,-1,-1,9,12,	  -1,-1,-1,25,-1,	4,9,4,22,24)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

    '/롯데닷컴 - 배송비 조건 30,000 이상 무료배송 // 구매자 전화 핸드폰 없음.. // 37 주문 반품 구분(주문,교환주문)
    elseif (SellSite="lotteCom") then
        xlPosArr = Array(2,0,-1,-1,-1,   15,-1,-1,-1,10,    13,14,11,12,12,   -1,-1,32,30,30,   29,5,27,4,-1    ,3,17,26,-1,42    ,-1,-1,37) ''29Col
                   ''Array(0,-1,-1,-1,-1,26,28,29,-1,30,31,32,33,34,34,42,43,5,9,-1,11,3,4,42,43,-1,35,38,20)
	    ArrayLen = UBound(xlPosArr)
	    skipString="출고지시일"
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in
''''				0				1				2				3				4
''''				주문번호, 		주문일, 		입금일, 		지불수단, 		주문자ID,
''''				5				6				7				8				9
''''				주문자, 		주문자전화,		주문자휴대전화,	주문자이메일, 	수령인,
''''				10				11				12				13				14
''''				수령인전화,		수령인핸드폰,	수령인Zip,		수령인addr1,	수령인addr2,
''''				15				16				17				18				19
''''				상품코드, 		옵션코드, 		수량, 			판매가, 		실판매가 XX소비자가,
''''				20				21				22				23				24
''''				정산액, 		상품명, 		옵션명, 		업체상품코드,	업체옵션코드,
''''				25				26				27				28				29
''''				주문디테일키, 	주문유의사항, 	상품요구사항1, 	배송비, 		ETC1컴,
''''				30				31				32				33				34
''''				ETC2(샵링커쇼핑몰),		ETC3(샵링커상품코드),			ETC4(샵링커주문번호),			국가코드,			ETC5(샵링커착불여부)
''''				35				36				37				38				39
''''				해외판매가,		해외배송비		해외실판매가	사이트별추가정보(해외)	미사용

    elseif (SellSite="lotteon") then
        xlPosArr = Array(1,0,-1,-1,-1,   10,9,9,-1,10,    11,11,13,12,12,   -1,-1,45,44,44,   -1,33,34,35,37    ,17,52,40,-1,-1    ,-1,-1,-1) ''29Col
	    ArrayLen = UBound(xlPosArr)
	    skipString="출고지시일"
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in


    '/롯데아이몰
    elseif (SellSite="lotteimall") then
       'xlPosArr = Array(5,0,-1,-1,-1,  38,39,40,-1,24,   30,31,25,26,26,   14,-1,18,17,17,   -1,10,12,8,-1,	 -1,36,-1,-1,-1)
       'xlPosArr = Array(6,0,-1,-1,-1,  40,41,42,-1,26,   32,33,27,28,28,   16,-1,20,19,19,   -1,12,14,10,-1,	 -1,38,-1,-1,-1)
	   xlPosArr = Array(6,0,-1,-1,-1,  41,42,43,-1,25,   31,32,26,27,27,   15,-1,19,18,18,   -1,11,13,9,-1,	 -1,39,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="주문번호"
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

	'/11번가_아이띵소
	elseif (SellSite="11stITS") then
	    xlPosArr = Array(2,46,4,-1,-1,  33,29,28,-1,12,   29,28,30,31,31,   -1,-1,10,38,38,   -1,6,7,-1,-1,	 -1,-1,-1,24,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

	'/GS SHOP
	elseif (SellSite="gseshop") then
	    'xlPosArr = Array(6,8,-1,-1,-1,  15,16,17,-1,10,  11,12,13,14,14,  42,-1,41,44,45,  -1,39,37,35,-1,  7,18,66,47,-1)  ''실판매가 추가 2014/03/17 (19) : 합계금액 수량으로 나누어야함.
		'xlPosArr = Array(6,8,-1,-1,-1,  15,16,17,-1,10,  11,12,13,14,14,  42,-1,41,44,45,  -1,39,37,35,-1,  7,18,68,47,-1)  ''실판매가 추가 2014/03/17 (19) : 합계금액 수량으로 나누어야함.
		'xlPosArr = Array(8,10,-1,-1,-1,  17,18,19,-1,12,  14,13,15,16,16,  44,-1,43,46,47,  -1,41,39,37,-1,  9,20,70,49,-1)  ''실판매가 추가 2014/03/17 (19) : 합계금액 수량으로 나누어야함.
		xlPosArr = Array(8,10,-1,-1,-1,  17,18,19,-1,12,  14,13,15,16,16,  44,-1,43,46,47,  -1,41,39,37,-1,  9,20,70,49,-1,  -1,-1,6)	''교환주문 제껴야함
	    ArrayLen = UBound(xlPosArr)
	    skipString = "상태" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

	'/Homeplus
	elseif (SellSite="homeplus") then
'	    xlPosArr = Array(6,8,-1,-1,-1,  15,16,17,-1,10,  11,12,13,14,14,  42,-1,41,44,45,  -1,39,37,35,-1,  7,18,66,47,-1)  ''실판매가 추가 2014/03/17 (19) : 합계금액 수량으로 나누어야함.
	    xlPosArr = Array(3,2,-1,-1,-1,  8,12,13,-1,9,  12,13,10,11,11,  16,-1,22,21,23,  -1,18,19,14,-1,  27,24,-1,-1,-1)  ''실판매가 추가 2014/03/17 (19) : 합계금액 수량으로 나누어야함.
	    ArrayLen = UBound(xlPosArr)
	    skipString = "상태" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "배송리스트"  '' sheet name Maybe filename in

	'/이지웰페어
	elseif (SellSite="ezwel") then
'	    xlPosArr = Array(3,2,-1,-1,-1,  8,12,13,-1,9,  12,13,10,11,11,  16,-1,22,21,23,  -1,18,19,14,-1,  -1,24,-1,-1,-1)
'	    xlPosArr = Array(1,6,-1,-1,-1,  3,5,4,-1,21,  23,22,24,25,25,  -1,-1,17,Array(10,11,12),10,  -1,9,15,8,-1,  2,26,-1,19,20)  '' 출고준비중 체크 추가 2015/03/02
	    xlPosArr = Array(1,6,-1,-1,-1,  3,5,4,-1,23,  25,24,26,27,27,  -1,-1,17,Array(10,11,12),10,  -1,9,15,8,-1,  2,28,-1,19,20)  '' 출고준비중 체크 추가 2015/03/02
	    ArrayLen = UBound(xlPosArr)
	    skipString = "상태" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "배송리스트"  '' sheet name Maybe filename in

	'/gsisuper
	elseif (SellSite="gsisuper") then
	    'xlPosArr = Array(2,1,-1,-1,-1,  14,15,16,-1,20,  21,22,23,24,25,  6,-1,8,10,9,  -1,7,-1,5,-1,  -1,26,-1,11,-1)
	    'xlPosArr = Array(3,1,-1,-1,-1,  4,5,6,-1,9,  10,11,12,13,13,  -1,-1,20,22,22,  -1,18,19,17,-1,  16,15,-1,-1,-1)
		xlPosArr = Array(3,1,-1,-1,-1,  16,17,18,-1,16,  17,18,19,20,20,  -1,-1,8,12,12,  -1,6,7,5,-1,  4,10,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "상태" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "주문내역"  '' sheet name Maybe filename in
''''				0				1				2				3				4
''''				주문번호, 		주문일, 		입금일, 		지불수단, 		주문자ID,
''''				5				6				7				8				9
''''				주문자, 		주문자전화,		주문자휴대전화,	주문자이메일, 	수령인,
''''				10				11				12				13				14
''''				수령인전화,		수령인핸드폰,	수령인Zip,		수령인addr1,	수령인addr2,
''''				15				16				17				18				19
''''				상품코드, 		옵션코드, 		수량, 			판매가, 		실판매가 XX소비자가,
''''				20				21				22				23				24
''''				정산액, 		상품명, 		옵션명, 		업체상품코드,	업체옵션코드,
''''				25				26				27				28				29
''''				주문디테일키, 	주문유의사항, 	상품요구사항1, 	배송비, 		ETC1컴,
''''				30				31				32				33				34
''''				ETC2(샵링커쇼핑몰),		ETC3(샵링커상품코드),			ETC4(샵링커주문번호),			국가코드,			ETC5(샵링커착불여부)
''''				35				36				37				38				39
''''				해외판매가,		해외배송비		해외실판매가	사이트별추가정보(해외)	미사용

'' 18 판매가 = 할인가
'' 19 실판매가  (쿠폰 반영 금액) 있는경우 꼭 넣어주실것. 2013/10/24
	'/GS25
	elseif (SellSite="GS25") then
	    xlPosArr = Array(3,1,-1,-1,-1,  9,10,11,-1,14,  16,17,18,19,20,  -1,-1,8,23,23,  -1,6,7,5,-1,  4,21,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문내역" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "e-카다로그상품발주_MD"  '' sheet name Maybe filename in

	elseif (SellSite="cjmall") then
	    'xlPosArr = Array(9,4,-1,-1,-1,   10,13,-1,-1,14,  15,16,17,50,50,  41,-1,22,24,-1,   -1,27,28,25,-1,   -1,34,-1,-1,-1)
		'xlPosArr = Array(8,3,-1,-1,-1,   9,10,-1,-1,11,  12,13,14,15,15,  34,-1,19,21,20,   -1,24,25,23,-1,   -1,27,-1,-1,-1)
		'xlPosArr = Array(10,4,-1,-1,-1,   11,12,-1,-1,13,  14,15,16,17,17,  40,-1,23,25,24,   -1,29,30,27,-1,   10,20,-1,-1,-1)
		xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "상태" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

'	'/프리비아
'	elseif (SellSite="privia") then
'	    xlPosArr = Array(3,3,-1,-1,-1,  5,28,27,-1,25,  28,27,37,26,26,  -1,-1,15,17,17,  -1,12,14,-1,-1,  -1,29,-1,19,-1)
'	    ArrayLen = UBound(xlPosArr)
'	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
'	    afile = sFilePath
'	    aSheetName = "first"  '' sheet name Maybe filename in

	'/프리비아		''2013-11-15 16:40 김진영 수정해봄
	elseif (SellSite="privia") then
	    xlPosArr = Array(3,2,-1,-1,-1,  5,29,28,-1,25,  29,28,26,27,27,  -1,-1,17,19,19,  -1,12,14,11,-1,  -1,30,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "first"  '' sheet name Maybe filename in

	'/momastore
	elseif (SellSite="momastore") then
	    'xlPosArr = Array(3,2,-1,-1,-1,  5,29,28,-1,25,  29,28,26,27,27,  -1,-1,17,19,19,  -1,12,14,11,-1,  -1,30,-1,21,-1)
	    xlPosArr = Array(1,0,-1,-1,-1,  5,6,6,-1,5,  6,6,7,8,9,  -1,-1,18,15,16,  -1,11,-1,10,-1,  -1,-1,-1,17,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "주문조회"  '' sheet name Maybe filename in

	'/엔조이뉴욕
	elseif (SellSite="NJOYNY") or (SellSite="itsNJOYNY") then
	    xlPosArr = Array(0,3,-1,-1,-1,  5,12,13,14,15,   17,18,19,20,20,   -1,-1,6,11,11,   26,8,10,9,-1,	 -1,33,-1,31,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "주문관리"  '' sheet name Maybe filename in

	'/티켓몬스터
	elseif (SellSite="ticketmonster") then
		'xlPosArr = Array(1,13,-1,-1,-1,  6,17,17,-1,15,   17,17,19,18,18,   -1,-1,11,10,10,   -1,8,9,-1,-1,	 -1,20,-1,-1,-1)
		'xlPosArr = Array(1,17,-1,-1,-1,  10,22,22,-1,19,   22,22,24,23,23,   29,-1,15,14,14,   -1,12,13,-1,-1,	 -1,25,-1,-1,-1)
		xlPosArr = Array(1,18,-1,-1,-1,  10,12,12,-1,20,   23,23,25,24,24,   30,-1,16,15,15,   -1,13,14,-1,-1,	 -1,26,-1,-1,-1)
		ArrayLen = UBound(xlPosArr)
		skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
		afile = sFilePath
		aSheetName = "주문관리"  '' sheet name Maybe filename in

	'/하프클럽
	elseif (SellSite="halfclub") then
		xlPosArr = Array(0,26,-1,-1,-1,  11,14,15,-1,13,   14,15,16,17,17,   2,-1,9,8,8,   -1,3,4,2,-1,	 1,18,-1,-1,-1)
		ArrayLen = UBound(xlPosArr)
		skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
		afile = sFilePath
		aSheetName = "주문관리"  '' sheet name Maybe filename in


	'/띵크어바웃유
	elseif (SellSite="thinkaboutyou") then
	    'xlPosArr = Array(1,13,-1,-1,-1,  6,17,17,-1,15,   17,17,19,18,18,   -1,-1,11,10,10,   -1,8,9,-1,-1,	 -1,20,-1,-1,-1)
	    xlPosArr = Array(1,0,-1,-1,-1,  4,5,5,6,17,   18,19,20,21,21,   -1,-1,11,12,15,   -1,9,10,-1,-1,	 -1,22,-1,27,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "주문관리"  '' sheet name Maybe filename in
''''				0				1				2				3				4
''''				주문번호, 		주문일, 		입금일, 		지불수단, 		주문자ID,
''''				5				6				7				8				9
''''				주문자, 		주문자전화,		주문자휴대전화,	주문자이메일, 	수령인,
''''				10				11				12				13				14
''''				수령인전화,		수령인핸드폰,	수령인Zip,		수령인addr1,	수령인addr2,
''''				15				16				17				18				19
''''				상품코드, 		옵션코드, 		수량, 			판매가, 		실판매가 XX소비자가,
''''				20				21				22				23				24
''''				정산액, 		상품명, 		옵션명, 		업체상품코드,	업체옵션코드,
''''				25				26				27				28				29
''''				주문디테일키, 	주문유의사항, 	상품요구사항1, 	배송비, 		ETC1컴,
''''				30				31				32				33				34
''''				ETC2(샵링커쇼핑몰),		ETC3(샵링커상품코드),			ETC4(샵링커주문번호),			국가코드,			ETC5(샵링커착불여부)
''''				35				36				37				38				39
''''				해외판매가,		해외배송비		해외실판매가	사이트별추가정보(해외)	미사용

'' 18 판매가 = 할인가
'' 19 실판매가  (쿠폰 반영 금액) 있는경우 꼭 넣어주실것. 2013/10/24

	'/어바웃펫
	elseif (SellSite="aboutpet") then
	    'xlPosArr = Array(0,42,-1,-1,-1,  3,5,5,4,45,   43,43,44,45,46,   -1,58,23,18,22,   -1,13,-1,10,-1,	 1,68,-1,13,-1)
		xlPosArr = Array(0,45,-1,-1,-1,  3,5,5,4,48,   49,49,50,51,52,   -1,66,22,17,30,   -1,13,-1,10,-1,	 1,76,-1,37,73)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문내역" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "order"  '' sheet name Maybe filename in

	'/쿠캣
	elseif (SellSite="cookatmall") then
	    'xlPosArr = Array(1,13,-1,-1,-1,  6,17,17,-1,15,   17,17,19,18,18,   -1,-1,11,10,10,   -1,8,9,-1,-1,	 -1,20,-1,-1,-1)
	    xlPosArr = Array(0,0,-1,-1,-1,  2,3,4,-1,2,   3,4,5,6,6,   -1,-1,11,10,10,   -1,9,-1,-1,-1,	 1,15,-1,13,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문내역" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "주문관리"  '' sheet name Maybe filename in


	'/맘큐
	elseif (SellSite="momQ") then
	    'xlPosArr = Array(1,0,-1,-1,-1,  4,5,5,6,17,   18,19,20,21,21,   -1,-1,11,12,15,   -1,9,10,-1,-1,	 -1,22,-1,27,-1)
	    xlPosArr = Array(3,32,-1,-1,-1,  16,18,19,-1,17,   18,19,20,21,22,   -1,-1,13,27,27,   -1,6,7,5,-1,	 -1,-1,-1,39,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "주문관리"  '' sheet name Maybe filename in

	'/가방팝	'/이니셜서비스
	elseif (SellSite="gabangpop") or (SellSite="itsGabangpop") then
	    xlPosArr = Array(5,31,-1,-1,-1,  20,23,24,-1,20, 23,24,21,22,22,   -1,-1,14,16,16,   -1,12,13,-1,-1,	 6,25,36,18,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "gabangpop_sample"  '' sheet name Maybe filename in

	'/무신사
	elseif (SellSite="musinsaITS") or (SellSite="itsMusinsa") then
	    xlPosArr = Array(5,32,-1,-1,-1,   21,24,25,-1,21,    24,25,22,23,23,   -1,-1,14,16,16,   -1,12,13,-1,-1,	 6,26,-1,18,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in

	'/GVG
	elseif (SellSite="GVG") then
	    xlPosArr = Array(0,1,-1,-1,-1,   8,12,13,-1,9,   12,13,10,11,11,   -1,-1,7,6,6,   -1,2,4,-1,-1, -1,14,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "GVG_주문서 엑셀양식"  '' sheet name Maybe filename in

	'/플레이어
	elseif (SellSite="player") or (SellSite="itsPlayer1") then
	    xlPosArr = Array(1,1,-1,-1,-1,   8,11,12,-1,7,   11,12,9,10,10,   -1,-1,21,24,24,   -1,17,17,-1,-1,   2,13,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in
''''				0				1				2				3				4
''''				주문번호, 		주문일, 		입금일, 		지불수단, 		주문자ID,
''''				5				6				7				8				9
''''				주문자, 		주문자전화,		주문자휴대전화,	주문자이메일, 	수령인,
''''				10				11				12				13				14
''''				수령인전화,		수령인핸드폰,	수령인Zip,		수령인addr1,	수령인addr2,
''''				15				16				17				18				19
''''				상품코드, 		옵션코드, 		수량, 			판매가, 		실판매가 XX소비자가,
''''				20				21				22				23				24
''''				정산액, 		상품명, 		옵션명, 		업체상품코드,	업체옵션코드,
''''				25				26				27				28				29
''''				주문디테일키, 	주문유의사항, 	상품요구사항1, 	배송비, 		ETC1컴,
''''				30				31				32				33				34
''''				ETC2(샵링커쇼핑몰),		ETC3(샵링커상품코드),			ETC4(샵링커주문번호),			국가코드,			ETC5(샵링커착불여부)
''''				35				36				37				38				39
''''				해외판매가,		해외배송비		해외실판매가	사이트별추가정보(해외)	미사용

'' 18 판매가 = 할인가
'' 19 실판매가  (쿠폰 반영 금액) 있는경우 꼭 넣어주실것. 2013/10/24

	'/인터파크
	elseif (SellSite="interpark") then
	    ''xlPosArr = Array(0,-1,-1,-1,-1,   26,28,29,-1,30,   31,32,33,34,34,   42,43,5,9,-1,   11,3,4,42,43, 1,35,38,20)
	    ''xlPosArr = Array(0,-1,-1,-1,-1,   26,28,29,-1,30,   31,32,33,34,34,   41,42,5,9,-1,   11,3,4,41,42,   1,35,38,20) ''주문디테일키 추가 20120831
	    ' xlPosArr = Array(0,-1,-1,-1,-1,   26,28,29,-1,30,   31,32,33,34,34,   41,42,5,9,8,   11,3,4,2,42,   1,35,38,20)    ''실판매가 추가 2014/01/14 (19) : 합계금액 수량으로 나누어야함.
		'xlPosArr = Array(0,-1,-1,-1,-1,   25,27,28,-1,29,   30,31,32,33,33,   42,43,6,9,8,   11,4,5,2,3,   1,36,39,19)    ''실판매가 추가 2014/01/14 (19) : 합계금액 수량으로 나누어야함.
		'xlPosArr = Array(0,-1,-1,-1,-1,   26,28,29,-1,30,   31,32,33,34,34,   43,44,6,9,8,   11,4,5,2,3,   1,37,40,20)    ''실판매가 추가 2014/01/14 (19) : 합계금액 수량으로 나누어야함.
		'xlPosArr = Array(0,-1,-1,-1,-1,   26,28,29,-1,30,   31,32,33,34,34,   43,44,5,9,8,   11,3,4,2,-1,   1,37,40,20)
		'xlPosArr = Array(0,-1,-1,-1,-1,   26,28,29,-1,30,   31,32,33,34,34,   43,44,6,9,8,   11,4,5,2,-1,   1,37,40,20)
		'xlPosArr = Array(0,-1,-1,-1,-1,   26,28,29,-1,30,   31,32,33,34,34,   43,44,5,9,8,   11,3,4,2,-1,   1,37,40,20)
		'xlPosArr = Array(0,-1,-1,-1,-1,   34,28,29,-1,30,   31,32,33,34,34,   43,44,5,9,8,   11,3,4,2,-1,   1,37,40,20)
        xlPosArr = Array(0,28,-1,-1,-1,   34,35,36,-1,19,   20,21,22,23,23,   25,32,7,15,-1,   27,5,6,4,-1,   1,24,-1,14)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "Sheet0" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

	'/반디앤루이스
	elseif (SellSite="bandinlunis") then
	    xlPosArr = Array(0,1,-1,-1,-1,   2,6,7,-1,5,   6,7,11,8,8,   -1,-1,4,-1,-1,   -1,3,10,-1,-1,   -1,9,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "List"  '' sheet name Maybe filename in

	'/민트샵		'/사용안함
	elseif (SellSite="mintstore") or (SellSite="itsMintstore") then
	    xlPosArr = Array(0,1,-1,-1,-1,   2,6,7,-1,2,   6,7,11,8,9,   -1,-1,5,-1,-1,   -1,3,4,-1,-1,   -1,3,4,10,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "List"  '' sheet name Maybe filename in

	'/교보 핫트랙스
	elseif (SellSite="hottracks") or (SellSite="itsHottracks") then
	    xlPosArr = Array(6,3,-1,-1,-1,   5,8,8,-1,7,   8,8,12,13,14,   -1,-1,18,-1,-1,   26,16,16,15,-1,   -1,9,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "hottracks_sample"  '' sheet name Maybe filename in

	'/별샵
	elseif (SellSite="byulshopITS") or (SellSite="itsByulshop") then
	    xlPosArr = Array(2,18,-1,-1,-1,   3,5,6,-1,4,   5,6,7,8,8,   -1,-1,14,15,15,   -1,10,10,-1,-1,   1,9,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "byulshop_order_sample"  '' sheet name Maybe filename in
	'/미미박스
	elseif (SellSite="itsMemebox") then
		'xlPosArr = Array(14,0,-1,-1,-1,   10,11,11,-1,10,   11,11,12,13,13,   -1,-1,7,9,9,   -1,5,6,2,-1,   -1,19,-1,-1,-1)
		'xlPosArr = Array(11,0,-1,-1,-1,   7,8,8,-1,7,   8,8,9,10,10,   -1,-1,5,6,6,   -1,3,4,2,-1,   -1,15,-1,-1,-1)
		xlPosArr = Array(4,0,-1,-1,-1,   16,17,17,-1,16,   17,17,19,20,20,   -1,-1,12,13,13,   -1,10,11,7,-1,   -1,24,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in

	'/SUHA
	elseif (SellSite="suhaITS") then
	    xlPosArr = Array(0,1,-1,-1,-1,   14,17,18,-1,8,   11,12,9,10,10,   -1,-1,7,6,6,   -1,2,4,3,-1,   -1,13,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "121.88.197.9_orders(1)"  '' sheet name Maybe filename in

	'/지마켓
	elseif (SellSite="gmarket") then
    	'xlPosArr = Array(2,3,31,-1,-1,   6,25,24,-1,13,  15,14,16,17,17,   -1,-1,10,4,4,   33,8,10,1,-1,   -1,18,-1,21,-1)
    	xlPosArr = Array(24,3,31,-1,-1,   6,26,25,-1,13,  15,14,16,17,17,   -1,-1,9,5,5,   33,8,10,1,-1,   2,18,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in

	'/옥션
	elseif (SellSite="auction1010") OR (SellSite="gmarket1010") then
    	'xlPosArr = Array(2,3,31,-1,-1,   6,25,24,-1,13,  15,14,16,17,17,   -1,-1,10,4,4,   33,8,10,1,-1,   -1,18,-1,21,-1)
    	xlPosArr = Array(23,3,33,-1,-1,   6,22,21,-1,12,  14,13,15,16,16,   34,-1,9,4,24,   25,8,10,1,-1,   2,17,-1,20,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in

	'/아이띵소 와디즈
	elseif (SellSite="itsWadiz") then
    	xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

	'/11번가
	elseif (SellSite="11st1010") then
'		xlPosArr = Array(2,46,4,-1,-1,  33,29,28,-1,12,   29,28,30,31,31,   -1,-1,10,38,38,   -1,6,7,-1,-1,	 -1,-1,-1,24,-1)
    	xlPosArr = Array(2,46,4,-1,-1,  33,29,28,-1,12,  29,28,30,31,31,   37,-1,10,38,40,   -1,6,7,36,-1,   3,32,-1,27,5, -1,41,39,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in
''''				0				1				2				3				4
''''				주문번호, 		주문일, 		입금일, 		지불수단, 		주문자ID,
''''				5				6				7				8				9
''''				주문자, 		주문자전화,		주문자휴대전화,	주문자이메일, 	수령인,
''''				10				11				12				13				14
''''				수령인전화,		수령인핸드폰,	수령인Zip,		수령인addr1,	수령인addr2,
''''				15				16				17				18				19
''''				상품코드, 		옵션코드, 		수량, 			판매가, 		실판매가 XX소비자가,
''''				20				21				22				23				24
''''				정산액, 		상품명, 		옵션명, 		업체상품코드,	업체옵션코드,
''''				25				26				27				28				29
''''				주문디테일키, 	주문유의사항, 	상품요구사항1, 	배송비, 		ETC1컴,
''''				30				31				32				33				34
''''				ETC2(샵링커쇼핑몰),		ETC3(샵링커상품코드),			ETC4(샵링커주문번호),			국가코드,			ETC5(샵링커착불여부)
''''				35				36				37				38				39
''''				해외판매가,		해외배송비		해외실판매가	사이트별추가정보(해외)	미사용

'' 18 판매가 = 할인가
'' 19 실판매가  (쿠폰 반영 금액) 있는경우 꼭 넣어주실것. 2013/10/24

	'/신세계TV쇼핑..공통양식으로 작업
	elseif (SellSite="shintvshopping") then
    	xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

	'/skstoa..공통양식으로 작업
	elseif (SellSite="skstoa") then
    	xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

	'/LFmall..공통양식으로 작업
	elseif (SellSite="LFmall") then
    	xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

	'/굿웨어몰..공통양식으로 작업
	elseif (SellSite="goodwearmall10") then
    	xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

	'/wconcept1010..공통양식으로 작업
	elseif (SellSite="wconcept1010") then
    	xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "Sheet"  '' sheet name Maybe filename in

	'/스토어팜
	elseif (SellSite="nvstorefarm") then
    	'xlPosArr = Array(2,3,31,-1,-1,   6,25,24,-1,13,  15,14,16,17,17,   -1,-1,10,4,4,   33,8,10,1,-1,   -1,18,-1,21,-1)
    	xlPosArr = Array(3,53,9,-1,-1,   4,40,40,-1,6,  37,37,41,39,39,   36,18,19,21,24,   -1,15,17,14,-1,   2,42,-1,33,-1,	22,23,20,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "발주발송관리"  '' sheet name Maybe filename in

	'/스토어팜 문방구
	elseif (SellSite="nvstoremoonbangu") then
    	'xlPosArr = Array(2,3,31,-1,-1,   6,25,24,-1,13,  15,14,16,17,17,   -1,-1,10,4,4,   33,8,10,1,-1,   -1,18,-1,21,-1)
    	xlPosArr = Array(3,53,9,-1,-1,   4,40,40,-1,6,  37,37,41,39,39,   36,18,19,21,24,   -1,15,17,14,-1,   2,42,-1,33,-1,	22,23,20,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "발주발송관리"  '' sheet name Maybe filename in

	'/스토어팜 캣앤독
	elseif (SellSite="Mylittlewhoopee") then
    	'xlPosArr = Array(2,3,31,-1,-1,   6,25,24,-1,13,  15,14,16,17,17,   -1,-1,10,4,4,   33,8,10,1,-1,   -1,18,-1,21,-1)
    	xlPosArr = Array(3,53,9,-1,-1,   4,40,40,-1,6,  37,37,41,39,39,   36,18,19,21,24,   -1,15,17,14,-1,   2,42,-1,33,-1,	22,23,20,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "발주발송관리"  '' sheet name Maybe filename in

	'/스토어팜 선물하기
	elseif (SellSite="nvstoregift") then
    	'xlPosArr = Array(2,3,31,-1,-1,   6,25,24,-1,13,  15,14,16,17,17,   -1,-1,10,4,4,   33,8,10,1,-1,   -1,18,-1,21,-1)
    	xlPosArr = Array(3,53,9,-1,-1,   4,40,40,-1,6,  37,37,41,39,39,   36,18,19,21,24,   -1,15,17,14,-1,   2,42,-1,33,-1,	22,23,20,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "발주발송관리"  '' sheet name Maybe filename in

	'/스토어팜클래스
	elseif (SellSite="nvstorefarmclass") then
    	xlPosArr = Array(1,55,13,-1,-1,   7,42,42,-1,9,  39,39,-1,-1,-1,   36,18,19,21,24,   -1,15,17,14,-1,   0,44,-1,-1,-1,	22,23,20,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = "발주발송관리"  '' sheet name Maybe filename in

	''cjmallITS		'/진영 2013/03/07추가../진영 2014-02-21 엑셀 폼이 바뀜에 따라 하단 수정../진영 2014-10-02 엑셀 폼이 바뀜에 따라 하단 수정../
	elseif (SellSite="cjmallITS") or (SellSite="itsCjmall") then
		'xlPosArr = Array(9,4,-1,-1,-1,   10,13,-1,-1,14,  15,16,17,50,50,  -1,-1,22,24,-1,   -1,27,28,26,-1,   -1,34,-1,-1,-1)
		xlPosArr = Array(8,3,-1,-1,-1,   9,10,-1,-1,11,  12,13,14,15,15,  -1,-1,19,21,-1,   -1,24,25,23,-1,   -1,27,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호"
	    afile = sFilePath
	    aSheetName = "Sheet1"  '' sheet name Maybe filename in
    '/힙합퍼
	elseif (SellSite="hiphoper") or (SellSite="itsHiphoper") then
	    xlPosArr = Array(5,-1,-1,-1,-1,   0,3,4,-1,0,   3,4,1,2,2,   -1,-1,10,11,11,   -1,8,7,6,-1,   -1,12,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "수령자" ''
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

    '/아이띵소_29cm
	elseif (SellSite="its29cm") then
	    xlPosArr = Array(0,1,-1,-1,-1,   2,3,4,5,6,   7,8,9,10,10,   -1,-1,17,16,16,   -1,14,15,13,-1,   -1,11,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

	'/위즈위드		'/이니셜서비스
    elseif (SellSite="wizwid") or (SellSite="itsWizwid") then
	    xlPosArr = Array(4,2,-1,-1,-1,   14,20,21,-1,15,   18,19,16,17,17,   -1,-1,11,13,13,   -1,7,9,6,-1,   3,23,29,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "택배사코드" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

	'/더블유컨셉	'/이니셜서비스
	elseif (SellSite="wconcept") or (SellSite="itsWconcept") then
	    If SellSite = "itsWconcept" Then
	    	xlPosArr = Array(2,1,-1,-1,-1,   13,19,20,-1,14,   17,18,15,16,16,   -1,-1,10,12,12,   -1,6,8,-1,-1,   3,21,27,-1,-1)
	    Else
	    	xlPosArr = Array(3,2,-1,-1,-1,   16,22,23,-1,17,   20,21,18,19,19,   -1,-1,13,15,15,   -1,7,9,-1,-1,   4,24,-1,-1,-1)
	    End If
	    ArrayLen = UBound(xlPosArr)
	    skipString = "택배사코드" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

	'/올레TV		'/사용안함
	elseif (SellSite="ollehtv") then
	    xlPosArr = Array(2,17,18,-1,-1,  9,11,12,-1,14,   11,12,15,16,16,   -1,-1,8,7,7,   -1,5,6,4,-1,   -1,-1,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString = "주문번호" ''Array("발송관리","◎생성시각","○발송관리","항목명에","주문번호")
	    afile = sFilePath
	    aSheetName = ""  '' sheet name Maybe filename in

	'/하나투어		'/사용안함
    elseif (SellSite="hanatour") then
        xlPosArr = Array(0,15,15,-1,-1		,2,4,5,-1,3		,4,5,16,16,16		,1,-1,8,9,9		,10,6,7,1,-1		,-1,17,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="주문번호"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in

	'/패션플러스
    elseif (SellSite="fashionplus") or (SellSite="itsFashionplus") then
    	xlPosArr = Array(1,6,6,-1,-1,     28,27,26,-1, 2,     27,26,25,23,23 ,     11,-1,13,14,14,     -1,12,4 ,11,-1,     -1,29,-1,24,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="주문번호"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in

	'/기프팅
    elseif (SellSite="giftting") then
    	xlPosArr = Array(1,13,13,-1,-1,     16,17,19,-1, 18,     17,19,20,21,21,     7,-1,22,23,23,     -1,5,6,-1,-1,     10,14,-1,24,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="관리용"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in


''''				0				1				2				3				4
''''				주문번호, 		주문일, 		입금일, 		지불수단, 		주문자ID,
''''				5				6				7				8				9
''''				주문자, 		주문자전화,		주문자휴대전화,	주문자이메일, 	수령인,
''''				10				11				12				13				14
''''				수령인전화,		수령인핸드폰,	수령인Zip,		수령인addr1,	수령인addr2,
''''				15				16				17				18				19
''''				상품코드, 		옵션코드, 		수량, 			판매가, 		실판매가 XX소비자가,
''''				20				21				22				23				24
''''				정산액, 		상품명, 		옵션명, 		업체상품코드,	업체옵션코드,
''''				25				26				27				28				29
''''				주문디테일키, 	주문유의사항, 	상품요구사항1, 	배송비, 		ETC1컴,
''''				30				31				32				33				34
''''				ETC2(샵링커쇼핑몰),		ETC3(샵링커상품코드),			ETC4(샵링커주문번호),			국가코드,			ETC5(샵링커착불여부)
''''				35				36				37				38				39
''''				해외판매가,		해외배송비		해외실판매가	사이트별추가정보(해외)	미사용

'' 18 판매가 = 할인가
'' 19 실판매가  (쿠폰 반영 금액) 있는경우 꼭 넣어주실것. 2013/10/24

	'/아이띵소_베네피아
    elseif (SellSite="itsbenepia") then
    	xlPosArr = Array(0,4,4,-1,-1,     5,22,23,-1, 14,     22,23,25,26,27,     -1,-1,10,11,11,     -1,6,8,7,-1,     1,28,-1,13,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="Sheet0"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in

	'/아이띵소_카카오톡스토어
    elseif (SellSite="itskakaotalkstore") then
    	'xlPosArr = Array(2,20,-1,-1,-1,     11,13,13,-1, 11,     13,13,17,16,16,     31,-1,6,21,23,     -1,4,5,3,-1,     -1,18,-1,37,-1)
		'xlPosArr = Array(2,19,-1,-1,-1,     10,12,12,-1, 10,     12,12,16,15,15,     30,-1,6,20,22,     -1,4,5,3,-1,     -1,17,-1,35,-1)
		xlPosArr = Array(2,19,-1,-1,-1,     10,12,12,-1, 10,     12,12,16,15,15,     30,-1,6,20,22,     -1,4,5,3,-1,     -1,17,-1,35,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="Sheet0"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in

	'/메이커스위드카카오
    elseif (SellSite="itsKaKaoMakers") then
    	'xlPosArr = Array(2,0,0,-1,-1,     12,13,14,-1, 12,     13,14,16,15,15,     7,-1,10,8,8,     -1,3,4,6,-1,     -1,17,-1,9,-1)
    	xlPosArr = Array(34,0,0,-1,-1,     11,13,13,-1, 11,     13,13,17,16,16,     -1,-1,6,22,22,     -1,4,5,3,-1,     2,18,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="Sheet0"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in

	'/카카오기프트
    ' elseif (SellSite="kakaogift") then
    ' 	xlPosArr = Array(2,0,0,-1,-1,     11,13,15,-1, 11,     13,15,17,16,16,     30,31,6,21,21,     -1,4,5,3,-1,     -1,18,-1,36,-1)
	'     ArrayLen = UBound(xlPosArr)
	'     skipString="관리용"
	'     afile = sFilePath
	'     aSheetName = maybeSheetName  '' sheet name Maybe filename in

	'/카카오기프트
    elseif (SellSite="kakaogift") then
    	xlPosArr = Array(3,1,1,-1,-1,     11,13,15,-1, 11,     13,15,17,16,16,     33,34,7,21,21,     25,5,6,4,-1,     -1,18,-1,38,22)
	    ArrayLen = UBound(xlPosArr)
	    skipString="관리용"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in
	'/아이띵소_카카오선물하기
    elseif (SellSite="itskakao") then
    	'xlPosArr = Array(2,0,0,-1,-1,     12,13,14,-1, 12,     13,14,16,15,15,     7,-1,10,8,8,     -1,3,4,6,-1,     -1,17,-1,9,-1)
    	'xlPosArr = Array(2,0,0,-1,-1,     11,13,13,-1, 11,     13,13,17,16,16,     30,-1,6,21,23,     -1,4,5,3,-1,     -1,18,-1,-1,-1)
		xlPosArr = Array(2,0,0,-1,-1,     10,12,12,-1, 10,     12,12,16,15,15,     29,-1,6,20,22,     -1,4,5,3,-1,     -1,17,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="Sheet0"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in

	'/아이띵소샵
    elseif (SellSite="ithinksoshop") then
'    	xlPosArr = Array(0,0,0,-1,-1,     11,12,12,-1, 11,     12,12,13,14,15,     -1,-1,9,10,4,     -1,7,8,6,-1,     1,2,-1,-1,-1)
    	xlPosArr = Array(1,3,4,-1,-1,     28,31,30,-1, 32,     34,33,27,25,26,     -1,-1,24,7,9,     -1,6,13,5,-1,     2,35,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="주문번호"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in

	'/위메크 프라이스
	elseif (SellSite="wemakeprice") then
	    'xlPosArr = Array(0,1,1,-1,-1,     6,15,15,-1,14,     15,15,16,17,17,     23,-1,12,13,13,     -1,9,11,8,10,     -1,22,19,-1,-1)
	    xlPosArr = Array(0,1,1,-1,-1,     6,14,14,-1,13,     14,14,15,16,16,     22,-1,11,12,12,     -1,8,10,7,9,     -1,21,18,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="주문ID"
	    afile = sFilePath
	    aSheetName = "Worksheet"  '' sheet name Maybe filename in

''''				0				1				2				3				4
''''				주문번호, 		주문일, 		입금일, 		지불수단, 		주문자ID,
''''				5				6				7				8				9
''''				주문자, 		주문자전화,		주문자휴대전화,	주문자이메일, 	수령인,
''''				10				11				12				13				14
''''				수령인전화,		수령인핸드폰,	수령인Zip,		수령인addr1,	수령인addr2,
''''				15				16				17				18				19
''''				상품코드, 		옵션코드, 		수량, 			판매가, 		실판매가 XX소비자가,
''''				20				21				22				23				24
''''				정산액, 		상품명, 		옵션명, 		업체상품코드,	업체옵션코드,
''''				25				26				27				28				29
''''				주문디테일키, 	주문유의사항, 	상품요구사항1, 	배송비, 		ETC1컴,
''''				30				31				32				33				34
''''				ETC2(샵링커쇼핑몰),		ETC3(샵링커상품코드),			ETC4(샵링커주문번호),			국가코드,			ETC5(샵링커착불여부)
''''				35				36				37				38				39
''''				해외판매가,		해외배송비		해외실판매가	사이트별추가정보(해외)	미사용

	'/신세계몰(SSG)
	elseif (SellSite="ssg") then
        'xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
        'xlPosArr = Array(8,35,-1,-1,-1,  36,38,39,-1,37,  38,39,40,41,41,   21,-1,29,30,30,   -1,20,23,22,-1,   11,47,24,-1,6, -1,-1,-1,-1,-1,	-1,-1,-1,7,-1)
		xlPosArr = Array(9,4,-1,-1,-1,  36,38,39,-1,37,  38,39,40,41,41,   21,-1,30,31,31,   -1,20,24,23,25,   12,47,-1,-1,7, -1,-1,-1,-1,-1,	-1,-1,-1,8,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="Sheet1"
	    afile = sFilePath
	    aSheetName = maybeSheetName

	'/쿠팡
	elseif (SellSite="coupang") then
	    'xlPosArr = Array(2,4,-1,-1,-1,	11,12,12,10,13,		14,14,15,16,16,	-1,-1,7,5,5,	-1,9,-1,-1,-1,	-1,17,-1,6,-1)
        'xlPosArr = Array(0,1,-1,-1,-1,  2,3,4,5,6,  7,8,9,10,10,   11,12,13,14,15,   -1,16,17,18,-1,   19,20,-1,21,-1)
		xlPosArr = Array(2,9,-1,-1,-1,	24,26,26,25,27,		28,28,29,30,30,	16,16,22,23,23,	-1,10,11,13,-1,	-1,31,14,20,1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="배송관리"
	    afile = sFilePath
	    aSheetName = maybeSheetName

	'/탐스 - 후이즈
	elseif (SellSite="5") then
	    xlPosArr = Array(0,2,42,39,6,5,9,10,11,13,16,17,14,15,15,22,-1,28,29,31,-1,24,25,-1,-1,-1,12,19,20,21)
	    ArrayLen = UBound(xlPosArr)
	    skipString="주문번호"
	    afile = sFilePath
	    aSheetName = "TOMS 주문내역"  '' sheet name Maybe filename in

	'/텐바이텐 - 인덱스추가형식
	elseif (SellSite="4") then
	    xlPosArr = Array(1,2,-1,-1,-1,3,4,5,6,7,8,9,10,11,12,15,-1,19,18,-1,-1,16,17,21,-1,0,13,20,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="일련번호"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in

    else
	    response.write "<script>alert('쇼핑몰 코드가 지정되지 않았습니다. -"&SellSite&"');</script>"
	    response.end
	end if

''ReDim xlRow(ArrayLen)
Dim xlRowALL

''rw "ArrayLen="&ArrayLen

dim ret : ret = fnGetXLFileArray(xlRowALL, afile, aSheetName, ArrayLen)

if (Not ret) or (Not IsArray(xlRowALL)) then
    response.write "<script>alert('파일이 올바르지 않거나 내용이 없습니다. "&Replace(Err.Description,"'","")&"');</script>"

    if (Err.Description="외부 테이블 형식이 잘못되었습니다.") then
        response.write "<script>alert('엑셀에서 Save As Excel 97 -2003 통합문서 형태로 저장후 사용하세요.');</script>"
    end if
    response.write "<script>history.back();</script>"
    response.end
end if

''데이터 처리.
Dim iLine, iResult
Dim paramInfo, retParamInfo, RetErr, retErrStr, sqlStr
Dim POS1,POS2,POS3, okCNT, bufItemName, bufItemSplit, bufrowObj, bufOneItemName, bufItemNo, bufOptionName
Dim errCNT, totErrMsg, tmpItemname
Dim rtitemid, rtitemoption, rtSellPrice
Dim t_addDlvPrice ,t_deliverytype ,t_sellcash ,t_defaultFreeBeasongLimit ,t_defaultDeliverPay, tempShippay
Dim tmppreAddr, tmpnextAddr
okCNT = 0 : errCNT = 0

Dim pcnt : pcnt = UBound(xlRowALL)
'    IF (sellsite="wemakeprice") then
'        for i=0 to pcnt
'            if IsObject(xlRowALL(i)) then
'                set iLine = xlRowALL(i)
'                bufItemName = iLine.FItemArray(21)
'                if (InStr(bufItemName,"] [")>0) then
'                    bufItemSplit = split(bufItemName,"] [")
'                    for k=LBound(bufItemSplit) to UBound(bufItemSplit)
'                        bufOneItemName = bufItemSplit(k)
'                        bufItemNo  = Trim(replace(Right(Replace(Replace(bufOneItemName,"]",""),"개",""),2)," ",""))
'
'                        IF K=0 then
'                            iLine.FItemArray(21) = bufOneItemName + "]"
'                            iLine.FItemArray(17) = bufItemNo
'    'rw iLine.FItemArray(21)
'    'rw iLine.FItemArray(17)
'                        ELSE
'                            set bufrowObj = new TXLRowObj
'                            bufrowObj.setArrayLength(UBound(iLine.FItemArray))
'
'                            For m=LBound(iLine.FItemArray) to UBound(iLine.FItemArray)
'                                bufrowObj.FItemArray(m) = iLine.FItemArray(m)
'                            Next
'
'                            bufrowObj.FItemArray(21)="[" + bufOneItemName
'                            bufrowObj.FItemArray(17)=bufItemNo
'    'rw "B:"&bufrowObj.FItemArray(21)
'    'rw "B:"&bufrowObj.FItemArray(17)
'                            ReDim Preserve xlRowALL(UBound(xlRowALL)+1)
'                            set xlRowALL(UBound(xlRowALL)) =  bufrowObj
'                        ENd IF
'                    next
'                end if
'             end if
'        Next
'    End IF

IF (sellsite<>"interpark") then
    dbget.BeginTrans
end if

    for i=0 to UBound(xlRowALL)
    ''rw UBound(xlRowALL)
        if (i>3000) then Exit For  ''''
        if IsObject(xlRowALL(i)) then
            set iLine = xlRowALL(i)
            ''날짜 형식 변경 - wemakePrice
            iLine.FItemArray(1) = Replace(iLine.FItemArray(1),".","-")
            iLine.FItemArray(1) = Replace(iLine.FItemArray(1),"/","-")
            iLine.FItemArray(2) = Replace(iLine.FItemArray(2),"/","-")

            if Len(iLine.FItemArray(1)) > 19 then
            	iLine.FItemArray(1) = Left(iLine.FItemArray(1), 19)
            end if

            if Len(iLine.FItemArray(2)) > 19 then
            	iLine.FItemArray(2) = Left(iLine.FItemArray(2), 19)
            end if

            if (iLine.FItemArray(9)="-") then iLine.FItemArray(9)=iLine.FItemArray(5)
            if (iLine.FItemArray(10)="-") then iLine.FItemArray(10)=iLine.FItemArray(6)
            if (iLine.FItemArray(11)="-") then iLine.FItemArray(11)=iLine.FItemArray(7)

''------------------------------------------
            IF (sellsite="shoplinker") then
                ''주문일/입금일 ''20130813125800
                iLine.FItemArray(1) = Left(iLine.FItemArray(1),4)&"-"&Mid(iLine.FItemArray(1),5,2)&"-"&Mid(iLine.FItemArray(1),7,2)&" "&Mid(iLine.FItemArray(1),9,2)&":"&Mid(iLine.FItemArray(1),11,2)&":"&Mid(iLine.FItemArray(1),13,2)
                iLine.FItemArray(2) = iLine.FItemArray(1)

                ''iLine.FItemArray(16) 옵션코드
'                if (iLine.FItemArray(16)="") and (iLine.FItemArray(24)<>"") then
'                    iLine.FItemArray(16)=getOptionCodByOption(iLine.FItemArray(15),iLine.FItemArray(24))
'                end if
'
'                if (iLine.FItemArray(16)="") and ((iLine.FItemArray(24)="NONE") or (iLine.FItemArray(24)="")) then
'                    iLine.FItemArray(16)="0000"
'                end if

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),replace(replace(iLine.FItemArray(22)," / FREE","")," FREE",""))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

                if (iLine.FItemArray(16)="") then
                    iLine.FItemArray(16)="0000"
                end if

                ''배송메세지 // gs / 매체구분,출고준수일

                if (iLine.FItemArray(26)=",,") then iLine.FItemArray(26)="" ''HOTTRACKS

                if (iLine.FItemArray(30)="HOTTRACKS") and (iLine.FItemArray(28)>2500) then ''배송비필드에 상품가가 들어가 있음
                    if (iLine.FItemArray(34)="유료배송") then
                        iLine.FItemArray(28)="2500"
                    else
                        iLine.FItemArray(28)="0"
                    end if
                end if

            end if

            IF (sellsite="cn10x10") then
            	'//제휴몰 상품명을 가지고 텐바이텐 상품명이 있나 확인해서 받아옴
                if iLine.FItemArray(23)<>"" then
                    iLine.FItemArray(15) = getItemIDByUpcheItemCode(sellsite,iLine.FItemArray(23))
                end if

				if (iLine.FItemArray(24)<>"") then
					if right(iLine.FItemArray(24),4)<>"0000" then
						iLine.FItemArray(24) = right(iLine.FItemArray(24),4)
						iLine.FItemArray(16)=getOptionCodByOption(iLine.FItemArray(15),iLine.FItemArray(24))
					else
						iLine.FItemArray(16)="0000"
					end if
				else
					iLine.FItemArray(16)="0000"
				end if

				iLine.FItemArray(5) = replace(replace(replace(iLine.FItemArray(5),"  ","!@@@!")," ",""),"!@@@!"," ")
				iLine.FItemArray(9) = replace(replace(replace(iLine.FItemArray(9),"  ","!@@@!")," ",""),"!@@@!"," ")
            END IF

			If (sellsite="cnglob10x10") Then
'On Error Resume Next
				tmpOverseasRealprice = iLine.FItemArray(19)
				If Lcase(iLine.FItemArray(30)) <> "gift" Then
					iLine.FItemArray(6) = replace(iLine.FItemArray(6), "+", "")						'전화번호에 + 공백으로 치환
					iLine.FItemArray(6) = replace(iLine.FItemArray(6), " ", "-")					'전화번호에 블랭크를 -로 치환

					iLine.FItemArray(7) = replace(iLine.FItemArray(7), "+", "")						'전화번호에 + 공백으로 치환
					iLine.FItemArray(7) = replace(iLine.FItemArray(7), " ", "-")					'전화번호에 블랭크를 -로 치환

					iLine.FItemArray(10) = replace(iLine.FItemArray(10), "+", "")					'전화번호에 + 공백으로 치환
					iLine.FItemArray(10) = replace(iLine.FItemArray(10), " ", "-")					'전화번호에 블랭크를 -로 치환

					iLine.FItemArray(11) = replace(iLine.FItemArray(11), "+", "")					'전화번호에 + 공백으로 치환
					iLine.FItemArray(11) = replace(iLine.FItemArray(11), " ", "-")					'전화번호에 블랭크를 -로 치환

					iLine.FItemArray(18) = Int(CLng((iLine.FItemArray(29) * iLine.FItemArray(18)) / iLine.FItemArray(17)) * 0.1) * 10		'엑셀의 주문품목 결제금액이 상품자체할인한 값 * 수량이 들어옴...따라서 (환율 * 주문품목 결제금액)/수량 후 원단위 절삭
					iLine.FItemArray(19) = Int(CLng((iLine.FItemArray(29) * iLine.FItemArray(19)) / iLine.FItemArray(17)) * 0.1) * 10		'실판매가인데..쿠폰 사용한 금액이 원본 엑셀에 없음..경주대리님에게 열추가후 수기 입력 요청함..
					iLine.FItemArray(37) = CDBL(FormatNumber(tmpOverseasRealprice / iLine.FItemArray(17),2))								'해외실판매가
					iLine.FItemArray(28) = Int(CLng(iLine.FItemArray(29) * iLine.FItemArray(28)) * 0.1) * 10	'배송비 = 환율 * 배송비 후 원단위 절삭

					overseasPrice			= iLine.FItemArray(35)
					overseasDeliveryPrice	= iLine.FItemArray(36)
					overseasRealPrice		= CDBL(FormatNumber(tmpOverseasRealprice / iLine.FItemArray(17),2))
					'######################## 김수현까지는 이 If문 #################################
'					If (iLine.FItemArray(24) <> "") Then
'						If right(iLine.FItemArray(24),4) <> "0000" Then
'							iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(24))
'						Else
'							iLine.FItemArray(16)="0000"
'						End if
'					Else
'						iLine.FItemArray(16)="0000"
'					End If
					'######################## 김수현까지는 이 If문 #################################

					'######################## 김수현 지나서는 이 If문  #################################
					If iLine.FItemArray(22) <> "" Then
						iLine.FItemArray(22) = Trim(iLine.FItemArray(22))
						iLine.FItemArray(22) = Replace(iLine.FItemArray(22), " ", "")			'띄어쓰기 치환
						'iLine.FItemArray(22) = Split(iLine.FItemArray(22), ":")(1)
					End If

					If (iLine.FItemArray(24) <> "") Then
						iLine.FItemArray(16) = getOptionCodeByMakeShopOptCode(iLine.FItemArray(15),iLine.FItemArray(24))
					Else
						iLine.FItemArray(16)="0000"
					End If

'					if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
'						if (iLine.FItemArray(22)<>"") then
'							iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
'						else
'							iLine.FItemArray(16)="0000"
'						end if
'					end if
					'######################## 김수현 지나서는 이 If문  #################################
				Else

				End If
'On Error Goto 0
			End If

			If (sellsite="cnhigo") Then
				tmpOverseasRealprice = ""
				tmpOverseasRealprice = iLine.FItemArray(19)
				iLine.FItemArray(18) = Int(CLng((iLine.FItemArray(29) * iLine.FItemArray(18)) / iLine.FItemArray(17)) * 0.1) * 10		'엑셀의 주문품목 결제금액이 상품자체할인한 값 * 수량이 들어옴...따라서 (환율 * 주문품목 결제금액)/수량 후 원단위 절삭
				iLine.FItemArray(19) = Int(CLng((iLine.FItemArray(29) * iLine.FItemArray(19)) / iLine.FItemArray(17)) * 0.1) * 10		'실판매가인데..쿠폰 사용한 금액이 원본 엑셀에 없음..경주대리님에게 열추가후 수기 입력 요청함..
				iLine.FItemArray(28) = Int(CLng(iLine.FItemArray(29) * iLine.FItemArray(28)) * 0.1) * 10	'배송비 = 환율 * 배송비 후 원단위 절삭

				overseasPrice			= iLine.FItemArray(35)
				overseasDeliveryPrice	= iLine.FItemArray(36)
				overseasRealPrice		= CDBL(FormatNumber(tmpOverseasRealprice / iLine.FItemArray(17),2))

				If iLine.FItemArray(16) = "" Then
					iLine.FItemArray(16)="0000"
				End If
			End If

			If (sellsite="cnugoshop") Then
				tmpOverseasRealprice = ""
				tmpOverseasRealprice = iLine.FItemArray(19)
				iLine.FItemArray(18) = Int(CLng((iLine.FItemArray(29) * iLine.FItemArray(18)) / iLine.FItemArray(17)) * 0.1) * 10		'엑셀의 주문품목 결제금액이 상품자체할인한 값 * 수량이 들어옴...따라서 (환율 * 주문품목 결제금액)/수량 후 원단위 절삭
				iLine.FItemArray(19) = Int(CLng((iLine.FItemArray(29) * iLine.FItemArray(19)) / iLine.FItemArray(17)) * 0.1) * 10		'실판매가인데..쿠폰 사용한 금액이 원본 엑셀에 없음..경주대리님에게 열추가후 수기 입력 요청함..
				iLine.FItemArray(28) = Int(CLng(iLine.FItemArray(29) * iLine.FItemArray(28)) * 0.1) * 10	'배송비 = 환율 * 배송비 후 원단위 절삭

				overseasPrice			= iLine.FItemArray(35)
				overseasDeliveryPrice	= iLine.FItemArray(36)
				overseasRealPrice		= CDBL(FormatNumber(tmpOverseasRealprice / iLine.FItemArray(17),2))

				If iLine.FItemArray(16) = "" Then
					iLine.FItemArray(16)="0000"
				End If
			End If

			If (sellsite="11stmy") Then
				Dim oDateArr, spOptName
				tmpOverseasRealprice = ""
				tmpOverseasRealprice = iLine.FItemArray(18) - iLine.FItemArray(19)		'판매가(iLine.FItemArray(18)) - 쿠폰가(iLine.FItemArray(19))를 tmpOverseasRealprice에 입력
				reserve01 = iLine.FItemArray(38)

				oDateArr = ""
				oDateArr		= Split(iLine.FItemArray(1), "-")
				If Ubound(oDateArr) > 0 Then
					iLine.FItemArray(1) = oDateArr(2)&"-"&oDateArr(1)&"-"&oDateArr(0)
				End If
				iLine.FItemArray(2) = iLine.FItemArray(1)
				iLine.FItemArray(18) = Int(CLng(iLine.FItemArray(29) * iLine.FItemArray(18)) * 0.1) * 10		'판매가 = (환율 * 객단가) 원단위 절삭
				iLine.FItemArray(19) = Int(CLng(iLine.FItemArray(29) * tmpOverseasRealprice) * 0.1) * 10		'실판매가 = (환율 * 객단가) 원단위 절삭
				iLine.FItemArray(28) = Int(CLng(iLine.FItemArray(29) * iLine.FItemArray(28)) * 0.1) * 10		'배송비 = 환율 * 배송비 후 원단위 절삭

				overseasPrice			= iLine.FItemArray(35)
				overseasDeliveryPrice	= iLine.FItemArray(36)
				overseasRealPrice		= tmpOverseasRealprice

				If (iLine.FItemArray(16) = "") and (iLine.FItemArray(15) <> "") then       ''옵션코드가 빈값.
					If (iLine.FItemArray(22) <> "") then
						spOptName = Trim(Split(iLine.FItemArray(22), "/")(0))
						iLine.FItemArray(16) = get11stOptionCodeByOptionName(iLine.FItemArray(15), Trim(Split(iLine.FItemArray(22), "/")(0)) )
						iLine.FItemArray(22) = spOptName
					Else
						iLine.FItemArray(16) = "0000"
					End If
				End if
			End If

			If (sellsite="zilingo") Then
				iLine.FItemArray(1) = CDate(iLine.FItemArray(1))
				iLine.FItemArray(2) = CDate(iLine.FItemArray(2))

				overseasPrice			= iLine.FItemArray(35)
				overseasDeliveryPrice	= 0
				overseasRealPrice		= iLine.FItemArray(35)

				iLine.FItemArray(28) = 0			'배송비
				iLine.FItemArray(36) = 0			'해외배송비

				tmpVal = getItemidOptionCodeByZilignoGoodno(iLine.FItemArray(23))
				iLine.FItemArray(15) = Split(tmpVal, "||")(0)
                iLine.FItemArray(16) = Split(tmpVal, "||")(1)
			End If

			If (sellsite="etsy") Then
				Dim splitYear, splitMonth, splitDay
				splitYear	= Split(iLine.FItemArray(1), "-")(2)
				splitMonth	= Split(iLine.FItemArray(1), "-")(0)
				splitDay	= Split(iLine.FItemArray(1), "-")(1)
				iLine.FItemArray(1) = splitYear & "-" & splitMonth & "-" & splitDay
				iLine.FItemArray(2) = splitYear & "-" & splitMonth & "-" & splitDay

				tmpOverseasRealprice = ""
				tmpOverseasRealprice = iLine.FItemArray(18) - iLine.FItemArray(19)		'판매가(iLine.FItemArray(18)) - 쿠폰가(iLine.FItemArray(19))를 tmpOverseasRealprice에 입력

				iLine.FItemArray(18) = Int(CLng(iLine.FItemArray(29) * iLine.FItemArray(18)) * 0.1) * 10		'판매가 = (환율 * 객단가) 원단위 절삭
				iLine.FItemArray(19) = Int(CLng(iLine.FItemArray(29) * tmpOverseasRealprice) * 0.1) * 10		'실판매가 = (환율 * 객단가) 원단위 절삭
				iLine.FItemArray(28) = Int(CLng(iLine.FItemArray(29) * iLine.FItemArray(28)) * 0.1) * 10		'배송비 = 환율 * 배송비 후 원단위 절삭

				iLine.FItemArray(14) = iLine.FItemArray(14) & " " & iLine.FItemArray(38) & " " & iLine.FItemArray(39)	'주소가 제각각 나눠져 있음

				overseasPrice			= iLine.FItemArray(35)
				overseasDeliveryPrice	= iLine.FItemArray(36)
				overseasRealPrice		= tmpOverseasRealprice

                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
			End If

            IF (sellsite="bandinlunis") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                iLine.FItemArray(18) = rtSellPrice
                iLine.FItemArray(19) = rtSellPrice

                iLine.FItemArray(1) = Trim(replace(iLine.FItemArray(1),"/","-"))    ''주문일
                iLine.FItemArray(5) = Trim(splitvalue(iLine.FItemArray(5),"(",0))   ''주문자

                if (iLine.FItemArray(6)="") then iLine.FItemArray(6)=iLine.FItemArray(7)
                if (iLine.FItemArray(10)="") then iLine.FItemArray(10)=iLine.FItemArray(11)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
                            iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            IF (sellsite="hottracks") or (SellSite="itsHottracks") then

            	'//상품명에 옵션명이 같이 들어가있슴.. 위치 계산해서 짤라냄
            	if instr(iLine.FItemArray(21),"(") > 0 then
            		iLine.FItemArray(21) = left(iLine.FItemArray(21), instr(iLine.FItemArray(21),"(")-2 )
            		iLine.FItemArray(22) = rtrim(replace(mid(iLine.FItemArray(22), instr(iLine.FItemArray(22),"(")+1 , 96 ),")",""))

            	'//옵션없음
            	else
            		iLine.FItemArray(21) = iLine.FItemArray(21)
            		iLine.FItemArray(22) = ""
            	end if

                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                iLine.FItemArray(18) = rtSellPrice
                iLine.FItemArray(19) = rtSellPrice

                iLine.FItemArray(1) = Trim(replace(iLine.FItemArray(1),"/","-"))    ''주문일
                iLine.FItemArray(5) = Trim(splitvalue(iLine.FItemArray(5),"(",0))   ''주문자

                if (iLine.FItemArray(6)="") then iLine.FItemArray(6)=iLine.FItemArray(7)
                if (iLine.FItemArray(10)="") then iLine.FItemArray(10)=iLine.FItemArray(11)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
                            iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

	        	'//우편번호 사이에 "-" 가 없을경우
	        	iLine.FItemArray(12)=Trim(iLine.FItemArray(12))
        		If Len(iLine.FItemArray(12)) <> 5 Then			'2015-10-13 14:26 김진영 수정..우편번호가 5자리가 아닐때 아래 IF문 실행
		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
		        	end if
		        End If
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            IF (sellsite="byulshopITS") or (SellSite="itsByulshop") then

            	'//상품명에 옵션명이 같이 들어가있슴.. 위치 계산해서 짤라냄
            	if instr(iLine.FItemArray(21),"(") > 0 then
            		iLine.FItemArray(21) = left(iLine.FItemArray(21), instr(iLine.FItemArray(21),"(")-1 )
            		iLine.FItemArray(22) = rtrim(replace(mid(iLine.FItemArray(22), instr(iLine.FItemArray(22),"(")+1 , 96 ),")",""))

            	'//옵션없음
            	else
            		iLine.FItemArray(21) = iLine.FItemArray(21)
            		iLine.FItemArray(22) = ""
            	end if

                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                iLine.FItemArray(1) = Trim(left(iLine.FItemArray(1),10))    ''주문일

                if (iLine.FItemArray(6)="") then iLine.FItemArray(6)=iLine.FItemArray(7)
                if (iLine.FItemArray(10)="") then iLine.FItemArray(10)=iLine.FItemArray(11)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
                            iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            IF (sellsite="itsMemebox") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice
                iLine.FItemArray(18) = Trim(iLine.FItemArray(18))
                iLine.FItemArray(17) = Trim(iLine.FItemArray(17))
                iLine.FItemArray(19) = Trim(iLine.FItemArray(19))

                if (iLine.FItemArray(17)<>"") then
                    if (iLine.FItemArray(17)>1) then
                        iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)/iLine.FItemArray(17))      '''갯수로 나눠야 단가가 나옴.
                    end if
				end if

                iLine.FItemArray(19) = iLine.FItemArray(18)
                iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            IF (sellsite="ithinksoshop") then
				Dim preAddr, nextAddr
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                iLine.FItemArray(1) = Trim(left(iLine.FItemArray(1),10))
                iLine.FItemArray(2) = Trim(left(iLine.FItemArray(2),10))

				'############ 2017-03-03 김진영 수정 ###############
                preAddr		= Trim(iLine.FItemArray(13))		'지번주소 전체
                nextAddr	= Trim(iLine.FItemArray(14))		'도로명주소 전체
				If (preAddr = "") AND (nextAddr <> "") Then		'만약 엑셀에 지번주소는 없고 도로명 주소만 있다면
					iLine.FItemArray(13) = nextAddr
					iLine.FItemArray(14) = nextAddr
				ElseIf (preAddr <> "") AND (nextAddr = "") Then	'만약 엑셀에 도로명 주소가 없고 지번 주소만 있다면
					iLine.FItemArray(13) = preAddr
					iLine.FItemArray(14) = preAddr
				End If
				'############ 2017-03-03 김진영 수정 끝 ###############

                if (iLine.FItemArray(6)="") then iLine.FItemArray(6)=iLine.FItemArray(7)
                if (iLine.FItemArray(10)="") then iLine.FItemArray(10)=iLine.FItemArray(11)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
                            iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="suhaITS") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice
                '/이니셜서비스일경우 등록자가 판매가격을 안넣는 경우에 소비자가로 대체
				if isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" then
					iLine.FItemArray(18) = rtSellPrice
				end if

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값. 상품코드가 매핑되어 있는경우
                    if (iLine.FItemArray(22)<>"단품-") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

			'cjmallITS 진영추가
            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
			IF (sellsite="cjmallITS")  or (SellSite="itsCjmall") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice
				iLine.FItemArray(6) = Replace(iLine.FItemArray(6),")","-")
				iLine.FItemArray(10)= Replace(iLine.FItemArray(10),")","-")
				iLine.FItemArray(11)= Replace(iLine.FItemArray(11),")","-")

'				Dim cjAddr1, cjAddr2
'				cjAddr1 = Split(iLine.FItemArray(13),"|")(0)
'				cjAddr2 = Split(iLine.FItemArray(13),"|")(1)
'				iLine.FItemArray(13) = cjAddr1
'				iLine.FItemArray(14) = cjAddr2
                iLine.FItemArray(1) = Trim(replace(iLine.FItemArray(1),"/","-"))    ''주문일

                if (iLine.FItemArray(10)="") then iLine.FItemArray(10)=iLine.FItemArray(11)
                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
                            iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="gmarket") then
				'########################################### 2015-07-03 일시적 사용 ###########################################
				sItemname = Split(iLine.FItemArray(22),".")(1)
				sItemname = Split(Trim(sItemname),":")(0)
                iLine.FItemArray(1) = left(iLine.FItemArray(1),10)
                iLine.FItemArray(2) = left(iLine.FItemArray(2),10)
				iLine.FItemArray(18) = 4900			'판매가
				iLine.FItemArray(19) = 4900			'실판매가
				iLine.FItemArray(21) = sItemname	'상품명
         		iLine.FItemArray(22) = mid(iLine.FItemArray(22),1,instr(iLine.FItemArray(22),"/")-1)
        		iLine.FItemArray(22) = mid(iLine.FItemArray(22),instr(iLine.FItemArray(22),":")+1,100)
				iLine.FItemArray(22) = left(iLine.FItemArray(22),2)		'옵션명

				Select Case sItemname
					Case "YELLOW"			iLine.FItemArray(15) = 849958
					Case "PINK"				iLine.FItemArray(15) = 849956
					Case "PASTEL PINK"		iLine.FItemArray(15) = 849954
					Case "PASTEL BLUE"		iLine.FItemArray(15) = 849949
					Case "ORANGE"			iLine.FItemArray(15) = 849948
					Case "GREEN"			iLine.FItemArray(15) = 849947
				End Select
				iLine.FItemArray(16) = getOptionCodByOptionNameGSShop(iLine.FItemArray(15),iLine.FItemArray(22))
                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""
				if iLine.FItemArray(28) <> "" then
					iLine.FItemArray(28) = replace(iLine.FItemArray(28),",","")
				end if
				'########################################### 2015-07-03 일시적 끝 ###########################################

'            	'//옵션없는경우
'            	if instr(iLine.FItemArray(22),":") = "" or instr(iLine.FItemArray(22),":") = "0" then
'            		iLine.FItemArray(22) = ""
'
'            	'//옵션있는경우 : 와 / 사이에 옵션을 추려냄
'            	else
'            		iLine.FItemArray(22) = mid(iLine.FItemArray(22),1,instr(iLine.FItemArray(22),"/")-1)
'            		iLine.FItemArray(22) = mid(iLine.FItemArray(22),instr(iLine.FItemArray(22),":")+1,100)
'            	end if
'
'            	'//옵션없는경우	수량처리
'            	if instr(iLine.FItemArray(17),":") = "" or instr(iLine.FItemArray(17),":") = "0" then
'            		iLine.FItemArray(17) = left(iLine.FItemArray(17), instr(iLine.FItemArray(17),"개")-1)
'
'            	'//옵션있는경우 수량처리 / 와 개 사이에 수량을 추려냄
'            	else
'            		iLine.FItemArray(17) = mid(iLine.FItemArray(17), instr(iLine.FItemArray(17),"/")+1,100)
'            		iLine.FItemArray(17) = left(iLine.FItemArray(17), len(iLine.FItemArray(17))-1)
'            		'iLine.FItemArray(17) = mid(iLine.FItemArray(17), instr(iLine.FItemArray(17),"/")+1 , (len(iLine.FItemArray(17))-instr(iLine.FItemArray(17),"/"))-1  )
'            	end if
'
'                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
'                iLine.FItemArray(1) = left(iLine.FItemArray(1),10)
'                iLine.FItemArray(2) = left(iLine.FItemArray(2),10)
'                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
'                iLine.FItemArray(15) = rtitemid
'                iLine.FItemArray(16) = rtitemoption
'                '//엑셀내에 존재함
'                'iLine.FItemArray(18) = rtSellPrice
'                'iLine.FItemArray(19) = rtSellPrice
'
'                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값. 상품코드가 매핑되어 있는경우
'                    if (iLine.FItemArray(22)<>"단품-") then
'                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
'                     else
'                        iLine.FItemArray(16)="0000"
'                     end if
'                end if
'
'                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""
'
'				if iLine.FItemArray(28) <> "" then
'					iLine.FItemArray(28) = replace(iLine.FItemArray(28),",","")
'				end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            IF (sellsite="mintstore") or (SellSite="itsMintstore") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                iLine.FItemArray(18) = rtSellPrice
                iLine.FItemArray(19) = rtSellPrice

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
                            iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="wizwid") or (SellSite="itsWizwid") then
            	iLine.FItemArray(0) = Trim(iLine.FItemArray(0))
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                iLine.FItemArray(23) = Trim(iLine.FItemArray(23))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption

                '//엑셀내에 존재함
                iLine.FItemArray(18) = rtSellPrice			'2017-05-11 김진영 주석 해제
                iLine.FItemArray(19) = rtSellPrice			'2017-05-11 김진영 주석 해제
                '/이니셜서비스일경우 등록자가 판매가격을 안넣는 경우에 소비자가로 대체
				if isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" then
					iLine.FItemArray(18) = rtSellPrice
				end if

                iLine.FItemArray(1) = Left(iLine.FItemArray(1),4)&"-"&Mid(iLine.FItemArray(1),5,2)&"-"&Mid(iLine.FItemArray(1),7,2)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값. 상품코드가 매핑되어 있는경우
                    if (iLine.FItemArray(22)<>"단품-") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="hanatour") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                iLine.FItemArray(1) = Left(iLine.FItemArray(1),4)&"-"&Mid(iLine.FItemArray(1),6,2)&"-"&Mid(iLine.FItemArray(1),9,2)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값. 상품코드가 매핑되어 있는경우
                    if (iLine.FItemArray(22)<>"단품-") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="fashionplus") or (SellSite="itsFashionplus") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                iLine.FItemArray(1) = "20" & Left(iLine.FItemArray(1),2)&"-"&Mid(iLine.FItemArray(1),4,2)&"-"&Mid(iLine.FItemArray(1),7,2)
				iLine.FItemArray(2) = "20" & Left(iLine.FItemArray(2),2)&"-"&Mid(iLine.FItemArray(2),4,2)&"-"&Mid(iLine.FItemArray(2),7,2)
                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값. 상품코드가 매핑되어 있는경우
                    if (iLine.FItemArray(22)<>"단품-") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""

                iLine.FItemArray(25) = iLine.FItemArray(0)

				'//주문번호가 - 로 나뉘면 -이전까지 짤라내서 주문번호로 입력
                if instr(iLine.FItemArray(0),"-") > 0 then
                	iLine.FItemArray(0) = left(iLine.FItemArray(0),instr(iLine.FItemArray(0),"-")-1)
                end if

				if iLine.FItemArray(28) <> "" then
					iLine.FItemArray(28) = replace(iLine.FItemArray(28),",","")
				end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="giftting") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                'CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)



                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice
				iLine.FItemArray(1) = Left(iLine.FItemArray(0),4)&"-"&Mid(iLine.FItemArray(0),5,2)&"-"&Mid(iLine.FItemArray(0),7,2)
				iLine.FItemArray(2) = Left(iLine.FItemArray(0),4)&"-"&Mid(iLine.FItemArray(0),5,2)&"-"&Mid(iLine.FItemArray(0),7,2)
				'iLine.FItemArray(2) = "20" & Left(iLine.FItemArray(2),2)&"-"&Mid(iLine.FItemArray(2),4,2)&"-"&Mid(iLine.FItemArray(2),7,2)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값. 상품코드가 매핑되어 있는경우
                   if (iLine.FItemArray(22)<>"") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

                if (iLine.FItemArray(16)="") then iLine.FItemArray(16)="0000"

                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""

				if iLine.FItemArray(28) <> "" then
					iLine.FItemArray(28) = replace(iLine.FItemArray(28),",","")
				end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="gsisuper") then
				iLine.FItemArray(1) = LEFT(iLine.FItemArray(1), 10)
				iLine.FItemArray(2) = LEFT(iLine.FItemArray(2), 10)
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
	        	iLine.FItemArray(12)=Trim(iLine.FItemArray(12))
        		If Len(iLine.FItemArray(12)) <> 5 Then
		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
		        	end if
		        End If

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="itsbenepia") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
	        	iLine.FItemArray(12)=Trim(iLine.FItemArray(12))
        		If Len(iLine.FItemArray(12)) <> 5 Then
		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
		        	end if
		        End If

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="itsKaKaoMakers") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
	        	iLine.FItemArray(12)=Trim(iLine.FItemArray(12))
        		If Len(iLine.FItemArray(12)) <> 5 Then
		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
		        	end if
		        End If

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="itskakaotalkstore") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
            	iLine.FItemArray(1) = Trim(left(iLine.FItemArray(1), 10))
            	iLine.FItemArray(2) = Trim(left(iLine.FItemArray(2), 10))
 				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))

                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
				'iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)/iLine.FItemArray(17))			'판매금액이 수량 합한 금액이 들어옴
				iLine.FItemArray(18) = rtSellPrice
				iLine.FItemArray(12)=Trim(iLine.FItemArray(12))

            	If (isnumeric(iLine.FItemArray(19))) Then
           			iLine.FItemArray(19) = CLNG(iLine.FItemArray(18)) - CLNG(iLine.FItemArray(19))
            	Else
            		iLine.FItemArray(19) = iLine.FItemArray(18)
            	End If

                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""

				if iLine.FItemArray(28) <> "" then
					iLine.FItemArray(28) = replace(iLine.FItemArray(28),",","")
				end if

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="itskakao") then
            	iLine.FItemArray(1) = Trim(left(iLine.FItemArray(1), 10))
            	iLine.FItemArray(2) = Trim(left(iLine.FItemArray(2), 10))
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
            	iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)/iLine.FItemArray(17))			'판매금액이 수량 합한 금액이 들어옴

            	If (isnumeric(iLine.FItemArray(19))) Then
            		If iLine.FItemArray(19) > 0 Then
            			iLine.FItemArray(19) = CLNG((iLine.FItemArray(18) - (iLine.FItemArray(19)) / iLine.FItemArray(17)))
            		End If
            	Else
            		iLine.FItemArray(19) = iLine.FItemArray(18)
            	End If

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
                    	iLine.FItemArray(22) = Trim(Split(iLine.FItemArray(22), ":")(1))
           				iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if


                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""

				if iLine.FItemArray(28) <> "" then
					iLine.FItemArray(28) = replace(iLine.FItemArray(28),",","")
				end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
			t_addDlvPrice=""
			t_deliverytype=""
			t_sellcash=""
			t_defaultFreeBeasongLimit=""
			t_defaultDeliverPay=""
            if (SellSite="kakaogift") then
				tempShippay = 0
				'2023-07-19 김진영 수정..
				iLine.FItemArray(5) = Trim(iLine.FItemArray(5))
				iLine.FItemArray(9) = Trim(iLine.FItemArray(9))
				If Len(iLine.FItemArray(5)) = 0 Then
					iLine.FItemArray(5) = "-"
					iLine.FItemArray(9) = "-"
				End If
            	iLine.FItemArray(1) = Trim(left(iLine.FItemArray(1), 10))
            	iLine.FItemArray(2) = Trim(left(iLine.FItemArray(2), 10))
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				iLine.FItemArray(27) = ""

				rtSellPrice = Clng(iLine.FItemArray(20))
                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값. 상품코드가 매핑되어 있는경우
					iLine.FItemArray(16)="0000"
                end if

				If instr(iLine.FItemArray(22),":") > 0 Then
					iLine.FItemArray(22) = Trim(Split(iLine.FItemArray(22), ":")(1))
				End If

                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""
				
				sqlStr = ""
				sqlStr = sqlStr & " SELECT TOP 1 isNull(k.addDlvPrice, 0) as addDlvPrice, i.deliverytype, i.sellcash, tu.defaultFreeBeasongLimit, tu.defaultDeliverPay " 
				sqlStr = sqlStr & " FROM db_item.dbo.tbl_item i  " 
				sqlStr = sqlStr & " LEFT JOIN [db_etcmall].[dbo].tbl_kakaoGift_regItem as k on i.itemid = k.itemid " 
				sqlStr = sqlStr & " JOIN [db_user].[dbo].tbl_user_c  as tu on i.makerid= tu.userid " 
				sqlStr = sqlStr & " WHERE i.itemid = '"& Trim(iLine.FItemArray(15)) &"' " 
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) then
					t_addDlvPrice = CLng(rsget("addDlvPrice"))
					t_deliverytype = rsget("deliverytype")
					t_sellcash = CLng(rsget("sellcash"))
					t_defaultFreeBeasongLimit = CLng(rsget("defaultFreeBeasongLimit"))
					t_defaultDeliverPay = CLng(rsget("defaultDeliverPay"))
				Else
					response.write "<script>alert('[입점제휴]제휴몰관리>>카카오기프트에 저장된 상품이 아닙니다.');</script>"
					response.end
				End If
				rsget.Close

				If t_addDlvPrice > 0 Then
					tempShippay = t_addDlvPrice * iLine.FItemArray(17)
				Else
					If t_deliverytype = "1" and t_sellcash < 30000 Then
						tempShippay = 2500 * iLine.FItemArray(17)
					ElseIf t_deliverytype = "9" and t_sellcash < t_defaultFreeBeasongLimit Then
						tempShippay = t_defaultDeliverPay * iLine.FItemArray(17)
					End If
				End If

				iLine.FItemArray(18) = (rtSellPrice - CLng(tempShippay)) / iLine.FItemArray(17)
				iLine.FItemArray(19) = (rtSellPrice - CLng(tempShippay)) / iLine.FItemArray(17)
				iLine.FItemArray(28) = CLng(tempShippay)
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="GVG") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
'				If chr(asc(iLine.FItemArray(6))) = "?" Then
'					iLine.FItemArray(6) = replace(iLine.FItemArray(6),LEFT(iLine.FItemArray(6),1),"")
'				End If
'				If chr(asc(iLine.FItemArray(7))) = "?" Then
'					iLine.FItemArray(7) = replace(iLine.FItemArray(7),LEFT(iLine.FItemArray(7),1),"")
'				End If
'				If chr(asc(iLine.FItemArray(10))) = "?" Then
'					iLine.FItemArray(10) = replace(iLine.FItemArray(10),LEFT(iLine.FItemArray(10),1),"")
'				End If
'				If chr(asc(iLine.FItemArray(11))) = "?" Then
'					iLine.FItemArray(11) = replace(iLine.FItemArray(11),LEFT(iLine.FItemArray(11),1),"")
'				End If

                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="11stITS") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="11st1010") then
            	rtitemoption = iLine.FItemArray(22)
            	beasongNum11st	= iLine.FItemArray(29)
            	Dim vMemo, rtrealSellPrice
	            iLine.FItemArray(1) = LEFT(Replace(iLine.FItemArray(1),"/","-"), 10)
	            iLine.FItemArray(2) = LEFT(Replace(iLine.FItemArray(2),"/","-"), 10)

				rtSellPrice = Clng(iLine.FItemArray(18)) + Clng(Clng(iLine.FItemArray(32)) / iLine.FItemArray(17))		'2017-06-30 김진영..옵션추가금액을 수량으로 나눔
				iLine.FItemArray(18) = rtSellPrice

				rtrealSellPrice = rtSellPrice - Clng((Clng(iLine.FItemArray(19))+Clng(iLine.FItemArray(31))) / iLine.FItemArray(17))
				iLine.FItemArray(19) = rtrealSellPrice

				iLine.FItemArray(16) = getOptionCodByOptionName11st(iLine.FItemArray(15), rtitemoption, vMemo)
				If vMemo <> "" then
					iLine.FItemArray(27) = vMemo
				End If
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="lotteimall") then
'                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
				Dim isOptAddLtimall
				sqlStr = ""
				sqlStr = sqlStr & " SELECT TOP 1 itemid, itemoption"
				sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] "
				sqlStr = sqlStr & " WHERE convert(varchar(20),itemid) + convert(varchar(20),itemoption) = '"&iLine.FItemArray(15)&"' "
				sqlStr = sqlStr & " and mallid = 'lotteimall' "
				rsget.Open sqlStr,dbget,1
				If (Not rsget.EOF) Then
					isOptAddLtimall = "Y"
					iLine.FItemArray(15) = rsget("itemid")
					iLine.FItemArray(16) = rsget("itemoption")
				Else
					isOptAddLtimall = "N"
				End If
				rsget.Close

				'If iLine.FItemArray(0) = "2016-10-22-G97886" Then
				'	iLine.FItemArray(15) = "1254995"
				'	iLine.FItemArray(16) = "Z220"
				'End If

				If isOptAddLtimall = "N" Then
	                if (iLine.FItemArray(16)="") then
	                    if (iLine.FItemArray(22)<>"") then
	                        iLine.FItemArray(16)=getOptionCodByOptionNameimall(iLine.FItemArray(15),iLine.FItemArray(22))
	                     else
	                        iLine.FItemArray(16)="0000"
	                     end if
	                end if
	            End If
                iLine.FItemArray(0) = replace(iLine.FItemArray(0),"-","")
            end if

'            rtitemid=""
'            rtitemoption=""
'            rtSellPrice=""
'            if (SellSite="gseshop") then
'                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
'                iLine.FItemArray(15) = rtitemid
'                iLine.FItemArray(16) = rtitemoption
'                '//엑셀내에 존재함
'                'iLine.FItemArray(18) = rtSellPrice
'                'iLine.FItemArray(19) = rtSellPrice
'
'				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
'
'                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
'                    if (iLine.FItemArray(22)<>"") then
'           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
'                     else
'                        iLine.FItemArray(16)="0000"
'                     end if
'                end if
'            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="gseshop") then
            	Dim isOptAddGS
'                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
				sqlStr = ""
				sqlStr = sqlStr & " SELECT TOP 1 itemid, itemoption"
				sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] "
				sqlStr = sqlStr & " WHERE convert(varchar(20),itemid) + convert(varchar(20),itemoption) = '"&iLine.FItemArray(15)&"' "
				sqlStr = sqlStr & " and mallid = 'gsshop' "
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) then
					isOptAddGS = "Y"
					iLine.FItemArray(15) = rsget("itemid")
					iLine.FItemArray(16) = rsget("itemoption")
				Else
					isOptAddGS = "N"
				End If
				rsget.Close

				If isOptAddGS = "N" Then
	                if (iLine.FItemArray(16)="") then
	                    if (iLine.FItemArray(22)<>"") then
	                        iLine.FItemArray(16)=getOptionCodByOptionNameGSShop(iLine.FItemArray(15),iLine.FItemArray(22))
	                     else
	                        iLine.FItemArray(16)="0000"
	                     end if
	                end if
	            End If

				If iLine.FItemArray(5) <> "" Then
					iLine.FItemArray(5) = LEFT(iLine.FItemArray(5), 20)
				End If

				If iLine.FItemArray(9) <> "" Then
					iLine.FItemArray(9) = LEFT(iLine.FItemArray(9), 20)
				End If

                '''가격 관련 수량이 1개 이상일때  ''2014-03-17 추가
                if (iLine.FItemArray(17)<>"") then
                    if (iLine.FItemArray(17)>1) then
                        if (iLine.FItemArray(19)<>"") then
                            iLine.FItemArray(19) = CLNG(iLine.FItemArray(19)/iLine.FItemArray(17))
                        end if
                    end if
                end if

                ''배송비 2014-03-17 추가
                if (iLine.FItemArray(28)="선불") then
                    iLine.FItemArray(28)=3000
                else
                    iLine.FItemArray(28)=0
                end if

                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

				If Right(iLine.FItemArray(1), 1) = ":" Then
					iLine.FItemArray(1) = iLine.FItemArray(1) & "00"
				End If

'				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
            	iLine.FItemArray(1) = dateconvert(iLine.FItemArray(1))
            	iLine.FItemArray(2) = dateconvert(iLine.FItemArray(1))
            end if

			rtitemid=""
			rtitemoption=""
			rtSellPrice=""
			If (SellSite="nvstorefarm") Then
				Dim isDisCountYn
				iLine.FItemArray(18) = Clng(iLine.FItemArray(18))
				iLine.FItemArray(32) = Clng(iLine.FItemArray(32))
				iLine.FItemArray(30) = Clng(iLine.FItemArray(30))
				iLine.FItemArray(31) = Clng(iLine.FItemArray(31))

				If (iLine.FItemArray(16) = "") Then
					iLine.FItemArray(16) = "0000"
				End if

				If (iLine.FItemArray(17) <> "") then
					rtSellPrice 	= iLine.FItemArray(18) + iLine.FItemArray(32)							'판매가 = 판매가 + 옵션추가금액

					sqlStr = ""
					sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
					sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_outmall_mustPriceItem "
					sqlStr = sqlStr & " WHERE mallgubun = '"& SellSite &"' "
					sqlStr = sqlStr & " and itemid = '"& Trim(iLine.FItemArray(15)) &"' "
					sqlStr = sqlStr & " and '"& LEFT(iLine.FItemArray(2), 10) &"' between startDate and endDate "
					rsget.CursorLocation = adUseClient
					rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
						If rsget("cnt") > 0 Then
							isDisCountYn = "Y"
						Else
							isDisCountYn = "N"
						End If
					rsget.Close

					If isDisCountYn = "Y" Then
						iLine.FItemArray(18) = rtSellPrice - CLng(iLine.FItemArray(30) / iLine.FItemArray(17))
					Else
						iLine.FItemArray(18) = rtSellPrice
					End If
				End If
				iLine.FItemArray(19) = CLng(iLine.FItemArray(19) / iLine.FItemArray(17))
				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
				iLine.FItemArray(2) = Left(iLine.FItemArray(2),10)
			End If

			rtitemid=""
			rtitemoption=""
			rtSellPrice=""
			If (SellSite="nvstoregift") Then
				iLine.FItemArray(18) = Clng(iLine.FItemArray(18))
				iLine.FItemArray(32) = Clng(iLine.FItemArray(32))
				iLine.FItemArray(30) = Clng(iLine.FItemArray(30))
				iLine.FItemArray(31) = Clng(iLine.FItemArray(31))

				If (iLine.FItemArray(16) = "") Then
					iLine.FItemArray(16) = "0000"
				End if

				If (iLine.FItemArray(17) <> "") then
					rtSellPrice 	= iLine.FItemArray(18) + iLine.FItemArray(32)							'판매가 = 판매가 + 옵션추가금액
					iLine.FItemArray(18) = rtSellPrice
					iLine.FItemArray(19) = CLng(iLine.FItemArray(19) / iLine.FItemArray(17))
				End If
				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
				iLine.FItemArray(2) = Left(iLine.FItemArray(2),10)
			End If

			rtitemid=""
			rtitemoption=""
			rtSellPrice=""
			If (SellSite="nvstoremoonbangu") Then
				iLine.FItemArray(18) = Clng(iLine.FItemArray(18))
				iLine.FItemArray(32) = Clng(iLine.FItemArray(32))
				iLine.FItemArray(30) = Clng(iLine.FItemArray(30))
				iLine.FItemArray(31) = Clng(iLine.FItemArray(31))

				If (iLine.FItemArray(16) = "") Then
					iLine.FItemArray(16) = "0000"
				End if

				If (iLine.FItemArray(17) <> "") then
					rtSellPrice 	= iLine.FItemArray(18) + iLine.FItemArray(32)							'판매가 = 판매가 + 옵션추가금액
					iLine.FItemArray(18) = rtSellPrice
					iLine.FItemArray(19) = CLng(iLine.FItemArray(19) / iLine.FItemArray(17))
				End If
				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
				iLine.FItemArray(2) = Left(iLine.FItemArray(2),10)
			End If

			rtitemid=""
			rtitemoption=""
			rtSellPrice=""
			If (SellSite="Mylittlewhoopee") Then
				iLine.FItemArray(18) = Clng(iLine.FItemArray(18))
				iLine.FItemArray(32) = Clng(iLine.FItemArray(32))
				iLine.FItemArray(30) = Clng(iLine.FItemArray(30))
				iLine.FItemArray(31) = Clng(iLine.FItemArray(31))

				If (iLine.FItemArray(16) = "") Then
					iLine.FItemArray(16) = "0000"
				End if

				If (iLine.FItemArray(17) <> "") then
					rtSellPrice 	= iLine.FItemArray(18) + iLine.FItemArray(32)							'판매가 = 판매가 + 옵션추가금액
					iLine.FItemArray(18) = rtSellPrice
					iLine.FItemArray(19) = CLng(iLine.FItemArray(19) / iLine.FItemArray(17))
				End If
				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
				iLine.FItemArray(2) = Left(iLine.FItemArray(2),10)
			End If

			rtitemid=""
			rtitemoption=""
			rtSellPrice=""
			If (SellSite="nvstorefarmclass") Then
				iLine.FItemArray(18) = Clng(iLine.FItemArray(18))
				iLine.FItemArray(32) = Clng(iLine.FItemArray(32))
				iLine.FItemArray(30) = Clng(iLine.FItemArray(30))
				iLine.FItemArray(31) = Clng(iLine.FItemArray(31))

				iLine.FItemArray(16) = getOptionCodByOptionNameClass(iLine.FItemArray(15), iLine.FItemArray(22))
				If (iLine.FItemArray(17) <> "") then
					rtSellPrice 	= iLine.FItemArray(18) + iLine.FItemArray(32)							'판매가 = 판매가 + 옵션추가금액
					iLine.FItemArray(18) = rtSellPrice
					iLine.FItemArray(19) = CLng(iLine.FItemArray(19) / iLine.FItemArray(17))
				End If
				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
				iLine.FItemArray(2) = Left(iLine.FItemArray(2),10)
			End If

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="auction1010") OR (SellSite="gmarket1010") then
				'옥션 엑셀의 내용이 trim처리가 전혀 안 되어있어서 전부 trim처리
				iLine.FItemArray(0) = trim(iLine.FItemArray(0))
				iLine.FItemArray(1) = trim(iLine.FItemArray(1))
				iLine.FItemArray(2) = trim(iLine.FItemArray(2))
				iLine.FItemArray(3) = trim(iLine.FItemArray(3))
				iLine.FItemArray(4) = trim(iLine.FItemArray(4))
				iLine.FItemArray(5) = trim(iLine.FItemArray(5))
				iLine.FItemArray(6) = trim(iLine.FItemArray(6))
				iLine.FItemArray(7) = trim(iLine.FItemArray(7))
				iLine.FItemArray(8) = trim(iLine.FItemArray(8))
				iLine.FItemArray(9) = trim(iLine.FItemArray(9))
				iLine.FItemArray(10) = trim(iLine.FItemArray(10))
				iLine.FItemArray(11) = trim(iLine.FItemArray(11))
				iLine.FItemArray(12) = trim(iLine.FItemArray(12))
				iLine.FItemArray(13) = trim(iLine.FItemArray(13))
				iLine.FItemArray(14) = trim(iLine.FItemArray(14))
				iLine.FItemArray(15) = trim(iLine.FItemArray(15))
				iLine.FItemArray(16) = trim(iLine.FItemArray(16))
				iLine.FItemArray(17) = CLNG(trim(iLine.FItemArray(17)))
				iLine.FItemArray(18) = trim(iLine.FItemArray(18))
				iLine.FItemArray(19) = CLNG(trim(iLine.FItemArray(19)))
				iLine.FItemArray(20) = CLNG(trim(iLine.FItemArray(20)))
				iLine.FItemArray(21) = trim(iLine.FItemArray(21))
				iLine.FItemArray(22) = trim(iLine.FItemArray(22))
				iLine.FItemArray(23) = trim(iLine.FItemArray(23))
				iLine.FItemArray(24) = trim(iLine.FItemArray(24))
				iLine.FItemArray(25) = trim(iLine.FItemArray(25))
				iLine.FItemArray(26) = trim(iLine.FItemArray(26))
				iLine.FItemArray(27) = trim(iLine.FItemArray(27))
				iLine.FItemArray(28) = trim(iLine.FItemArray(28))
				iLine.FItemArray(29) = trim(iLine.FItemArray(29))

                iLine.FItemArray(1) = left(iLine.FItemArray(1),10)
                iLine.FItemArray(2) = left(iLine.FItemArray(2),10)

                iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)/iLine.FItemArray(17))			'판매금액이 수량 합한 금액이 들어옴
                'iLine.FItemArray(19) 계산식(realsellprice)이 판매가 -(쿠폰가 / 수량)이어야 하나
				'(판매가 - 쿠폰가)/수량으로 계산되었음..						'2015-10-06 15시 경 김진영 수정
				'(서비스이용료+정산예정금액)/수량 = 실결제금액=realsellprice	'2015-10-07 17:45분 김진영 수정
                'iLine.FItemArray(19) = CLNG(iLine.FItemArray(18) - (iLine.FItemArray(19) / iLine.FItemArray(17)))		'ver) 판매가 - (쿠폰가 / 수량)
                iLine.FItemArray(19) = CLNG((iLine.FItemArray(19) + iLine.FItemArray(20)) / iLine.FItemArray(17))		'ver) (서비스이용료 + 정산예정금액) / 수량
				If LEFT(iLine.FItemArray(21), 4) = "텐바이텐" Then
					iLine.FItemArray(21) = Trim(replace(iLine.FItemArray(21), LEFT(iLine.FItemArray(21), 4), ""))
				End If

				If (iLine.FItemArray(23) = "B329873664") OR (iLine.FItemArray(23) = "B291782397")  Then				'특정 상품이면
					Dim spItemname
					iLine.FItemArray(16) = "0000"
	         		spItemname = mid(iLine.FItemArray(22),1,instr(iLine.FItemArray(22),"/")-1)
	        		spItemname = mid(spItemname,instr(spItemname,":")+1,100)
					iLine.FItemArray(22) = ""
					If iLine.FItemArray(23) = "B329873664" Then
						Select Case spItemname
							Case "애쉬"
								iLine.FItemArray(15) = 1443173
								iLine.FItemArray(21) = "퍼피웨건 애견유모차 클래식 - 애쉬"
							Case "캔디핑크"
								iLine.FItemArray(15) = 1444046
								iLine.FItemArray(21) = "퍼피웨건 애견유모차 클래식 - 캔디핑크"
							Case "다크그레이"
								iLine.FItemArray(15) = 1444047
								iLine.FItemArray(21) = "퍼피웨건 애견유모차 클래식 - 다크그레이"
						End Select
					ElseIf iLine.FItemArray(23) = "B291782397" Then
						Select Case spItemname
							Case "2016 탁상 달력"
								iLine.FItemArray(15) = 1401873
								iLine.FItemArray(21) = "[텐바이텐X응답하라1988] 2016 탁상 달력"
							Case "딱지 스티커"
								iLine.FItemArray(15) = 1401875
								iLine.FItemArray(21) = "[텐바이텐X응답하라1988] 딱지 스티커"
							Case "청춘 노트"
								iLine.FItemArray(15) = 1401877
								iLine.FItemArray(21) = "[텐바이텐X응답하라1988] 청춘시대 노트"
						End Select
					End If

	                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""
				Else													'기존 주문이면
	                iLine.FItemArray(16)=getOptionCodByOptionNameAuction(iLine.FItemArray(15),iLine.FItemArray(22), iLine.FItemArray(0))
	                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""
					If instr(iLine.FItemArray(22),"텍스트를 입력하세요") > 0 Then
'						Dim madeText
'						madeText = mid(iLine.FItemArray(22),instr(iLine.FItemArray(22),"텍스트를 입력하세요")+1,1000)

'						If instr(madeText, "텍스트를 입력하세요") > 0 Then
'							iLine.FItemArray(27) = Trim(Split(madeText, "：")(1))
'						Else
'							iLine.FItemArray(27) = ""
'						End If
						If instr(iLine.FItemArray(22),"텍스트를 입력하세요：") > 0 Then
							iLine.FItemArray(27) = Trim(Split(iLine.FItemArray(22), "텍스트를 입력하세요：")(1))
						ElseIf instr(iLine.FItemArray(22),"텍스트를 입력하세요:") > 0 Then
							iLine.FItemArray(27) = Trim(Split(iLine.FItemArray(22), "텍스트를 입력하세요:")(1))
						End If
					Else
						iLine.FItemArray(27) = ""
					End If
				End If

				if iLine.FItemArray(28) <> "" then
					iLine.FItemArray(28) = replace(iLine.FItemArray(28),",","")
				end if

'                iLine.FItemArray(16)=getOptionCodByOptionNameAuction(iLine.FItemArray(15),iLine.FItemArray(22), iLine.FItemArray(0))
'
'                if (iLine.FItemArray(26)="null") then iLine.FItemArray(26)=""
'
'				If instr(iLine.FItemArray(22),"텍스트를 입력하세요") > 0 Then
'					Dim madeText
'					madeText = mid(iLine.FItemArray(22),instr(iLine.FItemArray(22),"텍스트를 입력하세요")-1,1000)
'					If instr(madeText, "텍스트를 입력하세요") > 0 Then
'						iLine.FItemArray(27) = Trim(Split(madeText, "：")(1))
'					Else
'						iLine.FItemArray(27) = ""
'					End If
'				Else
'					iLine.FItemArray(27) = ""
'				End If
'
'				if iLine.FItemArray(28) <> "" then
'					iLine.FItemArray(28) = replace(iLine.FItemArray(28),",","")
'				end if
            end if

			rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="ezwel") then
            	iLine.FItemArray(1) = Trim(left(iLine.FItemArray(1),10))    ''주문일

				If LEFT(iLine.FItemArray(22), 3) = "선택:" Then
					iLine.FItemArray(22) = replace(iLine.FItemArray(22), LEFT(iLine.FItemArray(22), 3), "")
				End If

				If Right(iLine.FItemArray(22), 1) = "^" Then
					iLine.FItemArray(22) = replace(iLine.FItemArray(22), Right(iLine.FItemArray(22), 1), "")
				End If

'				'2017-07-19 김진영 추가..realsellprice관련..엑셀에 할인필드 있음..2017-07-24 김진영 할인한 값이 넘어옴;;
'				rtrealSellPrice = Clng(iLine.FItemArray(18)) - Clng(Clng(iLine.FItemArray(19)) / iLine.FItemArray(17))
'				iLine.FItemArray(19) = rtrealSellPrice

				sqlStr = ""
				sqlStr = sqlStr & " SELECT TOP 1 itemid FROM db_etcmall.dbo.tbl_ezwel_regitem WHERE ezwelgoodno = '"&iLine.FItemArray(23)&"' "
				rsget.Open sqlStr,dbget,1
				If (Not rsget.EOF) Then
					iLine.FItemArray(15) = rsget("itemid")
				End If
				rsget.Close

                if (iLine.FItemArray(16)="") then
                    if (iLine.FItemArray(22)<>"") then
                        iLine.FItemArray(16)=getOptionCodByOptionNameGSShop(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            If (SellSite="GS25") Then
                tmppreAddr		= Trim(iLine.FItemArray(13))			'지번주소 전체
                tmpnextAddr		= Trim(iLine.FItemArray(14))			'도로명주소 전체
				If (tmppreAddr = "") AND (tmpnextAddr <> "") Then		'만약 엑셀에 지번주소는 없고 도로명 주소만 있다면
					iLine.FItemArray(13) = tmpnextAddr
					iLine.FItemArray(14) = tmpnextAddr
				ElseIf (tmppreAddr <> "") AND (tmpnextAddr = "") Then	'만약 엑셀에 도로명 주소가 없고 지번 주소만 있다면
					iLine.FItemArray(13) = tmppreAddr
					iLine.FItemArray(14) = tmppreAddr
				ElseIf (tmppreAddr <> "") AND (tmpnextAddr <> "") Then
					iLine.FItemArray(13) = tmpnextAddr
					iLine.FItemArray(14) = tmpnextAddr
				End If

				'Select Case iLine.FItemArray(23)
				' 	Case "2800100203602"
				' 		iLine.FItemArray(15) = "3313868"
				' 		iLine.FItemArray(18) = "29000"
				' 	Case "2800100204449"
				' 		iLine.FItemArray(15) = "3471382"
				' 		iLine.FItemArray(18) = "45000"
				' 	Case "2800100204456"
				' 		iLine.FItemArray(15) = "4524679"
				' 		iLine.FItemArray(18) = "35000"
				' 	Case "2800100204463"
				' 		iLine.FItemArray(15) = "4890940"
				' 		iLine.FItemArray(18) = "55000"
				' 	Case "2800100204487"
				' 		iLine.FItemArray(15) = "4509495"
				' 		iLine.FItemArray(18) = "35000"
				' 	Case "2800100204494"
				' 		iLine.FItemArray(15) = "4509498"
				' 		iLine.FItemArray(18) = "35000"
				' 	Case "2800100204500"
				' 		iLine.FItemArray(15) = "3504305"
				' 		iLine.FItemArray(18) = "39000"
				' 	Case "2800100204517"
				' 		iLine.FItemArray(15) = "4728736"
				' 		iLine.FItemArray(18) = "29000"
				' End Select
'				iLine.FItemArray(16) = "0000"	'옵션코드

				Select Case iLine.FItemArray(23)
					Case "2800100218279"
						iLine.FItemArray(15) = "4568721"
						iLine.FItemArray(16) = "0015"
						iLine.FItemArray(18) = "32000"
					Case "2800100218286"
						iLine.FItemArray(15) = "4568721"
						iLine.FItemArray(16) = "0014"
						iLine.FItemArray(18) = "32000"
					Case "2800100218293"
						iLine.FItemArray(15) = "4495213"
						iLine.FItemArray(16) = "0011"
						iLine.FItemArray(18) = "22000"
					Case "2800100218309"
						iLine.FItemArray(15) = "3581268"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "14000"
					Case "2800100218316"
						iLine.FItemArray(15) = "4504295"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "11000"
					Case "2800100218323"
						iLine.FItemArray(15) = "3471386"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "17000"

					Case "2800100218330"
						iLine.FItemArray(15) = "5683080"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "22000"
					Case "2800100218347"
						iLine.FItemArray(15) = "5683121"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "45000"
					Case "2800100218354"
						iLine.FItemArray(15) = "5683124"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "45000"
					Case "2800100218361"
						iLine.FItemArray(15) = "5683125"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "35000"
					Case "2800100218378"
						iLine.FItemArray(15) = "5683126"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "43000"

					'2023-10-23 김진영 하단 코드 추가
					Case "2840000121644"
						iLine.FItemArray(15) = "4524590"
						iLine.FItemArray(16) = "0011"
						iLine.FItemArray(18) = "22000"
					Case "2840000121651"
						iLine.FItemArray(15) = "4524590"
						iLine.FItemArray(16) = "0012"
						iLine.FItemArray(18) = "22000"
					Case "2840000121675"
						iLine.FItemArray(15) = "4524590"
						iLine.FItemArray(16) = "0015"
						iLine.FItemArray(18) = "22000"
					Case "2840000121712"
						iLine.FItemArray(15) = "5661015"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "25000"
					Case "2840000121620"
						iLine.FItemArray(15) = "4568721"
						iLine.FItemArray(16) = "0013"
						iLine.FItemArray(18) = "54000"
					Case "2840000121521"
						iLine.FItemArray(15) = "5014914"
						iLine.FItemArray(16) = "0012"
						iLine.FItemArray(18) = "15000"
					Case "2840000121514"
						iLine.FItemArray(15) = "5014914"
						iLine.FItemArray(16) = "0013"
						iLine.FItemArray(18) = "15000"
					Case "2840000121507"
						iLine.FItemArray(15) = "4546794"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "29000"
					Case "2840000121569"
						iLine.FItemArray(15) = "5616003"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "8000"
					Case "2840000121538"
						iLine.FItemArray(15) = "5616004"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "15000"
					Case "2840000121743"
						iLine.FItemArray(15) = "5109313"
						iLine.FItemArray(16) = "0011"
						iLine.FItemArray(18) = "15000"
					Case "2840000121699"
						iLine.FItemArray(15) = "5109313"
						iLine.FItemArray(16) = "0013"
						iLine.FItemArray(18) = "15000"
					Case "2840000121583"
						iLine.FItemArray(15) = "4958612"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "49000"
					Case "2840000121668"
						iLine.FItemArray(15) = "5415524"
						iLine.FItemArray(16) = "0011"
						iLine.FItemArray(18) = "12000"
					Case "2840000121637"
						iLine.FItemArray(15) = "5415524"
						iLine.FItemArray(16) = "0012"
						iLine.FItemArray(18) = "12000"
					Case "2840000121613"
						iLine.FItemArray(15) = "5415524"
						iLine.FItemArray(16) = "0013"
						iLine.FItemArray(18) = "12000"
					Case "2840000121590"
						iLine.FItemArray(15) = "5415524"
						iLine.FItemArray(16) = "0014"
						iLine.FItemArray(18) = "12000"
					Case "2840000121705"
						iLine.FItemArray(15) = "3581268"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(18) = "24000"
				End Select
				iLine.FItemArray(19) = iLine.FItemArray(18)		'실판매가
				iLine.FItemArray(28) = "0"		'배송비
				iLine.FItemArray(26) = ""		'deliverymemo
            End If

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="homeplus") then
'                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                if (iLine.FItemArray(16)="") then
                    if (iLine.FItemArray(22)<>"") then
                        iLine.FItemArray(16)=getOptionCodByOptionNameGSShop(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

                If iLine.FItemArray(19) < 30000 Then
                	iLine.FItemArray(28)="2500"
                Else
                	iLine.FItemArray(28)="0"
                End If

                '''가격 관련 수량이 1개 이상일때  ''2014-03-17 추가
                if (iLine.FItemArray(17)<>"") then
                    if (iLine.FItemArray(17)>1) then
                        if (iLine.FItemArray(19)<>"") then
                            iLine.FItemArray(19) = CLNG(iLine.FItemArray(19)/iLine.FItemArray(17))
                        end if
                    end if
                end if
            end if

'            rtitemid=""
'            rtitemoption=""
'            rtSellPrice=""
'            if (SellSite="cjmall") then
'                if (iLine.FItemArray(16)="") then
'                    if (iLine.FItemArray(22)<>"") then
'                        iLine.FItemArray(16)=getOptionCodByOptionNameGSShop(iLine.FItemArray(15),iLine.FItemArray(22))
'                     else
'                        iLine.FItemArray(16)="0000"
'                     end if
'                end if
'				iLine.FItemArray(6) = Replace(iLine.FItemArray(6),")","-")
'				iLine.FItemArray(10)= Replace(iLine.FItemArray(10),")","-")
'				iLine.FItemArray(11)= Replace(iLine.FItemArray(11),")","-")
'
'				Dim cjmAddr1, cjmAddr2
'				cjmAddr1 = Split(iLine.FItemArray(13),"|")(0)
'				cjmAddr2 = Split(iLine.FItemArray(13),"|")(1)
'				iLine.FItemArray(13) = cjmAddr1
'				iLine.FItemArray(14) = cjmAddr2
'                iLine.FItemArray(1) = Trim(replace(iLine.FItemArray(1),"/","-"))    ''주문일
' 				if (iLine.FItemArray(10)="") then iLine.FItemArray(10)=iLine.FItemArray(11)
'            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
			IF (sellsite="cjmall") then
				'공통양식으로 변경 전..
                ' If instr(iLine.FItemArray(0),"-") > 0 Then
                ' 	iLine.FItemArray(0) = Split(iLine.FItemArray(0), "-")(0)
                ' End If

				' Dim oTenId
				' oTenId = iLine.FItemArray(15)
				' iLine.FItemArray(15) = Split(oTenId, "_")(0)
				' iLine.FItemArray(16) = Split(oTenId, "_")(1)

				' iLine.FItemArray(6) = Replace(iLine.FItemArray(6),")","-")
				' iLine.FItemArray(10)= Replace(iLine.FItemArray(10),")","-")
				' iLine.FItemArray(11)= Replace(iLine.FItemArray(11),")","-")
                ' iLine.FItemArray(1) = LEFT(Trim(replace(iLine.FItemArray(1),"/","-")), 10)    ''주문일
				' iLine.FItemArray(19) = iLine.FItemArray(19) / iLine.FItemArray(17)

                ' if (iLine.FItemArray(10)="") then iLine.FItemArray(10)=iLine.FItemArray(11)
                ' if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                '     if (iLine.FItemArray(22)<>"") then
                '             iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                '      else
                '         iLine.FItemArray(16)="0000"
                '      end if
                ' end if

				'공통양식으로 변경 후..2023-10-19 김진영 수정
				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				If (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") Then       ''옵션코드가 빈값.
					iLine.FItemArray(16)="0000"
				End If
				isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
				If isValid = "N" Then
					response.write "<script>alert('10x10 상품 코드와 10x10 옵션코드가 매칭이 된 것이 맞는 지 다시 확인하세요.');</script>"
					response.end
				End If

				If isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" Then
					iLine.FItemArray(18) = rtSellPrice
				End If

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)

            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="NJOYNY") or (SellSite="itsNJOYNY") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="ticketmonster") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="halfclub") then
            	iLine.FItemArray(15) = Split(iLine.FItemArray(23), "_")(0)
                iLine.FItemArray(16) = getOptionCodByOptionNameHalfClub(Split(iLine.FItemArray(23), "_")(0), iLine.FItemArray(21))
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="thinkaboutyou") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

	        	'//우편번호 사이에 "-" 가 없을경우
	        	iLine.FItemArray(12)=Trim(iLine.FItemArray(12))
        		If Len(iLine.FItemArray(12)) <> 5 Then			'2015-10-13 14:26 김진영 수정..우편번호가 5자리가 아닐때 아래 IF문 실행
		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
		        	end if
		        End If

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if


            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
			tmpItemname = ""
            if (SellSite="aboutpet") then
				If iLine.FItemArray(16) <> "0000" Then
					If InStr(iLine.FItemArray(21), "_") > 0 Then
						tmpItemname = Trim(Split(iLine.FItemArray(21), "_")(0))
					Else
						tmpItemname = Trim(iLine.FItemArray(21))
					End If
				Else
					tmpItemname = Trim(iLine.FItemArray(21))
				End If
				beasongNum11st	= Replace(iLine.FItemArray(29), ",", "")

				sqlStr = ""
				sqlStr = sqlStr & " SELECT TOP 1 itemid "
				sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_aboutpet_regitem "
				sqlStr = sqlStr & " WHERE RTRIM(LTRIM(itemname)) = '"& tmpItemname &"' "
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				If (Not rsget.EOF) Then
					iLine.FItemArray(15) = rsget("itemid")
				End If
				rsget.Close

				iLine.FItemArray(19) = Clng(iLine.FItemArray(19)/iLine.FItemArray(17))

	        	'//우편번호 사이에 "-" 가 없을경우
	        	iLine.FItemArray(12)=Trim(iLine.FItemArray(12))
        		If Len(iLine.FItemArray(12)) <> 5 Then			'2015-10-13 14:26 김진영 수정..우편번호가 5자리가 아닐때 아래 IF문 실행
		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
		        	end if
		        End If
            end if


            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="cookatmall") then
				iLine.FItemArray(1) = Left(iLine.FItemArray(0),4) &"-"& mid(iLine.FItemArray(0),5,2) &"-"& mid(iLine.FItemArray(0),7,2)
				Select Case iLine.FItemArray(21)
					Case "피너츠 프렌즈 유리컵_찰리브라운"
						iLine.FItemArray(15) = "3649588"
						iLine.FItemArray(16) = "0012"
						iLine.FItemArray(23) = "BR130144"
					Case "피너츠 프렌즈 유리컵_스누피"
						iLine.FItemArray(15) = "3649588"
						iLine.FItemArray(16) = "0011"
						iLine.FItemArray(23) = "BR130143"
					Case "피너츠 스누피와 친구들 머그컵_라이너스"
						iLine.FItemArray(15) = "2785591"
						iLine.FItemArray(16) = "0016"
						iLine.FItemArray(23) = "BR130142"
					Case "피너츠 스누피와 친구들 머그컵_샐리"
						iLine.FItemArray(15) = "2785591"
						iLine.FItemArray(16) = "0015"
						iLine.FItemArray(23) = "BR130141"
					Case "피너츠 스누피와 친구들 머그컵_루시"
						iLine.FItemArray(15) = "2785591"
						iLine.FItemArray(16) = "0014"
						iLine.FItemArray(23) = "BR130140"
					Case "피너츠 스누피와 친구들 머그컵_우드스탁"
						iLine.FItemArray(15) = "2785591"
						iLine.FItemArray(16) = "0012"
						iLine.FItemArray(23) = "BR130139"
					Case "피너츠 스누피와 친구들 머그컵_찰리브라운"
						iLine.FItemArray(15) = "2785591"
						iLine.FItemArray(16) = "0013"
						iLine.FItemArray(23) = "BR130138"
					Case "피너츠 스누피와 친구들 머그컵_스누피"
						iLine.FItemArray(15) = "2785591"
						iLine.FItemArray(16) = "0011"
						iLine.FItemArray(23) = "BR130137"
					Case "피너츠 스누피 레트로 토스터기"
						iLine.FItemArray(15) = "3471382"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(23) = "BR130136"
					Case "피너츠 스누피 샌드위치&와플메이커"
						iLine.FItemArray(15) = "2784156"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(23) = "BR130135"
					Case "디즈니 해피쏘우트 푸 머그컵"
						iLine.FItemArray(15) = "3701715"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(23) = "BR130155"
					Case "디즈니 위니 더 푸 머그컵_푸"
						iLine.FItemArray(15) = "3616014"
						iLine.FItemArray(16) = "0011"
						iLine.FItemArray(23) = "BR130156"
					Case "디즈니 위니 더 푸 머그컵_피글렛"
						iLine.FItemArray(15) = "3616014"
						iLine.FItemArray(16) = "0012"
						iLine.FItemArray(23) = "BR130157"
					Case "디즈니 위니 더 푸 머그컵_티거"
						iLine.FItemArray(15) = "3616014"
						iLine.FItemArray(16) = "0013"
						iLine.FItemArray(23) = "BR130158"
					Case "디즈니 위니 더 푸 머그컵_이요르"
						iLine.FItemArray(15) = "3616014"
						iLine.FItemArray(16) = "0014"
						iLine.FItemArray(23) = "BR130159"
					Case "디즈니 곰돌이 푸 시리얼 통 (3개세트)"
						iLine.FItemArray(15) = "3646836"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(23) = "BR130153"
					Case "디즈니 곰돌이 푸 시리얼 볼 세트 (스푼+시리얼 볼)"
						iLine.FItemArray(15) = "3581268"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(23) = "BR130154"
					Case "피너츠 리유저블 콜드컵 세트 5ea"
						iLine.FItemArray(15) = "2849183"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(23) = "BR130183"
					Case "산리오 헬로키티 빙수기"
						iLine.FItemArray(15) = "3530000"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(23) = "BR130195"
					Case "스누피 레트로 컵세트(2인조)"
						iLine.FItemArray(15) = "3471386"
						iLine.FItemArray(16) = "0000"
						iLine.FItemArray(23) = "CR130052"
					Case "스누피 레트로 브런치 플레이트_스누피브레드"
						iLine.FItemArray(15) = "3471393"
						iLine.FItemArray(16) = "0011"
						iLine.FItemArray(23) = "CR130053"
					Case "스누피 레트로 브런치 플레이트_우드스탁브레드"
						iLine.FItemArray(15) = "3471393"
						iLine.FItemArray(16) = "0012"
						iLine.FItemArray(23) = "CR130054"
				End Select

	        	'//우편번호 사이에 "-" 가 없을경우
	        	iLine.FItemArray(12)=Trim(iLine.FItemArray(12))
        		If Len(iLine.FItemArray(12)) <> 5 Then			'2015-10-13 14:26 김진영 수정..우편번호가 5자리가 아닐때 아래 IF문 실행
		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
		        	end if
		        End If
            end if


            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="momQ") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption

	        	'//우편번호 사이에 "-" 가 없을경우
	        	iLine.FItemArray(12)=Trim(iLine.FItemArray(12))
        		If Len(iLine.FItemArray(12)) <> 5 Then			'2015-10-13 14:26 김진영 수정..우편번호가 5자리가 아닐때 아래 IF문 실행
		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
		        	end if
		        End If

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
'            if (SellSite="privia") then
'				'옵션명뽑기..옵션/각인으로 들어와서 /를 스플릿함
'				iLine.FItemArray(22) = split(iLine.FItemArray(22),"/")(0)
'                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
'                iLine.FItemArray(15) = rtitemid
'                iLine.FItemArray(16) = rtitemoption
'
'                '//엑셀내에 존재함
'                'iLine.FItemArray(18) = rtSellPrice
'                'iLine.FItemArray(19) = rtSellPrice
'                '주문일시가 이상하게 넘어와서 치환
'				iLine.FItemArray(1) = Left(iLine.FItemArray(0),4) &"-"& mid(iLine.FItemArray(0),5,2) &"-"& mid(iLine.FItemArray(0),7,2)
'				'주문자명 뽑기
'				iLine.FItemArray(5) = split(iLine.FItemArray(5),"(")(0)
'
'                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
'                    if (iLine.FItemArray(22)<>"") then
'           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
'                     else
'                        iLine.FItemArray(16)="0000"
'                     end if
'                end if
'            end if
'2013-11-15 16:52 김진영 수정해봄
            if (SellSite="privia") then
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption

				'주문자명 뽑기
				If instr(iLine.FItemArray(5),"(") > 0 Then
					iLine.FItemArray(5) = split(iLine.FItemArray(5),"(")(0)
				End If

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="momastore") then
            	iLine.FItemArray(1) = Trim(left(iLine.FItemArray(1),10))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption

				'주문자명 뽑기
				If instr(iLine.FItemArray(5),"(") > 0 Then
					iLine.FItemArray(5) = split(iLine.FItemArray(5),"(")(0)
				End If

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

       		rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="its29cm") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)

                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice
                '/이니셜서비스일경우 등록자가 판매가격을 안넣는 경우에 소비자가로 대체
				if isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" then
					iLine.FItemArray(18) = rtSellPrice
				end if

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="gabangpop") or (SellSite="itsGabangpop") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)

                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice
                '/이니셜서비스일경우 등록자가 판매가격을 안넣는 경우에 소비자가로 대체
				if isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" then
					iLine.FItemArray(18) = rtSellPrice
				end if

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="musinsaITS") or (SellSite="itsMusinsa") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
				iLine.FItemArray(22) = replace(iLine.FItemArray(22),"NONE","")
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="celectory") then
                iLine.FItemArray(1) = replace(iLine.FItemArray(1),"'","")
                iLine.FItemArray(6) = replace(iLine.FItemArray(6),"'","")
                iLine.FItemArray(7) = replace(iLine.FItemArray(7),"'","")
                iLine.FItemArray(10) = replace(iLine.FItemArray(10),"'","")
                iLine.FItemArray(11) = replace(iLine.FItemArray(11),"'","")
                iLine.FItemArray(12) = replace(iLine.FItemArray(12),"'","")
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                iLine.FItemArray(22) = Trim(iLine.FItemArray(22))
                iLine.FItemArray(23) = Trim(iLine.FItemArray(23))

                CALL getEtcSiteNameOrCode2ItemCode(sellsite, iLine.FItemArray(23), iLine.FItemArray(21), iLine.FItemArray(22), rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption

                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="player") or (SellSite="itsPlayer1") then
            	'//상품명에 옵션명이 같이 들어가있슴.. 위치 계산해서 짤라냄
            	if instr(iLine.FItemArray(21),"(") > 0 then
            		iLine.FItemArray(21) = left(iLine.FItemArray(21), instr(iLine.FItemArray(21),"(")-1 )
            		iLine.FItemArray(22) = rtrim(replace(mid(iLine.FItemArray(22), instr(iLine.FItemArray(22),"(")+1 , 96 ),")",""))

            	'//옵션없음
            	else
            		iLine.FItemArray(21) = iLine.FItemArray(21)
            		iLine.FItemArray(22) = ""
            	end if


                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

				'//판매일이 안넘와와서, 주문번호 앞자리 10자리를 판매일로 처리
                iLine.FItemArray(1) = Left(iLine.FItemArray(0),4) &"-"& mid(iLine.FItemArray(0),5,2) &"-"& mid(iLine.FItemArray(0),7,2)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

            if (SellSite="wconcept") or (SellSite="itsWconcept") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice
                '/이니셜서비스일경우 등록자가 판매가격을 안넣는 경우에 소비자가로 대체
				if isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" then
					iLine.FItemArray(18) = rtSellPrice
				end if

                iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
           				temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if
            end if

			If (SellSite="hmall1010") Then

				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				If (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") Then       ''옵션코드가 빈값.
					iLine.FItemArray(16)="0000"
				End If
				isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
				If isValid = "N" Then
					response.write "<script>alert('10x10 상품 코드와 10x10 옵션코드가 매칭이 된 것이 맞는 지 다시 확인하세요.');</script>"
					response.end
				End If

				If isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" Then
					iLine.FItemArray(18) = rtSellPrice
				End If

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
			End If


			If (SellSite="lotteon") Then
				iLine.FItemArray(25) = Split(iLine.FItemArray(25), "_")(1)
				If instr(iLine.FItemArray(27),"텍스트를 입력하세요 :") > 0 Then
					iLine.FItemArray(27) = Trim(Split(iLine.FItemArray(27), "텍스트를 입력하세요 :")(1))
				Else
					iLine.FItemArray(27) = ""
				End If

				sqlStr = ""
				sqlStr = sqlStr & " select top 1 itemid, itemoption "
				sqlStr = sqlStr & " from db_item.dbo.tbl_OutMall_regedoption "
				sqlStr = sqlStr & " where outmallOptCode = '"& iLine.FItemArray(24) &"' "
				sqlStr = sqlStr & " and mallid = 'lotteon' "
				rw sqlStr
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				If (Not rsget.EOF) Then
					iLine.FItemArray(15) = rsget("itemid")
					iLine.FItemArray(16) = rsget("itemoption")
				End If
				rsget.Close

				if iLine.FItemArray(18) * iLine.FItemArray(17)  > 50000 then
					iLine.FItemArray(28) = 0
				else
					iLine.FItemArray(28) = 3000
				end if
			' 	isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
			' 	If isValid = "N" Then
			' 		response.write "<script>alert('10x10 상품 코드와 10x10 옵션코드가 매칭이 된 것이 맞는 지 다시 확인하세요.');</script>"
			' 		response.end
			' 	End If
			 	iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
			End If

			If (SellSite="shintvshopping") Then

				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				If (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") Then       ''옵션코드가 빈값.
					iLine.FItemArray(16)="0000"
				End If
				isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
				If isValid = "N" Then
					response.write "<script>alert('10x10 상품 코드와 10x10 옵션코드가 매칭이 된 것이 맞는 지 다시 확인하세요.');</script>"
					response.end
				End If

				If isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" Then
					iLine.FItemArray(18) = rtSellPrice
				End If

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
			End If

			If (SellSite="skstoa") Then

				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				If (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") Then       ''옵션코드가 빈값.
					iLine.FItemArray(16)="0000"
				End If
				isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
				If isValid = "N" Then
					response.write "<script>alert('10x10 상품 코드와 10x10 옵션코드가 매칭이 된 것이 맞는 지 다시 확인하세요.');</script>"
					response.end
				End If

				If isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" Then
					iLine.FItemArray(18) = rtSellPrice
				End If

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
			End If

			If (SellSite="LFmall") Then

				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				If (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") Then       ''옵션코드가 빈값.
					iLine.FItemArray(16)="0000"
				End If
				isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
				If isValid = "N" Then
					response.write "<script>alert('10x10 상품 코드와 10x10 옵션코드가 매칭이 된 것이 맞는 지 다시 확인하세요.');</script>"
					response.end
				End If

				If isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" Then
					iLine.FItemArray(18) = rtSellPrice
				End If

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
			End If

			If (SellSite="goodwearmall10") Then

				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				If (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") Then       ''옵션코드가 빈값.
					iLine.FItemArray(16)="0000"
				End If
				isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
				If isValid = "N" Then
					response.write "<script>alert('10x10 상품 코드와 10x10 옵션코드가 매칭이 된 것이 맞는 지 다시 확인하세요.');</script>"
					response.end
				End If

				If isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" Then
					iLine.FItemArray(18) = rtSellPrice
				End If

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
			End If

			If (SellSite="wconcept1010") Then

				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				If (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") Then       ''옵션코드가 빈값.
					iLine.FItemArray(16)="0000"
				End If
				isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
				If isValid = "N" Then
					response.write "<script>alert('10x10 상품 코드와 10x10 옵션코드가 매칭이 된 것이 맞는 지 다시 확인하세요.');</script>"
					response.end
				End If

				If isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" Then
					iLine.FItemArray(18) = rtSellPrice
				End If

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
			End If

			If (SellSite="itsWadiz") Then

				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				If (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") Then       ''옵션코드가 빈값.
					iLine.FItemArray(16)="0000"
				End If
				isValid = getIsValidItemIdOption(iLine.FItemArray(15), iLine.FItemArray(16))
				If isValid = "N" Then
					response.write "<script>alert('10x10 상품 코드와 10x10 옵션코드가 매칭이 된 것이 맞는 지 다시 확인하세요');</script>"
					response.end
				End If

				If isnull(iLine.FItemArray(18)) or iLine.FItemArray(18)="" Then
					iLine.FItemArray(18) = rtSellPrice
				End If

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
			End If

            '' CASE
            if (SellSite="hiphoper") or (SellSite="itsHiphoper") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                iLine.FItemArray(22) = replace(replace(Trim(replace(iLine.FItemArray(22),iLine.FItemArray(21),"")),"(",""),")","")
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

                '//수량이 합쳐져서 넘어옴 수량으로 나눈다
                if iLine.FItemArray(17) > 1 then
                	iLine.FItemArray(18) = iLine.FItemArray(18)/iLine.FItemArray(17)
                	iLine.FItemArray(19) = iLine.FItemArray(19)/iLine.FItemArray(17)
                end if
            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="ssg") then
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
				iLine.FItemArray(1) = left(Replace(iLine.FItemArray(1),".","-"),10)

				If Len(iLine.FItemArray(1)) = "8" Then
					iLine.FItemArray(1) = Left(iLine.FItemArray(1),4)&"-"&Mid(iLine.FItemArray(1),5,2)&"-"&Mid(iLine.FItemArray(1),7,2)
				End If
				iLine.FItemArray(22) = replace(iLine.FItemArray(22),"NONE","")

				If instr(iLine.FItemArray(22),"주문제작문구:") > 0 Then
					iLine.FItemArray(27) = Trim(Split(iLine.FItemArray(22), "주문제작문구:")(1))
				Else
					iLine.FItemArray(27) = ""
				End If

				beasongNum11st	= iLine.FItemArray(29)
				reserve01 = iLine.FItemArray(38)
				outmalloptionno = iLine.FItemArray(24)
				If outmalloptionno = "" Then
					outmalloptionno = "00000"
				End If

				if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
					if (iLine.FItemArray(22)<>"") then
						iLine.FItemArray(16)=getOptionCodByOptionNameSSG(iLine.FItemArray(15),iLine.FItemArray(22))
					else
						iLine.FItemArray(16)="0000"
					end if
				end if
            end if

            '' CASE
'            If (SellSite="coupang") then
'            	Dim tmpOptName
'            	tmpOptName = Trim(Replace(iLine.FItemArray(21),"선택1) 푸치바비 아이폰 5 5S 케이스 ",""))
'                iLine.FItemArray(21) = "Puchibabie 푸치바비_iPhone5/5S Case"
'                Select Case tmpOptName
'                	Case "무광_소프트그레이"
'                		iLine.FItemArray(22) = "01_Soft Grey"
'                		iLine.FItemArray(16) = "0011"
'                	Case "무광_소프트핑크"
'                		iLine.FItemArray(22) = "02_Soft Pink"
'                		iLine.FItemArray(16) = "0012"
'                	Case "무광_소프트민트"
'                		iLine.FItemArray(22) = "03_Soft Mint"
'                		iLine.FItemArray(16) = "0013"
'                	Case "유광_핫핑크"
'                		iLine.FItemArray(22) = "04_Hot Pink"
'                		iLine.FItemArray(16) = "0014"
'                	Case "유광_팝블루"
'                		iLine.FItemArray(22) = "05_Pop Blue"
'                		iLine.FItemArray(16) = "0015"
'                	Case "유광_팝스트라이프"
'                		iLine.FItemArray(22) = "06_Pop Stripe"
'                		iLine.FItemArray(16) = "0016"
'                	Case "무광_스위트프룻"
'                		iLine.FItemArray(22) = "07_Sweet Fruit"
'                		iLine.FItemArray(16) = "0017"
'                	Case "무광_파스텔블루"
'                		iLine.FItemArray(22) = "08_Pastel Blue"
'                		iLine.FItemArray(16) = "0018"
'                	Case "무광_그린로즈닷"
'                		iLine.FItemArray(22) = "09_Green Rose Dot"
'                		iLine.FItemArray(16) = "0019"
'                	Case "유광_화이트스케치"
'                		iLine.FItemArray(22) = "10_White Sketch"
'                		iLine.FItemArray(16) = "0020"
'                End Select
'                iLine.FItemArray(15) = "783540"
'                iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
'				iLine.FItemArray(18) = 13800
'				iLine.FItemArray(19) = 13800
'            end if

			If (SellSite="coupang") then
                ' iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                ' iLine.FItemArray(1) = left(Replace(iLine.FItemArray(1),".","-"),10)
				' iLine.FItemArray(22) = replace(iLine.FItemArray(22),"NONE","")
                ' CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                ' iLine.FItemArray(15) = rtitemid
                ' iLine.FItemArray(16) = rtitemoption
                ' '//엑셀내에 존재함
                ' 'iLine.FItemArray(18) = rtSellPrice
                ' 'iLine.FItemArray(19) = rtSellPrice

                ' if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                '     if (iLine.FItemArray(22)<>"") then
           		' 		temArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                '      else
                '         iLine.FItemArray(16)="0000"
                '      end if
                ' end if
				Dim spItemId, spItemOption
                spItemId        = Split(iLine.FItemArray(15), "_")(0)
                spItemOption    = Split(iLine.FItemArray(15), "_")(1)

            	iLine.FItemArray(15) = spItemId
				iLine.FItemArray(16) = spItemOption
				beasongNum11st	= iLine.FItemArray(29)
				outmalloptionno = iLine.FItemArray(27)
				iLine.FItemArray(27) = ""

				iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)
				iLine.FItemArray(0)	= Trim(iLine.FItemArray(0))
				iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
			end if

            if (SellSite="ollehtv") then
                iLine.FItemArray(21) = Trim(iLine.FItemArray(21))
                CALL getEtcSiteNameOrCode2ItemCode(sellsite,iLine.FItemArray(23),iLine.FItemArray(21),iLine.FItemArray(22),rtitemid, rtitemoption, rtSellPrice)
                iLine.FItemArray(15) = rtitemid
                iLine.FItemArray(16) = rtitemoption

                iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)/iLine.FItemArray(17))      '''갯수로 나눠야 단가가 나옴.
                iLine.FItemArray(19) = iLine.FItemArray(18)
                '//엑셀내에 존재함
                'iLine.FItemArray(18) = rtSellPrice
                'iLine.FItemArray(19) = rtSellPrice

                iLine.FItemArray(1) = Left(iLine.FItemArray(1),10)

                if (iLine.FItemArray(16)="") and (iLine.FItemArray(15)<>"") then       ''옵션코드가 빈값.
                    if (iLine.FItemArray(22)<>"") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

            end if

            rtitemid=""
            rtitemoption=""
            rtSellPrice=""
            if (SellSite="wemakeprice")  then
				Select Case Trim(iLine.FItemArray(24))
					Case "230200538"	iLine.FItemArray(16) = "0011"
					Case "230200539"	iLine.FItemArray(16) = "0012"
					Case "230200540"	iLine.FItemArray(16) = "0011"
					Case "230200541"	iLine.FItemArray(16) = "0012"
					Case "230200542"	iLine.FItemArray(16) = "0011"
					Case "230200543"	iLine.FItemArray(16) = "0012"
					Case "230200544"	iLine.FItemArray(16) = "0011"
					Case "230200545"	iLine.FItemArray(16) = "0012"
					Case "230200548"	iLine.FItemArray(16) = "0011"
					Case "230200549"	iLine.FItemArray(16) = "0012"
					Case "230200550"	iLine.FItemArray(16) = "0011"
					Case "230200551"	iLine.FItemArray(16) = "0012"
					Case "230200552"	iLine.FItemArray(16) = "0013"
					Case "230200553"	iLine.FItemArray(16) = "0014"
					Case "230200556"	iLine.FItemArray(16) = "0011"
					Case "230200557"	iLine.FItemArray(16) = "0012"
					Case "230200558"	iLine.FItemArray(16) = "0013"
					Case "230200559"	iLine.FItemArray(16) = "0014"
					Case "230200560"	iLine.FItemArray(16) = "0011"
					Case "230200561"	iLine.FItemArray(16) = "0012"
					Case "230200562"	iLine.FItemArray(16) = "0014"

					Case "222648154"	iLine.FItemArray(16) = "0011"
					Case "222648155"	iLine.FItemArray(16) = "0012"
					Case "222648156"	iLine.FItemArray(16) = "0013"
					Case "222648157"	iLine.FItemArray(16) = "0014"
					Case "222648158"	iLine.FItemArray(16) = "0011"
					Case "222648159"	iLine.FItemArray(16) = "0012"
					Case "222648160"	iLine.FItemArray(16) = "0013"
					Case "222648161"	iLine.FItemArray(16) = "0014"

					Case "235610341"	iLine.FItemArray(16) = "0011"
					Case "235610342"	iLine.FItemArray(16) = "0012"
					Case "235610343"	iLine.FItemArray(16) = "0013"
					Case "235610344"	iLine.FItemArray(16) = "0014"
					Case "235610345"	iLine.FItemArray(16) = "0015"
					Case "235610346"	iLine.FItemArray(16) = "0016"
					Case "235610335"	iLine.FItemArray(16) = "0011"
					Case "235610336"	iLine.FItemArray(16) = "0012"
					Case "235610337"	iLine.FItemArray(16) = "0013"
					Case "235610338"	iLine.FItemArray(16) = "0014"
					Case "235610339"	iLine.FItemArray(16) = "0015"
					Case "235610340"	iLine.FItemArray(16) = "0016"
					Case "235554716"	iLine.FItemArray(16) = "0011"
					Case "235554717"	iLine.FItemArray(16) = "0012"
					Case "235554718"	iLine.FItemArray(16) = "0013"
					Case "235554719"	iLine.FItemArray(16) = "0015"
					Case "235554720"	iLine.FItemArray(16) = "0016"
					Case "235554714"	iLine.FItemArray(16) = "0011"
					Case "235554715"	iLine.FItemArray(16) = "0012"
					Case "235554709"	iLine.FItemArray(16) = "0011"
					Case "235554710"	iLine.FItemArray(16) = "0012"
					Case "235554711"	iLine.FItemArray(16) = "0013"
					Case "235554712"	iLine.FItemArray(16) = "0014"
					Case "235554713"	iLine.FItemArray(16) = "0015"

					Case "237893189"	iLine.FItemArray(16) = "0011"
					Case "237893190"	iLine.FItemArray(16) = "0012"
					Case "237893191"	iLine.FItemArray(16) = "0013"
					Case "237893192"	iLine.FItemArray(16) = "0014"
					Case "237893193"	iLine.FItemArray(16) = "0015"
					Case "237893194"	iLine.FItemArray(16) = "0016"
					Case "237893195"	iLine.FItemArray(16) = "0011"
					Case "237893196"	iLine.FItemArray(16) = "0012"
					Case "237893197"	iLine.FItemArray(16) = "0013"
					Case "237893198"	iLine.FItemArray(16) = "0014"
					Case "237893199"	iLine.FItemArray(16) = "0015"
					Case "237893200"	iLine.FItemArray(16) = "0016"
					Case Else			iLine.FItemArray(16)="0000"
				End Select

'예전딜은 9700원 이상이면 무료배송 / 지금 하는 딜은 3만원이상이면 무료배송
'				if (iLine.FItemArray(18)>=9700) Then
'					iLine.FItemArray(28)="0"
'				else
'					iLine.FItemArray(28)="2500"
'				end if

				If (iLine.FItemArray(18)>=30000) Then
					iLine.FItemArray(28)="0"
				Else
					iLine.FItemArray(28)="2500"
				End If

                if (iLine.FItemArray(17)<>"") then
                    if (iLine.FItemArray(17)>1) then
                        iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)/iLine.FItemArray(17))      '''갯수로 나눠야 단가가 나옴.
                    end if
				end if
				iLine.FItemArray(19) = iLine.FItemArray(18)
            end if
'            IF (sellsite="wemakeprice") then
'                if (Left(iLine.FItemArray(21),Len("[상품선택 : 엄마 아빠 핸드폰줄 세트"))="[상품선택 : 엄마 아빠 핸드폰줄 세트") then
'                    iLine.FItemArray(15) = 475219
'                    iLine.FItemArray(21) = "엄마 아빠 핸드폰줄 세트"
'                    iLine.FItemArray(16) = "0000"
'                    iLine.FItemArray(22) = ""
'                    iLine.FItemArray(18) = 4400
'                    iLine.FItemArray(19) = 4400
'                elseif (Left(iLine.FItemArray(21),Len("[상품선택 : 카네이션 훈장 세트"))="[상품선택 : 카네이션 훈장 세트") then
'                    iLine.FItemArray(15) = 475372
'                    iLine.FItemArray(21) = "카네이션 훈장 세트"
'                    iLine.FItemArray(16) = "0000"
'                    iLine.FItemArray(22) = ""
'                    iLine.FItemArray(18) = 4350
'                    iLine.FItemArray(19) = 4350
'                elseif (Left(iLine.FItemArray(21),Len("[상품선택 : 카네이션 핸드폰줄(레드)"))="[상품선택 : 카네이션 핸드폰줄(레드)") then
'                    iLine.FItemArray(15) = 475218
'                    iLine.FItemArray(21) = "카네이션 핸드폰줄"
'                    iLine.FItemArray(16) = "0011"
'                    iLine.FItemArray(22) = "레드"
'                    iLine.FItemArray(18) = 4000
'                    iLine.FItemArray(19) = 4000
'                elseif (Left(iLine.FItemArray(21),Len("[상품선택 : 카네이션 핸드폰줄(핑크)"))="[상품선택 : 카네이션 핸드폰줄(핑크)") then
'                    iLine.FItemArray(15) = 475218
'                    iLine.FItemArray(21) = "카네이션 핸드폰줄"
'                    iLine.FItemArray(16) = "0012"
'                    iLine.FItemArray(22) = "핑크"
'                    iLine.FItemArray(18) = 4000
'                    iLine.FItemArray(19) = 4000
'                elseif (Left(iLine.FItemArray(21),Len("[상품선택 : 크리스탈 카네이션 브로치"))="[상품선택 : 크리스탈 카네이션 브로치") then
'                    iLine.FItemArray(15) = 475457
'                    iLine.FItemArray(21) = "크리스탈 카네이션 브로치"
'                    iLine.FItemArray(16) = "0000"
'                    iLine.FItemArray(22) = ""
'                    iLine.FItemArray(18) = 6750
'                    iLine.FItemArray(19) = 6750
'                elseif (Left(iLine.FItemArray(21),Len("[상품선택 : 진주 카네이션 브로치"))="[상품선택 : 진주 카네이션 브로치") then
'                    iLine.FItemArray(15) = 475459
'                    iLine.FItemArray(21) = "진주 카네이션 브로치"
'                    iLine.FItemArray(16) = "0000"
'                    iLine.FItemArray(22) = ""
'                    iLine.FItemArray(18) = 4250
'                    iLine.FItemArray(19) = 4250
'                end if
'            end if

            IF (sellsite="lotteCom") then
            	Dim isOptAddLotteCom
				sqlStr = ""
				sqlStr = sqlStr & " SELECT TOP 1 itemid, itemoption"
				sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] as m "
				sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_lotteAddOption_regItem as r on m.idx = r.midx "
				sqlStr = sqlStr & " WHERE IsNULL(r.LotteGoodNo, r.LotteTmpGoodNo)= '"&iLine.FItemArray(23)&"' "
				sqlStr = sqlStr & " and m.mallid = 'lotteCom' "
				rsget.Open sqlStr,dbget,1
				If (Not rsget.EOF) Then
					isOptAddLotteCom = "Y"
					iLine.FItemArray(15) = rsget("itemid")
					iLine.FItemArray(16) = rsget("itemoption")
				Else
					isOptAddLotteCom = "N"
				End If
				rsget.Close

				If isOptAddLotteCom = "N" Then
	                if (iLine.FItemArray(15)="") and (iLine.FItemArray(23)<>"") then
	                    iLine.FItemArray(15) = getItemIDByUpcheItemCode(sellsite,iLine.FItemArray(23))
	                end if

	                if (iLine.FItemArray(16)="") then
	                    if (iLine.FItemArray(22)<>"") then
	                        iLine.FItemArray(16)=getOptionCodByOptionNameLotte(iLine.FItemArray(15),iLine.FItemArray(22))
	                     else
	                        iLine.FItemArray(16)="0000"
	                     end if
	                end if
				End If

				iLine.FItemArray(0) = replace(iLine.FItemArray(0),"-","")				'2014-12-31 김진영 추가
                if (iLine.FItemArray(26)="") then
                    iLine.FItemArray(26) = iLine.FItemArray(29)
                else
                    iLine.FItemArray(26) = iLine.FItemArray(26)&VBCRLF&iLine.FItemArray(29)
                end if

                if (iLine.FItemArray(32)="교환주문") then
                    iLine.FItemArray(32)=3
                elseif (iLine.FItemArray(32)<>"주문") then
                    iLine.FItemArray(32)=9
                else
                    iLine.FItemArray(32)=0
                end if
            END IF

            IF (sellsite="dnshop") then
                if (iLine.FItemArray(16)="") then
                     if (iLine.FItemArray(22)<>"") and (iLine.FItemArray(22)<>"단품없음") then
                        iLine.FItemArray(16)=getOptionCodByOptionName(iLine.FItemArray(15),iLine.FItemArray(22))

                        ''if (iLine.FItemArray(16)="") then iLine.FItemArray(16)="0000"
                     else
                        iLine.FItemArray(16)="0000"
                     end if
                end if

				'2015-03-12 김진영 하단 주석처리, 바뀐 엑셀에는 상점번호가 없음
				''상점번호4 무료배송.
				'if (iLine.FItemArray(29)="4") then
				'   iLine.FItemArray(28) = -1
				'end if

            END IF

            IF (sellsite="interpark") then
                ''selldate
                iLine.FItemArray(1) = Left(iLine.FItemArray(0),4)&"-"&Mid(iLine.FItemArray(0),5,2)&"-"&Mid(iLine.FItemArray(0),7,2)

                ''주문자 ID가 끼어있음
                iLine.FItemArray(5) = ReplaceText(iLine.FItemArray(5),"(\()[\s\S]*(\))","")

                ''rw iLine.FItemArray(15)&"|"&iLine.FItemArray(16)
                if (iLine.FItemArray(15)=iLine.FItemArray(16)) then
                    iLine.FItemArray(16)="0000"
                end if

                '''옵션이 없는 케이스는 이상한 케이스..***
                if (iLine.FItemArray(16)="") then
                    IF (Trim(iLine.FItemArray(22))<>"") then
                        POS1 = InStr(iLine.FItemArray(22),"/")
                        bufOptionName = ""
                        IF (POS1>0) THEN bufOptionName=Mid(iLine.FItemArray(22),POS1+1,255)
                        bufOptionName = Trim(bufOptionName)
                        iLine.FItemArray(16) = getOptionCodByOptionName(iLine.FItemArray(15),bufOptionName)

                        'if iLine.FItemArray(16)="" then iLine.FItemArray(16)="0000"  ''' 옵션이 매핑 안되면 안들어가도록.. 2012/03/02
                        if iLine.FItemArray(16)="" then iLine.FItemArray(16)="0000" ''일단 넣음 주문입력시 수정하게끔.
                    else
                        iLine.FItemArray(16)="0000"
                    end if
                end if

                ''주문제작문구 추가 : 2012-09-14
                POS1 = InStr(iLine.FItemArray(22),"| 주문제작문구")
                IF (POS1<1) then
                    POS1 = InStr(iLine.FItemArray(22),"|주문제작문구")
                end if

                if (POS1>0) then
                    if (iLine.FItemArray(16) = "") then iLine.FItemArray(16) = "0000" ''20121219추가 ''|주문제작문구/tea party for two/ all that razz/ groovy grape

                    POS2 = InStr(Mid(iLine.FItemArray(22),pos1,512),"/")
                    if (POS2>0) then '' 27 :: 주문제작문구
                        iLine.FItemArray(27) = Trim(Mid(iLine.FItemArray(22),pos1+pos2,512))
                    end if
                end if


                '''383048 이상함.. // 옵션코드 쪽에 빈값이고 옵션명에 색상 / French Lilac" 옵션구분 / 옵션명 이런식으로 올라오는 경우 잇음.
                '''256712
                if (iLine.FItemArray(15)="383048") and ((iLine.FItemArray(16)="0000") or (iLine.FItemArray(16)="")) then
                    if (Trim(iLine.FItemArray(22))="옵션 / 색상선택 | 선택2 / Cobalt Blue") then
                        iLine.FItemArray(16)="0015"
                    elseif (Trim(iLine.FItemArray(22))="옵션 / 색상선택 | 선택2 / Ivory") then
                        iLine.FItemArray(16)="0014"
                    elseif (Trim(iLine.FItemArray(22))="옵션 / 색상선택 | 선택2 / Black") then
                        iLine.FItemArray(16)="0013"
                    elseif (Trim(iLine.FItemArray(22))="옵션 / 색상선택 | 선택2 / Orange Red") then
                        iLine.FItemArray(16)="0012"
                    elseif (Trim(iLine.FItemArray(22))="옵션 / 색상선택 | 선택2 / Brown") then
                        iLine.FItemArray(16)="0011"
                    end if
                end if

                if (iLine.FItemArray(15)="256712") and ((iLine.FItemArray(16)="0000") or (iLine.FItemArray(16)="")) then
                    if (Trim(iLine.FItemArray(22))="선택1 / 화이트+엽서세트선택안함") then
                        iLine.FItemArray(16)="Z310"
                    elseif (Trim(iLine.FItemArray(22))="선택1 / 내추럴+엽서세트선택안함") then
                        iLine.FItemArray(16)="Z210"
                    elseif (Trim(iLine.FItemArray(22))="선택1 / 블랙+엽서세트선택안함") then
                        iLine.FItemArray(16)="Z110"
                    end if
                end if
            End IF

            ''rw "@partnerItemID="&iLine.FItemArray(15)
            ''rw "@partnerItemName="&iLine.FItemArray(21)
            ''rw "@partnerOption="&iLine.FItemArray(16)
            ''rw "@partnerOptionName="&iLine.FItemArray(22)
            ''rw "@SellPrice="&iLine.FItemArray(18)
            ''rw "@RealSellPrice="&iLine.FItemArray(19)

            ''옵션가격 있는경우.

            iLine.FItemArray(17) = Replace(iLine.FItemArray(17),",","")
            iLine.FItemArray(18) = Replace(iLine.FItemArray(18),",","")
            iLine.FItemArray(19) = Replace(iLine.FItemArray(19),",","")



            if (sellsite="dnshop") then '' 2014/01/15 interpark 추가
                ''실매출 판매가-할인값 으로 변경 2014/03/10--------------------------- 매출금액이 해당 상품금액 합계가 아닌듯 함.
                if iLine.FItemArray(30)<>"" then
                    iLine.FItemArray(19) = iLine.FItemArray(18)-iLine.FItemArray(30)
                end if
                ''--------------------------------------------------------------------

            '''가격 관련 수량이 1개 이상일때  ''2011-06-29 추가
                if (iLine.FItemArray(17)<>"") then
                    if (iLine.FItemArray(17)>1) then
                        iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)/iLine.FItemArray(17))
                        if (iLine.FItemArray(19)<>"") then
                            iLine.FItemArray(19) = CLNG(iLine.FItemArray(19)/iLine.FItemArray(17))
                        end if
                    end if
                else
                    response.write "."&iLine.FItemArray(17)
                end if
            end if

            if (sellsite="interpark") then
                if (iLine.FItemArray(17)<>"") then
                    if (iLine.FItemArray(17)>1) then
                        if (iLine.FItemArray(19)<>"") then
                            iLine.FItemArray(19) = CLNG(iLine.FItemArray(19)/iLine.FItemArray(17))
                        end if
                    end if
                else
                    response.write "."&iLine.FItemArray(19)
                end if
            end if

			if (SellSite="lotteCom") then
'	            IF (UBound(iLine.FItemArray)>30) and (iLine.FItemArray(18)<>"") then
'	                IF (iLine.FItemArray(30)="") then iLine.FItemArray(30)="0"
'	                IF (iLine.FItemArray(31)="") then iLine.FItemArray(31)="0"
'
'	                iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)) + CLNG(iLine.FItemArray(30)) ''?? 주석처리 2013/10 서동석
'	            END IF
			end if

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

			dim countryCode
			if (SellSite="cn10x10") or (SellSite="cnglob10x10") or (SellSite="cnhigo")  or (SellSite = "11stmy") or (SellSite = "cnugoshop") or (SellSite = "zilingo") or (SellSite = "etsy") then
				countryCode = iLine.FItemArray(33)
			end if

			if ucase(countryCode)="" then countryCode="KR"

			If Sellsite = "cnglob10x10" Then

			End If

			Dim replaceMemo
			If iLine.FItemArray(26) <> "" Then
				replaceMemo = Replace(iLine.FItemArray(26), "&amp;", "&")
				replaceMemo = Replace(replaceMemo, "amp;", "&")
				replaceMemo = Replace(replaceMemo, "&nbsp;", " ")
				replaceMemo = Replace(replaceMemo, "nbsp;", " ")
				replaceMemo = Replace(replaceMemo, "&lt;", "<")
				replaceMemo = Replace(replaceMemo, "lt;", "<")
				replaceMemo = Replace(replaceMemo, "&gt;", ">")
				replaceMemo = Replace(replaceMemo, "gt;", ">")
				replaceMemo = Replace(replaceMemo, "&quot;", """")
				replaceMemo = Replace(replaceMemo, "quot;", """")
				iLine.FItemArray(26) = replaceMemo
			End If

			'주문자 / 수령인(영문이름Case) 길이 길때 강제로 Left 문자열 처리
			If iLine.FItemArray(5) <> "" Then
				iLine.FItemArray(5) = LEFT(iLine.FItemArray(5), 28)
			End If

			If iLine.FItemArray(15) <> "" Then
				iLine.FItemArray(15) = LEFT(iLine.FItemArray(15), 28)
			End If

IF (application("Svr_Info")	= "Dev") or C_ADMIN_AUTH then
    rw "@SellSite="&SellSite
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

    rw "@orderDlvPay="&iLine.FItemArray(28)
    if UBound(iLine.FItemArray)>=29 then
        rw "@etc1="&iLine.FItemArray(29)
    end if
    rw "@countryCode="&countryCode

    rw "@outMallGoodsNo="&iLine.FItemArray(23)
    IF (SellSite="shoplinker") THEN
        rw "@etc2(shoplinkermallname)="&iLine.FItemArray(30)
        rw "@etc3(shoplinkerPrdCode)="&iLine.FItemArray(31)
        rw "@etc4(shoplinkerOrderID)="&iLine.FItemArray(32)
        rw "@etc4(shoplinkerMallid)="&iLine.FItemArray(33)
    ENd IF
    rw "@overseasPrice="&overseasPrice
    rw "@overseasDeliveryPrice="&overseasDeliveryPrice
    rw "@overseasRealPrice="&overseasRealPrice
    rw "@reserve01="&reserve01
    rw "@beasongNum11st="&beasongNum11st
	rw "@outmalloptionno="&outmalloptionno
    rw "------------------------------------------------"
	' response.end
ENd IF

        IF (iLine.FItemArray(0)<>"") and (iLine.FItemArray(0)<>"20110430-927718") then
            IF (sellsite="lotteCom") then
                orderCsGbn = iLine.FItemArray(32)
            ELSEIF (sellsite="gseshop") then
                orderCsGbn = iLine.FItemArray(32)
				if (orderCsGbn <> "주문") then
					orderCsGbn = "3"
				else
					orderCsGbn = "0"
				end if
            ELSE
                orderCsGbn = "0"
            end if

            IF (SellSite="shoplinker") THEN  ''2013/09/16 추가 샵링커관련
                outMallGoodsNo=iLine.FItemArray(23)
                shoplinkermallname=iLine.FItemArray(30)
                shoplinkerPrdCode=iLine.FItemArray(31)
                shoplinkerOrderID=iLine.FItemArray(32)
                shoplinkerMallID =iLine.FItemArray(33)
'

               '' rw shoplinkermallname&":"&shoplinkerMallID
            ELSE
                outMallGoodsNo=iLine.FItemArray(23)
                shoplinkermallname=""
                shoplinkerPrdCode=""
                shoplinkerOrderID=""
                shoplinkerMallID =""
            ENd IF

            paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
                ,Array("@SellSite" , adVarchar	, adParamInput, 32, SellSite)	_
    			,Array("@OutMallOrderSerial"	, adVarchar	, adParamInput,32, iLine.FItemArray(0))	_
    			,Array("@SellDate"	,adDate, adParamInput,, iLine.FItemArray(1)) _
    			,Array("@PayType"	,adVarchar, adParamInput,32, iLine.FItemArray(3)) _
    			,Array("@Paydate"	,adDate, adParamInput,, iLine.FItemArray(2)) _
    			,Array("@matchItemID"	,adInteger, adParamInput,, iLine.FItemArray(15)) _
    			,Array("@matchItemOption"	,adVarchar, adParamInput,4, iLine.FItemArray(16)) _
    			,Array("@partnerItemID"	,adVarchar, adParamInput,32, iLine.FItemArray(15)) _
    			,Array("@partnerItemName"	,adVarchar, adParamInput,128, iLine.FItemArray(21)) _
    			,Array("@partnerOption"	,adVarchar, adParamInput,128, iLine.FItemArray(16)) _
    			,Array("@partnerOptionName"	,adVarchar, adParamInput,1024, iLine.FItemArray(22)) _
    			,Array("@OrderUserID"	,adVarchar, adParamInput,32, iLine.FItemArray(4)) _
    			,Array("@OrderName"	,adVarchar, adParamInput,32, iLine.FItemArray(5)) _
    			,Array("@OrderEmail"	,adVarchar, adParamInput,100, iLine.FItemArray(8)) _
    			,Array("@OrderTelNo"	,adVarchar, adParamInput,16, iLine.FItemArray(6)) _
    			,Array("@OrderHpNo"	,adVarchar, adParamInput,16, iLine.FItemArray(7)) _
    			,Array("@ReceiveName"	,adVarchar, adParamInput,32, iLine.FItemArray(9)) _
    			,Array("@ReceiveTelNo"	,adVarchar, adParamInput,16, iLine.FItemArray(10)) _
    			,Array("@ReceiveHpNo"	,adVarchar, adParamInput,16, iLine.FItemArray(11)) _
    			,Array("@ReceiveZipCode"	,adVarchar, adParamInput,20, iLine.FItemArray(12)) _
    			,Array("@ReceiveAddr1"	,adVarchar, adParamInput,128, iLine.FItemArray(13)) _
    			,Array("@ReceiveAddr2"	,adVarchar, adParamInput,512, iLine.FItemArray(14)) _
    			,Array("@SellPrice"	,adCurrency, adParamInput,, iLine.FItemArray(18)) _
    			,Array("@RealSellPrice"	,adCurrency, adParamInput,, iLine.FItemArray(19)) _
    			,Array("@ItemOrderCount"	,adInteger, adParamInput,, iLine.FItemArray(17)) _
    			,Array("@OrgDetailKey"	,adVarchar, adParamInput,32, iLine.FItemArray(25)) _
    			,Array("@DeliveryType"	,adInteger, adParamInput,, 0) _
    			,Array("@deliveryprice"	,adCurrency, adParamInput,, 0) _
    			,Array("@deliverymemo"	,adVarchar, adParamInput,400, iLine.FItemArray(26)) _
    			,Array("@requireDetail"	,adVarchar, adParamInput,1024, iLine.FItemArray(27)) _
    			,Array("@orderDlvPay"	,adCurrency, adParamInput,, iLine.FItemArray(28)) _
    			,Array("@orderCsGbn"	,adInteger, adParamInput,, orderCsGbn) _
    			,Array("@countryCode"	,adVarchar, adParamInput,2, countryCode) _
                ,Array("@outMallGoodsNo"	,adVarchar, adParamInput,20, outMallGoodsNo) _
    			,Array("@shoplinkerMallName" ,adVarchar, adParamInput,64, shoplinkermallname) _
    			,Array("@shoplinkerPrdCode"	,adVarchar, adParamInput,16, shoplinkerPrdCode) _
    			,Array("@shoplinkerOrderID"	,adVarchar, adParamInput,16, shoplinkerOrderID) _
    			,Array("@shoplinkerMallID"	,adVarchar, adParamInput,32, shoplinkerMallID) _
    			,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
				,Array("@overseasPrice"	,adCurrency, adParamInput,, overseasPrice) _
				,Array("@overseasDeliveryPrice"	,adCurrency, adParamInput,, overseasDeliveryPrice) _
				,Array("@overseasRealPrice"	,adCurrency, adParamInput,, overseasRealPrice) _
				,Array("@reserve01"	,adVarchar, adParamInput,32, reserve01) _
				,Array("@beasongNum11st"	,adVarchar, adParamInput,16, beasongNum11st) _
				,Array("@outmalloptionno"	,adVarchar, adParamInput,32, outmalloptionno) _
    		)

			If sellsite <> "nvstorefarmclass" Then
				If ( Trim(iLine.FItemArray(13)) = "") AND ( Trim(iLine.FItemArray(14)) = "") Then
					RetErr = -1
					retErrStr = "주소 누락 " & iLine.FItemArray(0)
					rw retErrStr
					set iLine = Nothing
					dbget.rollbackTrans
					response.write "<script>alert('주소 누락 건이 있습니다. 다시 확인하세요');</script>"
					dbget.close() : response.end
				End If
			End If

			If sellsite <> "cookatmall" and sellsite <> "cnglob10x10" and sellsite <> "cnhigo" and sellsite <> "11stmy" and sellsite <> "cnugoshop" and sellsite <> "kakaogift" and sellsite <> "etsy" and sellsite <> "nvstorefarmclass" Then
				If  (iLine.FItemArray(28) <> "") AND (isnumeric(iLine.FItemArray(28))) Then		'배송비가 5000이 넘으면 튕기게..
					'// CInt => CLng, skyer9, 2018-01-22
					If CLng(iLine.FItemArray(28)) > 5000 Then
		                RetErr = -1
		                retErrStr = "배송비 5000원 초과 " & iLine.FItemArray(0) & " 상품코드 =" & iLine.FItemArray(15)&" 옵션명 = "&iLine.FItemArray(22)
		                rw retErrStr
		                set iLine = Nothing
		                IF (sellsite<>"interpark") then
		                dbget.rollbackTrans
		                end if
		                response.write "<script>alert('배송비가 5000원이 넘습니다. 다시 확인하세요');</script>"
		                dbget.close() : response.end
		            End If
				End If
			End If

        	'//우편번호 사이에 "-" 가 없을경우
        	if (SellSite<>"cn10x10") and (SellSite<>"cnglob10x10") and (SellSite<>"cnhigo") and (SellSite <> "11stmy") and (SellSite <> "cnugoshop") and (SellSite <> "zilingo") and (SellSite <> "nvstorefarmclass") then
				'우편번호 치환..2015-12-23 16:08 김진영 우편번호가 5자리 미만일 때 튕기게..
        		If Len(iLine.FItemArray(12)) <= 4 Then	'wizwid의 우편번호가 4자리로 넘어옴..얼럿출력
	                RetErr = -1
	                retErrStr = "우편번호 5자리 미만"
	                rw retErrStr
	                set iLine = Nothing
	                IF (sellsite<>"interpark") then
	                dbget.rollbackTrans
	                end if
	                response.write "<script>alert('우편번호가 5자리 미만입니다. 다시 확인하세요');</script>"
	                dbget.close() : response.end
				Else
	        		If Len(iLine.FItemArray(12)) <> 5 Then			'2015-10-13 14:26 김진영 수정..우편번호가 5자리가 아닐때 아래 IF문 실행
			        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
			                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
			        	end if
			        End If
        		End If

'        		If Len(iLine.FItemArray(12)) = 4 Then			'2015-12-23 15:18 김진영 수정..wizwid의 우편번호가 4자리로 넘어옴..강제로 0붙임
'        			iLine.FItemArray(12) = CStr("0"&iLine.FItemArray(12))
'        		End If

'        		If Len(iLine.FItemArray(12)) <> 5 Then			'2015-10-13 14:26 김진영 수정..우편번호가 5자리가 아닐때 아래 IF문 실행
'		        	if instr(iLine.FItemArray(12),"-") = 0 or instr(iLine.FItemArray(12),"-") = "" then
'		                iLine.FItemArray(12) = left(iLine.FItemArray(12),3) &"-"& right(iLine.FItemArray(12),3)
'		        	end if
'		        End If
			end if

			If sellsite = "cnglob10x10" or sellsite = "cnhigo" or sellsite = "cnugoshop" Then
				If (len(iLine.FItemArray(10)) > 16) OR (len(iLine.FItemArray(11)) > 16) OR (len(iLine.FItemArray(6)) > 16) OR (len(iLine.FItemArray(7)) > 16) Then
	                RetErr = -1
	                retErrStr = "주문자 OR 수령인 전화번호 길이 16자리 초과"
	                rw retErrStr
	                set iLine = Nothing
	                IF (sellsite<>"interpark") then
	                dbget.rollbackTrans
	                end if
	                response.write "<script>alert('주문자 OR 수령인 전화번호 길이 16자리 초과');</script>"
	                dbget.close() : response.end
				End If
			End If

            If (SellSite="ezwel") Then
                If  (iLine.FItemArray(29) <> "출고준비중") then
                    RetErr = -1
	                retErrStr = "ezwell 엑셀 체크 주문상태 출고준비중 만 가능-"&iLine.FItemArray(0)&":"&iLine.FItemArray(29) ''채현아 요청 2015/03/02 추가
	                rw retErrStr

	                IF (sellsite<>"interpark") then
		                dbget.rollbackTrans
	                end if
	                response.write "<script>alert('"&"ezwell 엑셀 체크 주문상태 출고준비중 만 가능-"&iLine.FItemArray(0)&":"&iLine.FItemArray(29)&"');</script>"
	                dbget.close() : response.end
                end if
            end if

            if (iLine.FItemArray(16)<>"") and (iLine.FItemArray(15)<>"-1") and (iLine.FItemArray(15)<>"") then
                sqlStr = "db_temp.dbo.sp_TEN_xSite_TmpOrder_Insert"
                retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

                RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
                retErrStr  = GetValue(retParamInfo, "@retErrStr") ' 오류명

                if (RetErr<0) and (RetErr<>-1) then ''Break
                    rw retErrStr
                    set iLine = Nothing
                    IF (sellsite<>"interpark") then
                    dbget.rollbackTrans
                    end if
                    response.write "<script>alert('"&Replace("ERROR["&retErr&"]"& retErrStr,"'","")&"');</script>"
                    dbget.close() : response.end
                end if
            else
                RetErr = -1
                retErrStr = "상품코드 또는 옵션코드  매칭 실패" & iLine.FItemArray(0) & " 상품코드 =" & iLine.FItemArray(15)&" 옵션명 = "&iLine.FItemArray(22)
                rw retErrStr
                set iLine = Nothing
                IF (sellsite<>"interpark") then
                dbget.rollbackTrans
                end if
                response.write "<script>alert('"&Replace("ERROR["&retErr&"]"& retErrStr,"'","")&"');</script>"
                dbget.close() : response.end
            end if

            IF RetErr=0 then
                okCNT = okCNT +1
            ELSE
                'rw "RetErr:"&RetErr&":"&iLine.FItemArray(0)&":"&shoplinkerMallID
                'rw "retErrStr:"&retErrStr
                errCNT = errCNT + 1
                totErrMsg = totErrMsg + retErrStr + VbCRLF
            end if

        END IF

            IF (retErr)<>0 then

            END IF

            set iLine = Nothing
        end if

    Next
IF (sellsite<>"interpark") then
    dbget.CommitTrans
end if

''품절/가격 오류체크 ---------------------------------------------
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr
''-------------------------------------------------------------

IF errCNT<>0 then
    response.write "<script>alert('"&errCNT&"건 입력오류.\n\n"&Replace(totErrMsg,vbCRLF,"\n")&"')</script>"
end if
response.write "<script>alert('"&okCNT&"건 입력되었습니다.')</script>"
response.end
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
	conDB.Provider = "Microsoft.Jet.oledb.4.0"		'2017-10-30 김진영 하단으로 수정
	'conDB.Provider = "Microsoft.ACE.OLEDB.12.0"

	If SellSite = "gabangpop" or (SellSite="itsGabangpop") or SellSite = "musinsaITS" or (SellSite="itsMusinsa") Then
    	conDB.Properties("ExtEnded Properties").Value = "Excel 8.0;HDR=NO;IMEX=1"		'첫행까지 데이터(HDR), 필드속성무시(IMEX;숫자/텍스트)
    Else
		conDB.Properties("ExtEnded Properties").Value = "Excel 8.0;IMEX=1"  ''';IMEX=1 추가 2013/12/19
	End If

 ''   On Error Resume Next
        conDB.Open sFilePath

        IF (ERR) then
            fnGetXLFileArray=false
			'/이유를 알수 없는 서버단 에러남. "예기치 않은 오류. 외부 개체에 트랩 가능한 오류(C0000005)가 발생했습니다. 스크립트를 계속 실행할 수 없습니다"
			set conDB = nothing
            exit function
        End if
 ''  On Error Goto 0

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

''On Error Resume Next
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
					if Not IsArray(xlPosArr(i)) then
						'// 기존 로직
						if (xlPosArr(i)<0) then
							irowObj.FItemArray(i) = ""
						else
							'2019-10-11 15:05 김진영 gmartket1010 조건 추가
							If ((SellSite="gmarket1010") OR (SellSite="auction1010") OR (SellSite="hmall1010") OR (SellSite="gseshop") ) AND (i = 22) Then
								irowObj.FItemArray(i) = Replace(null2blank(Rs(xlPosArr(i))),"*","＊")
							Else
								irowObj.FItemArray(i) = Replace(null2blank(Rs(xlPosArr(i))),"*","")
							End If
						end if
					else
						'// 여러 필드를 합쳐야 할 때(예 : ezwel)
						tmpVal = 0
						for each tmpItem in xlPosArr(i)
							If (SellSite="ezwel") Then
								tmpVal = tmpVal + CLng(Trim(Replace(Replace(null2blank(Rs(tmpItem)),"*",""), "(상품할인)", "")))
							end if
						next
						irowObj.FItemArray(i) = tmpVal
					end if

                    ''rw irowObj.FItemArray(i)
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
''On Error Goto 0

    Set Rs = Nothing
    Set conDB = Nothing

    if Ubound(xlRowALL)< 1 then fnGetXLFileArray=false

End Function

Function AddTmpDbOrderData(ixlRowALL)
    AddTmpDbOrderData = false
end Function
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
