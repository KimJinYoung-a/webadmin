<% option Explicit %>
<!-- #include virtual="/designer/incSessionDesignerNoCache.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
'#######################################################
'   History : 2012.10.25 허진원 생성 - Tabs Upload 사용
'	Description : 파일 다운로드 처리
'#######################################################

	dim dfPath, FileNo, FileName, DestinationFolder, arrFileName

	FileNo = getNumeric(requestCheckVar(Request("fn"),8))

	if FileNo="" then Response.End
	 
	'// 품목번호별 다운로드 파일명
	Select Case FileNo
		Case "01": FileName = "텐바이텐_상품품목_01_의류.xls"
		Case "02": FileName = "텐바이텐_상품품목_02_구두_신발.xls"
		Case "03": FileName = "텐바이텐_상품품목_03_가방.xls"
		Case "04": FileName = "텐바이텐_상품품목_04_패션잡화(모자_벨트_액세서리).xls"
		Case "05": FileName = "텐바이텐_상품품목_05_침구류_커튼.xls"
		Case "06": FileName = "텐바이텐_상품품목_06_가구(침대_소파_싱크대_DIY제품).xls"
		Case "07": FileName = "텐바이텐_상품품목_07_영상가전(TV류).xls"
		Case "08": FileName = "텐바이텐_상품품목_08_가정용_전기제품(냉장고_세탁기_식기세척기_전자레인지).xls"
		Case "09": FileName = "텐바이텐_상품품목_09_계절가전(에어컨_온풍기).xls"
		Case "10": FileName = "텐바이텐_상품품목_10_사무용기기(컴퓨터_노트북_프린터).xls"
		Case "11": FileName = "텐바이텐_상품품목_11_광학기기(디지털카메라_캠코더).xls"
		Case "12": FileName = "텐바이텐_상품품목_12_소형전자(MP3_전자사전_등).xls"
		Case "13": FileName = "텐바이텐_상품품목_13_휴대폰.xls"
		Case "14": FileName = "텐바이텐_상품품목_14_내비게이션.xls"
		Case "15": FileName = "텐바이텐_상품품목_15_자동차용품(자동차부품_기타_자동차용품).xls"
		Case "16": FileName = "텐바이텐_상품품목_16_의료기기.xls"
		Case "17": FileName = "텐바이텐_상품품목_17_주방용품.xls"
		Case "18": FileName = "텐바이텐_상품품목_18_화장품.xls"
		Case "19": FileName = "텐바이텐_상품품목_19_귀금속_보석_시계류.xls"
		Case "20": FileName = "텐바이텐_상품품목_20_식품(농수산물).xls"
		Case "21": FileName = "텐바이텐_상품품목_21_가공식품.xls"
		Case "22": FileName = "텐바이텐_상품품목_22_건강기능식품.xls"
		Case "23": FileName = "텐바이텐_상품품목_23_영유아용품.xls"
		Case "24": FileName = "텐바이텐_상품품목_24_악기.xls"
		Case "25": FileName = "텐바이텐_상품품목_25_스포츠용품.xls"
		Case "26": FileName = "텐바이텐_상품품목_26_서적.xls"
		Case "27": FileName = "텐바이텐_상품품목_27_호텔_펜션_예약.xls"
		Case "28": FileName = "텐바이텐_상품품목_28_여행패키지.xls"
		Case "29": FileName = "텐바이텐_상품품목_29_항공권.xls"
		Case "30": FileName = "텐바이텐_상품품목_30_자동차_대여_서비스(렌터카).xls"
		Case "31": FileName = "텐바이텐_상품품목_31_물품대여_서비스(정수기,비데,공기청정기_등).xls"
		Case "32": FileName = "텐바이텐_상품품목_32_물품대여_서비스(서적,유아용품,행사용품_등).xls"
		Case "33": FileName = "텐바이텐_상품품목_33_디지털_콘텐츠(음원,게임,인터넷강의_등).xls"
		Case "34": FileName = "텐바이텐_상품품목_34_상품권_쿠폰.xls"
		Case "35": FileName = "텐바이텐_상품품목_35_기타.xls"
		Case "900": FileName = "텐바이텐_해외배송_정보.xls"
		Case "990": FileName = "텐바이텐_상품안전인증대상_정보.xls"
	End Select

	On Error Resume Next
	'파일 다운로드
	dfPath = server.mappath("/designer/itemmaster/itemInfoFile/")
	DestinationFolder = dfPath & "/" & fileName

	'// 다운로드 컨퍼넌트 선언 및 다운로드 전송
	Dim oDownload

	IF (application("Svr_Info")	= "Dev") then
	    Set oDownload = Server.CreateObject("TABS.Download")	   '' - TEST
	ELSE
	    Set oDownload = Server.CreateObject("TABSUpload4.Download")	''REAL
	END IF
	 
	oDownload.FilePath = DestinationFolder
	oDownload.FileName = fileName
	oDownload.TransferFile True

	Set oDownload = Nothing

    IF (ERR) then
		response.write "<script>alert('죄송합니다. 파일을 준비중입니다.')</script>"
		response.write "<script>self.close();</script>"
    End if
    On Error Goto 0
%>