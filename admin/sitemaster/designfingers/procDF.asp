<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  디자인핑거스 DB 처리
' History : 2008.03.14 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/designfingersCls.asp"-->
<%
Dim sMode, menupos, strSql, iCurrpage
Dim iDFSeq, ievt_code,sevt_link
Dim sDFType,sTitle,tContents,sPrizeDate,blnDisplay,blnOtherMall, sComment, blnMainDisplay
Dim arrItemid, sProdName, sProdSize, sProdColor, sProdJe, sProdGu, sProdSpe, sIsMovie, sOpenDate, sTag, edid, emktid

Dim iImgCode
Dim arrMainTop, arrTop, arrSmall, arrList, arr3dView, arrAdd, arrTxtAdd(30), intLoop, arrEventLeft, arrEventRight, arrplay
Dim arrInfo(10), tmp3dv
Dim sImgURL, strMsg

sMode=  requestCheckVar(request("sM"),1)
menupos = requestCheckVar(request("menupos"),10)
sOpenDate = request("opendate")
sTag = html2db(requestCheckVar(request("sTag"),100))
iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호

SELECT Case sMode
Case "I"
	'// 컨텐츠
	sDFType 		= requestCheckVar(request("selDFT"),10)
	sTitle 			= html2db(requestCheckVar(request("sT"),64))
	tContents 		= html2db(request("txtC"))
	sPrizeDate 		= requestCheckVar(request("dPD"),10)
	sComment		= html2db(requestCheckVar(request("sCom"),200))
	sProdName		= html2db(requestCheckVar(request("sPdN"),100))
	sProdSize		= html2db(requestCheckVar(request("sPdS"),100))
	sProdColor		= html2db(requestCheckVar(request("sPdC"),100))
	sProdJe			= html2db(requestCheckVar(request("sPdJ"),100))
	sProdGu			= html2db(requestCheckVar(request("sPdG"),100))
	sProdSpe		= html2db(requestCheckVar(request("sPdP"),100))
	blnDisplay 		= requestCheckVar(request("rdoD"),1)
	blnMainDisplay 	= requestCheckVar(request("rdoMD"),1)
	blnOtherMall	= requestCheckVar(request("rdoOM"),1)

	edid  		= requestCheckVar(Request("selDId"),32)		'담당 디자이너
	emktid 		= requestCheckVar(Request("selMKTId"),32)		'담당 MKT

	'//상품
	arrItemid	= request("arrI")

	'// 이미지
	arrTop 		= split(request("imgtop"),",")	'//top 타입
	arrplay 	= request("imgplay")			'//play 타입
	arrSmall 	= request("imgsmall")			'//small 타입  - 이미지 1개 등록
	arrList 	= request("imglist")			'//list 타입 	- 이미지 1개 등록
	arr3dView 	= split(request("img3dv"),",")	'//3dview 타입
	arrAdd 		= split(request("imgadd"),",")	'//add 타입
	arrMainTop	= request("imgmain_top")
	arrEventLeft	= request("imgeventLeft")
	arrEventRight	= request("imgeventRight")

	'//3dview 이미지 설명
	tmp3dv			=  split(request("sel3dv"),",")
	for intLoop = 0 to ubound(tmp3dv)
		arrInfo(intLoop) = trim(tmp3dv(intLoop))
	next

	sIsMovie = request("ismovie")

	'//add 이미지 맵
	for intLoop = 0 to 14
		arrTxtAdd(intLoop)	= html2db(request("tA"&(intLoop+1)))	'//add 타입 링크
	next

	dbget.beginTrans
	'// 컨텐츠 저장
	strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers]([DFType], [Title], [Contents], [PrizeDate], [Comment],[IsDisplay] ,[IsOtherMall], [Userid], [IsMainDisplay] " & _
			"	, [ProdName], [ProdSize], [ProdColor], [ProdJe], [ProdGu], [ProdSpe], [IsMovie], [OpenDate], [Tag], [designerid], [partMKTid]) "&_
			" VALUES("&sDFType&",'"&sTitle&"','"&tContents&"','"&sPrizeDate&"','"&sComment&"',"&blnDisplay&","&blnOtherMall&",'"&session("ssBctId")&"','"&blnMainDisplay&"'," & _
			"	'"&sProdName&"', '"&sProdSize&"', '"&sProdColor&"', '"&sProdJe&"', '"&sProdGu&"', '"&sProdSpe&"', '" & sIsMovie & "', '" & sOpenDate & "', '" & sTag & "', '" & edid & "', '" & emktid & "')"
		dbget.execute strSql

		strSql =" select SCOPE_IDENTITY()"
		rsget.Open strSql,dbget
 		IF Not rsget.EOF THEN
 			iDFSeq = rsget(0)
 		ELSE
 			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
 		END IF
 		rsget.close

	'//상품 저장 : 이벤트 상품 테이블에 저장 - 이벤트 코드는 항상 1로 세팅
	IF arrItemid <> "" THEN
		ievt_code = 1
		sevt_link = html2db("\designfingers\designfingers.asp?fingerid="&iDFSeq)
		strSql = "INSERT INTO [db_event].[dbo].[tbl_eventitem](evt_code, itemid, evtgroup_code, evtitem_linkurl)"&_
				" SELECT "&ievt_code&", itemid,"&iDFSeq&",'"&sevt_link&"' "&_
				"	FROM [db_item].[dbo].[tbl_item] WHERE itemid in ("&arrItemid&")"
		dbget.execute strSql
	END IF

	'//이미지저장 : 위치별 코드 구분으로 이미지저장
	' top- codeNo:2
	for intLoop = 0 To uBound(arrTop)
	if(trim(arrTop(intLoop)) <> "" ) THEN
		iImgCode = 2
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&","&(intLoop+1)&",'"&trim(arrTop(intLoop))&"')"
		dbget.execute strSql
	end if
	next

	' small - codeNo:3
	if(trim(arrSmall) <> "" ) THEN
		iImgCode = 3
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrSmall)&"')"
		dbget.execute strSql
	end if

	' list - codeNo:4
	if(trim(arrList) <> "" ) THEN
		iImgCode = 4
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrList)&"')"
		dbget.execute strSql
	end if

	' 3dview- codeNo:7
	for intLoop = 0 To uBound(arr3dView)
	if(trim(arr3dView(intLoop)) <> "" ) THEN
		iImgCode = 7
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL],[ImgDescCode])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&","&(intLoop+1)&",'"&trim(arr3dView(intLoop))&"',"&arrInfo(intLoop)&")"
		dbget.execute strSql
	end if
	next

	' play - codeNo:8
	if(trim(arrplay) <> "" ) THEN
		iImgCode = 8
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrplay)&"')"
		dbget.execute strSql
	end if

	' add- codeNo:5
	for intLoop = 0 To uBound(arrAdd)
	if(trim(arrAdd(intLoop)) <> "" OR trim(arrTxtAdd(intLoop)) <> "" ) THEN
		iImgCode = 5
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL],[link])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&","&(intLoop+1)&",'"&trim(arrAdd(intLoop))&"','"&arrTxtAdd(intLoop)&"')"
		dbget.execute strSql
	end if
	next

	' main_top - codeNo:21
	if(trim(arrMainTop) <> "" ) THEN
		iImgCode = 21
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrMainTop)&"')"
		dbget.execute strSql
	end if

	' eventLeft - codeNo:22
	if(trim(arrEventLeft) <> "" ) THEN
		iImgCode = 22
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrEventLeft)&"')"
		dbget.execute strSql
	end if

	' eventRight - codeNo:23
	if(trim(arrEventRight) <> "" ) THEN
		iImgCode = 23
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrEventRight)&"')"
		dbget.execute strSql
	end if

	IF Err.Number = 0 THEN
		dbget.CommitTrans

		Call sbAlertMsg ("등록되었습니다.", "/admin/sitemaster/designfingers/listDF.asp?menupos="&menupos&"&iC="&iCurrpage&"", "self")
		dbget.close()	:	response.End
	ELSE
		dbget.RollBackTrans
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
	END IF
	dbget.close()	:	response.End
Case "U"
	'// 컨텐츠
	iDFSeq			=  requestCheckVar(request("iDFS"),10)
	sDFType 		= requestCheckVar(request("selDFT"),10)
	sTitle 			= html2db(requestCheckVar(request("sT"),64))
	tContents 		= html2db(request("txtC"))
	sPrizeDate 		= requestCheckVar(request("dPD"),10)
	sComment		= html2db(requestCheckVar(request("sCom"),200))
	sProdName		= html2db(requestCheckVar(request("sPdN"),100))
	sProdSize		= html2db(requestCheckVar(request("sPdS"),100))
	sProdColor		= html2db(requestCheckVar(request("sPdC"),100))
	sProdJe			= html2db(requestCheckVar(request("sPdJ"),100))
	sProdGu			= html2db(requestCheckVar(request("sPdG"),100))
	sProdSpe		= html2db(requestCheckVar(request("sPdP"),100))
	blnDisplay 		= requestCheckVar(request("rdoD"),1)
	blnMainDisplay 	= requestCheckVar(request("rdoMD"),1)
	blnOtherMall	= requestCheckVar(request("rdoOM"),1)

	edid  		= requestCheckVar(Request("selDId"),32)		'담당 디자이너
	emktid 		= requestCheckVar(Request("selMKTId"),32)		'담당 MKT

	'//상품
	arrItemid	= request("arrI")

	'// 이미지
	arrTop 		= split(request("imgtop"),",")	'//top 타입
	arrplay 	= request("imgplay")			'//play 타입
	arrSmall 	= request("imgsmall")			'//small 타입  - 이미지 1개 등록
	arrList 	= request("imglist")			'//list 타입 	- 이미지 1개 등록
	arr3dView 	= split(request("img3dv"),",")	'//3dview 타입
	arrAdd 		= split(request("imgadd"),",")	'//add 타입
	arrMainTop	= request("imgmain_top")
	arrEventLeft	= request("imgeventLeft")
	arrEventRight	= request("imgeventRight")

	'//3dview 이미지 설명
	tmp3dv			=  split(request("sel3dv"),",")

	sIsMovie = request("ismovie")

	for intLoop = 0 to UBound(tmp3dv)
		arrInfo(intLoop) = trim(tmp3dv(intLoop))
	next

	'//add 이미지 맵
	for intLoop = 0 to 14
		arrTxtAdd(intLoop)	= html2db(request("tA"&(intLoop+1)))	'//add 타입 링크
	next

	dbget.beginTrans
	'// 컨텐츠 저장
	strSql = "UPDATE [db_sitemaster].[dbo].[tbl_designfingers] SET [DFType] = "&sDFType&", [Title]= '"&sTitle&"', [Contents]= '"&tContents&"' "&_
			", [PrizeDate]= '"&sPrizeDate&"', Comment = '"&sComment&"' ,[IsDisplay] ="&blnDisplay&",[IsOtherMall] = "&blnOtherMall&", [Userid]= '"&session("ssBctId")&"' "&_
			", [IsMainDisplay] = '"&blnMainDisplay&"', [ProdName] = '"&sProdName&"', [ProdSize] = '"&sProdSize&"', [ProdColor] = '"&sProdColor&"'" & _
			", [ProdJe] = '"&sProdJe&"', [ProdGu] = '"&sProdGu&"', [ProdSpe] = '"&sProdSpe&"', [IsMovie] = '" & sIsMovie & "', [OpenDate] = '" & sOpenDate & "', [Tag] = '" & sTag & "' " & _
			", [designerid] = '" & edid & "', [partMKTid] = '" & emktid & "' " & _
			" WHERE DFSeq = "&iDFSeq
		dbget.execute strSql

	'//상품 저장 : 이벤트 상품 테이블에 저장 - 이벤트 코드는 항상 1로 세팅
	IF arrItemid <> "" THEN
		ievt_code = 1
		sevt_link = html2db("\designfingers\designfingers.asp?fingerid="&iDFSeq)

		'기존등록 상품 삭제
		strSql = "DELETE From [db_event].[dbo].[tbl_eventitem] WHERE evt_code ="&ievt_code&" AND evtgroup_code ="&iDFSeq
		dbget.execute strSql

		'타 핑거스 등록상품인지 검사
		strSql = "Select evtgroup_code, itemid From [db_event].[dbo].[tbl_eventitem] Where itemid in ("&arrItemid&") and evt_code=" & ievt_code
		rsget.Open strSql,dbget, 1
 		IF Not(rsget.EOF or rsget.BOF) THEN
 			Do Until rsget.EOF
 				strMsg = strMsg & "- [" & rsget("itemid") & "]상품이 " & rsget("evtgroup_code") & "번 핑거스에 존재.\n"
 			rsget.MoveNext
 			Loop
 			dbget.RollBackTrans
 			strMsg = strMsg & "\n이미 다른 곳에 등록되어있어 새로 등록할 수 없습니다."
 			Call sbAlertMsg (strMsg, "back", "")
 			dbget.close()	:	response.End
 		END IF
 		rsget.close

		'상품 저장
		strSql = "INSERT INTO [db_event].[dbo].[tbl_eventitem](evt_code, itemid, evtgroup_code, evtitem_linkurl)"&_
				" SELECT "&ievt_code&", itemid,"&iDFSeq&",'"&sevt_link&"' "&_
				"	FROM [db_item].[dbo].[tbl_item] WHERE itemid in ("&arrItemid&")"
		dbget.execute strSql

	END IF

	'//기존이미지 삭제
		strSql = "Delete FRom [db_sitemaster].[dbo].[tbl_designfingers_image] WHERE DFSeq = "&iDFSeq&" and DFCodeSeq NOT IN (6, 24, 25) "
		dbget.execute strSql

	'//이미지저장 : 위치별 코드 구분으로 이미지저장
	' top- codeNo:2
	for intLoop = 0 To uBound(arrTop)
	if(trim(arrTop(intLoop)) <> "" ) THEN
		iImgCode = 2
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&","&(intLoop+1)&",'"&trim(arrTop(intLoop))&"')"
		dbget.execute strSql
	end if
	next

	' small - codeNo:3
	if(trim(arrSmall) <> "" ) THEN
		iImgCode = 3
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrSmall)&"')"
		dbget.execute strSql
	end if

	' list - codeNo:4
	if(trim(arrList) <> "" ) THEN
		iImgCode = 4
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrList)&"')"
		dbget.execute strSql
	end if

	' 3dview- codeNo:7
	for intLoop = 0 To uBound(arr3dView)
	if(trim(arr3dView(intLoop)) <> "" ) THEN
		iImgCode = 7
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL],[ImgDescCode])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&","&(intLoop+1)&",'"&trim(arr3dView(intLoop))&"',"&arrInfo(intLoop)&")"
		dbget.execute strSql
	end if
	next

	' play - codeNo:8
	if(trim(arrplay) <> "" ) THEN
		iImgCode = 8
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrplay)&"')"
		dbget.execute strSql
	end if

	' add- codeNo:5
	for intLoop = 0 To uBound(arrAdd)
	if(trim(arrAdd(intLoop)) <> "" OR trim(arrTxtAdd(intLoop)) <> "" ) THEN
		iImgCode = 5
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL],[link])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&","&(intLoop+1)&",'"&trim(arrAdd(intLoop))&"','"&arrTxtAdd(intLoop)&"')"
		dbget.execute strSql
	end if
	next

	' main_top - codeNo:21
	if(trim(arrMainTop) <> "" ) THEN
		iImgCode = 21
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrMainTop)&"')"
		dbget.execute strSql
	end if

	' eventLeft - codeNo:22
	if(trim(arrEventLeft) <> "" ) THEN
		iImgCode = 22
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrEventLeft)&"')"
		dbget.execute strSql
	end if

	' eventRight - codeNo:23
	if(trim(arrEventRight) <> "" ) THEN
		iImgCode = 23
		strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&trim(arrEventRight)&"')"
		dbget.execute strSql
	end if

	IF Err.Number = 0 THEN
		dbget.CommitTrans

		Call sbAlertMsg ("수정되었습니다.", "/admin/sitemaster/designfingers/listDF.asp?menupos="&menupos&"&iC="&iCurrpage&"", "self")
		dbget.close()	:	response.End
	ELSE
		dbget.RollBackTrans
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
	END IF
	dbget.close()	:	response.End
Case "B"	'//배너등록
	iImgCode	= 6
	iDFSeq		=  requestCheckVar(request("iDFS"),10)
	sImgURL		=  requestCheckVar(request("sIU"),100)

	'//기존이미지 삭제
		strSql = "Delete FRom [db_sitemaster].[dbo].[tbl_designfingers_image] WHERE DFSeq = "&iDFSeq&" and DFCodeSeq = 6 "
		dbget.execute strSql

	'//이미지 등록
	strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_image]([DFSeq], [DFCodeSeq], [DFImgID], [ImgURL])"&_
				" VALUES('"&iDFSeq&"',"&iImgCode&",1,'"&sImgURL&"')"
	dbget.execute strSql
	IF Err.Number = 0 THEN
		Call sbAlertMsg ("등록되었습니다.", "listbest.asp?menupos="&menupos, "opener")
	ELSE
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
	END IF

Case "C"	'//코드관리
	Dim iCodeSeq, sCodeDesc, iPCodeSeq,iCodeSort, blnUsing
	iCodeSeq 	= requestCheckVar(request("iCS"),10)
	sCodeDesc 	= requestCheckVar(request("sCD"),32)
	iPCodeSeq 	= requestCheckVar(request("selPCS"),10)
	iCodeSort 	= requestCheckVar(request("iCSort"),10)
	blnUsing 	= requestCheckVar(request("rdoU"),1)
	IF iCodeSeq ="" THEN
		strSql  ="INSERT INTO [db_sitemaster].[dbo].[tbl_designfingers_code]([PCodeSeq], [CodeDesc], [CodeSort], [IsUsing])"&_
				"VALUES("&iPCodeSeq&",'"&sCodeDesc&"',"&iCodeSort&","&blnUsing&")"
		dbget.execute strSql
	ELSE
		strSql  ="UPDATE [db_sitemaster].[dbo].[tbl_designfingers_code] SET PCodeSeq = "&iPCodeSeq&", CodeDesc = '"&sCodeDesc&"', CodeSort = "&iCodeSort&", IsUsing ="&blnUsing&_
				"	WHERE DFCodeSeq = "&iCodeSeq
		dbget.execute strSql
	END IF
	IF Err.Number = 0 THEN
		Call sbAlertMsg ("저장되었습니다.", "/admin/sitemaster/designfingers/popManageCode.asp?sPCS="&iPCodeSeq, "self")
	ELSE
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
	END IF
Case Else
	Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
END SELECT
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->


