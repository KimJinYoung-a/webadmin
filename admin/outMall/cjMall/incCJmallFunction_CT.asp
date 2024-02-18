<%
Dim isCJ_DebugMode : isCJ_DebugMode = True
Dim cjMallAPIURL

IF application("Svr_Info")="Dev" THEN
	cjMallAPIURL = "http://210.122.101.154:8110/IFPAServerAction.action"	'' 테스트서버
'	cjMallAPIURL = "http://210.122.101.154:8210/IFPAServerAction.action"	'' 개편될 CJ QA서버 URL
Else
	cjMallAPIURL = "http://api.cjmall.com/IFPAServerAction.action"			'' 실서버
End if

Public Pmode
Public OPTYN

Function getMode()
	getMode = Pmode
End Function

Function getOPTYN()
	getOPTYN = OPTYN
End Function

function getCjMallPrdNoByItemID(byval iitemid)
    dim ret
    dim sqlStr

    if iitemid="" then Exit function

    sqlStr = " select isNULL(cjmallprdno,'') as cjmallprdno from db_outMall.dbo.tbl_cjmall_regitem where itemid="&iitemid

    rsCTget.Open sqlStr, dbCTget
	If Not(rsCTget.EOF or rsCTget.BOF) Then
		ret = rsCTget("cjmallprdno")
	End If
	rsCTget.close

	getCjMallPrdNoByItemID = ret
end function

function getCjMallfirstItemoption(byval iitemid)
    dim ret
    dim sqlStr

    if iitemid="" then Exit function

    sqlStr = " select top 1 itemoption from db_outMall.dbo.tbl_OutMall_regedoption"
    sqlStr = sqlStr & " where mallid='"&CMALLNAME&"'"
    sqlStr = sqlStr & " and itemid="&iitemid

    rsCTget.Open sqlStr, dbCTget
	If Not(rsCTget.EOF or rsCTget.BOF) Then
		ret = rsCTget("itemoption")
	End If
	rsCTget.close

	if (ret="") then
	    ret = "0011"
	end if

	getCjMallfirstItemoption = ret
end function

Function getXMLString(byval iitemid, mode, paramA)
	Dim oCJMallItem
	Dim strRst, bufRET, buf1, notitemId, notmakerid

	SET oCJMallItem = new CCjmall
		oCJMallItem.FRectMode = mode
		oCJMallItem.FRectItemID = iitemid
	If (mode = "REG") Then
		oCJMallItem.getCJMallNotRegItemList
		If (oCJMallItem.FREsultCount > 0) Then
			getXMLString = oCJMallItem.FItemList(0).getCjmallItemRegXML
		End If
	ElseIf (mode = "LIST") Then '' 특정 날짜로 변경
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_07"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_07.xsd"">"
		strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
		strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
		strRst = strRst &"<tns:contents>"
		strRst = strRst &"	<tns:sinstDtFrom>"&iitemid&"</tns:sinstDtFrom>"
		strRst = strRst &"	<tns:sinstDtTo>"&iitemid&"</tns:sinstDtTo>"
		strRst = strRst &"	<tns:schnCd>30</tns:schnCd>"
		strRst = strRst &"	<tns:sinstTime>00:00:00</tns:sinstTime>"
		strRst = strRst &"</tns:contents>"
		strRst = strRst &"</tns:ifRequest>"
		getXMLString = strRst
	ElseIf (mode = "DayLIST") Then
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_07"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_07.xsd"">"
		strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
		strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
		strRst = strRst &"<tns:contents>"
		strRst = strRst &"	<tns:sinstDtFrom>"&Left(now - iitemid, 10)&"</tns:sinstDtFrom>"
		strRst = strRst &"	<tns:sinstDtTo>"&Left(now, 10)&"</tns:sinstDtTo>"
		strRst = strRst &"	<tns:schnCd>30</tns:schnCd>"
		strRst = strRst &"	<tns:sinstTime>00:00:00</tns:sinstTime>"
		strRst = strRst &"</tns:contents>"
		strRst = strRst &"</tns:ifRequest>"
		getXMLString = strRst
    elseif (mode="cjItemCheck") Then
        strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_07"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_07.xsd"">"
		strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
		strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
		strRst = strRst &"<tns:contents>"
		strRst = strRst &"	<tns:sinstDtFrom>"&"2013-04-01"&"</tns:sinstDtFrom>"
		strRst = strRst &"	<tns:sinstDtTo>"&Left(now, 10)&"</tns:sinstDtTo>"
		strRst = strRst &"	<tns:schnCd>30</tns:schnCd>"
		if (paramA="cjPrdno") then
		    strRst = strRst &"	<tns:itemCd>"&iitemid&"</tns:itemCd>"
		else
    		strRst = strRst &"	<tns:vpn>"&iitemid&"</tns:vpn>"
    	end if
		'strRst = strRst &"	<tns:sinstTime>00:00:00</tns:sinstTime>"
		strRst = strRst &"</tns:contents>"
		strRst = strRst &"</tns:ifRequest>"
		getXMLString = strRst
    ElseIf mode = "PRI" Then
        oCJMallItem.getCjmallEditedItemList
        If (oCJMallItem.FREsultCount > 0) Then
    		getXMLString = oCJMallItem.FItemList(0).getcjmallItemPriceModXML(paramA)
    	end if
    ElseIf mode = "PRI2" Then
        oCJMallItem.getCjmallEditedItemList
        If (oCJMallItem.FREsultCount > 0) Then
    		getXMLString = oCJMallItem.FItemList(0).getcjmallItemSellPriceModXML()
    	end if
	ElseIf (mode = "MDT") OR (mode = "EDT") OR (mode = "QTY") or (mode = "DateRes") Then
		oCJMallItem.getCjmallEditedItemList
		If (oCJMallItem.FREsultCount > 0) Then
			If mode = "MDT" Then
				Dim sqlStr, arrRows, isOptionExists, i
				Dim itemoption, optiontypename, optionname, optLimit, optlimityn, isUsing, optsellyn, preged, optNameDiff, forceExpired, oopt, ooptCd, YtoN, NtoY, DelOpt
				Dim validSellno
''2013-09-24 김진영 17:45 하단 프로시저 수정(단일상품일 때 자꾸 품절로 내려가서 단일상품의 경우 IF처리함)
				sqlStr = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_cjmall 'cjmall'," & iitemid
				rsget.CursorLocation = adUseClient
				rsget.CursorType = adOpenStatic
				rsget.LockType = adLockOptimistic
				rsget.Open sqlStr, dbget
				If Not(rsget.EOF or rsget.BOF) Then
					arrRows = rsget.getRows
				End If
				rsget.close
				isOptionExists = isArray(arrRows)
				bufRET = ""

				''2013-07-16 김진영 추가..제휴 등록 안 되야 되는 상품 추가(등록제외상품으로 등록된 시기보다 CJ에 미리 전시상품으로 등록되어있는 경우) saveSellYNItemResult에도 쿼리추가함
				sqlStr = "select count(*) as cnt from db_outMall.dbo.tbl_jaehyumall_not_in_itemid where mallgubun = 'cjmall' and itemid =" & iitemid
				rsget.Open sqlStr, dbget
				If Not(rsget.EOF or rsget.BOF) Then
					notitemId = rsget("cnt")
				End If
				rsget.close

				sqlStr = "select count(*) as cnt from db_AppWish.dbo.tbl_item as i join [db_outMall].dbo.tbl_jaehyumall_not_in_makerid as m on i.makerid = m.makerid where i.itemid = "& iitemid&" and m.mallgubun = 'cjmall'"
				rsget.Open sqlStr, dbget
				If Not(rsget.EOF or rsget.BOF) Then
					notmakerid = rsget("cnt")
				End If
				rsget.close

				If (isOptionExists) Then
					For i = 0 To UBound(ArrRows,2)
						itemoption			= ArrRows(1,i)
						optiontypename		= ArrRows(2,i)
						optionname			= Replace(Replace(db2Html(ArrRows(3,i)),":",""),",","")
						optLimit			= ArrRows(4,i)
						optlimityn			= ArrRows(5,i)
						isUsing				= ArrRows(6,i)
						optsellyn			= ArrRows(7,i)
						preged				= (ArrRows(11,i)=1)
						optNameDiff			= (ArrRows(12,i)=1)
						forceExpired		= (ArrRows(13,i)=1)
						oopt				= ArrRows(14,i)
						ooptCd				= ArrRows(15,i)
						YtoN				= (ArrRows(16,i)=1)
						NtoY				= (ArrRows(17,i)=1)
						DelOpt				= (ArrRows(18,i)=1)


						If ( notitemId > 0 OR notmakerid > 0) Then ''itemoption <> "0000" and
							strRst = ""
							strRst = strRst &"<tns:itemStates>"
							strRst = strRst &"<tns:typeCd>02</tns:typeCd>"							'!!!01=판매코드,02=단품코드
							strRst = strRst &"<tns:itemCd_zip>"&ooptCd&"</tns:itemCd_zip>"
							strRst = strRst &"<tns:chnCls>30</tns:chnCls>"
							strRst = strRst &"<tns:packInd>I</tns:packInd>"						'!!!A-진행, I-일시중단
							strRst = strRst &"</tns:itemStates>"
							OPTYN = "N"
							Pmode = "OptMod"
							bufRET = bufRET + strRst
                        ''기존 소스가 단품 1개에 대해서만 수정되는듯 ==>여러줄로 수정
						ElseIf (forceExpired) or (optNameDiff) or (DelOpt) or (isUsing="N") or (optsellyn="N") or (optlimityn = "Y" AND optLimit <= 5) Then			'한정이고 수량이 5개 이하인 경우 // (isUsing="N") or (optsellyn="N") or 추가 2013/05/31..''2013-12-04 13:30 김진영..optLimit < 5를 optLimit <= 5로 수정
							strRst = ""
							strRst = strRst &"<tns:itemStates>"
							strRst = strRst &"<tns:typeCd>02</tns:typeCd>"							'!!!01=판매코드,02=단품코드
							strRst = strRst &"<tns:itemCd_zip>"&ooptCd&"</tns:itemCd_zip>"
							strRst = strRst &"<tns:chnCls>30</tns:chnCls>"
							strRst = strRst &"<tns:packInd>I</tns:packInd>"						'!!!A-진행, I-일시중단
							strRst = strRst &"</tns:itemStates>"
							OPTYN = "N"
							Pmode = "OptMod"
							bufRET = bufRET + strRst
						ElseIf preged = False AND ooptCd = "" Then			'1.옵션이 추가되는경우 //확인요망. 단품 추가는 "EDT" 03_02 판매상품수정에서 할것.
							''bufRET = bufRET + oCJMallItem.FItemList(0).getcjmallItemModXML("unitY")
					    ELSEIF (preged) and (ooptCd<>"") then
					        strRst = ""
							strRst = strRst &"<tns:itemStates>"
							strRst = strRst &"<tns:typeCd>02</tns:typeCd>"							'!!!01=판매코드,02=단품코드
							strRst = strRst &"<tns:itemCd_zip>"&ooptCd&"</tns:itemCd_zip>"
							strRst = strRst &"<tns:chnCls>30</tns:chnCls>"
							strRst = strRst &"<tns:packInd>A</tns:packInd>"						'!!!A-진행, I-일시중단
							strRst = strRst &"</tns:itemStates>"
							OPTYN = "Y"
							Pmode = "OptMod"
							bufRET = bufRET + strRst
						Else
																		'3.옵션 가격이 변경되는경우


'						ElseIf YtoN = True OR NtoY = True OR DelOpt = True Then			'2.옵션 sellyn이 N 또는 Y로, itemoption테이블의 옵션이 강제 삭제 되는경우 (CJ단품코드를 받은 전제하임) ==> YtoN NtoY 필요없을듯 현재 상태 기준으로
'							strRst = ""
'							strRst = strRst &"<tns:itemStates>"
'							strRst = strRst &"<tns:typeCd>02</tns:typeCd>"							'!!!01=판매코드,02=단품코드
'							strRst = strRst &"<tns:itemCd_zip>"&ooptCd&"</tns:itemCd_zip>"
'							strRst = strRst &"<tns:chnCls>30</tns:chnCls>"
'							If YtoN = True Then
'								OPTYN = "Y"
'								strRst = strRst &"<tns:packInd>A</tns:packInd>"						'!!!A-진행, I-일시중단
'								Pmode = "OptMod"
'							ElseIf NtoY = True Then
'								OPTYN = "N"
'								strRst = strRst &"<tns:packInd>I</tns:packInd>"						'!!!A-진행, I-일시중단
'								Pmode = "OptMod"
'							ElseIf DelOpt = True Then
'								OPTYN = "N"
'								strRst = strRst &"<tns:packInd>I</tns:packInd>"						'!!!A-진행, I-일시중단
'								Pmode = "OptDel"
'							End If
'							strRst = strRst &"</tns:itemStates>"
'
'							bufRET = bufRET + trRst
						End If
					Next
				End If

				'' *********************************************************************************************
				'' 판매코드로 수정할 경우: 단품 전체 수정됨, 단품코드로 수정할 경우 판매코드 가격/상태는 수정안됨.
				buf1 = "<?xml version=""1.0"" encoding=""UTF-8""?>"
				buf1 = buf1 &"<tns:ifRequest xmlns:tns='http://www.example.org/ifpa' tns:ifId='IF_03_03' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation='http://www.example.org/ifpa ../IF_03_03.xsd'>"
				buf1 = buf1 &"<tns:vendorId>411378</tns:vendorId>"
				buf1 = buf1 &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
                buf1 = buf1 &"<tns:itemStates>"
				buf1 = buf1 &"<tns:typeCd>01</tns:typeCd>"							'!!!01=판매코드,02=단품코드
				buf1 = buf1 &"<tns:itemCd_zip>"&oCJMallItem.FItemList(0).FcjmallPrdNo&"</tns:itemCd_zip>"
				buf1 = buf1 &"<tns:chnCls>30</tns:chnCls>"
				buf1 = buf1 &"<tns:packInd>"&CHKIIF(oCJMallItem.FItemList(0).IsSoldOut,"I","A")&"</tns:packInd>"						'!!!A-진행, I-일시중단
				buf1 = buf1 &"</tns:itemStates>"
				getXMLString = buf1 & bufRET & "</tns:ifRequest>"

''				If bufRET = "" Then ''옵션이 없는경우는
''					''getXMLString = "MDT_NOT"							'수정할 것이 없는 경우
''
''					buf1 = "<?xml version=""1.0"" encoding=""UTF-8""? >"
''					buf1 = buf1 &"<tns:ifRequest xmlns:tns='http://www.example.org/ifpa' tns:ifId='IF_03_03' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation='http://www.example.org/ifpa ../IF_03_03.xsd'>"
''					buf1 = buf1 &"<tns:vendorId>411378</tns:vendorId>"
''					buf1 = buf1 &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
''                    buf1 = buf1 &"<tns:itemStates>"
''					buf1 = buf1 &"<tns:typeCd>01</tns:typeCd>"							'!!!01=판매코드,02=단품코드
''					buf1 = buf1 &"<tns:itemCd_zip>"&oCJMallItem.FItemList(0).FcjmallPrdNo&"</tns:itemCd_zip>"
''					buf1 = buf1 &"<tns:chnCls>30</tns:chnCls>"
''					buf1 = buf1 &"<tns:packInd>"&CHKIIF(oCJMallItem.FItemList(0).FSellYN="Y","A","I")&"</tns:packInd>"						'!!!A-진행, I-일시중단
''					buf1 = buf1 &"</tns:itemStates>"
''					getXMLString = buf1 & "</tns:ifRequest>"
''
''			    else
''			        buf1 = "<?xml version=""1.0"" encoding=""UTF-8""? >"
''					buf1 = buf1 &"<tns:ifRequest xmlns:tns='http://www.example.org/ifpa' tns:ifId='IF_03_03' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation='http://www.example.org/ifpa ../IF_03_03.xsd'>"
''					buf1 = buf1 &"<tns:vendorId>411378</tns:vendorId>"
''					buf1 = buf1 &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
''                    buf1 = buf1 &"<tns:itemStates>"
''					buf1 = buf1 &"<tns:typeCd>01</tns:typeCd>"							'!!!01=판매코드,02=단품코드
''					buf1 = buf1 &"<tns:itemCd_zip>"&oCJMallItem.FItemList(0).FcjmallPrdNo&"</tns:itemCd_zip>"
''					buf1 = buf1 &"<tns:chnCls>30</tns:chnCls>"
''					buf1 = buf1 &"<tns:packInd>"&CHKIIF(oCJMallItem.FItemList(0).FSellYN="Y","A","I")&"</tns:packInd>"						'!!!A-진행, I-일시중단
''					buf1 = buf1 &"</tns:itemStates>"
''					getXMLString = buf1 & bufRET & "</tns:ifRequest>"
''				End If
			ElseIf mode = "EDT" Then
				getXMLString = oCJMallItem.FItemList(0).getcjmallItemModXML("unitN")
			ElseIf mode = "QTY" Then
				getXMLString = oCJMallItem.FItemList(0).getcjmallItemQTYXML()
			ElseIf mode = "DateRes" Then
				getXMLString = oCJMallItem.FItemList(0).getcjmallItemDateXML()
			End If
		End If
	ELSEIF (mode="ORDLIST") or (mode="ORDCANCELLIST") or (mode="ORDLISTUP") then
	    strRst = ""
        strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
        strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_04_01"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_04_01.xsd"">"
        strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
        strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
        strRst = strRst &"<tns:contents>"
        IF (mode="ORDLIST") or (mode="ORDLISTUP") THEN
            strRst = strRst &"	<tns:instructionCls>"&"1"&"</tns:instructionCls>"
        ELSEIF (mode="ORDCANCELLIST") then
            strRst = strRst &"	<tns:instructionCls>"&"2"&"</tns:instructionCls>"
        END IF
        strRst = strRst &"	<tns:wbCrtDt>"&iitemid&"</tns:wbCrtDt>" ''조회날짜
        strRst = strRst &"</tns:contents>"
        strRst = strRst &"</tns:ifRequest>"
        getXMLString = strRst
    ELSEIF (mode="CSLIST") then
		'// CS내역일 경우 iitemid 는 날짜이다.
        strRst="<?xml version=""1.0"" encoding=""UTF-8""?>"
        strRst=strRst&"<tns:ifRequest tns:ifId=""IF_04_02"" xmlns:tns=""http://www.example.org/ifpa"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_04_02.xsd "">"
        strRst=strRst&"<tns:vendorId>411378</tns:vendorId>"
        strRst=strRst&"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
        strRst=strRst&"<tns:contents>"
        strRst=strRst&"<tns:wbCrtDt>"&iitemid&"</tns:wbCrtDt>"
        strRst=strRst&"<tns:autoFlag></tns:autoFlag>" ''조회조건 - 자동회수확정여부 Enum(""=전체, 0=N, 1=Y)
        strRst=strRst&"</tns:contents>"
        strRst=strRst&"</tns:ifRequest>"

        getXMLString = strRst
	ELSEIF (mode="commonCD") then
	    strRst = ""
        strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
        strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_02_01"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_02_01.xsd"">"
        strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
        strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
        strRst = strRst &"<tns:lgrpCd>"&iitemid&"</tns:lgrpCd>"
        strRst = strRst &"</tns:ifRequest>"
        getXMLString = strRst
	End If
	SET oCJMallItem = Nothing
End Function

Function getXMLSellyn(byval iitemid, mode, cmd)
	Dim oCJMallItem
	Dim strRst
	SET oCJMallItem = new CCjmall
		oCJMallItem.FRectMode = mode
		oCJMallItem.FRectItemID = iitemid
		''oCJMallItem.FRectMatchCateNotCheck = "on"
		oCJMallItem.getCjmallEditedItemList
		If (oCJMallItem.FREsultCount > 0) Then
			oCJMallItem.FItemList(0).FSellYN = cmd
			getXMLSellyn = oCJMallItem.FItemList(0).getcjmallItemSellStatusDTXML()
		End If
	SET oCJMallItem = Nothing
End Function

Function getOriginName2Code(iname, byref ioriginName)
	Dim sqlStr , retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT top 1 areacode, areaName" & VBCRLF
	sqlStr = sqlStr & " FROM db_temp.dbo.[tbl_cjmall_SourceAreaCode]" & VBCRLF
	sqlStr = sqlStr & " WHERE areaName='"&iname&"'"
	rsget.Open sqlStr,dbget,1
	if (Not rsget.Eof) then
		retVal = rsget("areacode")
		ioriginName = rsget("areaName")
	end if
	rsget.Close

	If (retVal = "") Then
		sqlStr = ""
		sqlStr = sqlStr & " SELECT top 1 areacode, areaName FROM db_temp.dbo.[tbl_cjmall_SourceAreaCode]" & VBCRLF
		sqlStr = sqlStr & " WHERE CharIndex('"&iname&"',areaName) > 0" & VBCRLF
		sqlStr = sqlStr & " or CharIndex(areaName,'"&iname&"') > 0" & VBCRLF
		rsget.Open sqlStr,dbget,1
		If (Not rsget.Eof) then
			retVal = rsget("areacode")
			ioriginName = rsget("areaName")
		End If
		rsget.Close
	End If

	If (retVal = "") Then
		retVal="000"
		ioriginName = "없음"
	End If

	getOriginName2Code=retVal
	Exit Function
End Function

Function getmakerName2Code(iname, byref ioriginName)
	Dim sqlStr , retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT top 1 code, makerName" & VBCRLF
	sqlStr = sqlStr & " FROM db_temp.dbo.tbl_cjmall_makerName" & VBCRLF
	sqlStr = sqlStr & " WHERE makerName='"&iname&"'"
	rsget.Open sqlStr,dbget,1
	if (Not rsget.Eof) then
		retVal = rsget("code")
		ioriginName = rsget("makerName")
	end if
	rsget.Close

	If (retVal = "") Then
		retVal="15210"
		ioriginName = "텐바이텐"
	End If

	getmakerName2Code = retVal
	Exit Function
End Function

Function regCjMallOneItem(byval iitemid, byRef ierrStr)
	''rw  "상품등록잠시중지"
	''regCjMallOneItem = False
	''Exit function
	''response.end
	Dim sqlStr, AssignedRow
	Dim mode : mode = "REG"
	Dim xmlStr : xmlStr = getXMLString(iitemid, mode, "") ''옵션 추가금액 있는 상품은 등록 불가하게..
	Dim cause
	If (xmlStr = "") Then
		ierrStr = "등록불가"
		''등록불가 사유를 뿌림..
		sqlStr = ""
		sqlStr = sqlStr & " SELECT i.itemid, isNULL(R.cjmallStatCD,-9) asCjmallStatCD " & VBCRLF
		sqlStr = sqlStr & " ,i.sellyn,i.limityn,i.limitno,i.limitsold, isnull(PD.CddKey, '') as CddKey, isnull(c.mapCnt,'') as mapCnt, isnull(N.itemid,'') as Nitemid " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_cjmall_regItem as R on i.itemid=R.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_jaehyumall_not_in_itemid as N on i.itemid=N.itemid and N.mallgubun = 'cjmall' " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_cjMall_prdDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid='cjmall' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small " & VBCRLF
		sqlStr = sqlStr & " WHERE i.itemid = "&iitemid
		rsget.Open sqlStr,dbget,1
		If Not(rsget.EOF or rsget.BOF) Then
			If (rsget("asCjmallStatCD") >= 3) Then
				ierrStr = ierrStr & " - 기등록상품"&" :: 상태["&rsget("asCjmallStatCD")&"]"
			End If

			If (rsget("sellyn") <> "Y") Then
			    ierrStr = ierrStr & " - 품절상태"
			End If

			If (rsget("mapCnt") = "0") Then
			    ierrStr = ierrStr & " - 카테고리미매칭상태"
			End If

			If (rsget("CddKey") = "") Then
			    ierrStr = ierrStr & " - 상품미분류상태"
			End If

			If (rsget("limityn") = "Y") and (rsget("limitno") - rsget("limitsold") < CMAXLIMITSELL) Then
				ierrStr = ierrStr & " - 한정수량 부족 ("&rsget("limitno")-rsget("limitsold")&") 개 남음"
				cause = "limitErr"
			End If

			If (rsget("Nitemid") <> "0") Then
			    ierrStr = ierrStr & " - 등록제외상품"
			End If

	    Else
			ierrStr = ierrStr & " - 상품조회불가"
	    End If
		rsget.Close

		''불가 사유를 못찾을 경우
		If (ierrStr = "등록불가") Then
			sqlStr = ""
			sqlStr = sqlStr & " SELECT itemid" & VBCRLF
			sqlStr = sqlStr & " ,count(*) as optCNT" & VBCRLF
			sqlStr = sqlStr & " ,sum(CASE WHEN optAddPrice > 0 then 1 ELSE 0 END) as optAddCNT" & VBCRLF
			sqlStr = sqlStr & " ,sum(CASE WHEN (optsellyn = 'N') or (optlimityn = 'Y' and (optlimitno - optlimitsold < 1)) then 1 ELSE 0 END) as optNotSellCnt" & VBCRLF
			sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_option" & VBCRLF
			sqlStr = sqlStr & " WHERE itemid="&iitemid & VBCRLF
			sqlStr = sqlStr & " and isusing='Y'" & VBCRLF
			sqlStr = sqlStr & " GROUP BY itemid"
			rsget.Open sqlStr,dbget,1
			If Not(rsget.EOF or rsget.BOF) Then
'				If (rsget("optAddCNT") > 0) Then
'					ierrStr = ierrStr & " - 옵션추가 금액 존재상품 등록불가"
'					cause = "optAddPrcExist"
'				End If

				If (rsget("optCnt") - rsget("optNotSellCnt") < 1) Then
					ierrStr = ierrStr & " - 옵션 판매가능상품 없음."
					cause = "noValidOpt"
				End If
			End If
			rsget.Close
		End If

		If (cause <> "") Then
			''제약조건 체크해야..
			sqlStr = ""
			sqlStr = sqlStr & "INSERT INTO db_temp.dbo.tbl_jaehyumall_not_in_itemid " & VBCRLF
			sqlStr = sqlStr & "(itemid, mallgubun, bigo) " & VBCRLF
			sqlStr = sqlStr & " SELECT i.itemid, '"&CMALLNAME&"', '"&cause&"'" & VBCRLF
			sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i" & VBCRLF
			sqlStr = sqlStr & " LEFT JOIN db_temp.dbo.tbl_jaehyumall_not_in_itemid as n on i.itemid = n.itemid " & VBCRLF
			sqlStr = sqlStr & " and n.mallgubun = '"&CMALLNAME&"' " & VBCRLF
			sqlStr = sqlStr & " WHERE i.itemid = "&iitemid & VBCRLF
			sqlStr = sqlStr & " and n.itemid is NULL"
			''' dbget.Execute sqlStr
			''' 넣지 않음.
		End If

		If (ierrStr <> "등록불가") Then
			ierrStr = iitemid &":"& ierrStr
		End If
		regCjMallOneItem = False
		Exit Function
    End If

    IF (isCJ_DebugMode) Then
        CALL XMLFileSave(xmlStr, mode, iitemid)
    End If

    ''등록예정으로 등록.
    sqlStr = ""
    sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_cjmall_regItem " & VBCRLF
    sqlStr = sqlStr & " (itemid, regdate, reguserid, cjmallStatCD, infodiv, cdmkey) " & VBCRLF
    sqlStr = sqlStr & " SELECT i.itemid, getdate(), '"&session("SSBctID")&"', 1, m.infodiv, m.cdmKey " & VBCRLF
    sqlStr = sqlStr & " FROM db_item.dbo.tbl_item i " & VBCRLF
    sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid " & VBCRLF
    sqlStr = sqlStr & " JOIN db_item.dbo.tbl_cjMall_prdDiv_mapping as m on i.cate_large = m.tencatelarge and i.cate_mid = m.tencatemid and i.cate_small = m.tencatesmall and c.infodiv = m.infodiv " & VBCRLF
    sqlStr = sqlStr & " JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as S on i.cate_large = S.tencatelarge and i.cate_mid = S.tencatemid and i.cate_small = S.tencatesmall " & VBCRLF
    sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_cjmall_regItem as R on i.itemid=R.itemid " & VBCRLF
    sqlStr = sqlStr & " WHERE i.itemid = "&iitemid & VBCRLF
    sqlStr = sqlStr & " and R.itemid is NULL and S.mallid = 'cjmall' "
    dbget.Execute sqlStr, AssignedRow

    IF (AssignedRow < 1) Then
    	sqlStr = ""
        sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_cjmall_regItem" & VBCRLF
        sqlStr = sqlStr & " SET cjmallStatCD = 1" & VBCRLF
        sqlStr = sqlStr & " WHERE itemid = "&iitemid & VBCRLF
        sqlStr = sqlStr & " and cjmallStatCD = 0"
        dbget.Execute sqlStr
    End IF

    AssignedRow = 0
    Dim retDoc, sURL
    sURL = cjMallAPIURL
    SET retDoc = xmlSend(sURL, xmlStr)
	    If (isCJ_DebugMode) Then
	        CALL XMLFileSave(retDoc.XML,"RET_"&mode, iitemid)
	    End If
	    regCjMallOneItem = saveCommonItemResult(retDoc, mode, iitemid)
    SET retDoc = Nothing
End Function

Function editCjmallOneItem(byval iitemid, byRef ierrStr)    ''상품 정보 수정
	Dim sqlStr, AssignedRow
	Dim mode : mode = "EDT"
	Dim xmlStr : xmlStr = getXMLString(iitemid, mode, "")

	If (xmlStr="") Then
		ierrStr = "수정불가"
		editCjmallOneItem = False
		Exit Function
	End If

	If (isCJ_DebugMode) Then
		CALL XMLFileSave(xmlStr, mode, iitemid)
	End If

	Dim retDoc, sURL
	sURL = cjMallAPIURL
	Set retDoc = xmlSend (sURL, xmlStr)
		If (isCJ_DebugMode) Then
			Call XMLFileSave(retDoc.XML,"RET_"&mode,iitemid)
		End If
		Call saveCommonItemResult(retDoc, mode, iitemid)
	Set retDoc = Nothing
End Function



Function editPriceCjmallOneItem(byval iitemid, byRef ierrStr)    ''상품 가격 수정
	Dim sqlStr, AssignedRow
	Dim mode : mode = "PRI"
	Dim xmlStr
	Dim isOptAddExists : isOptAddExists = false

	''if (iitemid=813141) then isOptAddExists = IsOptionAddPriceExistItem(iitemid)

	xmlStr = getXMLString(iitemid, mode, isOptAddExists)

	If (xmlStr="") Then
		ierrStr = "수정불가"
		editPriceCjmallOneItem = False
		Exit Function
	End If

	If (isCJ_DebugMode) Then
		CALL XMLFileSave(xmlStr, mode, iitemid)
	End If

	Dim retDoc, sURL
	sURL = cjMallAPIURL
	Set retDoc = xmlSend (sURL, xmlStr)
		If (isCJ_DebugMode) Then
			Call XMLFileSave(retDoc.XML,"RET_"&mode,iitemid)
		End If
		Call saveCommonItemResult(retDoc, mode, iitemid)
	Set retDoc = Nothing

	if (isOptAddExists) then
	    rw "isOptAddExists"
	    CALL editPriceCjmallOneItemByOption(iitemid,ierrStr,FALSE)
	end if
End Function

Function editPriceCjmallOneItemByOption(byval iitemid, byRef ierrStr, byval isRetry)    ''상품 단품 가격 수정
    Dim xmlStr
    Dim mode : mode = "PRI"
    Dim retDoc, sURL
    sURL = cjMallAPIURL
    xmlStr = getXMLString(iitemid, mode, false)
    If (isCJ_DebugMode) Then
		CALL XMLFileSave(xmlStr, mode, iitemid)
	End If
    Set retDoc = xmlSend (sURL, xmlStr)
	If (isCJ_DebugMode) Then
		Call XMLFileSave(retDoc.XML,"RET_"&mode,iitemid)
	End If
	if (isRetry) then
	    mode = "PRI_RE"
	end if
	Call saveCommonItemResult(retDoc, mode, iitemid)
    Set retDoc = Nothing
end function

Function editSellPriceCjmallOneItem(byval iitemid, byRef ierrStr)    ''상품 가격 수정
	Dim sqlStr, AssignedRow
	Dim mode : mode = "PRI2"
	Dim xmlStr
	Dim isOptAddExists : isOptAddExists = false

	xmlStr = getXMLString(iitemid, mode, isOptAddExists)

	If (xmlStr="") Then
		ierrStr = "수정불가"
		editSellPriceCjmallOneItem = False
		Exit Function
	End If

	If (isCJ_DebugMode) Then
		CALL XMLFileSave(xmlStr, mode, iitemid)
	End If

	Dim retDoc, sURL
	sURL = cjMallAPIURL
	Set retDoc = xmlSend (sURL, xmlStr)
		If (isCJ_DebugMode) Then
			Call XMLFileSave(retDoc.XML,"RET_"&mode,iitemid)
		End If
		Call saveCommonItemResult(retDoc, mode, iitemid)
	Set retDoc = Nothing

End Function

'총 기간 리스트 => 특정 날짜 리스트로 변경
Function listCjMallItem(theday)
	Dim sqlStr, AssignedRow
	Dim mode : mode = "LIST"
	Dim xmlStr : xmlStr = getXMLString(theday, mode, "")
	Dim cause
	If (xmlStr = "") Then
		listCjMallItem = False
		Exit Function
    End If

	If (isCJ_DebugMode) Then
		CALL XMLFileSave(xmlStr, mode, theday)
	End If

    Dim retDoc, sURL
    sURL = cjMallAPIURL
    SET retDoc = xmlSend(sURL, xmlStr)
	    If (isCJ_DebugMode) Then
	        CALL XMLFileSave(retDoc.XML, "RET_"&mode, theday)
	    End If
		listCjMallItem = saveListResult(retDoc, mode, "")
    SET retDoc = Nothing
End Function

'일정기간 리스트
Function daylistCjMallItem(sday)
	Dim sqlStr, AssignedRow
	Dim mode : mode = "DayLIST"
	Dim xmlStr : xmlStr = getXMLString(sday, mode, "")
	Dim cause
	If (xmlStr = "") Then
		daylistCjMallItem = False
		Exit Function
    End If

	If (isCJ_DebugMode) Then
		CALL XMLFileSave(xmlStr, mode, sday)
	End If

    Dim retDoc, sURL
    sURL = cjMallAPIURL
    SET retDoc = xmlSend(sURL, xmlStr)
	    If (isCJ_DebugMode) Then
	        CALL XMLFileSave(retDoc.XML, "RET_"&mode, sday)
	    End If
		daylistCjMallItem = saveListResult(retDoc, mode, "")
    SET retDoc = Nothing
End Function

'상품별 Check // itemid_option 방식으로 변경, CJ판매코드가 있을경우 Cj판매코드로 조회.
Function oneCjMallItemConfirm(iitemid, ierrStr)
	Dim sqlStr, AssignedRow
	Dim mode : mode = "cjItemCheck"
	Dim cause
	Dim cjMallPrdNo : cjMallPrdNo = getCjMallPrdNoByItemID(iitemid)
	Dim firstItemoption

	Dim xmlStr
	if (cjMallPrdNo<>"") then
	    xmlStr = getXMLString(cjMallPrdNo, mode, "cjPrdno")
	else
	    ''검색결과가 없으면 시간이 오래거림..?
	    firstItemoption = getCjMallfirstItemoption(iitemid)
		If iitemid = "899506" Then
			xmlStr = getXMLString(iitemid&"_Q"&firstItemoption, mode, "")
		Else
	    	xmlStr = getXMLString(iitemid&"_"&firstItemoption, mode, "")
	    End If
    end if

	If (xmlStr = "") Then
		oneCjMallItemConfirm = False
		Exit Function
    End If

	If (isCJ_DebugMode) Then
		CALL XMLFileSave(xmlStr, mode, iitemid)
	End If

    Dim retDoc, sURL
    sURL = cjMallAPIURL
    SET retDoc = xmlSend(sURL, xmlStr)
	    If (isCJ_DebugMode) Then
	        CALL XMLFileSave(retDoc.XML, "RET_"&mode, iitemid)
	    End If
	    ''rw retDoc.XML
		oneCjMallItemConfirm = saveListResult(retDoc, mode, iitemid)
    SET retDoc = Nothing

    ''cj코드로 재시도


    if (firstItemoption<>"0000") and (cjMallPrdNo="") then
        cjMallPrdNo = getCjMallPrdNoByItemID(iitemid)
        if (cjMallPrdNo<>"") then
            xmlStr = getXMLString(cjMallPrdNo, mode, "cjPrdno")

            If (xmlStr = "") Then
        		oneCjMallItemConfirm = False
        		Exit Function
            End If

            If (isCJ_DebugMode) Then
        		''CALL XMLFileSave(xmlStr, mode, iitemid)
        	End If

        	sURL = cjMallAPIURL
            SET retDoc = xmlSend(sURL, xmlStr)
        	    If (isCJ_DebugMode) Then
        	        ''CALL XMLFileSave(retDoc.XML, "RET_"&mode, iitemid)
        	    End If
        		oneCjMallItemConfirm = saveListResult(retDoc, mode, iitemid)
            SET retDoc = Nothing
        end if
    end if
End Function

'CJ주문내역 조회
Function getCjOrderList(imode,sday) ''"ORDLIST" , "ORDCANCELLIST"
    Dim mode : mode = imode
	Dim xmlStr : xmlStr = getXMLString(sday, mode, "")

	If (xmlStr = "") Then
		getCjOrderList = False
		Exit Function
    End If

    Dim retDoc, sURL
    sURL = cjMallAPIURL

    SET retDoc = xmlSend(sURL, xmlStr)
    'rw retDoc.XML
	    If (isCJ_DebugMode) Then
	        CALL XMLFileSave(retDoc.XML, "RET_"&mode, sday)
	    End If
		getCjOrderList = saveORDListResult(retDoc, mode, sday)
    SET retDoc = Nothing
End Function

'CJ CS내역 조회
Function getCjCsList(imode,sday)
    Dim mode : mode = imode
	Dim xmlStr : xmlStr = getXMLString(sday, mode, "")

	If (xmlStr = "") Then
		getCjCsList = False
		Exit Function
    End If

    Dim retDoc, sURL
    sURL = cjMallAPIURL

    SET retDoc = xmlSend(sURL, xmlStr)
	''response.write retDoc.XML
	    If (isCJ_DebugMode) Then
	        CALL XMLFileSave(retDoc.XML, "RET_"&mode, sday)
	    End If
		getCjCsList = saveCSListResult(retDoc, mode, sday)
    SET retDoc = Nothing
End Function

'CJ공통코드 조회
Function getcjCommonCodeList(ccd)
    Dim mode : mode = "commonCD"
	Dim xmlStr : xmlStr = getXMLString(ccd, mode, "")

	If (xmlStr = "") Then
		getcjCommonCodeList = False
		Exit Function
    End If

    Dim retDoc, sURL
    sURL = cjMallAPIURL

    SET retDoc = xmlSend(sURL, xmlStr)
	    If (isCJ_DebugMode) Then
	        CALL XMLFileSave(retDoc.XML, "RET_"&mode, sday)
	    End If
    SET retDoc = Nothing
End Function


Function editDTCjmallOneItem(byval iitemid, byRef ierrStr)      ''단품 재고정보 수정
	Dim sqlStr, AssignedRow
	Dim mode : mode = "MDT"
	Dim xmlStr : xmlStr = getXMLString(iitemid, mode, "")

	If xmlStr = "MDT_NOT" Then
		rw "단품 판매설정 - 수정할 것 없음"
		Exit Function
	End If

	Dim optMode : optMode = getMode()
	Dim	optgetOPTYN : optgetOPTYN = getOPTYN()

	If (isCJ_DebugMode) Then
		CALL XMLFileSave(xmlStr, mode, iitemid)
	End If

	Dim retDoc, sURL
	sURL = cjMallAPIURL
	Set retDoc = xmlSend (sURL, xmlStr)
		If (isCJ_DebugMode) Then
			Call XMLFileSave(retDoc.XML,"RET_"&mode,iitemid)
		End If

		If optMode = "OptMod" or optMode = "OptDel" Then
			If optMode = "OptDel" Then
				mode = "Del"
			End If
			Call saveSellYNItemResult(retDoc, mode, iitemid, optgetOPTYN)
		Else
			Call saveCommonItemResult(retDoc, mode, iitemid)
		End If
	Set retDoc = Nothing
End Function


'선택상품정보 수정용
Function editSellStatusCjmallOneItem(byval iitemid, byRef ierrStr, cmd)
	Dim sqlStr, AssignedRow
	Dim mode : mode = "SLD"
	Dim xmlStr : xmlStr = getXMLSellyn(iitemid,mode,cmd)

	If (xmlStr = "") Then
		ierrStr = "수정불가"
		editSellStatusCjmallOneItem = False
		Exit Function
	End If

	If (isCJ_DebugMode) Then
		CALL XMLFileSave(xmlStr, mode, iitemid)
	End If

	Dim retDoc, sURL
	sURL = cjMallAPIURL
	Set retDoc = xmlSend (sURL, xmlStr)
		If (isCJ_DebugMode) Then
			CALL XMLFileSave(retDoc.XML, "RET_"&mode, iitemid)
		End If
		Call saveSellYNItemResult(retDoc, mode, iitemid, cmd)
	Set retDoc = Nothing
End Function

''상품 수량 수정
Function editqtyCjmallOneItem(byval iitemid, byRef ierrStr, cmd)
	Dim sqlStr, AssignedRow
	Dim mode : mode = "QTY"
	Dim xmlStr : xmlStr = getXMLString(iitemid, mode, "")

	If (xmlStr="") Then
		ierrStr = "단품수량 - 수량 수정불가"
		editqtyCjmallOneItem = False
		Exit Function
	End If

	If (isCJ_DebugMode) Then
		CALL XMLFileSave(xmlStr, mode, iitemid)
	End If

	Dim retDoc, sURL
	sURL = cjMallAPIURL
	Set retDoc = xmlSend (sURL, xmlStr)
		If (isCJ_DebugMode) Then
			Call XMLFileSave(retDoc.XML,"RET_"&mode,iitemid)
		End If
		Call saveCommonItemResult(retDoc, mode, iitemid)
	Set retDoc = Nothing
End Function

''상품 수량 수정
Function editDateCjmallOneItem(byval iitemid, byRef ierrStr, cmd)
	Dim sqlStr, AssignedRow
	Dim mode : mode = "DateRes"
	Dim xmlStr : xmlStr = getXMLString(iitemid, mode, "")

	If (xmlStr="") Then
		ierrStr = "날짜 수정불가"
		editDateCjmallOneItem = False
		Exit Function
	End If

	If (isCJ_DebugMode) Then
		CALL XMLFileSave(xmlStr, mode, iitemid)
	End If

	Dim retDoc, sURL
	sURL = cjMallAPIURL
	Set retDoc = xmlSend (sURL, xmlStr)
		If (isCJ_DebugMode) Then
			Call XMLFileSave(retDoc.XML,"RET_"&mode,iitemid)
		End If
		Call saveCommonItemResult(retDoc, mode, iitemid)
	Set retDoc = Nothing
End Function


'선택 상품 일시중단/판매용
Function saveSellYNItemResult(retDoc, mode, prdno, smdcd)
	Dim errorMsg
	Dim sqlStr
	Dim AssignedRow, successYn
	Dim itemCd_zip, packInd, typeCd
	Dim Nodes, SubNodes, notitemId, notmakerid

	If mode = "MDT" Then
		Set Nodes = retDoc.getElementsByTagName("ns1:itemStates")
		If (Not (retDoc is Nothing)) Then
			IF application("Svr_Info")="Dev" THEN
				On Error Resume Next
			End If

			sqlStr = ""
			sqlStr = "select count(*) as cnt from db_outMall.dbo.tbl_jaehyumall_not_in_itemid where mallgubun = 'cjmall' and itemid =" & iitemid
			rsCTget.Open sqlStr, dbCTget
			If Not(rsCTget.EOF or rsCTget.BOF) Then
				notitemId = rsCTget("cnt")
			End If
			rsCTget.close

			sqlStr = ""
			sqlStr = "select count(*) as cnt from db_AppWish.dbo.tbl_item as i join [db_outMall].dbo.tbl_jaehyumall_not_in_makerid as m on i.makerid = m.makerid where i.itemid = "& iitemid&" and m.mallgubun = 'cjmall'"
			rsCTget.Open sqlStr, dbCTget
			If Not(rsCTget.EOF or rsCTget.BOF) Then
				notmakerid = rsCTget("cnt")
			End If
			rsCTget.close

			For each SubNodes in Nodes
			    errorMsg    = SubNodes.getElementsByTagName("ns1:errorMsg").item(0).text
			    typeCd      = SubNodes.getElementsByTagName("ns1:typeCd").item(0).text
				itemCd_zip 	= SubNodes.getElementsByTagName("ns1:itemCd_zip").item(0).text
				packInd		= SubNodes.getElementsByTagName("ns1:packInd").item(0).text
				successYn	= SubNodes.getElementsByTagName("ns1:successYn").item(0).text

				rw typeCd&","&itemCd_zip&","&packInd&","&CHKIIF(successYn<>"true",errorMsg,"")

                if (typeCd="01") then
                    sqlStr = ""                                                                 ''2013/06/20추가
                    sqlStr = sqlStr & " update db_outMall.dbo.tbl_cjmall_regItem" & VBCRLF
                    sqlStr = sqlStr & " set cjmallLastUpdate = getdate()" & VBCRLF
					If (notitemId > 0) OR (notmakerid > 0) Then
						sqlStr = sqlStr & " ,cjmallSellyn = 'N' " & VBCRLF
					Else
						sqlStr = sqlStr & " ,cjmallSellyn = 'Y' " & VBCRLF
					End If
                    sqlStr = sqlStr & " WHERE itemid = '"&prdno&"'  " & VBCRLF
                    sqlStr = sqlStr & " and cjmallPrdNo='"&itemCd_zip&"'" & VBCRLF
                    dbCTget.Execute sqlStr,AssignedRow
                elseif (typeCd="02") then
    				sqlStr = ""
    				sqlStr = sqlStr & " UPDATE [db_outMall].[dbo].tbl_OutMall_regedoption  " & VBCRLF
    				sqlStr = sqlStr & " SET outmallSellyn = '"&chkiif(packInd="A","Y","N")&"'" & VBCRLF
    				sqlStr = sqlStr & " , lastupdate = getdate() " & VBCRLF
    				sqlStr = sqlStr & " WHERE itemid = '"&prdno&"'  " & VBCRLF
    				sqlStr = sqlStr & " and outmallOptCode = '"&itemCd_zip&"' " & VBCRLF
    				sqlStr = sqlStr & " and mallid='"&CMALLNAME&"'"&VbCRLF

    				dbCTget.Execute sqlStr,AssignedRow
    			end if
			Next

			'2013-11-15 10:30 김진영 하단 쿼리 추가..// 이유 : 옵션전체가 다 바뀌었을 때 판매상태를 N으로 돌린다는 의미
			Dim sellynCnt
			sqlStr = ""
			sqlStr = sqlStr & " SELECT count(*) as cnt FROM db_outMall.dbo.tbl_Outmall_regedoption WHERE itemid="&prdno&" and mallid = 'cjmall' and outmallSellyn = 'Y' "
			rsCTget.Open sqlStr, dbCTget
				sellynCnt = rsCTget("cnt")
			rsCTget.Close
            sqlStr = ""
            sqlStr = sqlStr & " update db_outMall.dbo.tbl_cjmall_regItem set cjmallLastUpdate = getdate()" & VBCRLF
			If (sellynCnt > 0) Then
				sqlStr = sqlStr & " ,cjmallSellyn = 'Y' " & VBCRLF
			ElseIf (notitemId > 0) OR (notmakerid > 0) Then
				sqlStr = sqlStr & " ,cjmallSellyn = 'N' " & VBCRLF
			Else
				sqlStr = sqlStr & " ,cjmallSellyn = 'N' " & VBCRLF
			End If
            sqlStr = sqlStr & " WHERE itemid = '"&prdno&"'  " & VBCRLF
            dbCTget.Execute sqlStr
           '2013-11-15 10:30 김진영 하단 쿼리 추가..// 이유 : 옵션전체가 다 바뀌었을 때 판매상태를 N으로 돌린다는 의미 끝
		End If
		saveSellYNItemResult=true
	ElseIf mode = "Del" Then
		itemCd_zip = retDoc.getElementsByTagName("ns1:itemCd_zip").item(0).text
'		sqlStr = ""
'		sqlStr = sqlStr & " DELETE FROM [db_outMall].[dbo].tbl_OutMall_regedoption " & VBCRLF
'		sqlStr = sqlStr & " WHERE itemid = '"&prdno&"'  " & VBCRLF
'		sqlStr = sqlStr & " and outmallOptCode = '"&itemCd_zip&"' " & VBCRLF
'		sqlStr = sqlStr & " and mallid = 'cjmall' " & VBCRLF
'		dbCTget.Execute sqlStr,AssignedRow
		saveSellYNItemResult=true
	Else
		If (prdno <> "") Then
			sqlStr = ""
			sqlStr = sqlStr & " UPDATE R" & VBCRLF
			sqlStr = sqlStr & " SET cjmallSellYn = '"&smdcd&"'" & VBCRLF                        ''cjmallLastUpdate = getdate()" & VBCRLF
			sqlStr = sqlStr & " ,accFailCNT=0" & VBCRLF                 ''실패회수 초기화
			sqlStr = sqlStr & " FROM db_outMall.dbo.tbl_cjmall_regItem as R" & VBCRLF
			sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_item as i on R.itemid = i.itemid" & VBCRLF
			sqlStr = sqlStr & " WHERE R.itemid = "&prdno&""   & VBCRLF
			dbCTget.Execute sqlStr,AssignedRow
			saveSellYNItemResult=true
		End If
		''errorMsg = retDoc.getElementsByTagName("ns1:errorMsg").item(0).text
	End If

	If (isCJ_DebugMode) Then
		rw prdno &"_"&mode&"_"&errorMsg
	End If
End Function

'주문내역 저장용
Function saveORDListResult(retDoc, mode, sday)
    Dim Nodes, masterSubNodes, detailSubNodes, detailSubNodeItem, ErrNode, errorMsg
    Dim isErrExists : isErrExists = false
    Dim ordNo,custNm,custTelNo,custDeliveryCost

    Dim ordGSeq, ordDSeq, ordWSeq, ordDtlCls, ordDtlClsCd, wbCrtDt, outwConfDt, delivDtm, cnclInsDtm
    Dim oldordNo, toutYn, chnNm, receverNm, recvName, zipno, addr_1, addr_2, addr, telno, cellno
    Dim msgSpec, delvplnDt, packYn, itemNm, itemCd, unitCd, itemName, unitNm, contItemCd, wbIdNo
    Dim outwQty, realslAmt, outwAmt, delivInfo, promGiftSpec, cnclRsn, cnclRsnSpec, ordDtm, juminNum, dccouponCjhs, dccouponVendor
    Dim dtlCnt

	dim IsDetailItemAllCancel, IsCancelOrgOrder
	dim strSql

    Dim requireDetail, orderDlvPay, orderCsGbn, ierrStr, ierrCode
    dim succCnt : succCnt=0
    dim failCnt : failCnt=0
    dim skipCnt : skipCnt=0
    dim retVal

    Set Nodes = retDoc.getElementsByTagName("ns1:errorMsg")
    If (Not (Nodes is Nothing)) Then
        For each ErrNode in Nodes
            errorMsg = Nodes.item(0).text
            isErrExists = true
            rw "["&sday&"]"&errorMsg
        next
    end if

    if (Not isErrExists) then
        Set Nodes = retDoc.getElementsByTagName("ns1:instruction")

        If (Not (Nodes is Nothing)) Then
            For each masterSubNodes in Nodes
                ordNo       = masterSubNodes.getElementsByTagName("ns1:ordNo")(0).Text	        '주문번호
                custNm      = masterSubNodes.getElementsByTagName("ns1:custNm")(0).Text	        '주문자
                custTelNo   = masterSubNodes.getElementsByTagName("ns1:custTelNo")(0).Text	    '주문자 전화
                custDeliveryCost = masterSubNodes.getElementsByTagName("ns1:custDeliveryCost")(0).Text	'배송비

                Set detailSubNodes = masterSubNodes.getElementsByTagName("ns1:instructionDetail")

                ''rw ordNo&"|"&custNm&"|"&custTelNo&"|"&custDeliveryCost

                dtlCnt = 1
                If (Not (detailSubNodes is Nothing)) Then
                    For each detailSubNodeItem in detailSubNodes
                        requireDetail = ""
                        ierrStr = ""

                        ordGSeq = detailSubNodeItem.getElementsByTagName("ns1:ordGSeq")(0).Text	    '[ID:주문상품순번], 001
                        ordDSeq = detailSubNodeItem.getElementsByTagName("ns1:ordDSeq")(0).Text	    '[ID:주문상세순번], 001
                        ordWSeq = detailSubNodeItem.getElementsByTagName("ns1:ordWSeq")(0).Text	    '[ID:주문처리순번], 001
                        ordDtlCls = detailSubNodeItem.getElementsByTagName("ns1:ordDtlCls")(0).Text	        ' 주문정보 - 주문구분, 주문
                        ordDtlClsCd = detailSubNodeItem.getElementsByTagName("ns1:ordDtlClsCd")(0).Text	    ' 주문정보 - 주문구분코드, 10
                        wbCrtDt = detailSubNodeItem.getElementsByTagName("ns1:wbCrtDt")(0).Text	            ' 주문정보 - 지시일자, 2013-05-22+09:00
                        ''outwConfDt	'주문정보 - 출고확정일자
                        ''delivDtm	    '주문정보 - 배송완료일
                        ''cnclInsDtm	'주문정보 - 취소일자
                        ''oldordNo	    '주문정보 - 원주문번호
                        toutYn = detailSubNodeItem.getElementsByTagName("ns1:toutYn")(0).Text	            '주문정보 - 기출하구분(Y-기출하,N-정상출하), N
                        chnNm = detailSubNodeItem.getElementsByTagName("ns1:chnNm")(0).Text	                '주문정보 - 채널구분, INTERNET

                        if (mode<>"ORDLISTUP") then
                        receverNm = detailSubNodeItem.getElementsByTagName("ns1:receverNm")(0).Text	        '주문정보 - 인수자, 채현아
                        end if

                        'recvName	    '주문정보 - 수취인 영문명
                        zipno = detailSubNodeItem.getElementsByTagName("ns1:zipno")(0).Text	                '주문정보 - 우편번호, 03082
                        addr_1 = detailSubNodeItem.getElementsByTagName("ns1:addr_1")(0).Text	            '주문정보 - 주소, 서울시 종로구 대학로 57
                        addr_2 = detailSubNodeItem.getElementsByTagName("ns1:addr_2")(0).Text	            '주문정보 - 상세주소, 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐
                        'addr	        '주문정보 - 주소
                        telno = detailSubNodeItem.getElementsByTagName("ns1:telno")(0).Text	                '주문정보 - 인수자tel, 02)973-8514
                        cellno = detailSubNodeItem.getElementsByTagName("ns1:cellno")(0).Text	            '주문정보 - 인수자HP, 010)2715-8514
                        'msgSpec	    '주문정보 - 배송참고
                        'delvplnDt	    '주문정보 - 배송예정일
                        packYn = detailSubNodeItem.getElementsByTagName("ns1:packYn")(0).Text	            '상품정보 - 세트여부, 일반
                        'itemNm	        '상품정보 - 세트상품명
                        itemCd = detailSubNodeItem.getElementsByTagName("ns1:itemCd")(0).Text	            '상품정보 - 판매코드, 21899852
                        unitCd = detailSubNodeItem.getElementsByTagName("ns1:unitCd")(0).Text	            '상품정보 - 단품코드, 10047125217
                        itemName = detailSubNodeItem.getElementsByTagName("ns1:itemName")(0).Text	        '상품정보 - 판매상품명, 24K Gold 전자파차단스티커
                        unitNm = detailSubNodeItem.getElementsByTagName("ns1:unitNm")(0).Text	            '상품정보 - 단품상세, ES-01 잘될꺼야
                        contItemCd = detailSubNodeItem.getElementsByTagName("ns1:contItemCd")(0).Text	    '상품정보 - 협력사상품코드, 279751_0011
                        wbIdNo = detailSubNodeItem.getElementsByTagName("ns1:wbIdNo")(0).Text	            '상품정보 - 운송장식별번호, 20000420537940
                        outwQty = detailSubNodeItem.getElementsByTagName("ns1:outwQty")(0).Text	            '상품정보 - 수량, 1.0
                        realslAmt = detailSubNodeItem.getElementsByTagName("ns1:realslAmt")(0).Text	        '상품정보 - 판매가, 1800.0
                        outwAmt = detailSubNodeItem.getElementsByTagName("ns1:outwAmt")(0).Text	            '상품정보 - 고객결제가, 1800.0  :: 수량*판매가 인지, 수량*실판매가인지 확인 = 수량*실판매가
                        'delivInfo	    '기타정보 - 비고
                        'promGiftSpec	'기타정보 - 사은품내역
                        'juminNum       '주문정보-주민번호(아님), 발송 금지!
                        'cnclRsn	    '기타정보 - 교환/취소사유
                        'cnclRsnSpec	'기타정보 - 교환/취소사유상세
                        ordDtm = detailSubNodeItem.getElementsByTagName("ns1:ordDtm")(0).Text	            '주문정보-주문일시, 2013-05-22 15:05:02


                        ''필수로 안넘어오는정보들.
                        outwConfDt =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:outwConfDt")(0) Is Nothing)) Then
                            outwConfDt = detailSubNodeItem.getElementsByTagName("ns1:outwConfDt")(0).Text       '주문정보 - 출고확정일자
                        end if
                        delivDtm =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:delivDtm")(0) Is Nothing)) Then
                            delivDtm = detailSubNodeItem.getElementsByTagName("ns1:delivDtm")(0).Text        '주문정보 - 배송완료일
                        end if
                        cnclInsDtm =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:cnclInsDtm")(0) Is Nothing)) Then
                            cnclInsDtm = detailSubNodeItem.getElementsByTagName("ns1:cnclInsDtm")(0).Text        '주문정보 - 취소일자
                        end if
                        oldordNo =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:oldordNo")(0) Is Nothing)) Then
                            oldordNo = detailSubNodeItem.getElementsByTagName("ns1:oldordNo")(0).Text        '주문정보 - 원주문번호
                        end if
                        recvName =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:recvName")(0) Is Nothing)) Then
                            recvName = detailSubNodeItem.getElementsByTagName("ns1:recvName")(0).Text        '주문정보 - 수취인 영문명
                        end if
                        addr =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:addr")(0) Is Nothing)) Then
                            addr = detailSubNodeItem.getElementsByTagName("ns1:addr")(0).Text        '주문정보 - 주소all?
                        end if
                        msgSpec =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:msgSpec")(0) Is Nothing)) Then
                            msgSpec = detailSubNodeItem.getElementsByTagName("ns1:msgSpec")(0).Text        '주문정보 -배송참고
                        end if
                        delvplnDt =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:delvplnDt")(0) Is Nothing)) Then
                            delvplnDt = detailSubNodeItem.getElementsByTagName("ns1:delvplnDt")(0).Text        '주문정보 -배송예정일
                        end if
                        itemNm =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:itemNm")(0) Is Nothing)) Then
                            itemNm = detailSubNodeItem.getElementsByTagName("ns1:itemNm")(0).Text        '상품정보 -세트상품명
                        end if
                        juminNum =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:juminNum")(0) Is Nothing)) Then
                            juminNum = detailSubNodeItem.getElementsByTagName("ns1:juminNum")(0).Text       '주문정보-주민번호(아님), 발송 금지!
                        end if
                        dccouponCjhs =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:dccouponCjhs")(0) Is Nothing)) Then
                            dccouponCjhs = detailSubNodeItem.getElementsByTagName("ns1:dccouponCjhs")(0).Text       '주문정보 - 할인(당사부담)금액
                        end if
                        dccouponVendor =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:dccouponVendor")(0) Is Nothing)) Then
                            dccouponVendor = detailSubNodeItem.getElementsByTagName("ns1:dccouponVendor")(0).Text      '주문정보 - 할인(협력사부담)금액
                        end if

                        orderDlvPay = custDeliveryCost
                        if (dtlCnt>1) then
                            orderDlvPay = 0 ''첫번째 값만 넣음.
                        end if

                        orderCsGbn = ""
						if (toutYn <> "N") then
							'// 기출하 주문 스킵
							ordDtlClsCd = "99"
						end if
                        if (ordDtlClsCd="10") then
                            orderCsGbn="0"
                        elseif (ordDtlClsCd="20") then
                            orderCsGbn="2"  ''취소
                        end if

                        ''requireDetail = juminNum '' 주문제작문구
                        if (juminNum<>"") then                          ''2013/06/05 수정: 주문제작문구 아님?.
                            if (msgSpec<>"") then
                                msgSpec=msgSpec&VbCRLF&juminNum
                            else
                                msgSpec=juminNum
                            end if
                        end if

                        ierrCode = 0
                        ierrStr  = ""

                        if (mode="ORDLIST") then
                            if (orderCsGbn<>"") then

    							IsDetailItemAllCancel = False
    							IsCancelOrgOrder = False

    							if (orderCsGbn = "2") then
    								'// 취소
    								strSql = " select matchState, orderDlvPay, orgOrderCNT from db_temp.dbo.tbl_xSite_TMPOrder "
    								strSql = strSql + " where SellSite = 'cjmall' and OutMallOrderSerial = '" + CStr(ordNo) + "' and OrgDetailKey = '" & ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq & "' "
    								''rw strSql
    								rsget.Open strSql,dbget,1
    								if (Not rsget.Eof) then
    									if (CLng(outwQty) = rsget("orgOrderCNT")) then
    										'// 특정상품 전체취소
    										IsDetailItemAllCancel = True
    										if (rsget("matchState") = "I") then
    											'// 주문입력이전
    											IsCancelOrgOrder = True
    										end if
    									end if
    								end if
    								rsget.Close

    								if (IsDetailItemAllCancel and IsCancelOrgOrder) then
    									strSql = " update db_temp.dbo.tbl_xSite_TMPOrder set matchState = 'D' "
    									strSql = strSql + " where SellSite = 'cjmall' and OutMallOrderSerial = '" + CStr(ordNo) + "' and OrgDetailKey = '" & ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq & "' and matchState = 'I' "
    									''rw strSql
    									rsget.Open strSql, dbget, 1
    								end if
    							end if

                                '''899506_Q0055 ::
                                if (LEFT(splitvalue(contItemCd,"_",1),1)="Q") then
                                    contItemCd = replace(contItemCd,"Q","")
                                end if
                                if (outwQty<>"0") and (outwQty<>"1") and (outwQty<>"-1") and (outwQty<>"") then
                                    outwAmt = CLNG(outwAmt/outwQty)
                                end if
    							if (IsDetailItemAllCancel) then
    								'// 우선 수량 전체취소만 처리(수량 일부취소는 내역 입력되면 처리)
    								retVal = saveORDOneTmp(ordNo,ordDtm,splitvalue(contItemCd,"_",0),splitvalue(contItemCd,"_",1),itemName, unitNm _
    										, custNm , custTelNo, custTelNo _
    										, receverNm, telno, cellno, LEFT(zipno,3)&"-"&Right(zipno,3), addr_1, addr_2 _
    										, realslAmt, outwAmt, CLNG(outwQty), ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq & "-CA" _
    										, msgSpec, requireDetail, orderDlvPay, orderCsGbn _
    										, ierrCode, ierrStr)

    								'// 원주문 삭제되었으면 CS도 삭제
    								strSql = " if exists (select OutMallOrderSeq from db_temp.dbo.tbl_xSite_TMPOrder where SellSite = 'cjmall' and OutMallOrderSerial = '" + CStr(ordNo) + "' and OrgDetailKey = '" & ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq & "' and matchState = 'D') "
    								strSql = strSql + " begin "
    								strSql = strSql + " 	update db_temp.dbo.tbl_xSite_TMPOrder set matchState = 'D' where SellSite = 'cjmall' and OutMallOrderSerial = '" + CStr(ordNo) + "' and OrgDetailKey = '" & ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq & "-CA' and matchState = 'I' "
    								strSql = strSql + " end "
    								rsget.Open strSql, dbget, 1
    							else
    								retVal = saveORDOneTmp(ordNo,ordDtm,splitvalue(contItemCd,"_",0),splitvalue(contItemCd,"_",1),itemName, unitNm _
    										, custNm , custTelNo, custTelNo _
    										, receverNm, telno, cellno, LEFT(zipno,3)&"-"&Right(zipno,3), addr_1, addr_2 _
    										, realslAmt, outwAmt, CLNG(outwQty), ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq _
    										, msgSpec, requireDetail, orderDlvPay, orderCsGbn _
    										, ierrCode, ierrStr)
    							end if
                            else
                                retVal = false
                                ierrStr = "주문구분 [ordDtlClsCd="&ordDtlClsCd&"] 정의되지 않음"
                            end if
                        else
                            rw ordNo&"|"&ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq&"|"&realslAmt&"|"&outwAmt&"|"&outwAmt/outwQty

                            if (orderCsGbn<>"") then
                                sqlStr = " Update T"
                                sqlStr = sqlStr & " SET realSellPrice=(CASE WHEN SellPrice<>realSellPrice THEN realSellPrice ELSE "&outwAmt/outwQty&" END )"
                                sqlStr = sqlStr & " ,PRE_USE_UNITCOST=0"
                                sqlStr = sqlStr & " ,tenCpnUint=0"
                                sqlStr = sqlStr & " ,mallCpnUnit="&realslAmt-outwAmt/outwQty&""
                                sqlStr = sqlStr & " From db_temp.dbo.tbl_xSite_tmporder T"
                				sqlStr = sqlStr & " where sellsite='cjmall'"
                                sqlStr = sqlStr & " and outmallorderserial='"&ordNo&"'"
                                sqlStr = sqlStr & " and OrgDetailKey='"&ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq&"'"
                                sqlStr = sqlStr & " and mallCpnUnit is NULL" ''2014/02/01
''rw sqlStr
                				dbget.Execute sqlStr
            				end if
                        end if

                        dtlCnt = dtlCnt+1

                        if (retVal) then
                            succCnt = succCnt+1
                        else
                            failCnt = failCnt+1
                            if (ierrCode=-1) then skipCnt = skipCnt+1

                            if (mode="ORDCANCELLIST") then
                                rw "<font color='red'>["&ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq&"-CA]</font> "&ierrStr &" / "&custNm
                            else
                                rw "["&ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq&"] "&ierrStr &" / "&custNm
                            end if
                        end if

                    Next
                end if

                Set detailSubNodes = Nothing
            Next
        end if
    end if

    Set Nodes = Nothing
    rw succCnt & "건 입력"
    rw failCnt & "건 실패" & "("&skipCnt&" 건 skip)"

End function

'주문내역 저장용
Function saveCSListResult(retDoc, mode, sday)

	'// TODO : !!!!
    Exit function

    Dim Nodes, masterSubNodes, detailSubNodes, detailSubNodeItem, ErrNode, errorMsg
    Dim isErrExists : isErrExists = false
    Dim ordNo

    Dim ordGSeq, ordDSeq, ordWSeq, wbClsCd, wbCls
    Dim wbCrtDt, outwConfDt, confirmChk, cnclInsDtm
    Dim oldordNo, toutYn, chnNm, receverNm, recvName, zipno, addr_1, addr_2, addr, telno, cellno
    Dim msgSpec, delvplnDt, packYn, itemNm, itemCd, unitCd, itemName, unitNm, contItemCd, wbIdNo
    Dim outwQty, realslAmt, outwAmt, delivInfo, promGiftSpec, cnclRsn, cnclRsnSpec, ordDtm, juminNum, dccouponCjhs, dccouponVendor
    Dim dtlCnt

    Dim requireDetail, orderDlvPay, orderCsGbn, ierrStr, ierrCode
    dim succCnt : succCnt=0
    dim failCnt : failCnt=0
    dim skipCnt : skipCnt=0
    dim retVal

    Set Nodes = retDoc.getElementsByTagName("ns1:errorMsg")
    If (Not (Nodes is Nothing)) Then
        For each ErrNode in Nodes
            errorMsg = Nodes.item(0).text
            isErrExists = true
            rw "["&sday&"]"&errorMsg
        next
    end if

    if (Not isErrExists) then
        Set Nodes = retDoc.getElementsByTagName("ns1:instruction")

        If (Not (Nodes is Nothing)) Then
            For each masterSubNodes in Nodes
                ordNo       = masterSubNodes.getElementsByTagName("ns1:ordNo")(0).Text	        '주문번호

                Set detailSubNodes = masterSubNodes.getElementsByTagName("ns1:instructionDetail")

                ''rw ordNo&"|"&custNm&"|"&custTelNo&"|"&custDeliveryCost

                dtlCnt = 1
                If (Not (detailSubNodes is Nothing)) Then
                    For each detailSubNodeItem in detailSubNodes
                        requireDetail = ""
                        ierrStr = ""

                        ordGSeq = detailSubNodeItem.getElementsByTagName("ns1:ordGSeq")(0).Text	    '[ID:주문상품순번], 001
                        ordDSeq = detailSubNodeItem.getElementsByTagName("ns1:ordDSeq")(0).Text	    '[ID:주문상세순번], 001
                        ordWSeq = detailSubNodeItem.getElementsByTagName("ns1:ordWSeq")(0).Text	    '[ID:주문처리순번], 001

                        wbClsCd = detailSubNodeItem.getElementsByTagName("ns1:wbClsCd")(0).Text	        ' 주문정보 - 진행구분코드
                        ''------------------------------------------------------------------------------------------------------------
                        ordDtlClsCd = detailSubNodeItem.getElementsByTagName("ns1:ordDtlClsCd")(0).Text	    ' 주문정보 - 주문구분코드, 10
                        wbCrtDt = detailSubNodeItem.getElementsByTagName("ns1:wbCrtDt")(0).Text	            ' 주문정보 - 지시일자, 2013-05-22+09:00
                        ''outwConfDt	'주문정보 - 출고확정일자
                        ''delivDtm	    '주문정보 - 배송완료일
                        ''cnclInsDtm	'주문정보 - 취소일자
                        ''oldordNo	    '주문정보 - 원주문번호
                        toutYn = detailSubNodeItem.getElementsByTagName("ns1:toutYn")(0).Text	            '주문정보 - 기출하구분(Y-기출하,N-정상출하), N
                        chnNm = detailSubNodeItem.getElementsByTagName("ns1:chnNm")(0).Text	                '주문정보 - 채널구분, INTERNET
                        receverNm = detailSubNodeItem.getElementsByTagName("ns1:receverNm")(0).Text	        '주문정보 - 인수자, 채현아
                        'recvName	    '주문정보 - 수취인 영문명
                        zipno = detailSubNodeItem.getElementsByTagName("ns1:zipno")(0).Text	                '주문정보 - 우편번호, 03082
                        addr_1 = detailSubNodeItem.getElementsByTagName("ns1:addr_1")(0).Text	            '주문정보 - 주소, 서울시 종로구 대학로 57
                        addr_2 = detailSubNodeItem.getElementsByTagName("ns1:addr_2")(0).Text	            '주문정보 - 상세주소, 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐
                        'addr	        '주문정보 - 주소
                        telno = detailSubNodeItem.getElementsByTagName("ns1:telno")(0).Text	                '주문정보 - 인수자tel, 02)973-8514
                        cellno = detailSubNodeItem.getElementsByTagName("ns1:cellno")(0).Text	            '주문정보 - 인수자HP, 010)2715-8514
                        'msgSpec	    '주문정보 - 배송참고
                        'delvplnDt	    '주문정보 - 배송예정일
                        packYn = detailSubNodeItem.getElementsByTagName("ns1:packYn")(0).Text	            '상품정보 - 세트여부, 일반
                        'itemNm	        '상품정보 - 세트상품명
                        itemCd = detailSubNodeItem.getElementsByTagName("ns1:itemCd")(0).Text	            '상품정보 - 판매코드, 21899852
                        unitCd = detailSubNodeItem.getElementsByTagName("ns1:unitCd")(0).Text	            '상품정보 - 단품코드, 10047125217
                        itemName = detailSubNodeItem.getElementsByTagName("ns1:itemName")(0).Text	        '상품정보 - 판매상품명, 24K Gold 전자파차단스티커
                        unitNm = detailSubNodeItem.getElementsByTagName("ns1:unitNm")(0).Text	            '상품정보 - 단품상세, ES-01 잘될꺼야
                        contItemCd = detailSubNodeItem.getElementsByTagName("ns1:contItemCd")(0).Text	    '상품정보 - 협력사상품코드, 279751_0011
                        wbIdNo = detailSubNodeItem.getElementsByTagName("ns1:wbIdNo")(0).Text	            '상품정보 - 운송장식별번호, 20000420537940
                        outwQty = detailSubNodeItem.getElementsByTagName("ns1:outwQty")(0).Text	            '상품정보 - 수량, 1.0
                        realslAmt = detailSubNodeItem.getElementsByTagName("ns1:realslAmt")(0).Text	        '상품정보 - 판매가, 1800.0
                        outwAmt = detailSubNodeItem.getElementsByTagName("ns1:outwAmt")(0).Text	            '상품정보 - 고객결제가, 1800.0  :: 수량*판매가 인지, 수량*실판매가인지 확인
                        'delivInfo	    '기타정보 - 비고
                        'promGiftSpec	'기타정보 - 사은품내역
                        'juminNum       '주문정보-주민번호(아님), 발송 금지!
                        'cnclRsn	    '기타정보 - 교환/취소사유
                        'cnclRsnSpec	'기타정보 - 교환/취소사유상세
                        ordDtm = detailSubNodeItem.getElementsByTagName("ns1:ordDtm")(0).Text	            '주문정보-주문일시, 2013-05-22 15:05:02


                        ''필수로 안넘어오는정보들.
                        wbCls =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:wbCls")(0) Is Nothing)) Then
                            wbCls = detailSubNodeItem.getElementsByTagName("ns1:wbCls")(0).Text       '주문정보 - 진행구분
                        end if

                        confirmChk =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:confirmChk")(0) Is Nothing)) Then
                            confirmChk = detailSubNodeItem.getElementsByTagName("ns1:confirmChk")(0).Text        '주문정보 - 협력사실제회수확인 0,1
                        end if
                        ''-------------------------------------------------------------------------------------------------------------------------
                        cnclInsDtm =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:cnclInsDtm")(0) Is Nothing)) Then
                            cnclInsDtm = detailSubNodeItem.getElementsByTagName("ns1:cnclInsDtm")(0).Text        '주문정보 - 취소일자
                        end if
                        oldordNo =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:oldordNo")(0) Is Nothing)) Then
                            oldordNo = detailSubNodeItem.getElementsByTagName("ns1:oldordNo")(0).Text        '주문정보 - 원주문번호
                        end if
                        recvName =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:recvName")(0) Is Nothing)) Then
                            recvName = detailSubNodeItem.getElementsByTagName("ns1:recvName")(0).Text        '주문정보 - 수취인 영문명
                        end if
                        addr =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:addr")(0) Is Nothing)) Then
                            addr = detailSubNodeItem.getElementsByTagName("ns1:addr")(0).Text        '주문정보 - 주소all?
                        end if
                        msgSpec =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:msgSpec")(0) Is Nothing)) Then
                            msgSpec = detailSubNodeItem.getElementsByTagName("ns1:msgSpec")(0).Text        '주문정보 -배송참고
                        end if
                        delvplnDt =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:delvplnDt")(0) Is Nothing)) Then
                            delvplnDt = detailSubNodeItem.getElementsByTagName("ns1:delvplnDt")(0).Text        '주문정보 -배송예정일
                        end if
                        itemNm =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:itemNm")(0) Is Nothing)) Then
                            itemNm = detailSubNodeItem.getElementsByTagName("ns1:itemNm")(0).Text        '상품정보 -세트상품명
                        end if
                        juminNum =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:juminNum")(0) Is Nothing)) Then
                            juminNum = detailSubNodeItem.getElementsByTagName("ns1:juminNum")(0).Text       '주문정보-주민번호(아님), 발송 금지!
                        end if
                        dccouponCjhs =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:dccouponCjhs")(0) Is Nothing)) Then
                            dccouponCjhs = detailSubNodeItem.getElementsByTagName("ns1:dccouponCjhs")(0).Text       '주문정보 - 할인(당사부담)금액
                        end if
                        dccouponVendor =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:dccouponVendor")(0) Is Nothing)) Then
                            dccouponVendor = detailSubNodeItem.getElementsByTagName("ns1:dccouponVendor")(0).Text      '주문정보 - 할인(협력사부담)금액
                        end if

                        orderDlvPay = custDeliveryCost
                        if (dtlCnt>1) then
                            orderDlvPay = 0 ''첫번째 값만 넣음.
                        end if

                        orderCsGbn = ""
                        if (ordDtlClsCd="10") then
                            orderCsGbn="0"
                        elseif (ordDtlClsCd="20") then
                            orderCsGbn="2"  ''취소
                        end if

                        requireDetail = juminNum '' 주문제작문구
                        ierrCode = 0
                        ierrStr  = ""

                        if (orderCsGbn<>"") then
                            retVal = saveCsOneTmp(ordNo,ordDtm,splitvalue(contItemCd,"_",0),splitvalue(contItemCd,"_",1),itemName, unitNm _
                                    , custNm , custTelNo, custTelNo _
                                    , receverNm, telno, cellno, LEFT(zipno,3)&"-"&Right(zipno,3), addr_1, addr_2 _
                                    , realslAmt, realslAmt, CLNG(outwQty), ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq _
                                    , msgSpec, requireDetail, orderDlvPay, orderCsGbn _
                                    , ierrCode, ierrStr)
                        else
                            retVal = false
                            ierrStr = "주문구분 [ordDtlClsCd="&ordDtlClsCd&"] 정의되지 않음"
                        end if

                        dtlCnt = dtlCnt+1

                        if (retVal) then
                            succCnt = succCnt+1
                        else
                            failCnt = failCnt+1
                            if (ierrCode=-1) then skipCnt = skipCnt+1

                            if (mode="ORDCANCELLIST") then
                                rw "<font color='red'>["&ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq&"]</font> "&ierrStr &" / "&custNm
                            else
                                rw "["&ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq&"] "&ierrStr &" / "&custNm
                            end if
                        end if

                    Next
                end if

                Set detailSubNodes = Nothing
            Next
        end if
    end if

    Set Nodes = Nothing
    rw succCnt & "건 입력"
    rw failCnt & "건 실패" & "("&skipCnt&" 건 skip)"

End function

function saveORDOneTmp(OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName _
        , OrderName, OrderTelNo, OrderHpNo _
        , ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
        , SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
        , deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
        , byref ierrCode, byref ierrStr )
    dim paramInfo, retParamInfo
    dim SellSite : SellSite = "cjmall"
    dim PayType  : PayType  = "50"
    dim sqlStr
	dim countryCode

	if countryCode="" then countryCode="KR"

    saveORDOneTmp =false

    OrderTelNo = replace(OrderTelNo,")","-")
    OrderHpNo = replace(OrderHpNo,")","-")
    ReceiveTelNo = replace(ReceiveTelNo,")","-")
    ReceiveHpNo = replace(ReceiveHpNo,")","-")

    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
        ,Array("@SellSite" , adVarchar	, adParamInput, 32, SellSite)	_
		,Array("@OutMallOrderSerial"	, adVarchar	, adParamInput,32, OutMallOrderSerial)	_
		,Array("@SellDate"	,adDate, adParamInput,, SellDate) _
		,Array("@PayType"	,adVarchar, adParamInput,32, PayType) _
		,Array("@Paydate"	,adDate, adParamInput,, SellDate) _
		,Array("@matchItemID"	,adInteger, adParamInput,, matchItemID) _
		,Array("@matchItemOption"	,adVarchar, adParamInput,4, matchItemOption) _
		,Array("@partnerItemID"	,adVarchar, adParamInput,32, matchItemID) _
		,Array("@partnerItemName"	,adVarchar, adParamInput,128, partnerItemName) _
		,Array("@partnerOption"	,adVarchar, adParamInput,128, matchItemOption) _
		,Array("@partnerOptionName"	,adVarchar, adParamInput,128, partnerOptionName) _
		,Array("@OrderUserID"	,adVarchar, adParamInput,32, "") _
		,Array("@OrderName"	,adVarchar, adParamInput,32, OrderName) _
		,Array("@OrderEmail"	,adVarchar, adParamInput,100, "") _
		,Array("@OrderTelNo"	,adVarchar, adParamInput,16, OrderTelNo) _
		,Array("@OrderHpNo"	,adVarchar, adParamInput,16, OrderHpNo) _
		,Array("@ReceiveName"	,adVarchar, adParamInput,32, ReceiveName) _
		,Array("@ReceiveTelNo"	,adVarchar, adParamInput,16, ReceiveTelNo) _
		,Array("@ReceiveHpNo"	,adVarchar, adParamInput,16, ReceiveHpNo) _
		,Array("@ReceiveZipCode"	,adVarchar, adParamInput,7, ReceiveZipCode) _
		,Array("@ReceiveAddr1"	,adVarchar, adParamInput,128, ReceiveAddr1) _
		,Array("@ReceiveAddr2"	,adVarchar, adParamInput,512, ReceiveAddr2) _
		,Array("@SellPrice"	,adCurrency, adParamInput,, SellPrice) _
		,Array("@RealSellPrice"	,adCurrency, adParamInput,, RealSellPrice) _
		,Array("@ItemOrderCount"	,adInteger, adParamInput,, ItemOrderCount) _
		,Array("@OrgDetailKey"	,adVarchar, adParamInput,32, OrgDetailKey) _
		,Array("@DeliveryType"	,adInteger, adParamInput,, 0) _
		,Array("@deliveryprice"	,adCurrency, adParamInput,, 0) _
		,Array("@deliverymemo"	,adVarchar, adParamInput,400, deliverymemo) _
		,Array("@requireDetail"	,adVarchar, adParamInput,400, requireDetail) _
		,Array("@orderDlvPay"	,adCurrency, adParamInput,, orderDlvPay) _
		,Array("@orderCsGbn"	,adInteger, adParamInput,, orderCsGbn) _
    	,Array("@countryCode"	,adVarchar, adParamInput,2, countryCode) _
    	,Array("@outMallGoodsNo"	,adVarchar, adParamInput,16, "") _
    	,Array("@shoplinkerMallName" ,adVarchar, adParamInput,64, "") _
    	,Array("@shoplinkerPrdCode"	,adVarchar, adParamInput,16, "") _
    	,Array("@shoplinkerOrderID"	,adVarchar, adParamInput,16, "") _
    	,Array("@shoplinkerMallID"	,adVarchar, adParamInput,32, "") _
		,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
	)

    if (matchItemOption<>"") and (matchItemID<>"-1") and (matchItemID<>"") then
        sqlStr = "db_temp.dbo.sp_TEN_xSite_TmpOrder_Insert"
        retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

        ierrCode = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
        ierrStr  = GetValue(retParamInfo, "@retErrStr")   ' 에러메세지
    else
        ierrCode = -999
        ierrStr = "상품코드 또는 옵션코드  매칭 실패" & OrgDetailKey & " 상품코드 =" & matchItemID&" 옵션명 = "&partnerOptionName
        rw "["&ierrCode&"]"&retErrStr
        dbget.close() : response.end
    end if

    saveORDOneTmp = (ierrCode=0)
end function

'리스트 호출용 // 딜레이 time 있음.. // 단품 가격 제대로 안넘어오나.. (22211226)
Function saveListResult(retDoc, mode, iitemid)
	Dim errorMsg, strSql
	Dim Nodes, SubNodes
	Dim XitemCd, Xstatus, XslCls, XHapvpn, Xvpn, XunitCd, Xitemcode
	Dim uprItemNm, itemNm, slprc,exLeadtm, zClassId, packInd, purchvat, taxYn
	Dim OverLapNo
	Dim SelOK, AssignedItemCnt, AssignedRow
	SelOK = 0
	AssignedItemCnt = 0

	Set Nodes = retDoc.getElementsByTagName("ns1:unit")

	If (Not (retDoc is Nothing)) Then
		On Error Resume Next
		For each SubNodes in Nodes
			XitemCd = SubNodes.getElementsByTagName("ns1:itemCd")(0).Text	'판매코드
			Xstatus = SubNodes.getElementsByTagName("ns1:status")(0).Text	'결재상태
			XslCls 	= SubNodes.getElementsByTagName("ns1:slCls")(0).Text	'판매구분(상태)
			XHapvpn	= SubNodes.getElementsByTagName("ns1:vpn")(0).Text		'업체상품코드
			XunitCd = SubNodes.getElementsByTagName("ns1:unitCd")(0).Text	'단품코드

			uprItemNm= SubNodes.getElementsByTagName("ns1:uprItemNm")(0).Text	'판매상품명
			itemNm  = SubNodes.getElementsByTagName("ns1:itemNm")(0).Text	'단품상세
			slprc   = SubNodes.getElementsByTagName("ns1:slprc")(0).Text	'판매가
			exLeadtm= SubNodes.getElementsByTagName("ns1:exLeadtm")(0).Text	'리드타임(L/T)
			packInd = SubNodes.getElementsByTagName("ns1:packInd")(0).Text
			purchvat = SubNodes.getElementsByTagName("ns1:purchvat")(0).Text ''매입가 vat포함?
			taxYn    = SubNodes.getElementsByTagName("ns1:taxYn")(0).Text
'			zClassId= SubNodes.getElementsByTagName("ns1:zClassId")(0).Text	'hsk

        'rw "마진:"&purchvat*1.1&"/"&slprc&":"&CHKIIF(slprc<>0,purchvat*1.1/slprc*100,"")&":"&taxYn
        ''if (taxYn="Y") then purchvat=purchvat*1.1
			Xvpn 	= Split(XHapvpn, "_")(0)
			Xitemcode = Split(XHapvpn, "_")(1)

			If (OverLapNo <> Xvpn) Then
				strSql = ""
				strSql = strSql & " UPDATE R " & VBCRLF
				strSql = strSql & " SET cjmallregdate = isNULL(cjmallregdate,getdate())" & VBCRLF
				strSql = strSql & " , cjmallPrdNo = '"&XitemCd&"'" & VBCRLF
				If  Xstatus = "A" Then	'승인완료 이고 판매중일 때 (Xstatus A:승인완료, XslCls A:진행, I:일시중단) ''AND XslCls = "A"
					strSql = strSql & " , cjmallStatCd = 7" & VBCRLF
				End If
				strSql = strSql & " , lastStatCheckDate = getdate()" & VBCRLF                               ''수정
				strSql = strSql & " FROM db_item.dbo.tbl_cjmall_regitem as R " & VBCRLF
				strSql = strSql & " INNER JOIN db_item.dbo.tbl_item as i on R.itemid=i.itemid " & VBCRLF
				strSql = strSql & " WHERE i.itemid = '"&Xvpn&"' "
				dbget.Execute strSql

				if (OverLapNo<>"") then
				    ''상품상태 update

				    strSql = ""
				    strSql = strSql & " update R"
                    strSql = strSql & " set cjmallsellyn=(CASE WHEN T.SellCNT>0 THEN 'Y' ELSE 'N' END)"
                    strSql = strSql & " ,cjMallPrice=(CASE WHEN T.mayItemPrice>0 then T.mayItemPrice ELSE R.cjMallPrice END)"
                    strSql = strSql & " from db_item.dbo.tbl_cjmall_regItem R"
                    strSql = strSql & " 	Join ("
                    strSql = strSql & " 		select itemid, count(*) as optCNT"
                    strSql = strSql & " 		, sum(CASE WHEN outmallsellyn='Y' THEN 1 ELSE 0 END) as SellCNT"
                    strSql = strSql & " 		, min(outmallAddPrice) as mayItemPrice"
                    strSql = strSql & " 		from db_item.dbo.tbl_OutMall_regedoption"
                    strSql = strSql & " 		where itemid="&Xvpn&""
                    strSql = strSql & " 		and mallid='cjmall'"
                    strSql = strSql & " 		group by itemid"
                    strSql = strSql & " 	) T on R.itemid=T.itemid"
                    strSql = strSql & " where R.itemid="&Xvpn&""

                    dbget.Execute strSql
				    AssignedItemCnt = AssignedItemCnt + 1
				end if

			End If

			'If Xitemcode <> "" AND Xitemcode <> "0000" Then
			If Xitemcode <> "" Then
				strSql = ""
				strSql = strSql & " UPDATE db_item.dbo.tbl_OutMall_regedoption  " & VBCRLF
				strSql = strSql & " SET outmallOptCode = '"&XunitCd&"' " & VBCRLF
				strSql = strSql & " , outmallsellyn='"&CHKIIF(XslCls="I","N","Y")&"'" & VBCRLF
				if (Xitemcode<>"0000") then
				    strSql = strSql & " , outmallOptName='"&html2DB(itemNm)&"'"& VBCRLF
			    end if
			    strSql = strSql & " , outmallAddPrice="&slprc& VBCRLF
			    strSql = strSql & " , outmallleadTime='"&exLeadtm&"'"& VBCRLF
				strSql = strSql & " , checkdate = getdate() " & VBCRLF
				strSql = strSql & " , outmallsuppPrc="&purchvat*1.1& VBCRLF
				strSql = strSql & " WHERE itemid = '"&Xvpn&"' and itemoption = '"&Xitemcode&"' " & VBCRLF
				strSql = strSql & " and mallid='"&CMALLNAME&"'" ''' 추가
				dbget.Execute strSql, AssignedRow

'				rw "진영*1.1:"&Round((slprc) /1.1 - (12/100) * ((slprc)/1.1))*1.1
'				rw "purchvat*1.1:"&purchvat*1.1


				''기존 내역에 없을경우
				if (AssignedRow<1) then

				    sqlStr = ""
					sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
					sqlStr = sqlStr & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice,outmallleadTime, outmallsuppPrc,lastUpdate) " & VBCRLF
					sqlStr = sqlStr & " VALUES " & VBCRLF
					sqlStr = sqlStr & " ('"&Xvpn&"'"& VBCRLF
					sqlStr = sqlStr & ",  '"&Xitemcode&"'"& VBCRLF
					sqlStr = sqlStr & ", '"&CMALLNAME&"'"& VBCRLF
					sqlStr = sqlStr & ", '"&XunitCd&"'"& VBCRLF
					sqlStr = sqlStr & ", '"&html2db(CHKIIF(Xitemcode<>"0000", itemNm, "단일상품"))&"'"& VBCRLF
					sqlStr = sqlStr & ", '"&CHKIIF(XslCls="I", "N", "Y")&"'"& VBCRLF
					sqlStr = sqlStr & ", '"&"N"&"'"& VBCRLF
					sqlStr = sqlStr & ", '0'"& VBCRLF
					sqlStr = sqlStr & ", '"&slprc&"'"& VBCRLF
					sqlStr = sqlStr & ", '"&exLeadtm&"'"& VBCRLF
					sqlStr = sqlStr & ", "&purchvat*1.1&""& VBCRLF
					sqlStr = sqlStr & ", getdate()) "

					dbget.Execute sqlStr
				end if
			End If
			OverLapNo = Xvpn

			SelOK = SelOK + 1
			rw XHapvpn&"|"&XunitCd&"|"&Xstatus&"|"&XslCls&"|"&uprItemNm&"|"&itemNm&"|"&slprc&"|"&purchvat*1.1&"|"&exLeadtm&"|"&packInd
		Next

		if (OverLapNo<>"") then
		    ''상품상태 update
		    strSql = ""
		    strSql = strSql & " update R"
            strSql = strSql & " set cjmallsellyn=(CASE WHEN T.SellCNT>0 THEN 'Y' ELSE 'N' END)"
            strSql = strSql & " ,cjMallPrice=(CASE WHEN T.mayItemPrice>0 then T.mayItemPrice ELSE R.cjMallPrice END)"
            strSql = strSql & " ,regedOptCnt=isNULL(T.regedOptCnt,0)"
            strSql = strSql & " from db_item.dbo.tbl_cjmall_regItem R"
            strSql = strSql & " 	Join ("
            strSql = strSql & " 		select itemid, count(*) as optCNT"
            strSql = strSql & " 		, sum(CASE WHEN outmallsellyn='Y' THEN 1 ELSE 0 END) as SellCNT"
            ''strSql = strSql & " 		, sum(CASE WHEN outmallsellyn='Y' and itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
            strSql = strSql & " 		, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
            strSql = strSql & " 		, min(outmallAddPrice) as mayItemPrice"
            strSql = strSql & " 		from db_item.dbo.tbl_OutMall_regedoption"
            strSql = strSql & " 		where itemid="&Xvpn&""
            strSql = strSql & " 		and mallid='cjmall'"
            strSql = strSql & " 		group by itemid"
            strSql = strSql & " 	) T on R.itemid=T.itemid"
            strSql = strSql & " where R.itemid="&Xvpn&""

            dbget.Execute strSql
		    AssignedItemCnt = AssignedItemCnt + 1
		end if

        ''업데이트 된 tbl_OutMall_regedoption 기준으로 tbl_cjmall_regitem 의 판매상태 update 필요
        '' 849637 CASE 확인 상품명, 승인대기

    	If SelOK = 0 Then
    	    if (iitemid<>"") then
    	        ''체크실패시 반복되지 않도록
    	        strSql = strSql & " update R"
                strSql = strSql & " set lastStatCheckDate = getdate()" & VBCRLF
                strSql = strSql & " from db_item.dbo.tbl_cjmall_regItem R" & VBCRLF
                strSql = strSql & " where itemid="&iitemid
                dbget.Execute strSql
    	    end if
    		rw iitemid & "::"&"검색 결과 없음"
    		saveListResult = false
    	End If

	End If
	on Error Goto 0

	Set Nodes = Nothing

    if (AssignedItemCnt>0) then
	    rw "상품 "&AssignedItemCnt&" 건 sync"
	    saveListResult = true
    end if
End Function

'그외 용
Function saveCommonItemResult(retDoc, mode, prdno)
	Dim errorMsg
	Dim sqlStr
	Dim AssignedRow, successYn
	Dim Titemid, Tlimitno, Tlimitsold, Titemoption, Toptionname, Toptlimitno, Toptlimitsold, Toptsellyn, Toptlimityn, Toptaddprice, Tlimityn, Tsellyn, Titemsu, Tsellcash
	Dim Nodes, OneNode, SubNodes
	Dim typeCD, itemCD_ZIP, newUnitRetail, newUnitCost, packInd
	Dim unitCd, strDt, endDt, availSupQty, notitemId, notmakerid
    Dim ierrStr

	successYn = false
	errorMsg = ""

'2013-07-08 김진영 하단 If문 주석
'	If (Not (retDoc is Nothing)) Then
'	    if (retDoc.getElementsByTagName("ns1:successYn").item(0).text="false") then  ''2013/06/20 15시경부터 이상.
'	        successYn= false
'	        errorMsg = mode & "ERR_" & retDoc.getElementsByTagName("ns1:errorMsg").item(0).text
'	    else
'   		errorMsg = retDoc.getElementsByTagName("ns1:errorMsg").item(0).text
' 	 		If Left(errorMsg,4) = "[성공]" Then
' 				successYn= true
'			End If
'		end if
'	End If

'2013-07-08 김진영 하단 If문 추가수정
'2014-09-15 11:12 김진영 on error 문 추가 / cj API 변경으로 Left => instr로 변경
	If (Not (retDoc is Nothing)) Then
		on Error resume next
			errorMsg = retDoc.getElementsByTagName("ns1:errorMsg").item(0).text
		on Error Goto 0

	    If Instr(errorMsg, "I:[성공]") > 0 Then
			successYn= true
	    else
			successYn= false
    	end if
	End If

	If (successYn = true) Then
	'성공이고 mode=REG면
		If mode = "REG" Then
	'reged옵션 테이블에 데이터 꼽기
			sqlStr = ""
			sqlStr = sqlStr & " SELECT i.itemid, i.limitno ,i.limitsold, i.sellcash, o.itemoption, o.optionname, o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, o.optaddprice, i.sellyn, i.limityn " & VBCRLF
			sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i " & VBCRLF
			sqlStr = sqlStr & " LEFT JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' " & VBCRLF
			sqlStr = sqlStr & " WHERE i.itemid = "&prdno&" " & VBCRLF
			sqlStr = sqlStr & " ORDER BY o.itemoption ASC "
			rsget.Open sqlStr,dbget,1
			If Not(rsget.EOF or rsget.BOF) Then
				If rsget.RecordCount = 1 AND IsNull(rsget("itemoption")) Then
					Titemid			= rsget("itemid")
					Titemoption 	= "0000"
					Toptionname		= "단일상품"
					Tlimitno		= rsget("limitno")
					Tlimitsold		= rsget("limitsold")
					Tlimityn		= rsget("limityn")
					Tsellyn			= rsget("sellyn")

					If (Tlimityn="Y") then
						If (Tlimitno - Tlimitsold) < 5 Then
							Titemsu = 0
						Else
							Titemsu = Tlimitno - Tlimitsold - 5
						End If
					Else
						Titemsu = 999
					End If

					sqlStr = ""
					sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
					sqlStr = sqlStr & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
					sqlStr = sqlStr & " VALUES " & VBCRLF
					sqlStr = sqlStr & " ('"&Titemid&"',  '"&Titemoption&"', 'cjmall', '', '"&html2db(Toptionname)&"', '"&rsget("sellyn")&"', '"&rsget("limityn")&"', '"&Titemsu&"', '0', getdate()) "
					dbget.Execute sqlStr
				Else
					For i = 1 to rsget.RecordCount
						Titemid			= rsget("itemid")
						Tlimitno		= rsget("limitno")
						Tlimitsold		= rsget("limitsold")
						Titemoption		= rsget("itemoption")
						Toptionname		= rsget("optionname")
						Toptlimitno		= rsget("optlimitno")
						Toptlimitsold	= rsget("optlimitsold")
						Toptsellyn		= rsget("optsellyn")
						Toptlimityn		= rsget("optlimityn")
						Toptaddprice	= rsget("optaddprice")
						Tsellcash		= rsget("sellcash")

						If (Toptlimityn="Y") then
							If (Toptlimitno - Toptlimitsold) < 5 Then
								Titemsu = 0
							Else
								Titemsu = Toptlimitno - Toptlimitsold - 5
							End If
						Else
							Titemsu = 999
						End If

						If Left(Titemoption, 1) = "Z" Then
							Toptionname = Replace(Toptionname, ",", "/")
						End If

						sqlStr = ""
						sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
						sqlStr = sqlStr & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
						sqlStr = sqlStr & " VALUES " & VBCRLF
						sqlStr = sqlStr & " ('"&Titemid&"',  '"&Titemoption&"', 'cjmall', '', '"&html2db(Toptionname)&"', '"&Toptsellyn&"', '"&Toptlimityn&"', '"&Titemsu&"', '"&Toptaddprice+Tsellcash&"', getdate()) "
						dbget.Execute sqlStr
						rsget.MoveNext
					Next
				End If
			End If
			rsget.Close
		ElseIf mode = "EDT" Then  '' 결과값으로 저장요망
			sqlStr = ""
			sqlStr = sqlStr & " SELECT i.itemid, i.limitno ,i.limitsold, i.sellcash, o.itemoption, o.optionname, o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, o.optaddprice, R.outmallOptCode " & VBCRLF
			sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i " & VBCRLF
			sqlStr = sqlStr & " LEFT JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' " & VBCRLF
			sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_OutMall_regedoption as R on i.itemid = R.itemid and R.itemoption = o.itemoption and R.mallid='"&CMALLNAME&"'" & VBCRLF   ''R.mallid='"&CMALLNAME&"' 추가
			sqlStr = sqlStr & " WHERE i.itemid = "&prdno&" " & VBCRLF
			sqlStr = sqlStr & " ORDER BY o.itemoption ASC "
			rsget.Open sqlStr,dbget,1
			If Not(rsget.EOF or rsget.BOF) Then
				For i = 1 to rsget.RecordCount
					If IsNull(rsget("outmallOptCode")) Then
						Titemid			= rsget("itemid")
						Tlimitno		= rsget("limitno")
						Tlimitsold		= rsget("limitsold")
						Titemoption		= rsget("itemoption")
						Toptionname		= rsget("optionname")
						Toptlimitno		= rsget("optlimitno")
						Toptlimitsold	= rsget("optlimitsold")
						Toptsellyn		= rsget("optsellyn")
						Toptlimityn		= rsget("optlimityn")
						Toptaddprice	= rsget("optaddprice")
						Tsellcash		= rsget("sellcash")
						If Left(Titemoption, 1) = "Z" Then
							Toptionname = Replace(Toptionname, ",", "/")
						End If
						sqlStr = ""
						sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
						sqlStr = sqlStr & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
						sqlStr = sqlStr & " VALUES " & VBCRLF
						sqlStr = sqlStr & " ('"&Titemid&"',  '"&Titemoption&"', 'cjmall', '', '"&html2db(Toptionname)&"', '"&Toptsellyn&"', '"&Toptlimityn&"', '"&Toptlimitno&"', '"&Toptaddprice+Tsellcash&"', getdate()) "
						'dbget.Execute sqlStr
						''중복오류
					End If
					rsget.MoveNext
				Next
			End If
			rsget.Close
	    ElseIf mode = "MDT" THEN
	        rw mode
            Set Nodes = retDoc.getElementsByTagName("ns1:itemStates")
	        If (Not (Nodes is Nothing)) Then
                For each OneNode in Nodes
                    If (Not (OneNode is Nothing)) Then
                        successYn=OneNode.getElementsByTagName("ns1:successYn").item(0).text
                        if (successYn="true") then
                            errorMsg        = OneNode.getElementsByTagName("ns1:errorMsg").item(0).text
                            typeCd          = OneNode.getElementsByTagName("ns1:typeCd").item(0).text
                            itemCd_zip      = OneNode.getElementsByTagName("ns1:itemCd_zip").item(0).text
                            packInd         = OneNode.getElementsByTagName("ns1:packInd").item(0).text

                            rw typeCd&","&itemCd_zip&","&packInd&","&CHKIIF(successYn<>"true",errorMsg,"")

                            if (typeCd="01") then

                            elseif (typeCd="02") then
                                sqlStr = "UpDate R"&VbCRLF
								sqlStr = sqlStr & " SET outmallsellyn='"&CHKIIF(packInd="A","Y","N")&"'"&VbCRLF
            				    sqlStr = sqlStr & " , lastupdate = getdate() " & VBCRLF
            				    sqlStr = sqlStr & " from db_item.dbo.tbl_OutMall_regedoption R"&VbCRLF
            				    sqlStr = sqlStr & "  where mallid='"&CMALLNAME&"'"&VbCRLF
            				    sqlStr = sqlStr & "  and itemid="&prdno&VbCRLF
            				    sqlStr = sqlStr & "  and outmallOptCode='"&itemCd_zip&"'"&VbCRLF
            				    'dbget.Execute sqlStr
            				end if

                        end if
        			end if
			    Next
		    END IF
	    ElseIf mode = "QTY" THEN
	        Set Nodes = retDoc.getElementsByTagName("ns1:ltSupplyPlans")
	        If (Not (Nodes is Nothing)) Then
                For each OneNode in Nodes
                    If (Not (OneNode is Nothing)) Then
                        successYn=OneNode.getElementsByTagName("ns1:successYn").item(0).text
                        if (successYn="true") then
                            errorMsg        = OneNode.getElementsByTagName("ns1:errorMsg").item(0).text
                            unitCd          = OneNode.getElementsByTagName("ns1:unitCd").item(0).text
                            strDt           = OneNode.getElementsByTagName("ns1:strDt").item(0).text
                            endDt           = OneNode.getElementsByTagName("ns1:endDt").item(0).text
                            availSupQty     = OneNode.getElementsByTagName("ns1:availSupQty").item(0).text

                            if (strDt=endDt) then
                                availSupQty=0
                            end if

                            rw unitCd&","&strDt&","&endDt&","&availSupQty&","&CHKIIF(successYn<>"true",errorMsg,"")

                            sqlStr = "UpDate R"&VbCRLF
        				    sqlStr = sqlStr & " SET outmalllimitno="&availSupQty&VbCRLF
        				    sqlStr = sqlStr & " ,outmalllimityn='Y'" ''강제로 지정
        				    sqlStr = sqlStr & " from db_item.dbo.tbl_OutMall_regedoption R"&VbCRLF
        				    sqlStr = sqlStr & "  where mallid='"&CMALLNAME&"'"&VbCRLF
        				    sqlStr = sqlStr & "  and itemid="&prdno&VbCRLF
        				    sqlStr = sqlStr & "  and outmallOptCode='"&unitCd&"'"&VbCRLF
        				    dbget.Execute sqlStr

                        end if
        			end if
			    Next
		    END IF
        ElseIf mode = "PRI" or mode = "PRI2" or mode = "PRI_RE" THEN
            'rw retDoc.saveHTML
            'itemPrices-itemPrice-typeCD,itemCD_ZIP,chnCls,effectiveDate,newUnitRetail, newUnitCost
            '          -ifResult-successYn,errorMsg
            Set Nodes = retDoc.getElementsByTagName("ns1:itemPrices")
            If (Not (Nodes is Nothing)) Then
                For each OneNode in Nodes
                    If (Not (OneNode is Nothing)) Then
                        successYn=OneNode.getElementsByTagName("ns1:successYn").item(0).text
                        if (successYn="true" or successYn="false") then
                            errorMsg        = OneNode.getElementsByTagName("ns1:errorMsg").item(0).text
            				typeCD 	        = OneNode.getElementsByTagName("ns1:typeCD").item(0).text
            				itemCD_ZIP		= OneNode.getElementsByTagName("ns1:itemCD_ZIP").item(0).text
            				newUnitRetail	= OneNode.getElementsByTagName("ns1:newUnitRetail").item(0).text
            				newUnitCost	    = OneNode.getElementsByTagName("ns1:newUnitCost").item(0).text

            				rw typeCD&","&itemCD_ZIP&","&newUnitRetail&","&newUnitRetail&","&CHKIIF(successYn<>"true",errorMsg,"")
            				if ((successYn="false") and (errorMsg="[이미등록된 자료가 이거나 중복된 단품이 존재합니다.]")) then
            				    rw "<font color=red>"&errorMsg&"</font>"
            				    IF (mode="PRI") then  ''mode<>PRI_RE 무한루프 방지
            				        rw "RE_TRY"
            				        ierrStr =""
            				        CALL editPriceCjmallOneItemByOption(prdno,ierrStr,TRUE)
            				        rw ierrStr
            				        Exit function
            				    END IF
            			    end if

            				if (successYn="true") then
                				if (typeCD="01") then

                				elseif (typeCD="02") then                   ''단품
                					If mode = "PRI" Then
	                				    sqlStr = "UpDate R"&VbCRLF
	                				    sqlStr = sqlStr & " SET outmallAddPrice="&newUnitRetail&VbCRLF
	                				    sqlStr = sqlStr & " ,lastupdate=getdate()"&VbCRLF
	                				    sqlStr = sqlStr & " ,checkdate=getdate()"&VbCRLF
	                				    sqlStr = sqlStr & " from db_item.dbo.tbl_OutMall_regedoption R"&VbCRLF
	                				    sqlStr = sqlStr & "  where mallid='"&CMALLNAME&"'"&VbCRLF
	                				    sqlStr = sqlStr & "  and itemid="&prdno&VbCRLF
	                				    sqlStr = sqlStr & "  and outmallOptCode='"&itemCD_ZIP&"'"&VbCRLF
	                				    dbget.Execute sqlStr
	                				End If
                				end if
                		    end if
        			    end if
        			end if
			    Next
		    END IF

		End If

		If (prdno <> "") Then
			Dim MustPrice
			sqlStr = ""
			sqlStr = sqlStr & " SELECT sellcash, buycash, orgprice FROM db_item.dbo.tbl_item where itemid = "&prdno&"  " & VBCRLF
			rsget.Open sqlStr,dbget,1
			If Not(rsget.EOF or rsget.BOF) Then
				If CLng(10000 - rsget("buycash") / rsget("sellcash") * 100 * 100) / 100 < 15 Then
					MustPrice = rsget("orgprice")
				Else
					MustPrice = rsget("sellcash")
				End If
			End If
			rsget.Close

			sqlStr = ""
			sqlStr = sqlStr & " UPDATE R" & VBCRLF
			sqlStr = sqlStr & " SET accFailCNT=0" & VBCRLF                 ''실패회수 초기화

			''가격 수정, 상품 품절은 cjmallLastUpdate 하지 않음. ''2013-10-31 김진영//가격수정(PRI)도 판매가격이 변경전으로 돌아가서 불편(채현아)..cjmallLastUpdate 수정되게 변경
			If (mode = "REG") or (mode = "MDT") or (mode = "EDT") or (mode = "PRI") Then
			    sqlStr = sqlStr & " , cjmallLastUpdate = getdate()" & VBCRLF
            end if

			If (mode = "REG") Then
				sqlStr = sqlStr & " ,cjmallStatCd=(CASE WHEN isNULL(cjmallStatCd, -1) < 3 then 3 ELSE cjmallStatCd END)"        ''임시등록완료(등록 후 승인대기)
				sqlStr = sqlStr & " ,cjmallRegdate=isNULL(cjmallRegdate,getdate())" & VbCrlf
			End If

			If (mode = "PRI") or (mode = "REG") Then
				sqlStr = sqlStr & " ,cjmallPrice = '"&MustPrice&"'" & VBCRLF
				If mode = "PRI" Then
					sqlStr = sqlStr & " ,lastPriceCheckDate = getdate()" & VBCRLF
				End If
			End If

			If (mode = "SLD") Then
				sqlStr = sqlStr & " ,cjmallSellYn = 'N'" & VBCRLF
			Else
				If (mode = "MDT") or (mode = "REG") Then
					sqlStr = sqlStr & " ,cjmallSellYn = i.sellyn" & VBCRLF              ''MDT 일경우 결과로 저장할지 확인.
				End If
			End If
			sqlStr = sqlStr & " FROM db_item.dbo.tbl_cjmall_regItem as R" & VBCRLF
			sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid" & VBCRLF
			sqlStr = sqlStr & " WHERE R.itemid = "&prdno&""   & VBCRLF
			dbget.Execute sqlStr,AssignedRow
			saveCommonItemResult=true
		End If
	Else
		Call Fn_AcctFailTouch(CMALLNAME, prdno, errorMsg)
		'' Err Log Insert
		sqlStr = ""
		sqlStr = sqlStr & " INSERT into db_log.dbo.tbl_interparkEdit_log" & VBCRLF
		sqlStr = sqlStr & " (itemid, interParkPrdNo, sellcash, buycash, sellyn, ErrMsg, logdate, mallid)" & VBCRLF
		sqlStr = sqlStr & " SELECT "&prdno & VBCRLF
		sqlStr = sqlStr & " ,'' ,i.sellcash, i.buycash, i.sellyn" & VBCRLF
		sqlStr = sqlStr & " ,convert(varchar(100), '"&html2db(errorMsg)&"')" & VBCRLF
		sqlStr = sqlStr & " ,getdate(), '"&CMALLNAME&"'" & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i WHERE i.itemid="&prdno&"" & VBCRLF
		dbget.Execute sqlStr

		If (mode = "REG") Then
			sqlStr = ""
			sqlStr = sqlStr & " UPDATE R" & VBCRLF
			sqlStr = sqlStr & " SET cjmallStatCd = -1"                   '''등록실패
			sqlStr = sqlStr & " FROM db_item.dbo.tbl_cjmall_regitem as R" & VBCRLF
			sqlStr = sqlStr & " WHERE R.itemid = "&prdno&"" & VBCRLF
			sqlStr = sqlStr & " and cjmallStatCd = 1"                    ''전송
			dbget.Execute sqlStr
		End If

		If (mode = "EDT") Then
		    rw "["&errorMsg&"]"
    		if (Trim(errorMsg)="1번째 단품:유효하지 않은 값입니다.[단품정보-협력사상품코드(Vpn)]가 이미 존재합니다. 새로운 값을 사용하세요.") then
    		    sqlStr = ""
    			sqlStr = sqlStr & " UPDATE R" & VBCRLF
    			sqlStr = sqlStr & " SET lastStatCheckDate=NULL"                   '''등록실패
    			sqlStr = sqlStr & " FROM db_item.dbo.tbl_cjmall_regitem as R" & VBCRLF
    			sqlStr = sqlStr & " WHERE R.itemid = "&prdno&"" & VBCRLF
    			dbget.Execute sqlStr
    		end if
        end if
	End If

	If (isCJ_DebugMode) Then
		rw prdno &"_"&mode&"_"&errorMsg
	End If
End Function

Function Fn_AcctFailTouch(iMallID, iitemid, iLastErrStr)
	Dim strSql
	iLastErrStr = html2db(iLastErrStr)

	If (iMallID = "lotteCom") Then
		strSql = "Update R" & VBCRLF
		strSql = strSql &" SET accFailCnt=accFailCnt+1" & VBCRLF
		strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" From db_item.dbo.tbl_lotte_regItem R" & VBCRLF
		strSql = strSql &" where itemid="&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "lotteimall") Then
		strSql = "Update R"&VbCRLF
		strSql = strSql &" SET accFailCnt=accFailCnt+1" & VBCRLF
		strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" From db_item.dbo.tbl_LTiMall_regItem R" & VBCRLF
		strSql = strSql &" where itemid="&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "interpark") Then
		strSql = "Update R"&VbCRLF
		strSql = strSql &" SET accFailCnt=accFailCnt+1" & VBCRLF
		strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" From db_item.dbo.tbl_interpark_reg_item R" & VBCRLF
		strSql = strSql &" where itemid="&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "cjmall") Then
		strSql = ""
		strSql = strSql & "UPDATE R"&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_item.dbo.tbl_cjmall_regitem as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	End If
End Function

Function XMLSend(url, xmlStr)
	Dim poster, SendDoc, retDoc
	Set SendDoc = server.createobject("MSXML2.DomDocument.3.0")
		SendDoc.async = False
		SendDoc.LoadXML(xmlStr)

	Set poster = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		poster.open "POST", url, false
		poster.setRequestHeader "CONTENT_TYPE", "text/xml"
		poster.setTimeouts 5000,90000,90000,90000  ''2013/07/22 추가
		poster.send SendDoc

	Set retDoc = server.createobject("MSXML2.DomDocument.3.0")
		retDoc.async = False
		retDoc.LoadXML(poster.responseTEXT)

	''response.write poster.responseTEXT

	Set XMLSend = retDoc
	Set SendDoc = Nothing
	Set poster = Nothing
End Function

Function getCurrDateTimeFormat()
    dim nowtimer : nowtimer= timer()
    getCurrDateTimeFormat = left(now(),10)&"_"&nowtimer
End Function

Sub CheckFolderCreate(sFolderPath)
	Dim objfile
	Set objfile=Server.CreateObject("Scripting.FileSystemObject")
		IF NOT  objfile.FolderExists(sFolderPath) THEN
			objfile.CreateFolder sFolderPath
		END IF
	Set objfile=Nothing
End Sub

Function XMLFileSave(xmlStr, mode, iitemid)
   Exit function  ''로그 안남김

	Dim fso,tFile
	Dim opath
	Select Case mode
		Case "REG", "RET_REG"
			opath = "/admin/etc/cjmall/xmlFiles/INSERT/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
		Case "LIST", "DayLIST", "RET_LIST", "RET_DayLIST", "commonCD", "RET_commonCD", "RET_SONGJANG", "cjItemCheck", "RET_cjItemCheck"
			opath = "/admin/etc/cjmall/xmlFiles/SELECT/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
		Case "EDT", "RET_EDT", "MDT", "RET_MDT", "PRI", "RET_PRI", "PRI2", "RET_PRI2", "QTY", "RET_QTY", "DateRes", "RET_DateRes"
			opath = "/admin/etc/cjmall/xmlFiles/UPDATE/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
		Case "SLD", "RET_SLD"
			opath = "/admin/etc/cjmall/xmlFiles/UPDATE_SellStatus/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
	    Case "RET_ORDLIST", "RET_ORDCANCELLIST", "RET_CSLIST"
	        opath = "/admin/etc/cjmall/xmlFiles/ORDER/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
	End Select

	Dim defaultPath : defaultPath = server.mappath(opath) + "\"
	Dim fileName
	If mode = "LIST" or mode = "DayLIST" Then
		fileName = mode &"_"& getCurrDateTimeFormat& ".xml"
	Else
		fileName = mode &"_"& getCurrDateTimeFormat&"_"&iitemid&".xml"
	End If

	CALL CheckFolderCreate(defaultPath)
	''debug
	Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(defaultPath & FileName )
			tFile.Write(xmlStr)
			tFile.Close
		Set tFile = nothing
	Set fso = nothing
End Function

function getLastOrderInputDT()
    dim sqlStr
    sqlStr = "select top 1 convert(varchar(10),selldate,21) as lastOrdInputDt"
    sqlStr = sqlStr&" from db_temp.dbo.tbl_XSite_TMpOrder"
    sqlStr = sqlStr&" where sellsite='cjmall'"
    sqlStr = sqlStr&" order by selldate desc"

    rsget.Open sqlStr,dbget,1
	if (Not rsget.Eof) then
		getLastOrderInputDT = rsget("lastOrdInputDt")
	end if
	rsget.Close

end function

function getLastOrderInputDTUp()
    dim sqlStr
    sqlStr = " select top 1 convert(varchar(10),selldate,21) as lastOrdInputDt"
    sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_tmporder"
    sqlStr = sqlStr & " where sellsite='cjmall'"
    sqlStr = sqlStr & " and  convert(varchar(10),selldate,21)>("
    sqlStr = sqlStr & " 	select top 1 convert(varchar(10),selldate,21) as slDt from db_temp.dbo.tbl_xSite_tmporder"
    sqlStr = sqlStr & " 	where sellsite='cjmall'"
    sqlStr = sqlStr & " 	and tenCpnUint is Not NULL"
    sqlStr = sqlStr & " 	order by selldate desc"
    sqlStr = sqlStr & " ) order by selldate"

    rsget.Open sqlStr,dbget,1
	if (Not rsget.Eof) then
		getLastOrderInputDTUp = rsget("lastOrdInputDt")
	end if
	rsget.Close

    'getLastOrderInputDTUp="2013-06-14"
end function

function IsOptionAddPriceExistItem(iitemid)
    dim sqlStr
    IsOptionAddPriceExistItem=false
    sqlStr = "select count(*) as CNT from db_item.dbo.tbl_item_option"&VbCRLF
    sqlStr = sqlStr&" where itemid="&iitemid&VbCRLF
    sqlStr = sqlStr&" and optAddprice<>0"&VbCRLF

    rsget.Open sqlStr,dbget,1
	if (Not rsget.Eof) then
		IsOptionAddPriceExistItem = rsget("CNT")>0
	end if
	rsget.Close

end function
%>
