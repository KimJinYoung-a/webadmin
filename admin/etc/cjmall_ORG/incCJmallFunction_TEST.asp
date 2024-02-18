<%
Dim isCJ_DebugMode : isCJ_DebugMode = True
Dim cjMallAPIURL

IF application("Svr_Info")="Dev" THEN
	cjMallAPIURL = "http://210.122.101.154:8110/IFPAServerAction.action"	'' �׽�Ʈ����
Else
	cjMallAPIURL = "http://api.cjmall.com/IFPAServerAction.action"			'' �Ǽ���
End if

Public Pmode
Public OPTYN

Function getMode()
	getMode = Pmode
End Function

Function getOPTYN()
	getOPTYN = OPTYN
End Function

Function getXMLString(byval iitemid, mode)
	Dim oCJMallItem
	Dim strRst
	SET oCJMallItem = new CCjmall
		oCJMallItem.FRectMode = mode
		oCJMallItem.FRectItemID = iitemid
	If (mode = "REG") Then
		oCJMallItem.getCJMallNotRegItemList
		If (oCJMallItem.FREsultCount > 0) Then
			getXMLString = oCJMallItem.FItemList(0).getCjmallItemRegXML
		End If
	ElseIf (mode = "LIST") Then
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_07"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_07.xsd"">"
		strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
		strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
		strRst = strRst &"<tns:contents>"
		strRst = strRst &"	<tns:sinstDtFrom>2013-04-05</tns:sinstDtFrom>"
		strRst = strRst &"	<tns:sinstDtTo>"&Left(now, 10)&"</tns:sinstDtTo>"
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
	ElseIf (mode = "MDT") OR (mode = "EDT") OR (mode = "PRI") OR (mode = "QTY") Then
		oCJMallItem.getCjmallEditedItemList
		If (oCJMallItem.FREsultCount > 0) Then
			If mode = "MDT" Then
				Dim sqlStr, arrRows, isOptionExists, i
				Dim itemoption, optiontypename, optionname, optLimit, optlimityn, isUsing, optsellyn, preged, optNameDiff, forceExpired, oopt, ooptCd, YtoN, NtoY, DelOpt
				Dim validSellno

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

						If optlimityn = "Y" AND optLimit < 5 Then			'�����̰� ������ 5�� �̸��� ���
							strRst = ""
							strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
							strRst = strRst &"<tns:ifRequest xmlns:tns='http://www.example.org/ifpa' tns:ifId='IF_03_03' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation='http://www.example.org/ifpa ../IF_03_03.xsd'>"
							strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
							strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
							strRst = strRst &"<tns:itemStates>"
							strRst = strRst &"	<tns:typeCd>02</tns:typeCd>"							'!!!01=�Ǹ��ڵ�,02=��ǰ�ڵ�
							strRst = strRst &"	<tns:itemCd_zip>"&ooptCd&"</tns:itemCd_zip>"
							strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
							strRst = strRst &"	<tns:packInd>I</tns:packInd>"						'!!!A-����, I-�Ͻ��ߴ�
							strRst = strRst &"</tns:itemStates>"
							strRst = strRst &"</tns:ifRequest>"
							OPTYN = "N"
							Pmode = "OptMod"
							getXMLString = strRst
						ElseIf preged = False AND ooptCd = "" Then			'1.�ɼ��� �߰��Ǵ°��
							getXMLString = oCJMallItem.FItemList(0).getcjmallItemModXML("unitY")
						ElseIf YtoN = True OR NtoY = True OR DelOpt = True Then			'2.�ɼ� sellyn�� N �Ǵ� Y��, itemoption���̺��� �ɼ��� ���� ���� �Ǵ°�� (CJ��ǰ�ڵ带 ���� ��������)
							strRst = ""
							strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
							strRst = strRst &"<tns:ifRequest xmlns:tns='http://www.example.org/ifpa' tns:ifId='IF_03_03' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation='http://www.example.org/ifpa ../IF_03_03.xsd'>"
							strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
							strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
							strRst = strRst &"<tns:itemStates>"
							strRst = strRst &"	<tns:typeCd>02</tns:typeCd>"							'!!!01=�Ǹ��ڵ�,02=��ǰ�ڵ�
							strRst = strRst &"	<tns:itemCd_zip>"&ooptCd&"</tns:itemCd_zip>"
							strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
							If YtoN = True Then
								OPTYN = "Y"
								strRst = strRst &"	<tns:packInd>A</tns:packInd>"						'!!!A-����, I-�Ͻ��ߴ�
								Pmode = "OptMod"
							ElseIf NtoY = True Then
								OPTYN = "N"
								strRst = strRst &"	<tns:packInd>I</tns:packInd>"						'!!!A-����, I-�Ͻ��ߴ�
								Pmode = "OptMod"
							ElseIf DelOpt = True Then
								OPTYN = "N"
								strRst = strRst &"	<tns:packInd>I</tns:packInd>"						'!!!A-����, I-�Ͻ��ߴ�
								Pmode = "OptDel"
							End If
							strRst = strRst &"</tns:itemStates>"
							strRst = strRst &"</tns:ifRequest>"
							getXMLString = strRst
						Else
																		'3.�ɼ� ������ ����Ǵ°��
						End If
					Next
				End If
				If getXMLString = "" Then
					getXMLString = "MDT_NOT"							'������ ���� ���� ���
				End If
			ElseIf mode = "EDT" Then
				getXMLString = oCJMallItem.FItemList(0).getcjmallItemModXML("unitN")
			ElseIf mode = "PRI" Then
				getXMLString = oCJMallItem.FItemList(0).getcjmallItemPriceModXML()
			ElseIf mode = "QTY" Then
				getXMLString = oCJMallItem.FItemList(0).getcjmallItemQTYXML()
			End If
		End If
	ELSEIF (mode="ORDLIST") or (mode="ORDCANCELLIST") then
	    strRst = ""
        strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
        strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_04_01"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_04_01.xsd"">"
        strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
        strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
        strRst = strRst &"<tns:contents>"
        IF (mode="ORDLIST") THEN
            strRst = strRst &"	<tns:instructionCls>"&"1"&"</tns:instructionCls>"
        ELSEIF (mode="ORDCANCELLIST") then
            strRst = strRst &"	<tns:instructionCls>"&"2"&"</tns:instructionCls>"
        END IF
        strRst = strRst &"	<tns:wbCrtDt>"&iitemid&"</tns:wbCrtDt>" ''��ȸ��¥
        strRst = strRst &"</tns:contents>"
        strRst = strRst &"</tns:ifRequest>"
        getXMLString = strRst
    ELSEIF (mode="CSLIST") then
		'// CS������ ��� iitemid �� ��¥�̴�.
		strRst = ""
        strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"
        strRst = strRst & "<tns:ifRequest tns:ifId=""IF_04_02"" xmlns:tns=""http://www.example.org/ifpa"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_04_02.xsd "">"
        strRst = strRst & "<tns:vendorId>411378</tns:vendorId>"
        strRst = strRst & "<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
        strRst = strRst & "<tns:contents>"
        strRst = strRst & "<tns:wbCrtDt>"&iitemid&"</tns:wbCrtDt>"
        strRst = strRst & "<tns:autoFlag></tns:autoFlag>" 			''// ��ȸ���� - �ڵ�ȸ��Ȯ������ Enum(""=��ü, 0=N, 1=Y)
        strRst = strRst & "</tns:contents>"
        strRst = strRst & "</tns:ifRequest>"

		''response.write strRst
		''dbget.close()
		''response.end
        getXMLString = strRst
	ELSEIF (mode="CSORDLIST") or (mode="CSORDCANCELLIST") then
		'// CS������ ��� iitemid �� ��¥�̴�.
	    strRst = ""
        strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"
        strRst = strRst & "<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_04_01"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_04_01.xsd"">"
        strRst = strRst & "<tns:vendorId>411378</tns:vendorId>"
        strRst = strRst & "<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
        strRst = strRst & "<tns:contents>"
        IF (mode="CSORDLIST") THEN
            strRst = strRst &"	<tns:instructionCls>"&"1"&"</tns:instructionCls>"
        ELSEIF (mode="CSORDCANCELLIST") then
            strRst = strRst &"	<tns:instructionCls>"&"2"&"</tns:instructionCls>"
        END IF
        strRst = strRst & "	<tns:wbCrtDt>"&iitemid&"</tns:wbCrtDt>" ''��ȸ��¥
        strRst = strRst & "</tns:contents>"
        strRst = strRst & "</tns:ifRequest>"
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
		oCJMallItem.FRectMatchCateNotCheck = "on"
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
		ioriginName = "����"
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
		ioriginName = "�ٹ�����"
	End If

	getmakerName2Code = retVal
	Exit Function
End Function

Function regCjMallOneItem(byval iitemid, byRef ierrStr)
	''rw  "��ǰ����������"
	''regCjMallOneItem = False
	''Exit function
	''response.end
	Dim sqlStr, AssignedRow
	Dim mode : mode = "REG"
	Dim xmlStr : xmlStr = getXMLString(iitemid, mode) ''�ɼ� �߰��ݾ� �ִ� ��ǰ�� ��� �Ұ��ϰ�..
	Dim cause
	If (xmlStr = "") Then
		ierrStr = "��ϺҰ�"
		''��ϺҰ� ������ �Ѹ�..
		sqlStr = ""
		sqlStr = sqlStr & " SELECT i.itemid, isNULL(R.cjmallStatCD,-9) asCjmallStatCD " & VBCRLF
		sqlStr = sqlStr & " ,i.sellyn,i.limityn,i.limitno,i.limitsold, isnull(PD.CddKey, '') as CddKey, isnull(c.mapCnt,'') as mapCnt " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_cjmall_regItem as R on i.itemid=R.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_cjMall_prdDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid='cjmall' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small " & VBCRLF
		sqlStr = sqlStr & " WHERE i.itemid = "&iitemid
		rsget.Open sqlStr,dbget,1
		If Not(rsget.EOF or rsget.BOF) Then
			If (rsget("asCjmallStatCD") >= 3) Then
				ierrStr = ierrStr & " - ���ϻ�ǰ"&" :: ����["&rsget("asCjmallStatCD")&"]"
			End If

			If (rsget("sellyn") <> "Y") Then
			    ierrStr = ierrStr & " - ǰ������"
			End If

			If (rsget("mapCnt") = "0") Then
			    ierrStr = ierrStr & " - ī�װ��̸�Ī����"
			End If

			If (rsget("CddKey") = "") Then
			    ierrStr = ierrStr & " - ��ǰ�̺з�����"
			End If

			If (rsget("limityn") = "Y") and (rsget("limitno") - rsget("limitsold") < CMAXLIMITSELL) Then
				ierrStr = ierrStr & " - �������� ���� ("&rsget("limitno")-rsget("limitsold")&") �� ����"
				cause = "limitErr"
			End If
	    Else
			ierrStr = ierrStr & " - ��ǰ��ȸ�Ұ�"
	    End If
		rsget.Close

		''�Ұ� ������ ��ã�� ���
		If (ierrStr = "��ϺҰ�") Then
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
				If (rsget("optAddCNT") > 0) Then
					ierrStr = ierrStr & " - �ɼ��߰� �ݾ� �����ǰ ��ϺҰ�"
					cause = "optAddPrcExist"
				End If

				If (rsget("optCnt") - rsget("optNotSellCnt") < 1) Then
					ierrStr = ierrStr & " - �ɼ� �ǸŰ��ɻ�ǰ ����."
					cause = "noValidOpt"
				End If
			End If
			rsget.Close
		End If

		If (cause <> "") Then
			''�������� üũ�ؾ�..
			sqlStr = ""
			sqlStr = sqlStr & "INSERT INTO db_temp.dbo.tbl_jaehyumall_not_in_itemid " & VBCRLF
			sqlStr = sqlStr & "(itemid, mallgubun, bigo) " & VBCRLF
			sqlStr = sqlStr & " SELECT i.itemid, '"&CMALLNAME&"', '"&cause&"'" & VBCRLF
			sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i" & VBCRLF
			sqlStr = sqlStr & " LEFT JOIN db_temp.dbo.tbl_jaehyumall_not_in_itemid as n on i.itemid = n.itemid " & VBCRLF
			sqlStr = sqlStr & " and n.mallgubun = '"&CMALLNAME&"' " & VBCRLF
			sqlStr = sqlStr & " WHERE i.itemid = "&iitemid & VBCRLF
			sqlStr = sqlStr & " and n.itemid is NULL"
			dbget.Execute sqlStr
		End If

		If (ierrStr <> "��ϺҰ�") Then
			ierrStr = iitemid &":"& ierrStr
		End If
		regCjMallOneItem = False
		Exit Function
    End If

    IF (isCJ_DebugMode) Then
        CALL XMLFileSave(xmlStr, mode, iitemid)
    End If

    ''��Ͽ������� ���.
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

Function editCjmallOneItem(byval iitemid, byRef ierrStr)    ''��ǰ ���� ����
	Dim sqlStr, AssignedRow
	Dim mode : mode = "EDT"
	Dim xmlStr : xmlStr = getXMLString(iitemid, mode)

	If (xmlStr="") Then
		ierrStr = "�����Ұ�"
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

Function editPriceCjmallOneItem(byval iitemid, byRef ierrStr)    ''��ǰ ���� ����
	Dim sqlStr, AssignedRow
	Dim mode : mode = "PRI"
	Dim xmlStr : xmlStr = getXMLString(iitemid, mode)

	If (xmlStr="") Then
		ierrStr = "�����Ұ�"
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
End Function

'�� �Ⱓ ����Ʈ
Function listCjMallItem()
	Dim sqlStr, AssignedRow
	Dim mode : mode = "LIST"
	Dim xmlStr : xmlStr = getXMLString(iitemid, mode)
	Dim cause
	If (xmlStr = "") Then
		listCjMallItem = False
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
		listCjMallItem = saveListResult(retDoc, mode)
    SET retDoc = Nothing
End Function

'�����Ⱓ ����Ʈ
Function daylistCjMallItem(sday)
	Dim sqlStr, AssignedRow
	Dim mode : mode = "DayLIST"
	Dim xmlStr : xmlStr = getXMLString(sday, mode)
	Dim cause
	If (xmlStr = "") Then
		daylistCjMallItem = False
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
		daylistCjMallItem = saveListResult(retDoc, mode)
    SET retDoc = Nothing
End Function

'CJ�ֹ����� ��ȸ
Function getCjOrderList(imode,sday) ''"ORDLIST" , "ORDCANCELLIST"
    Dim mode : mode = imode
	Dim xmlStr : xmlStr = getXMLString(sday, mode)

	If (xmlStr = "") Then
		getCjOrderList = False
		Exit Function
    End If

    Dim retDoc, sURL
    sURL = cjMallAPIURL

    SET retDoc = xmlSend(sURL, xmlStr)
    ''response.write retDoc.XML
	    If (isCJ_DebugMode) Then
	        CALL XMLFileSave(retDoc.XML, "RET_"&mode, sday)
	    End If
		getCjOrderList = saveORDListResult(retDoc, mode, sday)
    SET retDoc = Nothing
End Function

'CJ CS���� ��ȸ(ȸ������)
Function getCjCsList(imode,sday)
    Dim mode : mode = imode
	Dim xmlStr : xmlStr = getXMLString(sday, mode)

	''rw xmlStr
	If (xmlStr = "") Then
		getCjCsList = False
		Exit Function
    End If

    Dim retDoc, sURL
    sURL = cjMallAPIURL

	''rw sURL
    SET retDoc = xmlSend(sURL, xmlStr)
	''response.write retDoc.XML
	    If (isCJ_DebugMode) Then
	        CALL XMLFileSave(retDoc.XML, "RET_"&mode, sday)
	    End If
		getCjCsList = saveCSListResult(retDoc, mode, sday)
    SET retDoc = Nothing
End Function

'CJ CS���� ��ȸ(�ֹ�����)
Function getCjCsListInOrder(imode,sday)
    Dim mode : mode = imode
	Dim xmlStr : xmlStr = getXMLString(sday, mode)

	If (xmlStr = "") Then
		getCjCsListInOrder = False
		Exit Function
    End If

    Dim retDoc, sURL
    sURL = cjMallAPIURL

	''rw sURL
    SET retDoc = xmlSend(sURL, xmlStr)
	''response.write retDoc.XML
	    If (isCJ_DebugMode) Then
	        CALL XMLFileSave(retDoc.XML, "RET_"&mode, sday)
	    End If
		getCjCsListInOrder = saveCSOrderListResult(retDoc, mode, sday)
    SET retDoc = Nothing
End Function

'CJ�����ڵ� ��ȸ
Function getcjCommonCodeList(ccd)
    Dim mode : mode = "commonCD"
	Dim xmlStr : xmlStr = getXMLString(ccd, mode)

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


Function editDTCjmallOneItem(byval iitemid, byRef ierrStr)      ''��ǰ ������� ����
	Dim sqlStr, AssignedRow
	Dim mode : mode = "MDT"
	Dim xmlStr : xmlStr = getXMLString(iitemid, mode)

	If xmlStr = "MDT_NOT" Then
		rw "������ �� ����"
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


'���û�ǰ���� ������
Function editSellStatusCjmallOneItem(byval iitemid, byRef ierrStr, cmd)
	Dim sqlStr, AssignedRow
	Dim mode : mode = "SLD"
	Dim xmlStr : xmlStr = getXMLSellyn(iitemid,mode,cmd)

	If (xmlStr = "") Then
		ierrStr = "�����Ұ�"
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

''��ǰ ���� ����
Function editqtyCjmallOneItem(byval iitemid, byRef ierrStr, cmd)
	Dim sqlStr, AssignedRow
	Dim mode : mode = "QTY"
	Dim xmlStr : xmlStr = getXMLString(iitemid, mode)

	If (xmlStr="") Then
		ierrStr = "���� �����Ұ�"
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

'���� ��ǰ �Ͻ��ߴ�/�Ǹſ�
Function saveSellYNItemResult(retDoc, mode, prdno, smdcd)
	Dim errorMsg
	Dim sqlStr
	Dim AssignedRow, successYn
	Dim itemCd_zip
	successYn = false

	If (Not (retDoc is Nothing)) Then
		errorMsg = retDoc.getElementsByTagName("ns1:errorMsg").item(0).text
		If Left(errorMsg,4) = "[����]" Then
			successYn= true
		End If
	End If

	If (successYn = true) Then
		If mode = "MDT" Then
			itemCd_zip = retDoc.getElementsByTagName("ns1:itemCd_zip").item(0).text
			sqlStr = ""
			sqlStr = sqlStr & " UPDATE [db_item].[dbo].tbl_OutMall_regedoption SET " & VBCRLF
			sqlStr = sqlStr & " outmallSellyn = '"&smdcd&"', lastupdate = getdate() " & VBCRLF
			sqlStr = sqlStr & " WHERE itemid = '"&prdno&"'  " & VBCRLF
			sqlStr = sqlStr & " and outmallOptCode = '"&itemCd_zip&"' " & VBCRLF
			sqlStr = sqlStr & " and mallid = 'cjmall' " & VBCRLF
			dbget.Execute sqlStr,AssignedRow
			saveSellYNItemResult=true
		ElseIf mode = "Del" Then
			itemCd_zip = retDoc.getElementsByTagName("ns1:itemCd_zip").item(0).text
			sqlStr = ""
			sqlStr = sqlStr & " DELETE FROM [db_item].[dbo].tbl_OutMall_regedoption " & VBCRLF
			sqlStr = sqlStr & " WHERE itemid = '"&prdno&"'  " & VBCRLF
			sqlStr = sqlStr & " and outmallOptCode = '"&itemCd_zip&"' " & VBCRLF
			sqlStr = sqlStr & " and mallid = 'cjmall' " & VBCRLF
			dbget.Execute sqlStr,AssignedRow
			saveSellYNItemResult=true
		Else
			If (prdno <> "") Then
				sqlStr = ""
				sqlStr = sqlStr & " UPDATE R" & VBCRLF
				sqlStr = sqlStr & " SET cjmallLastUpdate = getdate()" & VBCRLF
				sqlStr = sqlStr & " ,cjmallSellYn = '"&smdcd&"'" & VBCRLF
				sqlStr = sqlStr & " ,accFailCNT=0" & VBCRLF                 ''����ȸ�� �ʱ�ȭ
				sqlStr = sqlStr & " FROM db_item.dbo.tbl_cjmall_regItem as R" & VBCRLF
				sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid" & VBCRLF
				sqlStr = sqlStr & " WHERE R.itemid = "&prdno&""   & VBCRLF
				dbget.Execute sqlStr,AssignedRow
				saveSellYNItemResult=true
			End If
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
	End If

	If (isCJ_DebugMode) Then
		rw prdno &"_"&errorMsg
	End If
End Function

'�ֹ����� �����
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
                ordNo       = masterSubNodes.getElementsByTagName("ns1:ordNo")(0).Text	        '�ֹ���ȣ
                custNm      = masterSubNodes.getElementsByTagName("ns1:custNm")(0).Text	        '�ֹ���
                custTelNo   = masterSubNodes.getElementsByTagName("ns1:custTelNo")(0).Text	    '�ֹ��� ��ȭ
                custDeliveryCost = masterSubNodes.getElementsByTagName("ns1:custDeliveryCost")(0).Text	'��ۺ�

                Set detailSubNodes = masterSubNodes.getElementsByTagName("ns1:instructionDetail")

                ''rw ordNo&"|"&custNm&"|"&custTelNo&"|"&custDeliveryCost

                dtlCnt = 1
                If (Not (detailSubNodes is Nothing)) Then
                    For each detailSubNodeItem in detailSubNodes
                        requireDetail = ""
                        ierrStr = ""

                        ordGSeq = detailSubNodeItem.getElementsByTagName("ns1:ordGSeq")(0).Text	    '[ID:�ֹ���ǰ����], 001
                        ordDSeq = detailSubNodeItem.getElementsByTagName("ns1:ordDSeq")(0).Text	    '[ID:�ֹ��󼼼���], 001
                        ordWSeq = detailSubNodeItem.getElementsByTagName("ns1:ordWSeq")(0).Text	    '[ID:�ֹ�ó������], 001
                        ordDtlCls = detailSubNodeItem.getElementsByTagName("ns1:ordDtlCls")(0).Text	        ' �ֹ����� - �ֹ�����, �ֹ�
                        ordDtlClsCd = detailSubNodeItem.getElementsByTagName("ns1:ordDtlClsCd")(0).Text	    ' �ֹ����� - �ֹ������ڵ�, 10
                        wbCrtDt = detailSubNodeItem.getElementsByTagName("ns1:wbCrtDt")(0).Text	            ' �ֹ����� - ��������, 2013-05-22+09:00
                        ''outwConfDt	'�ֹ����� - ���Ȯ������
                        ''delivDtm	    '�ֹ����� - ��ۿϷ���
                        ''cnclInsDtm	'�ֹ����� - �������
                        ''oldordNo	    '�ֹ����� - ���ֹ���ȣ
                        toutYn = detailSubNodeItem.getElementsByTagName("ns1:toutYn")(0).Text	            '�ֹ����� - �����ϱ���(Y-������,N-��������), N
                        chnNm = detailSubNodeItem.getElementsByTagName("ns1:chnNm")(0).Text	                '�ֹ����� - ä�α���, INTERNET
                        receverNm = detailSubNodeItem.getElementsByTagName("ns1:receverNm")(0).Text	        '�ֹ����� - �μ���, ä����
                        'recvName	    '�ֹ����� - ������ ������
                        zipno = detailSubNodeItem.getElementsByTagName("ns1:zipno")(0).Text	                '�ֹ����� - �����ȣ, 110809
                        addr_1 = detailSubNodeItem.getElementsByTagName("ns1:addr_1")(0).Text	            '�ֹ����� - �ּ�, ���� ���α� ������
                        addr_2 = detailSubNodeItem.getElementsByTagName("ns1:addr_2")(0).Text	            '�ֹ����� - ���ּ�, 1-45���� �������� 6��
                        'addr	        '�ֹ����� - �ּ�
                        telno = detailSubNodeItem.getElementsByTagName("ns1:telno")(0).Text	                '�ֹ����� - �μ���tel, 02)973-8514
                        cellno = detailSubNodeItem.getElementsByTagName("ns1:cellno")(0).Text	            '�ֹ����� - �μ���HP, 010)2715-8514
                        'msgSpec	    '�ֹ����� - �������
                        'delvplnDt	    '�ֹ����� - ��ۿ�����
                        packYn = detailSubNodeItem.getElementsByTagName("ns1:packYn")(0).Text	            '��ǰ���� - ��Ʈ����, �Ϲ�
                        'itemNm	        '��ǰ���� - ��Ʈ��ǰ��
                        itemCd = detailSubNodeItem.getElementsByTagName("ns1:itemCd")(0).Text	            '��ǰ���� - �Ǹ��ڵ�, 21899852
                        unitCd = detailSubNodeItem.getElementsByTagName("ns1:unitCd")(0).Text	            '��ǰ���� - ��ǰ�ڵ�, 10047125217
                        itemName = detailSubNodeItem.getElementsByTagName("ns1:itemName")(0).Text	        '��ǰ���� - �ǸŻ�ǰ��, 24K Gold ���������ܽ�ƼĿ
                        unitNm = detailSubNodeItem.getElementsByTagName("ns1:unitNm")(0).Text	            '��ǰ���� - ��ǰ��, ES-01 �ߵɲ���
                        contItemCd = detailSubNodeItem.getElementsByTagName("ns1:contItemCd")(0).Text	    '��ǰ���� - ���»��ǰ�ڵ�, 279751_0011
                        wbIdNo = detailSubNodeItem.getElementsByTagName("ns1:wbIdNo")(0).Text	            '��ǰ���� - �����ĺ���ȣ, 20000420537940
                        outwQty = detailSubNodeItem.getElementsByTagName("ns1:outwQty")(0).Text	            '��ǰ���� - ����, 1.0
                        realslAmt = detailSubNodeItem.getElementsByTagName("ns1:realslAmt")(0).Text	        '��ǰ���� - �ǸŰ�, 1800.0
                        outwAmt = detailSubNodeItem.getElementsByTagName("ns1:outwAmt")(0).Text	            '��ǰ���� - ��������, 1800.0  :: ����*�ǸŰ� ����, ����*���ǸŰ����� Ȯ��
                        'delivInfo	    '��Ÿ���� - ���
                        'promGiftSpec	'��Ÿ���� - ����ǰ����
                        'juminNum       '�ֹ�����-�ֹι�ȣ(�ƴ�), �߼� ����!
                        'cnclRsn	    '��Ÿ���� - ��ȯ/��һ���
                        'cnclRsnSpec	'��Ÿ���� - ��ȯ/��һ�����
                        ordDtm = detailSubNodeItem.getElementsByTagName("ns1:ordDtm")(0).Text	            '�ֹ�����-�ֹ��Ͻ�, 2013-05-22 15:05:02


                        ''�ʼ��� �ȳѾ����������.
                        outwConfDt =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:outwConfDt")(0) Is Nothing)) Then
                            outwConfDt = detailSubNodeItem.getElementsByTagName("ns1:outwConfDt")(0).Text       '�ֹ����� - ���Ȯ������
                        end if
                        delivDtm =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:delivDtm")(0) Is Nothing)) Then
                            delivDtm = detailSubNodeItem.getElementsByTagName("ns1:delivDtm")(0).Text        '�ֹ����� - ��ۿϷ���
                        end if
                        cnclInsDtm =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:cnclInsDtm")(0) Is Nothing)) Then
                            cnclInsDtm = detailSubNodeItem.getElementsByTagName("ns1:cnclInsDtm")(0).Text        '�ֹ����� - �������
                        end if
                        oldordNo =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:oldordNo")(0) Is Nothing)) Then
                            oldordNo = detailSubNodeItem.getElementsByTagName("ns1:oldordNo")(0).Text        '�ֹ����� - ���ֹ���ȣ
                        end if
                        recvName =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:recvName")(0) Is Nothing)) Then
                            recvName = detailSubNodeItem.getElementsByTagName("ns1:recvName")(0).Text        '�ֹ����� - ������ ������
                        end if
                        addr =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:addr")(0) Is Nothing)) Then
                            addr = detailSubNodeItem.getElementsByTagName("ns1:addr")(0).Text        '�ֹ����� - �ּ�all?
                        end if
                        msgSpec =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:msgSpec")(0) Is Nothing)) Then
                            msgSpec = detailSubNodeItem.getElementsByTagName("ns1:msgSpec")(0).Text        '�ֹ����� -�������
                        end if
                        delvplnDt =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:delvplnDt")(0) Is Nothing)) Then
                            delvplnDt = detailSubNodeItem.getElementsByTagName("ns1:delvplnDt")(0).Text        '�ֹ����� -��ۿ�����
                        end if
                        itemNm =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:itemNm")(0) Is Nothing)) Then
                            itemNm = detailSubNodeItem.getElementsByTagName("ns1:itemNm")(0).Text        '��ǰ���� -��Ʈ��ǰ��
                        end if
                        juminNum =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:juminNum")(0) Is Nothing)) Then
                            juminNum = detailSubNodeItem.getElementsByTagName("ns1:juminNum")(0).Text       '�ֹ�����-�ֹι�ȣ(�ƴ�), �߼� ����!
                        end if
                        dccouponCjhs =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:dccouponCjhs")(0) Is Nothing)) Then
                            dccouponCjhs = detailSubNodeItem.getElementsByTagName("ns1:dccouponCjhs")(0).Text       '�ֹ����� - ����(���δ�)�ݾ�
                        end if
                        dccouponVendor =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:dccouponVendor")(0) Is Nothing)) Then
                            dccouponVendor = detailSubNodeItem.getElementsByTagName("ns1:dccouponVendor")(0).Text      '�ֹ����� - ����(���»�δ�)�ݾ�
                        end if

                        orderDlvPay = custDeliveryCost
                        if (dtlCnt>1) then
                            orderDlvPay = 0 ''ù��° ���� ����.
                        end if

                        orderCsGbn = ""
                        if (ordDtlClsCd="10") then
                            orderCsGbn="0"
                        elseif (ordDtlClsCd="20") then
                            orderCsGbn="2"  ''���
                        end if

                        requireDetail = juminNum '' �ֹ����۹���
                        ierrCode = 0
                        ierrStr  = ""

                        if (orderCsGbn<>"") and (masterSubNodes.getElementsByTagName("ns1:oldordNo")(0) is Nothing) then

							IsDetailItemAllCancel = False
							IsCancelOrgOrder = False

							if (orderCsGbn = "2") then
								'// ���
								strSql = " select matchState, orderDlvPay, orgOrderCNT from db_temp.dbo.tbl_xSite_TMPOrder "
								strSql = strSql + " where SellSite = 'cjmall' and OutMallOrderSerial = '" + CStr(ordNo) + "' and OrgDetailKey = '" & ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq & "' "
								''rw strSql
								rsget.Open strSql,dbget,1
								if (Not rsget.Eof) then
									if (CLng(outwQty) = rsget("orgOrderCNT")) then
										'// Ư����ǰ ��ü���
										IsDetailItemAllCancel = True
										if (rsget("matchState") = "I") then
											'// �ֹ��Է�����
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

							if (IsDetailItemAllCancel) then
								'// �켱 ���� ��ü��Ҹ� ó��(���� �Ϻ���Ҵ� ���� �ԷµǸ� ó��)
								retVal = saveORDOneTmp(ordNo,ordDtm,splitvalue(contItemCd,"_",0),splitvalue(contItemCd,"_",1),itemName, unitNm _
										, custNm , custTelNo, custTelNo _
										, receverNm, telno, cellno, LEFT(zipno,3)&"-"&Right(zipno,3), addr_1, addr_2 _
										, realslAmt, realslAmt, CLNG(outwQty), ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq & "-CA" _
										, msgSpec, requireDetail, orderDlvPay, orderCsGbn _
										, ierrCode, ierrStr)

								'// ���ֹ� �����Ǿ����� CS�� ����
								strSql = " if exists (select OutMallOrderSeq from db_temp.dbo.tbl_xSite_TMPOrder where SellSite = 'cjmall' and OutMallOrderSerial = '" + CStr(ordNo) + "' and OrgDetailKey = '" & ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq & "' and matchState = 'D') "
								strSql = strSql + " begin "
								strSql = strSql + " 	update db_temp.dbo.tbl_xSite_TMPOrder set matchState = 'D' where SellSite = 'cjmall' and OutMallOrderSerial = '" + CStr(ordNo) + "' and OrgDetailKey = '" & ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq & "-CA' and matchState = 'I' "
								strSql = strSql + " end "
								rsget.Open strSql, dbget, 1
							else
								retVal = saveORDOneTmp(ordNo,ordDtm,splitvalue(contItemCd,"_",0),splitvalue(contItemCd,"_",1),itemName, unitNm _
										, custNm , custTelNo, custTelNo _
										, receverNm, telno, cellno, LEFT(zipno,3)&"-"&Right(zipno,3), addr_1, addr_2 _
										, realslAmt, realslAmt, CLNG(outwQty), ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq _
										, msgSpec, requireDetail, orderDlvPay, orderCsGbn _
										, ierrCode, ierrStr)
							end if
                        elseif Not (masterSubNodes.getElementsByTagName("ns1:oldordNo")(0) is Nothing) then
							''������ �ֹ� ó�� : oldordNo	�ֹ����� - ���ֹ���ȣ
							''��ǰ���� �Ͽ��ٰ� �ٽ� ��ǰ�� öȸ�ϴ� ��� ��
                            retVal = false
                            ierrStr = "������ �ֹ� [oldordNo="&masterSubNodes.getElementsByTagName("ns1:oldordNo")(0).Text&"] "
						else
                            retVal = false
                            ierrStr = "�ֹ����� [ordDtlClsCd="&ordDtlClsCd&"] ���ǵ��� ����"
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
    rw succCnt & "�� �Է�"
    rw failCnt & "�� ����" & "("&skipCnt&" �� skip)"

End function

'CS���� �����(ȸ������)
Function saveCSListResult(retDoc, mode, sday)

	'' '// TODO : !!!!
    '' Exit function

    Dim Nodes, masterSubNodes, detailSubNodes, detailSubNodeItem, ErrNode, errorMsg
    Dim isErrExists : isErrExists = false

	dim ordNo, custNm, ordGSeq, ordDSeq, ordWSeq, wbClsCd, wbCls, confirmChk, wbCrtDt, delivDtm, chnNm
	dim recvNm, zipno, addr, telno, cellno, packYn
	dim itemNm, itemCd, unitCd, compOrdNo, contItemCd, itemName, unitNm, recallConfQty, slPrc, custAmt, outwAmt, wbIdNo, wbNo
	dim deliComp, autoFlag, msgSpec, vocNm, claimNote, wbProg, exceptRsnFlg, recvBefCnclYn

    Dim requireDetail, orderDlvPay, orderCsGbn, ierrStr, ierrCode, dtlCnt
    dim succCnt : succCnt=0
    dim failCnt : failCnt=0
    dim skipCnt : skipCnt=0
    dim retVal

	dim strSql

	dim SellSite, OutMallOrderSerial, OrgDetailKey, CSDetailKey

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
                ordNo       = masterSubNodes.getElementsByTagName("ns1:ordNo")(0).Text	        '�ֹ���ȣ
				custNm		= ""

                Set detailSubNodes = masterSubNodes.getElementsByTagName("ns1:instructionDetail")

                dtlCnt = 1
                If (Not (detailSubNodes is Nothing)) Then
                    For each detailSubNodeItem in detailSubNodes

						ordGSeq = ""
						ordDSeq = ""
						ordWSeq = ""
						wbClsCd = ""
						wbCls = ""
						confirmChk = ""
						wbCrtDt = ""
						delivDtm = ""
						chnNm = ""
						recvNm = ""
						zipno = ""
						addr = ""
						telno = ""
						cellno = ""
						packYn = ""
						itemNm = ""
						itemCd = ""
						unitCd = ""
						compOrdNo = ""
						contItemCd = ""
						itemName = ""
						unitNm = ""
						recallConfQty = ""
						slPrc = ""
						custAmt = ""
						outwAmt = ""
						wbIdNo = ""
						wbNo = ""
						deliComp = ""
						autoFlag = ""
						msgSpec = ""
						vocNm = ""
						claimNote = ""
						wbProg = ""
						exceptRsnFlg = ""
						recvBefCnclYn = ""

                        requireDetail = ""
                        ierrStr = ""

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:ordGSeq")(0) Is Nothing)) Then
							'// �ֹ����� - �ֹ���ǰ����
                            ordGSeq = detailSubNodeItem.getElementsByTagName("ns1:ordGSeq")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:ordDSeq")(0) Is Nothing)) Then
							'// �ֹ����� - �ֹ��󼼼���
                            ordDSeq = detailSubNodeItem.getElementsByTagName("ns1:ordDSeq")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:ordWSeq")(0) Is Nothing)) Then
							'// �ֹ����� - �ֹ�ó������
                            ordWSeq = detailSubNodeItem.getElementsByTagName("ns1:ordWSeq")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:wbClsCd")(0) Is Nothing)) Then
							'// �ֹ����� - ���౸���ڵ�
                            wbClsCd = detailSubNodeItem.getElementsByTagName("ns1:wbClsCd")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:wbCls")(0) Is Nothing)) Then
							'// �ֹ����� - ���౸��
                            wbCls = detailSubNodeItem.getElementsByTagName("ns1:wbCls")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:confirmChk")(0) Is Nothing)) Then
							'// �ֹ����� - ���»����ȸ��Ȯ��
                            confirmChk = detailSubNodeItem.getElementsByTagName("ns1:confirmChk")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:wbCrtDt")(0) Is Nothing)) Then
							'// �ֹ����� - ȸ����������
                            wbCrtDt = detailSubNodeItem.getElementsByTagName("ns1:wbCrtDt")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:delivDtm")(0) Is Nothing)) Then
							'// �ֹ����� - ȸ����������
                            delivDtm = detailSubNodeItem.getElementsByTagName("ns1:delivDtm")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:chnNm")(0) Is Nothing)) Then
							'// �ֹ����� - ä�α���
                            chnNm = detailSubNodeItem.getElementsByTagName("ns1:chnNm")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:recvNm")(0) Is Nothing)) Then
							'// �ֹ����� - �μ���
                            recvNm = detailSubNodeItem.getElementsByTagName("ns1:recvNm")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:zipno")(0) Is Nothing)) Then
							'// �ֹ����� - �����ȣ
                            zipno = detailSubNodeItem.getElementsByTagName("ns1:zipno")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:addr")(0) Is Nothing)) Then
							'// �ֹ����� - �ּ�
                            addr = detailSubNodeItem.getElementsByTagName("ns1:addr")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:telno")(0) Is Nothing)) Then
							'// �ֹ����� - �μ���tel
                            telno = detailSubNodeItem.getElementsByTagName("ns1:telno")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:cellno")(0) Is Nothing)) Then
							'// �ֹ����� - �μ���HP
                            cellno = detailSubNodeItem.getElementsByTagName("ns1:cellno")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:packYn")(0) Is Nothing)) Then
							'// ��ǰ���� - ��Ʈ����
                            packYn = detailSubNodeItem.getElementsByTagName("ns1:packYn")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:itemNm")(0) Is Nothing)) Then
							'// ��ǰ���� - ��Ʈ��ǰ��
                            itemNm = detailSubNodeItem.getElementsByTagName("ns1:itemNm")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:unitNm")(0) Is Nothing)) Then
							'// ��ǰ���� - ��ǰ��
                            unitNm = detailSubNodeItem.getElementsByTagName("ns1:unitNm")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:recallConfQty")(0) Is Nothing)) Then
							'// ��ǰ���� - ȸ������
                            recallConfQty = detailSubNodeItem.getElementsByTagName("ns1:recallConfQty")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:slPrc")(0) Is Nothing)) Then
							'// ��ǰ���� - �ǸŰ�
                            slPrc = detailSubNodeItem.getElementsByTagName("ns1:slPrc")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:custAmt")(0) Is Nothing)) Then
							'// ��ǰ���� - ��������
                            custAmt = detailSubNodeItem.getElementsByTagName("ns1:custAmt")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:outwAmt")(0) Is Nothing)) Then
							'// ��ǰ���� - ȸ���ݾ�
                            outwAmt = detailSubNodeItem.getElementsByTagName("ns1:outwAmt")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:wbIdNo")(0) Is Nothing)) Then
							'// ������� - �����ĺ���ȣ
                            wbIdNo = detailSubNodeItem.getElementsByTagName("ns1:wbIdNo")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:wbNo")(0) Is Nothing)) Then
							'// ������� - ������ȣ
                            wbNo = detailSubNodeItem.getElementsByTagName("ns1:wbNo")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:deliComp")(0) Is Nothing)) Then
							'// ������� - �ù��
                            deliComp = detailSubNodeItem.getElementsByTagName("ns1:deliComp")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:autoFlag")(0) Is Nothing)) Then
							'// ������� - �ڵ�ȸ��Ȯ������
                            autoFlag = detailSubNodeItem.getElementsByTagName("ns1:autoFlag")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:msgSpec")(0) Is Nothing)) Then
							'// ������� - ����
                            msgSpec = detailSubNodeItem.getElementsByTagName("ns1:msgSpec")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:vocNm")(0) Is Nothing)) Then
							'// ��Ÿ���� - ��ǰ����
                            vocNm = detailSubNodeItem.getElementsByTagName("ns1:vocNm")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:claimNote")(0) Is Nothing)) Then
							'// ��Ÿ���� - ��ǰ������
                            claimNote = detailSubNodeItem.getElementsByTagName("ns1:claimNote")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:wbProg")(0) Is Nothing)) Then
							'// ��Ÿ���� - ó������
                            wbProg = detailSubNodeItem.getElementsByTagName("ns1:wbProg")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:exceptRsnFlg")(0) Is Nothing)) Then
							'// ��Ÿ���� - û�࿹��
                            exceptRsnFlg = detailSubNodeItem.getElementsByTagName("ns1:exceptRsnFlg")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:recvBefCnclYn")(0) Is Nothing)) Then
							'// ��Ÿ���� - �μ�����ҿ���
                            recvBefCnclYn = detailSubNodeItem.getElementsByTagName("ns1:recvBefCnclYn")(0).Text
                        end if

						'// �����Ͻ�
						'// 2013-05-28+09:00 => 2013-05-28 09:00:00
						wbCrtDt = Replace(wbCrtDt, "+", " ")
						wbCrtDt = wbCrtDt + ":00"

                        orderCsGbn = ""
                        if (wbClsCd="30") then
							'// ��ǰ
                            orderCsGbn="A004"
                        elseif (wbClsCd="45") then
							'// ��ȯ(���ϻ�ǰ �ٸ���ǰ ���??)
                            orderCsGbn="A011"
                        end if

                        ierrCode = 0
                        ierrStr  = ""

                        if (orderCsGbn<>"") then
							SellSite			= "cjmall"
							OutMallOrderSerial	= ordNo
							OrgDetailKey		= CStr(ordNo) & "-" & CStr(ordGSeq) & "-" & CStr(ordDSeq) & "-" & "001"		'// ���ֹ� ������Ű
							CSDetailKey			= CStr(ordNo) & "-" & CStr(ordGSeq) & "-" & CStr(ordDSeq) & "-" & CStr(ordWSeq)

							''response.write ordGSeq & "<br>"
							retVal = True

							strSql = " select idx from db_temp.dbo.tbl_xSite_TMPCS where SellSite = '" + CStr(SellSite) + "' and OutMallOrderSerial = '" + CStr(OutMallOrderSerial) + "' and OrgDetailKey = '" + CStr(OrgDetailKey) + "' and CSDetailKey = '" + CStr(CSDetailKey) + "' "
							rsget.Open strSql,dbget,1
							if (Not rsget.Eof) then
								ierrCode = -1
								retVal = false
							end if
							rsget.Close


							if (ierrCode <> -1) then
								strSql = " if not exists (select idx from db_temp.dbo.tbl_xSite_TMPCS where SellSite = '" + CStr(SellSite) + "' and OutMallOrderSerial = '" + CStr(OutMallOrderSerial) + "' and OrgDetailKey = '" + CStr(OrgDetailKey) + "' and CSDetailKey = '" + CStr(CSDetailKey) + "') "
								strSql = strSql + " begin "
								strSql = strSql + " insert into db_temp.dbo.tbl_xSite_TMPCS(divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
								strSql = strSql + " , OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) "
								strSql = strSql + " values('" + CStr(orderCsGbn) + "', '" + html2db(CStr(vocNm)) + "', '" + html2db(CStr(SellSite)) + "', '" + html2db(CStr(OutMallOrderSerial)) + "', '', '', '', '', '" + html2db(CStr(recvNm)) + "', '" + html2db(CStr(telno)) + "', '" + html2db(CStr(cellno)) + "', '" + html2db(CStr(zipno)) + "', '" + html2db(CStr(addr)) + "', '', '" + html2db(CStr(msgSpec)) + "' "
								strSql = strSql + " , '" + html2db(CStr(wbCrtDt)) + "', '" + html2db(CStr(OrgDetailKey)) + "', '" + html2db(CStr(CSDetailKey)) + "', " + CStr(recallConfQty) + ") "
								strSql = strSql + " end "
								''rw strSql
								rsget.Open strSql, dbget, 1
							end if
                        else
                            retVal = false
                            ierrStr = "CS���� [wbClsCd="&wbClsCd&"] ���ǵ��� ����"
                        end if

                        dtlCnt = dtlCnt+1

                        if (retVal) then
                            succCnt = succCnt+1
                        else
                            failCnt = failCnt+1
                            if (ierrCode=-1) then skipCnt = skipCnt+1

                            if (mode="CSLIST") then
                                rw "<font color='red'>["&ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq&"]</font> "&ierrStr & " "
                            else
                                rw "["&ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq&"] "&ierrStr & " "
                            end if
                        end if

                    Next
                end if

                Set detailSubNodes = Nothing
            Next
        end if
    end if

    Set Nodes = Nothing
    rw succCnt & "�� �Է�"
    rw failCnt & "�� ����" & "("&skipCnt&" �� skip)"

	if (succCnt > 0) then
		strSql = " update c "
		strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
		strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
		strSql = strSql + " , c.OrderName = o.OrderName "
		strSql = strSql + " from "
		strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
		strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
		strSql = strSql + " on "
		strSql = strSql + " 	1 = 1 "
		strSql = strSql + " 	and c.SellSite = o.SellSite "
		strSql = strSql + " 	and c.OutMallOrderSerial = Replace(o.OutMallOrderSerial, '-', '') "
		strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
		strSql = strSql + " where "
		strSql = strSql + " 	1 = 1 "
		strSql = strSql + " 	and c.orderserial is NULL "
		strSql = strSql + " 	and o.orderserial is not NULL "
		strSql = strSql + " 	and c.sellsite = 'cjmall' "
		''rw strSql
		rsget.Open strSql, dbget, 1
	end if

End function

'CS���� �����(�ֹ�����)
Function saveCSOrderListResult(retDoc, mode, sday)

	'' '// TODO : !!!!
    ''Exit function

    Dim Nodes, masterSubNodes, detailSubNodes, detailSubNodeItem, ErrNode, errorMsg
    Dim isErrExists : isErrExists = false

	dim ordNo, custNm, sendName, custTelNo, custDeliveryCost, oldordNo
	dim ordGSeq, ordDSeq, ordWSeq, ordDtlCls, ordDtlClsCd, wbCrtDt, outwConfDt, delivDtm, cnclInsDtm, toutYn, chnNm, receverNm, recvName, zipno, addr_1, addr_2, addr, telno, cellno, msgSpec
	dim delvplnDt, packYn, itemNm, itemCd, unitCd, itemName, unitNm, contItemCd, wbIdNo, outwQty
	dim realslAmt, outwAmt, delivInfo, promGiftSpec, cnclRsn, cnclRsnSpec, ordDtm, juminNum, dccouponCjhs, dccouponVendor

    Dim requireDetail, orderDlvPay, orderCsGbn, ierrStr, ierrCode, dtlCnt
    dim succCnt : succCnt=0
    dim failCnt : failCnt=0
    dim skipCnt : skipCnt=0
    dim retVal

	dim strSql

	dim SellSite, OutMallOrderSerial, OrgDetailKey, CSDetailKey

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
                ordNo       = masterSubNodes.getElementsByTagName("ns1:ordNo")(0).Text	        '�ֹ���ȣ
                custNm      = masterSubNodes.getElementsByTagName("ns1:custNm")(0).Text	        '�ֹ���
                custTelNo   = masterSubNodes.getElementsByTagName("ns1:custTelNo")(0).Text	    '�ֹ��� ��ȭ

				oldordNo = ""
                if (Not (masterSubNodes.getElementsByTagName("ns1:oldordNo")(0) Is Nothing)) Then
					'// �ֹ����� - ���ֹ���ȣ(������ �ֹ��Է½�)
                    ''oldordNo = detailSubNodeItem.getElementsByTagName("ns1:oldordNo")(0).Text
                end if

                Set detailSubNodes = masterSubNodes.getElementsByTagName("ns1:instructionDetail")

                dtlCnt = 1
                If (Not (detailSubNodes is Nothing)) Then
                    For each detailSubNodeItem in detailSubNodes

						ordGSeq			= ""
						ordDSeq			= ""
						ordWSeq			= ""
						ordDtlCls		= ""
						ordDtlClsCd		= ""
						wbCrtDt			= ""
						outwConfDt		= ""
						delivDtm		= ""
						cnclInsDtm		= ""
						oldordNo		= ""
						toutYn			= ""
						chnNm			= ""
						receverNm		= ""
						recvName		= ""
						zipno			= ""
						addr_1			= ""
						addr_2			= ""
						addr			= ""
						telno			= ""
						cellno			= ""
						msgSpec			= ""
						delvplnDt		= ""
						packYn			= ""
						itemNm			= ""
						itemCd			= ""
						unitCd			= ""
						itemName		= ""
						unitNm			= ""
						contItemCd		= ""
						wbIdNo			= ""
						outwQty			= ""
						realslAmt		= ""
						outwAmt			= ""
						delivInfo		= ""
						promGiftSpec	= ""
						cnclRsn			= ""
						cnclRsnSpec		= ""
						ordDtm			= ""
						juminNum		= ""
						dccouponCjhs	= ""
						dccouponVendor	= ""

                        requireDetail = ""
                        ierrStr = ""

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:ordGSeq")(0) Is Nothing)) Then
							'// �ֹ����� - �ֹ���ǰ����
                            ordGSeq = detailSubNodeItem.getElementsByTagName("ns1:ordGSeq")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:ordDSeq")(0) Is Nothing)) Then
							'// �ֹ����� - �ֹ��󼼼���
                            ordDSeq = detailSubNodeItem.getElementsByTagName("ns1:ordDSeq")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:ordWSeq")(0) Is Nothing)) Then
							'// �ֹ����� - �ֹ�ó������
                            ordWSeq = detailSubNodeItem.getElementsByTagName("ns1:ordWSeq")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:ordDtlCls")(0) Is Nothing)) Then
							'// �ֹ����� - �ֹ�����
                            ordDtlCls = detailSubNodeItem.getElementsByTagName("ns1:ordDtlCls")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:ordDtlClsCd")(0) Is Nothing)) Then
							'// �ֹ����� - �ֹ������ڵ�
                            ordDtlClsCd = detailSubNodeItem.getElementsByTagName("ns1:ordDtlClsCd")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:wbCrtDt")(0) Is Nothing)) Then
							'// �ֹ����� - ��������
                            wbCrtDt = detailSubNodeItem.getElementsByTagName("ns1:wbCrtDt")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:outwConfDt")(0) Is Nothing)) Then
							'// �ֹ����� - ���Ȯ������
                            outwConfDt = detailSubNodeItem.getElementsByTagName("ns1:outwConfDt")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:delivDtm")(0) Is Nothing)) Then
							'// �ֹ����� - ��ۿϷ���
                            delivDtm = detailSubNodeItem.getElementsByTagName("ns1:delivDtm")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:cnclInsDtm")(0) Is Nothing)) Then
							'// �ֹ����� - �������
                            cnclInsDtm = detailSubNodeItem.getElementsByTagName("ns1:cnclInsDtm")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:oldordNo")(0) Is Nothing)) Then
							'// �ֹ����� - ���ֹ���ȣ(�������ֹ��� ���)
                            oldordNo = detailSubNodeItem.getElementsByTagName("ns1:oldordNo")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:toutYn")(0) Is Nothing)) Then
							'// �ֹ����� - �����ϱ���
                            toutYn = detailSubNodeItem.getElementsByTagName("ns1:toutYn")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:chnNm")(0) Is Nothing)) Then
							'// �ֹ����� - ä�α���
                            chnNm = detailSubNodeItem.getElementsByTagName("ns1:chnNm")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:receverNm")(0) Is Nothing)) Then
							'// �ֹ����� - �μ���
                            receverNm = detailSubNodeItem.getElementsByTagName("ns1:receverNm")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:recvName")(0) Is Nothing)) Then
							'// �ֹ����� - ������ ������
                            recvName = detailSubNodeItem.getElementsByTagName("ns1:recvName")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:zipno")(0) Is Nothing)) Then
							'// �ֹ����� - �����ȣ
                            zipno = detailSubNodeItem.getElementsByTagName("ns1:zipno")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:addr_1")(0) Is Nothing)) Then
							'// �ֹ����� - �ּ�
                            addr_1 = detailSubNodeItem.getElementsByTagName("ns1:addr_1")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:addr_2")(0) Is Nothing)) Then
							'// �ֹ����� - ���ּ�
                            addr_2 = detailSubNodeItem.getElementsByTagName("ns1:addr_2")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:addr")(0) Is Nothing)) Then
							'// �ֹ����� - �ּ�
                            addr = detailSubNodeItem.getElementsByTagName("ns1:addr")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:telno")(0) Is Nothing)) Then
							'// �ֹ����� - �μ���tel
                            telno = detailSubNodeItem.getElementsByTagName("ns1:telno")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:cellno")(0) Is Nothing)) Then
							'// �ֹ����� - �μ���HP
                            cellno = detailSubNodeItem.getElementsByTagName("ns1:cellno")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:msgSpec")(0) Is Nothing)) Then
							'// �ֹ����� - �������
                            msgSpec = detailSubNodeItem.getElementsByTagName("ns1:msgSpec")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:delvplnDt")(0) Is Nothing)) Then
							'// �ֹ����� - ��ۿ�����
                            delvplnDt = detailSubNodeItem.getElementsByTagName("ns1:delvplnDt")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:packYn")(0) Is Nothing)) Then
							'// ��ǰ���� - ��Ʈ����
                            packYn = detailSubNodeItem.getElementsByTagName("ns1:packYn")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:itemNm")(0) Is Nothing)) Then
							'// ��ǰ���� - ��Ʈ��ǰ��
                            itemNm = detailSubNodeItem.getElementsByTagName("ns1:itemNm")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:itemCd")(0) Is Nothing)) Then
							'// ��ǰ���� - �Ǹ��ڵ�
                            itemCd = detailSubNodeItem.getElementsByTagName("ns1:itemCd")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:unitCd")(0) Is Nothing)) Then
							'// ��ǰ���� - ��ǰ�ڵ�
                            unitCd = detailSubNodeItem.getElementsByTagName("ns1:unitCd")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:itemName")(0) Is Nothing)) Then
							'// ��ǰ���� - �ǸŻ�ǰ��
                            itemName = detailSubNodeItem.getElementsByTagName("ns1:itemName")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:unitNm")(0) Is Nothing)) Then
							'// ��ǰ���� - ��ǰ��
                            unitNm = detailSubNodeItem.getElementsByTagName("ns1:unitNm")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:contItemCd")(0) Is Nothing)) Then
							'// ��ǰ���� - ���»��ǰ�ڵ�
                            contItemCd = detailSubNodeItem.getElementsByTagName("ns1:contItemCd")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:wbIdNo")(0) Is Nothing)) Then
							'// ��ǰ���� - �����ĺ���ȣ
                            wbIdNo = detailSubNodeItem.getElementsByTagName("ns1:wbIdNo")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:outwQty")(0) Is Nothing)) Then
							'// ��ǰ���� - ����
                            outwQty = detailSubNodeItem.getElementsByTagName("ns1:outwQty")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:realslAmt")(0) Is Nothing)) Then
							'// ��ǰ���� - �ǸŰ�
                            realslAmt = detailSubNodeItem.getElementsByTagName("ns1:realslAmt")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:outwAmt")(0) Is Nothing)) Then
							'// ��ǰ���� - ��������
                            outwAmt = detailSubNodeItem.getElementsByTagName("ns1:outwAmt")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:delivInfo")(0) Is Nothing)) Then
							'// ��Ÿ���� - ���
                            delivInfo = detailSubNodeItem.getElementsByTagName("ns1:delivInfo")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:promGiftSpec")(0) Is Nothing)) Then
							'// ��Ÿ���� - ����ǰ����
                            promGiftSpec = detailSubNodeItem.getElementsByTagName("ns1:promGiftSpec")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:cnclRsn")(0) Is Nothing)) Then
							'// ��Ÿ���� - ��ȯ/��һ���
                            cnclRsn = detailSubNodeItem.getElementsByTagName("ns1:cnclRsn")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:cnclRsnSpec")(0) Is Nothing)) Then
							'// ��Ÿ���� - ��ȯ/��һ�����
                            cnclRsnSpec = detailSubNodeItem.getElementsByTagName("ns1:cnclRsnSpec")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:ordDtm")(0) Is Nothing)) Then
							'// �ֹ�����-�ֹ��Ͻ�
                            ordDtm = detailSubNodeItem.getElementsByTagName("ns1:ordDtm")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:juminNum")(0) Is Nothing)) Then
							'// �ֹ�����-�ֹι�ȣ
                            juminNum = detailSubNodeItem.getElementsByTagName("ns1:juminNum")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:dccouponCjhs")(0) Is Nothing)) Then
							'// �ֹ����� - ����(���δ�)�ݾ�
                            dccouponCjhs = detailSubNodeItem.getElementsByTagName("ns1:dccouponCjhs")(0).Text
                        end if

                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:dccouponVendor")(0) Is Nothing)) Then
							'// �ֹ����� - ����(���»�δ�)�ݾ�
                            dccouponVendor = detailSubNodeItem.getElementsByTagName("ns1:dccouponVendor")(0).Text
                        end if


						'' '// �����Ͻ�
						'' '// 2013-05-28+09:00 => 2013-05-28 09:00:00
						wbCrtDt = Replace(wbCrtDt, "+", " ")
						wbCrtDt = wbCrtDt + ":00"


						retVal = false

						ierrStr = ""
						orderCsGbn = ""
						if (ordDtlClsCd = "10") and (toutYn = "N") then
							'// �����ֹ� ��ŵ
							ierrCode = -1
						elseif (ordDtlClsCd = "10") and (toutYn <> "N") then
							'// ������ �ֹ�
							orderCsGbn = "A009"
							retVal = True
							cnclRsnSpec = "�������ֹ�"
						else
							if (ordDtlClsCd = "20") then
								'// �ֹ����
								orderCsGbn = "A008"
								retVal = True

								if (cnclRsnSpec = "") then
									cnclRsnSpec = "�ܼ�����"
								end if
							elseif (ordDtlClsCd = "40") then
								'// ��ȯ���(���ϻ�ǰ)
								orderCsGbn = "A000"
								retVal = True

								if (cnclRsnSpec = "") then
									cnclRsnSpec = "�ܼ�����"
								end if
							else
								'// ����
								ierrStr = "<font color=red>���� : ó�� �ʵ� CS�����Դϴ�.(" + CStr(ordDtlClsCd) + ")</font>"
								ierrCode = 1
							end if

						end if

						if (orderCsGbn<>"") then
							SellSite			= "cjmall"
							OutMallOrderSerial	= ordNo
							OrgDetailKey		= CStr(ordNo) & "-" & CStr(ordGSeq) & "-" & CStr(ordDSeq) & "-" & "001"		'// ���ֹ� ������Ű
							CSDetailKey			= CStr(ordNo) & "-" & CStr(ordGSeq) & "-" & CStr(ordDSeq) & "-" & CStr(ordWSeq)

							'' 	''response.write ordGSeq & "<br>"
							retVal = True
							ierrCode = 0

							strSql = " select idx from db_temp.dbo.tbl_xSite_TMPCS where SellSite = '" + CStr(SellSite) + "' and OutMallOrderSerial = '" + CStr(OutMallOrderSerial) + "' and OrgDetailKey = '" + CStr(OrgDetailKey) + "' and CSDetailKey = '" + CStr(CSDetailKey) + "' "
							rsget.Open strSql,dbget,1
							if (Not rsget.Eof) then
								ierrCode = -1
								retVal = false
							end if
							rsget.Close

							if (ierrCode <> -1) then
								strSql = " if not exists (select idx from db_temp.dbo.tbl_xSite_TMPCS where SellSite = '" + CStr(SellSite) + "' and OutMallOrderSerial = '" + CStr(OutMallOrderSerial) + "' and OrgDetailKey = '" + CStr(OrgDetailKey) + "' and CSDetailKey = '" + CStr(CSDetailKey) + "') "
								strSql = strSql + " begin "
								strSql = strSql + " insert into db_temp.dbo.tbl_xSite_TMPCS(divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
								strSql = strSql + " , OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) "
								strSql = strSql + " values('" + CStr(orderCsGbn) + "', '" + html2db(CStr(cnclRsnSpec)) + "', '" + html2db(CStr(SellSite)) + "', '" + html2db(CStr(OutMallOrderSerial)) + "', '', '', '', '', '" + html2db(CStr(receverNm)) + "', '" + html2db(CStr(telno)) + "', '" + html2db(CStr(cellno)) + "', '" + html2db(CStr(zipno)) + "', '" + html2db(CStr(addr_1) & " " & CStr(addr_2)) + "', '', '" + html2db(CStr(wbIdNo)) + "' "
								strSql = strSql + " , '" + html2db(CStr(wbCrtDt)) + "', '" + html2db(CStr(OrgDetailKey)) + "', '" + html2db(CStr(CSDetailKey)) + "', " + CStr(outwQty) + ") "
								strSql = strSql + " end "
								''rw strSql
								rsget.Open strSql, dbget, 1

								if orderCsGbn = "A009" then
									strSql = " update db_temp.dbo.tbl_xSite_TMPCS "
									strSql = strSql + " set orgOutMallOrderSerial = '" + CStr(oldordNo) + "' "
									strSql = strSql + " where SellSite = '" + CStr(SellSite) + "' and OutMallOrderSerial = '" + CStr(OutMallOrderSerial) + "' and OrgDetailKey = '" + CStr(OrgDetailKey) + "' and CSDetailKey = '" + CStr(CSDetailKey) + "' "
									''rw strSql
									rsget.Open strSql, dbget, 1
								end if
							end if
						end if


                        dtlCnt = dtlCnt+1

                        if (retVal) then
                            succCnt = succCnt+1
                        else
							if (ierrCode <> 0) then
								failCnt = failCnt+1
								if (ierrCode=-1) then skipCnt = skipCnt+1
							end if
                        end if

						if (ordDtlClsCd <> "10") then
							rw "["&ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq&"] "&ierrStr & " "
						end if

                    Next
                end if

                Set detailSubNodes = Nothing
            Next
        end if
    end if

    Set Nodes = Nothing
    rw succCnt & "�� �Է�"
    rw failCnt & "�� ����" & "("&skipCnt&" �� skip)"

	if (succCnt > 0) then
		strSql = " update c "
		strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
		strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
		strSql = strSql + " , c.OrderName = o.OrderName "
		strSql = strSql + " from "
		strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
		strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
		strSql = strSql + " on "
		strSql = strSql + " 	1 = 1 "
		strSql = strSql + " 	and c.SellSite = o.SellSite "
		strSql = strSql + " 	and c.OutMallOrderSerial = Replace(o.OutMallOrderSerial, '-', '') "
		strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
		strSql = strSql + " where "
		strSql = strSql + " 	1 = 1 "
		strSql = strSql + " 	and c.orderserial is NULL "
		strSql = strSql + " 	and o.orderserial is not NULL "
		strSql = strSql + " 	and c.sellsite = 'cjmall' "
		''rw strSql
		rsget.Open strSql, dbget, 1
	end if

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
		,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
	)

    if (matchItemOption<>"") and (matchItemID<>"-1") and (matchItemID<>"") then
        sqlStr = "db_temp.dbo.sp_TEN_xSite_TmpOrder_Insert"
        retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

        ierrCode = GetValue(retParamInfo, "@RETURN_VALUE") ' �����ڵ�
        ierrStr  = GetValue(retParamInfo, "@retErrStr")   ' �����޼���
    else
        ierrCode = -999
        ierrStr = "��ǰ�ڵ� �Ǵ� �ɼ��ڵ�  ��Ī ����" & OrgDetailKey & " ��ǰ�ڵ� =" & matchItemID&" �ɼǸ� = "&partnerOptionName
        rw "["&ierrCode&"]"&retErrStr
        dbget.close() : response.end
    end if

    saveORDOneTmp = (ierrCode=0)
end function

'����Ʈ ȣ���
Function saveListResult(retDoc, mode)
	Dim errorMsg, strSql
	Dim Nodes, SubNodes
	Dim XitemCd, Xstatus, XslCls, XHapvpn, Xvpn, XunitCd, Xitemcode
	Dim OverLapNo
	Dim SelOK
	SelOK = 0
	Set Nodes = retDoc.getElementsByTagName("ns1:unit")

	If (Not (retDoc is Nothing)) Then
		IF application("Svr_Info")="Dev" THEN
			On Error Resume Next
		End If

		For each SubNodes in Nodes
			XitemCd = SubNodes.getElementsByTagName("ns1:itemCd")(0).Text	'�Ǹ��ڵ�
			Xstatus = SubNodes.getElementsByTagName("ns1:status")(0).Text	'�������
			XslCls 	= SubNodes.getElementsByTagName("ns1:slCls")(0).Text	'�Ǹű���
			XHapvpn	= SubNodes.getElementsByTagName("ns1:vpn")(0).Text		'��ü��ǰ�ڵ�
			XunitCd = SubNodes.getElementsByTagName("ns1:unitCd")(0).Text	'��ǰ�ڵ�
			Xvpn 	= Split(XHapvpn, "_")(0)
			Xitemcode = Split(XHapvpn, "_")(1)

			If OverLapNo <> Xvpn Then
				strSql = ""
				strSql = strSql & " UPDATE R " & VBCRLF
				strSql = strSql & " SET cjmallregdate = getdate()" & VBCRLF
				strSql = strSql & " , cjmallPrdNo = '"&XitemCd&"'" & VBCRLF
				If  Xstatus = "A" AND XslCls = "A" Then	'���οϷ� �̰� �Ǹ����� �� (Xstatus A:���οϷ�, XslCls A:����, I:�Ͻ��ߴ�)
					strSql = strSql & " , cjmallStatCd = 7" & VBCRLF
				End If
				strSql = strSql & " , cjmallLastUpdate = getdate()" & VBCRLF
				strSql = strSql & " FROM db_item.dbo.tbl_cjmall_regitem as R " & VBCRLF
				strSql = strSql & " INNER JOIN db_item.dbo.tbl_item as i on R.itemid=i.itemid " & VBCRLF
				strSql = strSql & " WHERE i.itemid = '"&Xvpn&"' "
				dbget.Execute strSql
				SelOK = SelOK + 1
			End If

			'If Xitemcode <> "" AND Xitemcode <> "0000" Then
			If Xitemcode <> "" Then
				strSql = ""
				strSql = strSql & " UPDATE db_item.dbo.tbl_OutMall_regedoption SET " & VBCRLF
				strSql = strSql & " outmallOptCode = '"&XunitCd&"' " & VBCRLF
				strSql = strSql & " , checkdate = getdate() " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&Xvpn&"' and itemoption = '"&Xitemcode&"' " & VBCRLF
				dbget.Execute strSql
				SelOK = SelOK + 1
			End If
			OverLapNo = Xvpn
		Next

	If SelOK = 0 Then
		rw "�˻� ��� ����"
	End If

	End If

	Set Nodes = Nothing
End Function

'�׿� ��
Function saveCommonItemResult(retDoc, mode, prdno)
	Dim errorMsg
	Dim sqlStr
	Dim AssignedRow, successYn
	Dim Titemid, Tlimitno, Tlimitsold, Titemoption, Toptionname, Toptlimitno, Toptlimitsold, Toptsellyn, Toptlimityn, Toptaddprice, Tlimityn, Tsellyn, Titemsu
	successYn = false

	If (Not (retDoc is Nothing)) Then
		errorMsg = retDoc.getElementsByTagName("ns1:errorMsg").item(0).text
		If Left(errorMsg,4) = "[����]" Then
			successYn= true
		End If
	End If

	If (successYn = true) Then
	'�����̰� mode=REG��
		If mode = "REG" Then
	'reged�ɼ� ���̺� ������ �ű�
			sqlStr = ""
			sqlStr = sqlStr & " SELECT i.itemid, i.limitno ,i.limitsold, o.itemoption, o.optionname, o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, o.optaddprice, i.sellyn, i.limityn " & VBCRLF
			sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i " & VBCRLF
			sqlStr = sqlStr & " LEFT JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' " & VBCRLF
			sqlStr = sqlStr & " WHERE i.itemid = "&prdno&" " & VBCRLF
			sqlStr = sqlStr & " ORDER BY o.itemoption ASC "
			rsget.Open sqlStr,dbget,1
			If Not(rsget.EOF or rsget.BOF) Then
				If rsget.RecordCount = 1 AND IsNull(rsget("itemoption")) Then
					Titemid			= rsget("itemid")
					Titemoption 	= "0000"
					Toptionname		= "���ϻ�ǰ"
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
					sqlStr = sqlStr & " ('"&Titemid&"',  '"&Titemoption&"', 'cjmall', '', '"&Toptionname&"', '"&rsget("sellyn")&"', '"&rsget("limityn")&"', '"&Titemsu&"', '0', getdate()) "
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
						sqlStr = sqlStr & " ('"&Titemid&"',  '"&Titemoption&"', 'cjmall', '', '"&Toptionname&"', '"&Toptsellyn&"', '"&Toptlimityn&"', '"&Titemsu&"', '"&Toptaddprice&"', getdate()) "
						dbget.Execute sqlStr
						rsget.MoveNext
					Next
				End If
			End If
			rsget.Close
		ElseIf mode = "MDT" Then
			sqlStr = ""
			sqlStr = sqlStr & " SELECT i.itemid, i.limitno ,i.limitsold, o.itemoption, o.optionname, o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, o.optaddprice, R.outmallOptCode " & VBCRLF
			sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i " & VBCRLF
			sqlStr = sqlStr & " LEFT JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' " & VBCRLF
			sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_OutMall_regedoption as R on i.itemid = R.itemid and R.itemoption = o.itemoption " & VBCRLF
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
						If Left(Titemoption, 1) = "Z" Then
							Toptionname = Replace(Toptionname, ",", "/")
						End If
						sqlStr = ""
						sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
						sqlStr = sqlStr & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
						sqlStr = sqlStr & " VALUES " & VBCRLF
						sqlStr = sqlStr & " ('"&Titemid&"',  '"&Titemoption&"', 'cjmall', '', '"&Toptionname&"', '"&Toptsellyn&"', '"&Toptlimityn&"', '"&Toptlimitno&"', '"&Toptaddprice&"', getdate()) "
						dbget.Execute sqlStr
					End If
					rsget.MoveNext
				Next
			End If
			rsget.Close

		End If

		If (prdno <> "") Then
			sqlStr = ""
			sqlStr = sqlStr & " UPDATE R" & VBCRLF
			sqlStr = sqlStr & " SET cjmallLastUpdate = getdate()" & VBCRLF

			If (mode = "REG") Then
				sqlStr = sqlStr & " ,cjmallStatCd=(CASE WHEN isNULL(cjmallStatCd, -1) < 3 then 3 ELSE cjmallStatCd END)"        ''�ӽõ�ϿϷ�(��� �� ���δ��)
				sqlStr = sqlStr & " ,cjmallRegdate=isNULL(cjmallRegdate,getdate())" & VbCrlf
			End If

			If (mode = "PRI") or (mode = "REG") Then
				sqlStr = sqlStr & " ,cjmallPrice = i.sellcash" & VBCRLF
			End If

			If (mode = "SLD") Then
				sqlStr = sqlStr & " ,cjmallSellYn = 'N'" & VBCRLF
			Else
				If (mode = "MDT") or (mode = "REG") Then
					sqlStr = sqlStr & " ,cjmallSellYn = i.sellyn" & VBCRLF
				End If
			End If

			sqlStr = sqlStr & " ,accFailCNT=0" & VBCRLF                 ''����ȸ�� �ʱ�ȭ
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
			sqlStr = sqlStr & " SET cjmallStatCd = -1"                   '''��Ͻ���
			sqlStr = sqlStr & " FROM db_item.dbo.tbl_cjmall_regitem as R" & VBCRLF
			sqlStr = sqlStr & " WHERE R.itemid = "&prdno&"" & VBCRLF
			sqlStr = sqlStr & " and cjmallStatCd = 1"                    ''����
			dbget.Execute sqlStr
		End If
	End If

	If (isCJ_DebugMode) Then
		rw prdno &"_"&errorMsg
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
		poster.setTimeouts 5000,90000,90000,90000  ''2013/07/25 �߰�
		poster.send SendDoc

	Set retDoc = server.createobject("MSXML2.DomDocument.3.0")
		retDoc.async = False
		retDoc.LoadXML(poster.responseTEXT)

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
	Dim fso,tFile
	Dim opath
	Select Case mode
		Case "REG", "RET_REG"
			opath = "/admin/etc/cjmall/xmlFiles/INSERT/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
		Case "LIST", "DayLIST", "RET_LIST", "RET_DayLIST", "commonCD", "RET_commonCD", "RET_SONGJANG"
			opath = "/admin/etc/cjmall/xmlFiles/SELECT/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
		Case "EDT", "RET_EDT", "MDT", "RET_MDT", "PRI", "RET_PRI", "QTY", "RET_QTY"
			opath = "/admin/etc/cjmall/xmlFiles/UPDATE/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
		Case "SLD", "RET_SLD"
			opath = "/admin/etc/cjmall/xmlFiles/UPDATE_SellStatus/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
	    Case "RET_ORDLIST", "RET_ORDCANCELLIST", "RET_CSLIST", "RET_CSORDLIST", "RET_CSORDCANCELLIST"
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

function getLastCSInputDT(mode)
    dim sqlStr
    sqlStr = "select top 1 convert(varchar(10),LastCheckDate,21) as lastCSInputDt"
    sqlStr = sqlStr&" from db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "

	sqlStr = sqlStr&" where sellsite = 'cjmall' and csGubun = '" + CStr(mode) + "' "

    rsget.Open sqlStr,dbget,1
	if (Not rsget.Eof) then
		getLastCSInputDT = rsget("lastCSInputDt")
	else
		getLastCSInputDT = "2013-05-30"
	end if
	rsget.Close

end function

function UpdateLastCSInputDT(mode, dt)
    dim sqlStr

	sqlStr = " update db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "
	sqlStr = sqlStr & " set LastCheckDate = '" + CStr(dt) + "' "
	sqlStr = sqlStr & " where sellsite = 'cjmall' and csGubun = '" + CStr(mode) + "' "
	dbget.Execute sqlStr

end function

%>
