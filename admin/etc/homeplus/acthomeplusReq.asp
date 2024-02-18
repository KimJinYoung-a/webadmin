<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 600 %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/homeplus/homepluscls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'#################################### 홈플러스 기본 정보 Setting ####################################
Public homeplusAPIURL
Public strInterface
Public homeplusVenderID
Public homepluspasswd

IF application("Svr_Info") = "Dev" THEN
	homeplusAPIURL = "http://112.108.7.201:7006/services/API2?wsdl"
	strInterface = "http://112.108.7.201:7006/api/services/API2"
	homeplusVenderID = "292811"
	homepluspasswd = "qwer1234"
Else
	homeplusAPIURL = "http://api.direct.homeplus.co.kr:17004/services/API2?wsdl"
	strInterface = "http://api.direct.homeplus.co.kr:17004/api/services/API2"
	homeplusVenderID = "292811"
	homepluspasswd = "cube1010!!"
End if
'#####################################################################################################

'################################### 각종 Function Setting  ##########################################
Function HomeplusOneItemReg(iitemid, strParam, byRef iErrStr, iSellCash, ihomeplusSellYn, ilimityn, ilimitno, ilimitsold, iitemname, mode)
	Dim xmlStr : xmlStr = strParam
	Dim objXML, xmlDOM, SubNodes, strSql
	Dim retCode, homegoodNo, iMessage, optlist, s_ITEMNO, i_ITEMNO, s_OPTION_NAME
	Dim AssignedRow
	Dim Tlimitno, Tlimitsold, Tlimityn, Titemsu
	If (xmlStr = "") Then
		HomeplusOneItemReg = false
		Exit Function
    End If

	Set objXML = Server.CreateObject("Msxml2.XMLHTTP.3.0")
		objXML.open "post", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#" &mode
		objXML.send(xmlStr)

	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
'				response.write objXML.ResponseText
'				response.end
		If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
'					response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		End If

		retCode		= xmlDOM.selectSingleNode("soapenv:Envelope/soapenv:Body/createNewProductResponse/ns1:createNewProductReturn/ns1:code").text
		homegoodNo	= xmlDOM.selectSingleNode("soapenv:Envelope/soapenv:Body/createNewProductResponse/ns1:createNewProductReturn/ns1:i_STYLENO").text
		iMessage	= xmlDOM.selectSingleNode("soapenv:Envelope/soapenv:Body/createNewProductResponse/ns1:createNewProductReturn/ns1:message").text

		If retCode = "E0000" Then	'성공(E0000)
			'상품존재여부 확인
			strSql = "SELECT COUNT(itemid) FROM db_outmall.dbo.tbl_homeplus_regItem WHERE itemid='" & iitemid & "'"
			rsCTget.Open strSql, dbCTget, 1
			If rsCTget(0) > 0 Then
				'// 존재 -> 수정
				strSql = ""
				strSql = strSql & " UPDATE R" & VbCRLF
				strSql = strSql & "	Set homeplusLastUpdate = getdate() "  & VbCRLF
				strSql = strSql & "	, homeplusGoodNo = '" & homegoodNo & "'"  & VbCRLF
				strSql = strSql & "	, homeplusPrice = " &iSellCash& VbCRLF
				strSql = strSql & "	, accFailCnt = 0"& VbCRLF
				strSql = strSql & "	, homeplusRegdate = isNULL(homeplusRegdate, getdate())"
				If (homegoodNo <> "") Then
				    strSql = strSql & "	, homeplusstatCD = '7'"& VbCRLF					'등록완료(임시)
				Else
					strSql = strSql & "	, homeplusstatCD = '1'"& VbCRLF					'전송시도
				End If
				strSql = strSql & "	From db_outmall.dbo.tbl_homeplus_regItem R"& VbCRLF
				strSql = strSql & " Where R.itemid = '" & iitemid & "'"
				dbCTget.Execute(strSql)
			Else
				'// 없음 -> 신규등록
				strSql = ""
				strSql = strSql & " INSERT INTO db_outmall.dbo.tbl_homeplus_regItem "
				strSql = strSql & " (itemid, regitemname, reguserid, homeplusRegdate, homeplusLastUpdate, homeplusGoodNo, homeplusPrice, homeplusSellYn, homeplusStatCd) VALUES " & VbCRLF
				strSql = strSql & " ('" & iitemid & "'" & VBCRLF
				strSql = strSql & " , '" & iitemname & "'" &_
				strSql = strSql & " , '" & session("ssBctId") & "'" &_
				strSql = strSql & " , getdate(), getdate()" & VBCRLF
				strSql = strSql & " , '" & homegoodNo & "'" & VBCRLF
				strSql = strSql & " , '" & iSellCash & "'" & VBCRLF
				strSql = strSql & " , '" & ihomeplusSellYn & "'" & VBCRLF
				If (homegoodNo <> "") Then
				    strSql = strSql & ",'7'"											'등록완료(임시)
				Else
				    strSql = strSql & ",'1'"											'전송시도
				End If
				strSql = strSql & ")"
				dbCTget.Execute(strSql)
				actCnt = actCnt + 1
			End If
			rsCTget.Close

			Set optlist = xmlDOM.getElementsByTagName("ns1:ITEMRESULT")
				For each SubNodes in optlist
					s_ITEMNO		= Trim(SubNodes.getElementsByTagName("ns1:s_ITEMNO").item(0).text)		'텐바이텐 옵션코드
					i_ITEMNO		= Trim(SubNodes.getElementsByTagName("ns1:i_ITEMNO").item(0).text)		'홈플러스 옵션코드
					s_OPTION_NAME	= Trim(SubNodes.getElementsByTagName("ns1:s_OPTION_NAME").item(0).text)	'옵션명
					If s_ITEMNO = "0000" Then
						Tlimitno		= ilimitno
						Tlimitsold		= ilimitsold
						Tlimityn		= ilimityn
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
						sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_OutMall_regedoption " & VBCRLF
						sqlStr = sqlStr & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
						sqlStr = sqlStr & " VALUES " & VBCRLF
						sqlStr = sqlStr & " ('"&iitemid&"',  '"&s_ITEMNO&"', 'homeplus', '"&i_ITEMNO&"', '"&html2db(s_OPTION_NAME)&"', 'Y', '"&ilimityn&"', '"&Titemsu&"', '0', getdate()) "
						dbCTget.Execute sqlStr
					Else
						sqlStr = ""
						sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_OutMall_regedoption " & VBCRLF
						sqlStr = sqlStr & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
						sqlStr = sqlStr & " SELECT itemid, itemoption, 'homeplus', '"&i_ITEMNO&"', optionname, optsellyn, 'Y', " & VBCRLF
						sqlStr = sqlStr & " Case WHEN optlimityn = 'Y' AND optlimitno - optlimitsold <= 5 THEN '0' " & VBCRLF
						sqlStr = sqlStr & " 	 WHEN optlimityn = 'Y' AND optlimitno - optlimitsold > 5 THEN optlimitno - optlimitsold - 5 " & VBCRLF
						sqlStr = sqlStr & " 	 WHEN optlimityn = 'N' THEN '999' End " & VBCRLF
						sqlStr = sqlStr & " , '0', getdate() " & VBCRLF
						sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_item_option " & VBCRLF
						sqlStr = sqlStr & " WHERE itemid= '"&iitemid&"' " & VBCRLF
						sqlStr = sqlStr & " and itemoption = '"& s_ITEMNO &"' "
						dbCTget.Execute sqlStr
					End If
				Next
			Set optlist = nothing
			strSql = ""
			strSql = strSql & " UPDATE R " & VBCRLF
			strSql = strSql & " SET regedOptCnt = isNULL(T.regedOptCnt,0) " & VBCRLF
			strSql = strSql & " FROM db_outmall.dbo.tbl_homeplus_regItem R " & VBCRLF
			strSql = strSql & " Join ( " & VBCRLF
			strSql = strSql & " 	SELECT R.itemid, count(*) as CNT, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt "
			strSql = strSql & " 	FROM db_outmall.dbo.tbl_homeplus_regItem R " & VBCRLF
			strSql = strSql & " 	JOIN db_outmall.dbo.tbl_OutMall_regedoption Ro on R.itemid = Ro.itemid and Ro.mallid = 'homeplus' and Ro.itemid = " &iitemid & VBCRLF
			strSql = strSql & " 	GROUP BY R.itemid " & VBCRLF
			strSql = strSql & " ) T on R.itemid = T.itemid " & VBCRLF
			dbCTget.Execute strSql
			HomeplusOneItemReg = true
			Set objXML = Nothing
			Set xmlDOM = Nothing
			rw "[" & iitemid & "]:"&iMessage
		Else						'실패(E)
		    iErrStr =  "상품 등록중 오류 [" & iitemid & "]:"&iMessage
			Set objXML = Nothing
			Set xmlDOM = Nothing
		    Exit Function
		End If
		On Error Goto 0
	End If
End Function

Function HomeplusOneItemEdit(iitemid, iHomeplusGoodNo, byRef iErrStr, strParam, mode)
	Dim xmlStr : xmlStr = strParam
	Dim objXML, xmlDOM, retCode, iMessage
	Dim strRst, strSql

	If (xmlStr = "") Then
		HomeplusOneItemEdit = false
		Exit Function
    End If

	Set objXML = Server.CreateObject("Msxml2.XMLHTTP.3.0")
		objXML.open "post", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#" &mode
		objXML.send(xmlStr)

	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
'			response.write objXML.ResponseText
'			response.end
		retCode		= xmlDOM.selectSingleNode("soapenv:Envelope/soapenv:Body/updateProductResponse/ns1:updateProductReturn/ns1:code").text
		iMessage	= xmlDOM.selectSingleNode("soapenv:Envelope/soapenv:Body/updateProductResponse/ns1:updateProductReturn/ns1:message").text

		If retCode = "E0000" Then	'성공(E0000)이면 실패횟수 초기화
			strSql = ""
			strSql = strSql & " UPDATE db_outmall.dbo.tbl_homeplus_regItem " & VbCRLF
			strSql = strSql & " SET accFailCnt = 0 " & VbCRLF
			strSql = strSql & " WHERE itemid='" & iitemid & "'"
			dbCTget.Execute(strSql)

			HomeplusOneItemEdit = true
			Set objXML = Nothing
			Set xmlDOM = Nothing
			rw "[" & iitemid & "]:"&iMessage
		Else						'실패(E)
			iErrStr =  "정보 수정 중 오류 [" & iitemid & "]:"&iMessage
			Set objXML = Nothing
			Set xmlDOM = Nothing
		    Exit Function
		End If
	Else
	    iErrStr =  "[" & iitemid & "]:상품 수정 중 통신 오류 "
		Set objXML = Nothing
		Set xmlDOM = Nothing
	    Exit Function
	End If
End Function

Function HomeplusOneItemOPTEdit(iitemid, iHomeplusGoodNo, byRef iErrStr, strParam, mode)
	Dim xmlStr : xmlStr = strParam
	Dim objXML, xmlDOM, retCode, iMessage
	Dim strRst, strSql

	If (xmlStr = "") Then
		HomeplusOneItemOPTEdit = false
		Exit Function
    End If

	Set objXML = Server.CreateObject("Msxml2.XMLHTTP.3.0")
		objXML.open "post", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#" &mode
		objXML.send(xmlStr)

	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
'			response.write objXML.ResponseText
'			response.end
		retCode		= xmlDOM.selectSingleNode("soapenv:Envelope/soapenv:Body/updateProductItemResponse/ns1:updateProductItemReturn/ns1:code").text
		iMessage	= xmlDOM.selectSingleNode("soapenv:Envelope/soapenv:Body/updateProductItemResponse/ns1:updateProductItemReturn/ns1:message").text

		If retCode = "E0000" Then	'성공(E0000)이면 실패횟수 초기화
			strSql = ""
			strSql = strSql & " UPDATE db_outmall.dbo.tbl_homeplus_regItem " & VbCRLF
			strSql = strSql & " SET accFailCnt = 0 " & VbCRLF
			strSql = strSql & " , homeplusLastUpdate = getdate() " & VbCRLF
			strSql = strSql & " WHERE itemid='" & iitemid & "'"
			dbCTget.Execute(strSql)

			HomeplusOneItemOPTEdit = true
			Set objXML = Nothing
			Set xmlDOM = Nothing
			rw "[" & iitemid & "]:"&iMessage
		Else						'실패(E)
			iErrStr =  "아이템 정보 외 수정 중 오류 [" & iitemid & "]:"&iMessage
			Set objXML = Nothing
			Set xmlDOM = Nothing
		    Exit Function
		End If
	Else
	    iErrStr =  "[" & iitemid & "]:아이템 정보 외 수정 중 통신 오류 "
		Set objXML = Nothing
		Set xmlDOM = Nothing
	    Exit Function
	End If
End Function

Function HomeplusOneItemIMGEdit(iitemid, iHomeplusGoodNo, byRef iErrStr, strParam, mode)
	Dim xmlStr : xmlStr = strParam
	Dim objXML, xmlDOM, retCode, iMessage
	Dim strRst, strSql

	If (xmlStr = "") Then
		HomeplusOneItemIMGEdit = false
		Exit Function
    End If

	Set objXML = Server.CreateObject("Msxml2.XMLHTTP.3.0")
		objXML.open "post", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#" &mode
		objXML.send(xmlStr)

	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
'			response.write objXML.ResponseText
'			response.end
		retCode		= xmlDOM.selectSingleNode("soapenv:Envelope/soapenv:Body/updateImageResponse/ns1:updateImageReturn/ns1:code").text
		iMessage	= xmlDOM.selectSingleNode("soapenv:Envelope/soapenv:Body/updateImageResponse/ns1:updateImageReturn/ns1:message").text

		If retCode = "E0000" Then	'성공(E0000)
			HomeplusOneItemIMGEdit = true
			Set objXML = Nothing
			Set xmlDOM = Nothing
			rw "[" & iitemid & "]:"&iMessage
		Else						'실패(E)
			iErrStr =  "이미지 수정 중 오류 [" & iitemid & "]:"&iMessage
			Set objXML = Nothing
			Set xmlDOM = Nothing
		    Exit Function
		End If
	Else
	    iErrStr =  "[" & iitemid & "]:상품 이미지 수정 중 통신 오류 "
		Set objXML = Nothing
		Set xmlDOM = Nothing
	    Exit Function
	End If
End Function

Function HomeplusOneItemView(iitemid, iHomeplusGoodNo, byRef iErrStr, strParam, mode)
	Dim xmlStr : xmlStr = strParam
	Dim objXML, xmlDOM, retCode, iMessage, regedItemStatus, actCnt
	Dim strSql, oneProdInfo, SubNodes, AssignedRow, StockQty
	Dim hplOptStatus, hplOptno, regedOpt10x10OptNo, regedOpt10x10OptNm

	If (xmlStr = "") Then
		HomeplusOneItemView = false
		Exit Function
    End If

	Set objXML = Server.CreateObject("Msxml2.XMLHTTP.3.0")
		objXML.open "post", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#" &mode
		objXML.send(xmlStr)

	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
'			response.write objXML.ResponseText
'			response.end
		retCode			= xmlDOM.selectSingleNode("soapenv:Envelope/soapenv:Body/searchProductResponse/ns1:searchProductReturn/ns1:code").text
		iMessage		= xmlDOM.selectSingleNode("soapenv:Envelope/soapenv:Body/searchProductResponse/ns1:searchProductReturn/ns1:message").text
		regedItemStatus	= xmlDOM.selectSingleNode("soapenv:Envelope/soapenv:Body/searchProductResponse/ns1:searchProductReturn/ns1:SALE").text

		If retCode = "E0000" Then	'성공(E0000)
			Set oneProdInfo = XMLDom.SelectNodes("soapenv:Envelope/soapenv:Body/searchProductResponse/ns1:searchProductReturn/ns1:ITEMRESULT/ns1:ITEMRESULT")
				For Each SubNodes In oneProdInfo
					hplOptStatus			= SubNodes.SelectSingleNode("ns1:SALE").text
					hplOptno				= SubNodes.SelectSingleNode("ns1:i_ITEMNO").text
					regedOpt10x10OptNo		= SubNodes.SelectSingleNode("ns1:s_ITEMNO").text
					regedOpt10x10OptNm		= SubNodes.SelectSingleNode("ns1:s_OPTION_NAME").text

					If oHomeplus.FItemList(i).FoptionCnt = 0 Then		'단품이라면
						StockQty = oHomeplus.FItemList(i).GetHomeplusLmtQty
					Else												'옵션이라면
						If oHomeplus.FItemList(i).FLimityn = "Y" Then
							strSql = ""
							strSql = strSql & " SELECT CASE WHEN (optlimitno - optlimitsold) <= 5 Then '0' Else (optlimitno - optlimitsold - 5) End as StockQty "
							strSql = strSql & " FROM db_AppWish.dbo.tbl_item_option  "
							strSql = strSql & " WHERE itemid='"&iitemid&"' and itemoption = '"&regedOpt10x10OptNo&"' "
					        rsCTget.Open strSql, dbCTget
							If Not(rsCTget.EOF or rsCTget.BOF) Then
								StockQty = rsCTget("StockQty")
							Else
								StockQty = 0
							End If
							rsCTget.Close
						Else
							StockQty = 999
						End If
					End If
					'1.없으면 인서트 있으면 업데이트
					strSql = ""
					strSql = strSql & " IF Exists(SELECT * FROM db_outmall.dbo.tbl_OutMall_regedoption where itemid='"&iitemid&"' and itemoption = '"&regedOpt10x10OptNo&"' and outmallOptCode = '"&hplOptno&"' and mallid = 'homeplus') "
					strSql = strSql & " BEGIN"& VbCRLF
					strSql = strSql & " UPDATE oP "
				    strSql = strSql & " SET outmallOptName='"&html2DB(regedOpt10x10OptNm)&"'"&VbCRLF
					strSql = strSql & " ,outmallOptCode='"&hplOptno&"'"&VbCRLF
				    strSql = strSql & " ,lastupdate=getdate()"&VbCRLF
				    strSql = strSql & " ,outMallSellyn='"&Chkiif(hplOptStatus="true", "Y", "N")&"'"&VbCRLF
				    strSql = strSql & " ,outmalllimityn='Y'"&VbCRLF
				    strSql = strSql & " ,outMallLimitNo="&StockQty&VbCRLF
				    strSql = strSql & " ,checkdate=getdate()"&VbCRLF
				    strSql = strSql & " FROM db_outmall.dbo.tbl_OutMall_regedoption oP"&VbCRLF
				    strSql = strSql & " WHERE itemid="&iitemid&VbCRLF
				    strSql = strSql & " and convert(int, outmallOptCode)='"&hplOptno&"'"&VbCRLF				'개편전 outmallOptCode는 001,002,003 이렇게 들어가있으나 개편 후엔 1,2,3이렇게 변함
				    strSql = strSql & " and mallid='homeplus'"&VbCRLF
					strSql = strSql & " END ELSE "
					strSql = strSql & " BEGIN"& VbCRLF
					strSql = strSql & " INSERT INTO db_outmall.dbo.tbl_OutMall_regedoption "
			        strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastupdate)"
			        strSql = strSql & " VALUES ('"&iitemid&"', '"&regedOpt10x10OptNo&"', 'homeplus', '"&hplOptno&"', '"&html2DB(regedOpt10x10OptNm)&"', '"&Chkiif(hplOptStatus="true", "Y", "N")&"', 'Y', '"&StockQty&"', '', getdate())"
					strSql = strSql & " END "
				    dbCTget.Execute strSql, AssignedRow
					actCnt = actCnt+AssignedRow
'					rw "retCode : " &retCode
'					rw "iMessage : "&iMessage
'					rw "regedItemStatus : "&regedItemStatus
'					rw "hplOptStatus : "&hplOptStatus
'					rw "hplOptno : "&hplOptno
'					rw "regedOpt10x10OptNo : "&regedOpt10x10OptNo
'					rw "regedOpt10x10OptNm : "&regedOpt10x10OptNm
'					rw "----------------------------"
				Next
'				2.regedItemStatus에 따라 상품판매 상태 변경 / regrdOptcnt수 변경
				If (actCnt > 0) Then
					strSql = " update R"   &VbCRLF
					strSql = strSql & " set regedOptCnt=isNULL(T.regedOptCnt,0)"   &VbCRLF
					strSql = strSql & " ,homeplusSellYn = '"&Chkiif(regedItemStatus="true", "Y", "N")&"'"   &VbCRLF
					strSql = strSql & " from db_outmall.dbo.tbl_homeplus_regItem R"   &VbCRLF
					strSql = strSql & " 	Join ("   &VbCRLF
					strSql = strSql & " 		select R.itemid,count(*) as CNT "
					strSql = strSql & " 		, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
					strSql = strSql & "        from db_outmall.dbo.tbl_homeplus_regItem R"   &VbCRLF
					strSql = strSql & " 			Join db_outmall.dbo.tbl_OutMall_regedoption Ro"   &VbCRLF
					strSql = strSql & " 			on R.itemid=Ro.itemid"   &VbCRLF
					strSql = strSql & " 			and Ro.mallid='homeplus'"   &VbCRLF
					strSql = strSql & "             and Ro.itemid="&iitemid&VbCRLF
					strSql = strSql & " 		group by R.itemid"   &VbCRLF
					strSql = strSql & " 	) T on R.itemid=T.itemid"   &VbCRLF
					dbCTget.Execute strSql
				End If
			HomeplusOneItemView = true
			Set objXML = Nothing
			Set xmlDOM = Nothing
			rw "[" & iitemid & "]:"&iMessage

			strSql = ""
			strSql = strSql & " SELECT count(*) as cnt FROM db_outmall.dbo.tbl_OutMall_regedoption where itemid='"&iitemid&"' and outmallSellyn = 'Y' and mallid = 'homeplus' "
			rsCTget.Open strSql, dbCTget
			If rsCTget("cnt") = 0 Then
				strSql = ""
				strSql = strSql & " UPDATE oP "
			    strSql = strSql & " SET homeplusSellYn ='N'"&VbCRLF
			    strSql = strSql & " FROM db_outmall.dbo.tbl_homeplus_regitem oP"&VbCRLF
			    strSql = strSql & " WHERE itemid="&iitemid&VbCRLF
				dbCTget.Execute strSql
				rw "["&iitemid&"]:옵션이 전부 품절이므로 품절로 변경합니다"
			End If
			rsCTget.Close
		Else						'실패(E)
			iErrStr =  "상품 조회 중 오류 [" & iitemid & "]:"&iMessage
			Set objXML = Nothing
			Set xmlDOM = Nothing
		    Exit Function
		End If
	Else
	    iErrStr =  "[" & iitemid & "]:상품 조회 중 통신 오류 "
		Set objXML = Nothing
		Set xmlDOM = Nothing
	    Exit Function
	End If
End Function

Function HomeplusOneItemSellStatEdit(iitemid, iHomeplusGoodNo, ichgSellYn, byRef iErrStr, strParam, mode)
    Dim xmlStr : xmlStr = strParam
    Dim objXML, xmlDOM, retCode, iMessage
    Dim strRst, strSql

	If (xmlStr = "") Then
		HomeplusOneItemSellStatEdit = false
		Exit Function
    End If

	Set objXML = Server.CreateObject("Msxml2.XMLHTTP.3.0")
		objXML.open "post", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#" &mode
		objXML.send(xmlStr)

	If objXML.Status = "200" Then
		Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
		retCode		= xmlDOM.selectSingleNode("soapenv:Envelope/soapenv:Body/setProductStatusResponse/ns1:setProductStatusReturn/ns1:code").text
		iMessage	= xmlDOM.selectSingleNode("soapenv:Envelope/soapenv:Body/setProductStatusResponse/ns1:setProductStatusReturn/ns1:message").text

		If retCode = "E0000" Then	'성공(E0000)
			strSql = ""
			strSql = strSql & " UPDATE db_outmall.dbo.tbl_homeplus_regItem " & VbCRLF
			strSql = strSql & " SET homeplusLastUpdate = getdate() " & VbCRLF
			strSql = strSql & " ,homeplusSellYn = '" & ichgSellYn & "'" & VbCRLF
			strSql = strSql & " WHERE itemid='" & iitemid & "'"
			dbCTget.Execute(strSql)
			HomeplusOneItemSellStatEdit = true
			Set objXML = Nothing
			Set xmlDOM = Nothing
			rw "[" & iitemid & "]:"&iMessage
		Else						'실패(E)
		    iErrStr =  "상품 상태변경 중 오류 [" & iitemid & "]:"&iMessage
			Set objXML = Nothing
			Set xmlDOM = Nothing
		    Exit Function
		End If
		On Error Goto 0
	End If
End Function

Function getXMLString(mode)
	Dim strRst
	If mode = "login" Then
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/""  xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:ns1=""http://xml.apache.org/axis/"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:"&mode&" xmlns:m=""" & strInterface & """>"
		strRst = strRst & "			<venderId>"&homeplusVenderID&"</venderId>"
		strRst = strRst & "			<passwd>"&homepluspasswd&"</passwd>"
		strRst = strRst & "		</m:"&mode&">"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
	ElseIf mode = "getCategories" Then
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/""  xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<SOAP-ENV:Body>"
		strRst = strRst & "		<m:"&mode&" xmlns:m=""" & strInterface & """></m:"&mode&">"
		strRst = strRst & "	</SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
	End If
	getXMLString = strRst
End Function

'로그인 API
Function HomeplusLoginAPI()
    Dim mode : mode = "login"
	Dim xmlStr : xmlStr = getXMLString(mode)
	Dim objXML, xmlDOM
	If (xmlStr = "") Then
		HomeplusLoginAPI = false
		Exit Function
    End If
    On Error Resume Next
	Set objXML = server.CreateObject("WinHttp.WinHttpRequest.5.1")
		objXML.open "POST", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#"&mode
		objXML.setTimeouts 5000,90000,90000,90000
		objXML.send(xmlStr)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.ValidateOnParse= True
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				If Err <> 0 then
					Response.Write "<script language=javascript>alert('홈플러스 로그인에 실패했습니다.');history.back();</script>"
					Response.End
					HomeplusLoginAPI = false
				End If

				If xmlDOM.getElementsByTagName("ns1:code").item(0).text = "E0000" Then
					HomeplusLoginAPI = true
				Else
					Response.Write "<script language=javascript>alert('홈플러스 로그인에 실패했습니다.');history.back();</script>"
					Response.End
				End If
				On Error Goto 0
			Set xmlDOM = nothing
		Else
			Response.Write "<script language=javascript>alert('홈플러스 로그인중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');history.back();</script>"
			Response.End
			HomeplusLoginAPI = false
		End If
	Set objXML = nothing
End Function

'로그인 API(XMLHTTP로는 최초 로그인때 time out이 걸려서 수정)
'ServerXMLHTTP 이걸로는 세션이 유지가 안 되서 두번 호출;;
'더 좋은 방법이 있으면 수정;
Function HomeplusLoginAPI2()
    Dim mode : mode = "login"
	Dim xmlStr : xmlStr = getXMLString(mode)
	Dim objXML, xmlDOM
	Dim confirmLogin
	If (xmlStr = "") Then
		HomeplusLoginAPI2 = false
		Exit Function
    End If

    On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#"&mode
		objXML.setTimeouts 5000,90000,90000,90000
		objXML.send(xmlStr)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.ValidateOnParse= True
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				If xmlDOM.getElementsByTagName("ns1:code").item(0).text = "E0000" Then
					confirmLogin = "Y"
				End If
				On Error Goto 0
			Set xmlDOM = nothing
		End If
	Set objXML = nothing

	If confirmLogin = "Y" Then
		Set objXML = Server.CreateObject("Msxml2.XMLHTTP.3.0")
			objXML.open "POST", "" & homeplusAPIURL & "", False
			objXML.setRequestHeader "CONTENT_TYPE", "text/xml; charset=utf-8"
			objXML.setRequestHeader "Content-Length", Len(xmlStr)
			objXML.setRequestHeader "SOAPAction", strInterface & "#"&mode
			objXML.send(xmlStr)
			If objXML.Status = "200" Then
				Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
					xmlDOM.async = False
					xmlDOM.ValidateOnParse= True
					xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
					If Err <> 0 then
						Response.Write "<script language=javascript>alert('홈플러스 로그인에 실패했습니다.');history.back();</script>"
						Response.End
						HomeplusLoginAPI2 = false
					End If

					If xmlDOM.getElementsByTagName("ns1:code").item(0).text = "E0000" Then
						HomeplusLoginAPI2 = true
					Else
						Response.Write "<script language=javascript>alert('홈플러스 로그인에 실패했습니다.');history.back();</script>"
						Response.End
					End If
					On Error Goto 0
				Set xmlDOM = nothing
			Else
				Response.Write "<script language=javascript>alert('홈플러스 로그인중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');history.back();</script>"
				Response.End
				HomeplusLoginAPI2 = false
			End If
		Set objXML = nothing
	End If
End Function

'카테고리 API
Function HomeplusCategoryAPI()
    Dim mode : mode = "getCategories"
	Dim xmlStr : xmlStr = getXMLString(mode)
	Dim objXML, xmlDOM, retCode, hplist, SubNodes, strSql
	Dim hDIVISION, hGROUP, hDEPT, hCLASS, hSUBCLASS, hDIV_NAME, hGROUP_NAME, hDEPT_NAME, hCLASS_NAME, hSUB_NAME, hCATEGORY_ID, hCATEGORY_NAME
	Dim AssignedRow

	If (xmlStr = "") Then
		HomeplusCategoryAPI = false
		Exit Function
    End If
	Set objXML = Server.CreateObject("Msxml2.XMLHTTP.3.0")
		objXML.open "post", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(xmlStr)
		objXML.setRequestHeader "SOAPAction", strInterface & "#"&mode
		objXML.send(xmlStr)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				On Error Resume Next
					'response.write objXML.ResponseText
					retCode = xmlDOM.getElementsByTagName("ns1:code").item(0).text
					iMessage = xmlDOM.getElementsByTagName("ns1:message").item(0).text
rw retCode
rw iMessage
'					If retCode = "E0000" Then
'						strSql = ""
'						strSql = strSql & " DELETE FROM db_outmall.dbo.tbl_homeplus_dftcategory "
'						dbCTget.Execute(strSql)
'						Set hplist = xmlDOM.getElementsByTagName("ns1:list")
'							For each SubNodes in hplist
'								hDIVISION		= Trim(SubNodes.getElementsByTagName("ns1:DIVISION").item(0).text)		'최상위 분류코드
'								hGROUP			= Trim(SubNodes.getElementsByTagName("ns1:GROUP").item(0).text)			'DIVISION 하위 분류 코드
'								hDEPT			= Trim(SubNodes.getElementsByTagName("ns1:DEPT").item(0).text)			'GROUP 하위 분류 코드
'								hCLASS			= Trim(SubNodes.getElementsByTagName("ns1:CLASS").item(0).text)			'DEPT 하위 분류 코드
'								hSUBCLASS		= Trim(SubNodes.getElementsByTagName("ns1:SUBCLASS").item(0).text)		'CLASS 하위 분류 코드
'								hDIV_NAME		= Trim(SubNodes.getElementsByTagName("ns1:DIV_NAME").item(0).text)		'DIVISION 분류명
'								hGROUP_NAME		= Trim(SubNodes.getElementsByTagName("ns1:GROUP_NAME").item(0).text)	'GROUP 분류명
'								hDEPT_NAME		= Trim(SubNodes.getElementsByTagName("ns1:DEPT_NAME").item(0).text)		'DEPT 분류명
'								hCLASS_NAME		= Trim(SubNodes.getElementsByTagName("ns1:CLASS_NAME").item(0).text)	'CLASS 분류명
'								hSUB_NAME		= Trim(SubNodes.getElementsByTagName("ns1:SUB_NAME").item(0).text)		'SUBCLASS 분류명
'								hCATEGORY_ID	= Trim(SubNodes.getElementsByTagName("ns1:CATEGORY_ID").item(0).text)	'상품정보제공고시를 위한 카테고리 아이디
'								hCATEGORY_NAME	= Trim(SubNodes.getElementsByTagName("ns1:CATEGORY_NAME").item(0).text)	'상품정보제공고시를 위한 카테고리 명
'
'								strSql = ""
'								strSql = strSql & " INSERT INTO db_outmall.dbo.tbl_homeplus_dftcategory (hDIVISION, hGROUP, hDEPT, hCLASS, hSUBCLASS, hDIV_NAME, hGROUP_NAME, hDEPT_NAME, hCLASS_NAME, hSUB_NAME, hCATEGORY_ID, hCATEGORY_NAME) VALUES " & VBCRLF
'								strSql = strSql & " ('"&db2html(hDIVISION)&"', '"&db2html(hGROUP)&"', '"&db2html(hDEPT)&"', '"&db2html(hCLASS)&"', '"&db2html(hSUBCLASS)&"', '"&db2html(hDIV_NAME)&"', '"&hGROUP_NAME&"', '"&db2html(hDEPT_NAME)&"', '"&db2html(hCLASS_NAME)&"', '"&db2html(hSUB_NAME)&"', '"&db2html(hCATEGORY_ID)&"', '"&db2html(hCATEGORY_NAME)&"')" & VBCRLF
'								dbCTget.Execute strSql, AssignedRow
'								actCnt = actCnt+AssignedRow
'							Next
'						Set hplist = nothing
'					End If
				On Error Goto 0
			Set xmlDOM = nothing
		Else
			HomeplusCategoryAPI = false
		End If
	Set objXML = nothing
End Function
'#####################################################################################################
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),20)
Dim arrItemid : arrItemid = request("cksel")
Dim alertMsg, iMessage, actCnt, sqlStr, retErrStr
Dim oHomeplus, i, strParam, iErrStr, ret1, chgSellYn, iitemid
chgSellYn = request("chgSellYn")
actCnt = 0

If (cmdparam = "RegSelect") Then				'선택상품 실제 등록
    arrItemid = Trim(arrItemid)
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
		dbCTget.Close: Response.End
	End If

	'## 선택상품 목록 접수
	Set oHomeplus = new CHomeplus
		oHomeplus.FPageSize	= 20
		oHomeplus.FRectItemID	= arrItemid
		oHomeplus.getHomeplusNotRegItemList

	    If (oHomeplus.FResultCount < 1) Then
	        arrItemid = split(arrItemid,",")
	        For i = LBound(arrItemid) to UBound(arrItemid)
	            CALL Fn_AcctFailTouch("homeplus",arrItemid(i),"등록가능상품 없음 :등록조건 확인: 판매Y, 옵션추가액...")
	        Next

	        If (IsAutoScript) Then
	            rw "S_ERR|등록가능상품 없음 :등록조건 확인: 판매Y, 할인..."
	            dbCTget.Close: Response.End
	        Else
	            Response.Write "<script language=javascript>alert('등록가능상품 없음 :등록조건 확인: 판매Y, 옵션추가액...');</script>"
				dbCTget.Close: Response.End
			End If
		End If

		For i = 0 to (oHomeplus.FResultCount - 1)
			If (oHomeplus.FItemList(i).FhDIVISION) = "" Then		'만약 기준카테고리 매핑을 하지 않은 카테고리라면
				Response.Write "<script language=javascript>alert('기준카테고리 매칭을 하지 않은 상품번호: [" & oHomeplus.FItemList(i).Fitemid & "]');</script>"
				dbCTget.Close: Response.End
			End If

			If (oHomeplus.FItemList(i).FbrandDepthCode = "") AND (oHomeplus.FItemList(i).FdepthCode = "") Then
				Response.Write "<script language=javascript>alert('전시카테고리 매칭을 하지 않은 상품번호: [" & oHomeplus.FItemList(i).Fitemid & "]');</script>"
				dbCTget.Close: Response.End
			End If

			sqlStr = ""
			sqlStr = sqlStr & " IF NOT Exists(SELECT * FROM db_outmall.dbo.tbl_homeplus_regItem where itemid="&oHomeplus.FItemList(i).Fitemid&")"
			sqlStr = sqlStr & " BEGIN"& VbCRLF
			sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_homeplus_regItem "
	        sqlStr = sqlStr & " (itemid, regdate, reguserid, homeplusstatCD, regitemname)"
	        sqlStr = sqlStr & " VALUES ("&oHomeplus.FItemList(i).Fitemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oHomeplus.FItemList(i).FItemName)&"')"
			sqlStr = sqlStr & " END "
			dbCTget.Execute sqlStr

			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oHomeplus.FItemList(i).checkTenItemOptionValid Then
			    On Error Resume Next
				'//상품등록 파라메터
				strParam = oHomeplus.FItemList(i).getHomeplusItemRegXML()

				If Err <> 0 Then
				    rw Err.Description
					Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oHomeplus.FItemList(i).Fitemid & "]');</script>"
					dbCTget.Close: Response.End
				End If

	            On Error Goto 0
	            iErrStr = ""
	            If HomeplusLoginAPI2() Then
					ret1 = HomeplusOneItemReg(oHomeplus.FItemList(i).FItemid, strParam, iErrStr, oHomeplus.FItemList(i).FSellCash, oHomeplus.FItemList(i).getHomeplusSellYn, oHomeplus.FItemList(i).FLimityn, oHomeplus.FItemList(i).FLimitNo, oHomeplus.FItemList(i).FLimitSold, html2db(oHomeplus.FItemList(i).FItemName), "createNewProduct")
		            If (ret1) Then
		                actCnt = actCnt+1
		            Else
		                CALL Fn_AcctFailTouch("homeplus", oHomeplus.FItemList(i).Fitemid, iErrStr)
		                retErrStr = retErrStr & iErrStr
		                rw iErrStr
		            End If
		        End If
			Else
				CALL Fn_AcctFailTouch("homeplus", oHomeplus.FItemList(i).Fitemid, iErrStr)
				iErrStr = "["&oHomeplus.FItemList(i).Fitemid&"] 옵션검사 실패"
				retErrStr = retErrStr & iErrStr
			End If
		Next
	Set oHomeplus = Nothing
    If (retErrStr <> "") Then
        Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
    End If
ElseIf (cmdparam = "EditSelect") Then				'정보 수정
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
		dbCTget.Close: Response.End
	End If
	'## 수정할 상품 목록 접수
	Set oHomeplus = new CHomeplus
		oHomeplus.FPageSize	= 20
		oHomeplus.FRectItemID	= arrItemid
		oHomeplus.getHomeplusEditedItemList
		For i = 0 to (oHomeplus.FResultCount - 1)
			On Error Resume Next
			strParam = oHomeplus.FItemList(i).getHomeplusItemEditXML()
		    iErrStr = ""
 			If HomeplusLoginAPI2() Then
				If (HomeplusOneItemEdit(oHomeplus.FItemList(i).Fitemid, oHomeplus.FItemList(i).FHomeplusGoodNo, iErrStr, strParam, "updateProduct")) Then
					actCnt = actCnt+1
				Else
	                CALL Fn_AcctFailTouch("homeplus", oHomeplus.FItemList(i).Fitemid, iErrStr)
	                retErrStr = retErrStr & iErrStr
					rw iErrStr
				End If
			End If
		Next
	Set oHomeplus = Nothing
	If (retErrStr<>"") Then
		Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
	End If
ElseIf (cmdparam = "EditItemSelect") Then			'정보외 수정
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
		dbCTget.Close: Response.End
	End If
	'## 수정할 상품 목록 접수
	Set oHomeplus = new CHomeplus
		oHomeplus.FPageSize	= 20
		oHomeplus.FRectItemID	= arrItemid
		oHomeplus.getHomeplusEditedItemList
		For i = 0 to (oHomeplus.FResultCount - 1)
			On Error Resume Next
 			If HomeplusLoginAPI2() Then
				strParam = ""
				If (oHomeplus.FItemList(i).FmaySoldOut = "Y") OR (oHomeplus.FItemList(i).IsSoldOutLimit5Sell) Then
					chgSellYn = "N"
				    iErrStr = ""
				    strParam = oHomeplus.FItemList(i).getHomeplusItemSellYNXML(chgSellYn)
					If (HomeplusOneItemSellStatEdit(oHomeplus.FItemList(i).Fitemid, oHomeplus.FItemList(i).FHomeplusGoodNo, chgSellYn, iErrStr, strParam, "setProductStatus")) Then
						rw "["&oHomeplus.FItemList(i).Fitemid&"]"&"품절처리"
						actCnt = actCnt+1
					Else
						rw "["&oHomeplus.FItemList(i).Fitemid&"]"&iErrStr
					End If
					sqlStr = ""
					sqlStr = sqlStr & " UPDATE db_outmall.dbo.tbl_homeplus_regItem SET lastStatCheckDate=getdate() WHERE itemid = '"&oHomeplus.FItemList(i).Fitemid&"' "
					dbCTget.Execute sqlStr
				Else

					If (oHomeplus.FItemList(i).FHomeplusSellYn = "N" AND oHomeplus.FItemList(i).IsSoldOut = False) Then
						iErrStr = ""
						chgSellYn = "Y"
						strParam = oHomeplus.FItemList(i).getHomeplusItemSellYNXML(chgSellYn)
						Call (HomeplusOneItemSellStatEdit(oHomeplus.FItemList(i).Fitemid, oHomeplus.FItemList(i).FHomeplusGoodNo, chgSellYn, iErrStr, strParam, "setProductStatus"))
						sqlStr = ""
						sqlStr = sqlStr & " UPDATE db_outmall.dbo.tbl_homeplus_regItem SET lastStatCheckDate=getdate() WHERE itemid = '"&oHomeplus.FItemList(i).Fitemid&"' "
						dbCTget.Execute sqlStr
					End If

	 				strParam = ""
					strParam = oHomeplus.FItemList(i).getHomeplusItemEditOPTXML()
				    iErrStr = ""
					Call (HomeplusOneItemOPTEdit(oHomeplus.FItemList(i).Fitemid, oHomeplus.FItemList(i).FHomeplusGoodNo, iErrStr, strParam, "updateProduct"))
					If iErrStr <> "" Then
						rw iErrStr
					End If

					strParam = ""
					strParam = oHomeplus.FItemList(i).getHomeplusItemViewXML()
				    iErrStr = ""
					If (HomeplusOneItemView(oHomeplus.FItemList(i).Fitemid, oHomeplus.FItemList(i).FHomeplusGoodNo, iErrStr, strParam, "searchProduct")) Then
						actCnt = actCnt+1
					Else
		                CALL Fn_AcctFailTouch("homeplus", oHomeplus.FItemList(i).Fitemid, iErrStr)
		                retErrStr = retErrStr & iErrStr
						rw iErrStr
					End If

					If Err <> 0 Then
						response.write Err.description
						Response.Write "<script language=javascript>alert('텐바이텐 HomeplusOneItemView 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & oHomeplus.FItemList(i).Fitemid & "]');</script>"
						dbCTget.Close: Response.End
					End If
			        on Error Goto 0
		        End If
			End If
		Next
	Set oHomeplus = Nothing
	If (retErrStr<>"") Then
		Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
	End If
ElseIf (cmdparam = "EditImgSelect") Then			'이미지 수정
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
		dbCTget.Close: Response.End
	End If
	'## 수정할 상품 목록 접수
	Set oHomeplus = new CHomeplus
		oHomeplus.FPageSize	= 20
		oHomeplus.FRectItemID	= arrItemid
		oHomeplus.getHomeplusEditedItemList
		For i = 0 to (oHomeplus.FResultCount - 1)
			On Error Resume Next
			strParam = oHomeplus.FItemList(i).getHomeplusItemEditImgXML()
		    iErrStr = ""
 			If HomeplusLoginAPI2() Then
				If (HomeplusOneItemIMGEdit(oHomeplus.FItemList(i).Fitemid, oHomeplus.FItemList(i).FHomeplusGoodNo, iErrStr, strParam, "updateImage")) Then
					actCnt = actCnt+1
				Else
	                CALL Fn_AcctFailTouch("homeplus", oHomeplus.FItemList(i).Fitemid, iErrStr)
	                retErrStr = retErrStr & iErrStr
					rw iErrStr
				End If
			End If
		Next
	Set oHomeplus = Nothing
	If (retErrStr<>"") Then
		Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
	End If

ElseIf (cmdparam = "EditSellYn") Then				'판매상태 수정
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
		dbCTget.Close: Response.End
	End If
	'## 수정할 상품 목록 접수
	Set oHomeplus = new CHomeplus
		oHomeplus.FPageSize	= 20
		oHomeplus.FRectItemID	= arrItemid
		oHomeplus.getHomeplusEditedItemList

		If (chgSellYn="N") and (oHomeplus.FResultCount < 1) and (arrItemid = "") Then
		    oHomeplus.getHomeplusreqExpireItemList
		End If

		For i = 0 to (oHomeplus.FResultCount - 1)
			strParam = oHomeplus.FItemList(i).getHomeplusItemSellYNXML(chgSellYn)
		    iErrStr = ""
 			If HomeplusLoginAPI2() Then
				If (HomeplusOneItemSellStatEdit(oHomeplus.FItemList(i).Fitemid, oHomeplus.FItemList(i).FHomeplusGoodNo, chgSellYn, iErrStr, strParam, "setProductStatus")) Then
					actCnt = actCnt+1
				Else
					rw "["&iitemid&"]"&iErrStr
				End If
			End If
			retErrStr = retErrStr & iErrStr
		Next
	Set oHomeplus = Nothing
	If (retErrStr<>"") Then
		Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
	End If
ElseIf (cmdparam="ViewSelect") Then
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
		dbCTget.Close: Response.End
	End If
	'## 수정할 상품 목록 접수
	Set oHomeplus = new CHomeplus
		oHomeplus.FPageSize	= 20
		oHomeplus.FRectItemID	= arrItemid
		oHomeplus.getHomeplusEditedItemList

		On Error Resume Next
		strParam = oHomeplus.FItemList(i).getHomeplusItemViewXML()
		For i = 0 to (oHomeplus.FResultCount - 1)
		    iErrStr = ""

			If HomeplusLoginAPI2() Then
				If (HomeplusOneItemView(oHomeplus.FItemList(i).Fitemid, oHomeplus.FItemList(i).FHomeplusGoodNo, iErrStr, strParam, "searchProduct")) Then
					actCnt = actCnt+1
				Else
	                CALL Fn_AcctFailTouch("homeplus", oHomeplus.FItemList(i).Fitemid, iErrStr)
	                retErrStr = retErrStr & iErrStr
					rw iErrStr
				End If
			End If
		Next

	If (retErrStr<>"") Then
		Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
	End If
ElseIf (cmdparam="CategoryView") Then
	If HomeplusLoginAPI2() Then
		HomeplusCategoryAPI()
	End If
Else
	rw "미지정 ["&cmdparam&"]"
End If

If Err.Number = 0 Then
	If (IsAutoScript) then
		rw "OK|"& iMessage & "<br>"& actCnt & "건이 처리되었습니다."
	Else
		Response.Write "<script language=javascript>alert('" & iMessage & "\n"& actCnt & "건이 처리되었습니다.');</script>"
	End if
Else
	If (IsAutoScript) then
		rw "S_ERR|처리 중에 오류가 발생했습니다"
	Else
		Response.Write "<script language=javascript>alert('자료 저장 중에 오류가 발생했습니다.\n관리자에게 문의해주세요.');</script>"
	End if
End If
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->