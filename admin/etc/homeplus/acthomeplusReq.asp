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
'#################################### Ȩ�÷��� �⺻ ���� Setting ####################################
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

'################################### ���� Function Setting  ##########################################
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

		If retCode = "E0000" Then	'����(E0000)
			'��ǰ���翩�� Ȯ��
			strSql = "SELECT COUNT(itemid) FROM db_outmall.dbo.tbl_homeplus_regItem WHERE itemid='" & iitemid & "'"
			rsCTget.Open strSql, dbCTget, 1
			If rsCTget(0) > 0 Then
				'// ���� -> ����
				strSql = ""
				strSql = strSql & " UPDATE R" & VbCRLF
				strSql = strSql & "	Set homeplusLastUpdate = getdate() "  & VbCRLF
				strSql = strSql & "	, homeplusGoodNo = '" & homegoodNo & "'"  & VbCRLF
				strSql = strSql & "	, homeplusPrice = " &iSellCash& VbCRLF
				strSql = strSql & "	, accFailCnt = 0"& VbCRLF
				strSql = strSql & "	, homeplusRegdate = isNULL(homeplusRegdate, getdate())"
				If (homegoodNo <> "") Then
				    strSql = strSql & "	, homeplusstatCD = '7'"& VbCRLF					'��ϿϷ�(�ӽ�)
				Else
					strSql = strSql & "	, homeplusstatCD = '1'"& VbCRLF					'���۽õ�
				End If
				strSql = strSql & "	From db_outmall.dbo.tbl_homeplus_regItem R"& VbCRLF
				strSql = strSql & " Where R.itemid = '" & iitemid & "'"
				dbCTget.Execute(strSql)
			Else
				'// ���� -> �űԵ��
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
				    strSql = strSql & ",'7'"											'��ϿϷ�(�ӽ�)
				Else
				    strSql = strSql & ",'1'"											'���۽õ�
				End If
				strSql = strSql & ")"
				dbCTget.Execute(strSql)
				actCnt = actCnt + 1
			End If
			rsCTget.Close

			Set optlist = xmlDOM.getElementsByTagName("ns1:ITEMRESULT")
				For each SubNodes in optlist
					s_ITEMNO		= Trim(SubNodes.getElementsByTagName("ns1:s_ITEMNO").item(0).text)		'�ٹ����� �ɼ��ڵ�
					i_ITEMNO		= Trim(SubNodes.getElementsByTagName("ns1:i_ITEMNO").item(0).text)		'Ȩ�÷��� �ɼ��ڵ�
					s_OPTION_NAME	= Trim(SubNodes.getElementsByTagName("ns1:s_OPTION_NAME").item(0).text)	'�ɼǸ�
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
		Else						'����(E)
		    iErrStr =  "��ǰ ����� ���� [" & iitemid & "]:"&iMessage
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

		If retCode = "E0000" Then	'����(E0000)�̸� ����Ƚ�� �ʱ�ȭ
			strSql = ""
			strSql = strSql & " UPDATE db_outmall.dbo.tbl_homeplus_regItem " & VbCRLF
			strSql = strSql & " SET accFailCnt = 0 " & VbCRLF
			strSql = strSql & " WHERE itemid='" & iitemid & "'"
			dbCTget.Execute(strSql)

			HomeplusOneItemEdit = true
			Set objXML = Nothing
			Set xmlDOM = Nothing
			rw "[" & iitemid & "]:"&iMessage
		Else						'����(E)
			iErrStr =  "���� ���� �� ���� [" & iitemid & "]:"&iMessage
			Set objXML = Nothing
			Set xmlDOM = Nothing
		    Exit Function
		End If
	Else
	    iErrStr =  "[" & iitemid & "]:��ǰ ���� �� ��� ���� "
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

		If retCode = "E0000" Then	'����(E0000)�̸� ����Ƚ�� �ʱ�ȭ
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
		Else						'����(E)
			iErrStr =  "������ ���� �� ���� �� ���� [" & iitemid & "]:"&iMessage
			Set objXML = Nothing
			Set xmlDOM = Nothing
		    Exit Function
		End If
	Else
	    iErrStr =  "[" & iitemid & "]:������ ���� �� ���� �� ��� ���� "
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

		If retCode = "E0000" Then	'����(E0000)
			HomeplusOneItemIMGEdit = true
			Set objXML = Nothing
			Set xmlDOM = Nothing
			rw "[" & iitemid & "]:"&iMessage
		Else						'����(E)
			iErrStr =  "�̹��� ���� �� ���� [" & iitemid & "]:"&iMessage
			Set objXML = Nothing
			Set xmlDOM = Nothing
		    Exit Function
		End If
	Else
	    iErrStr =  "[" & iitemid & "]:��ǰ �̹��� ���� �� ��� ���� "
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

		If retCode = "E0000" Then	'����(E0000)
			Set oneProdInfo = XMLDom.SelectNodes("soapenv:Envelope/soapenv:Body/searchProductResponse/ns1:searchProductReturn/ns1:ITEMRESULT/ns1:ITEMRESULT")
				For Each SubNodes In oneProdInfo
					hplOptStatus			= SubNodes.SelectSingleNode("ns1:SALE").text
					hplOptno				= SubNodes.SelectSingleNode("ns1:i_ITEMNO").text
					regedOpt10x10OptNo		= SubNodes.SelectSingleNode("ns1:s_ITEMNO").text
					regedOpt10x10OptNm		= SubNodes.SelectSingleNode("ns1:s_OPTION_NAME").text

					If oHomeplus.FItemList(i).FoptionCnt = 0 Then		'��ǰ�̶��
						StockQty = oHomeplus.FItemList(i).GetHomeplusLmtQty
					Else												'�ɼ��̶��
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
					'1.������ �μ�Ʈ ������ ������Ʈ
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
				    strSql = strSql & " and convert(int, outmallOptCode)='"&hplOptno&"'"&VbCRLF				'������ outmallOptCode�� 001,002,003 �̷��� �������� ���� �Ŀ� 1,2,3�̷��� ����
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
'				2.regedItemStatus�� ���� ��ǰ�Ǹ� ���� ���� / regrdOptcnt�� ����
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
				rw "["&iitemid&"]:�ɼ��� ���� ǰ���̹Ƿ� ǰ���� �����մϴ�"
			End If
			rsCTget.Close
		Else						'����(E)
			iErrStr =  "��ǰ ��ȸ �� ���� [" & iitemid & "]:"&iMessage
			Set objXML = Nothing
			Set xmlDOM = Nothing
		    Exit Function
		End If
	Else
	    iErrStr =  "[" & iitemid & "]:��ǰ ��ȸ �� ��� ���� "
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

		If retCode = "E0000" Then	'����(E0000)
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
		Else						'����(E)
		    iErrStr =  "��ǰ ���º��� �� ���� [" & iitemid & "]:"&iMessage
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

'�α��� API
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
					Response.Write "<script language=javascript>alert('Ȩ�÷��� �α��ο� �����߽��ϴ�.');history.back();</script>"
					Response.End
					HomeplusLoginAPI = false
				End If

				If xmlDOM.getElementsByTagName("ns1:code").item(0).text = "E0000" Then
					HomeplusLoginAPI = true
				Else
					Response.Write "<script language=javascript>alert('Ȩ�÷��� �α��ο� �����߽��ϴ�.');history.back();</script>"
					Response.End
				End If
				On Error Goto 0
			Set xmlDOM = nothing
		Else
			Response.Write "<script language=javascript>alert('Ȩ�÷��� �α����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');history.back();</script>"
			Response.End
			HomeplusLoginAPI = false
		End If
	Set objXML = nothing
End Function

'�α��� API(XMLHTTP�δ� ���� �α��ζ� time out�� �ɷ��� ����)
'ServerXMLHTTP �̰ɷδ� ������ ������ �� �Ǽ� �ι� ȣ��;;
'�� ���� ����� ������ ����;
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
						Response.Write "<script language=javascript>alert('Ȩ�÷��� �α��ο� �����߽��ϴ�.');history.back();</script>"
						Response.End
						HomeplusLoginAPI2 = false
					End If

					If xmlDOM.getElementsByTagName("ns1:code").item(0).text = "E0000" Then
						HomeplusLoginAPI2 = true
					Else
						Response.Write "<script language=javascript>alert('Ȩ�÷��� �α��ο� �����߽��ϴ�.');history.back();</script>"
						Response.End
					End If
					On Error Goto 0
				Set xmlDOM = nothing
			Else
				Response.Write "<script language=javascript>alert('Ȩ�÷��� �α����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');history.back();</script>"
				Response.End
				HomeplusLoginAPI2 = false
			End If
		Set objXML = nothing
	End If
End Function

'ī�װ� API
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
'								hDIVISION		= Trim(SubNodes.getElementsByTagName("ns1:DIVISION").item(0).text)		'�ֻ��� �з��ڵ�
'								hGROUP			= Trim(SubNodes.getElementsByTagName("ns1:GROUP").item(0).text)			'DIVISION ���� �з� �ڵ�
'								hDEPT			= Trim(SubNodes.getElementsByTagName("ns1:DEPT").item(0).text)			'GROUP ���� �з� �ڵ�
'								hCLASS			= Trim(SubNodes.getElementsByTagName("ns1:CLASS").item(0).text)			'DEPT ���� �з� �ڵ�
'								hSUBCLASS		= Trim(SubNodes.getElementsByTagName("ns1:SUBCLASS").item(0).text)		'CLASS ���� �з� �ڵ�
'								hDIV_NAME		= Trim(SubNodes.getElementsByTagName("ns1:DIV_NAME").item(0).text)		'DIVISION �з���
'								hGROUP_NAME		= Trim(SubNodes.getElementsByTagName("ns1:GROUP_NAME").item(0).text)	'GROUP �з���
'								hDEPT_NAME		= Trim(SubNodes.getElementsByTagName("ns1:DEPT_NAME").item(0).text)		'DEPT �з���
'								hCLASS_NAME		= Trim(SubNodes.getElementsByTagName("ns1:CLASS_NAME").item(0).text)	'CLASS �з���
'								hSUB_NAME		= Trim(SubNodes.getElementsByTagName("ns1:SUB_NAME").item(0).text)		'SUBCLASS �з���
'								hCATEGORY_ID	= Trim(SubNodes.getElementsByTagName("ns1:CATEGORY_ID").item(0).text)	'��ǰ����������ø� ���� ī�װ� ���̵�
'								hCATEGORY_NAME	= Trim(SubNodes.getElementsByTagName("ns1:CATEGORY_NAME").item(0).text)	'��ǰ����������ø� ���� ī�װ� ��
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

If (cmdparam = "RegSelect") Then				'���û�ǰ ���� ���
    arrItemid = Trim(arrItemid)
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbCTget.Close: Response.End
	End If

	'## ���û�ǰ ��� ����
	Set oHomeplus = new CHomeplus
		oHomeplus.FPageSize	= 20
		oHomeplus.FRectItemID	= arrItemid
		oHomeplus.getHomeplusNotRegItemList

	    If (oHomeplus.FResultCount < 1) Then
	        arrItemid = split(arrItemid,",")
	        For i = LBound(arrItemid) to UBound(arrItemid)
	            CALL Fn_AcctFailTouch("homeplus",arrItemid(i),"��ϰ��ɻ�ǰ ���� :������� Ȯ��: �Ǹ�Y, �ɼ��߰���...")
	        Next

	        If (IsAutoScript) Then
	            rw "S_ERR|��ϰ��ɻ�ǰ ���� :������� Ȯ��: �Ǹ�Y, ����..."
	            dbCTget.Close: Response.End
	        Else
	            Response.Write "<script language=javascript>alert('��ϰ��ɻ�ǰ ���� :������� Ȯ��: �Ǹ�Y, �ɼ��߰���...');</script>"
				dbCTget.Close: Response.End
			End If
		End If

		For i = 0 to (oHomeplus.FResultCount - 1)
			If (oHomeplus.FItemList(i).FhDIVISION) = "" Then		'���� ����ī�װ� ������ ���� ���� ī�װ����
				Response.Write "<script language=javascript>alert('����ī�װ� ��Ī�� ���� ���� ��ǰ��ȣ: [" & oHomeplus.FItemList(i).Fitemid & "]');</script>"
				dbCTget.Close: Response.End
			End If

			If (oHomeplus.FItemList(i).FbrandDepthCode = "") AND (oHomeplus.FItemList(i).FdepthCode = "") Then
				Response.Write "<script language=javascript>alert('����ī�װ� ��Ī�� ���� ���� ��ǰ��ȣ: [" & oHomeplus.FItemList(i).Fitemid & "]');</script>"
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

			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oHomeplus.FItemList(i).checkTenItemOptionValid Then
			    On Error Resume Next
				'//��ǰ��� �Ķ����
				strParam = oHomeplus.FItemList(i).getHomeplusItemRegXML()

				If Err <> 0 Then
				    rw Err.Description
					Response.Write "<script language=javascript>alert('�ٹ����� ��ǰ���� ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oHomeplus.FItemList(i).Fitemid & "]');</script>"
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
				iErrStr = "["&oHomeplus.FItemList(i).Fitemid&"] �ɼǰ˻� ����"
				retErrStr = retErrStr & iErrStr
			End If
		Next
	Set oHomeplus = Nothing
    If (retErrStr <> "") Then
        Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
    End If
ElseIf (cmdparam = "EditSelect") Then				'���� ����
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbCTget.Close: Response.End
	End If
	'## ������ ��ǰ ��� ����
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
ElseIf (cmdparam = "EditItemSelect") Then			'������ ����
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbCTget.Close: Response.End
	End If
	'## ������ ��ǰ ��� ����
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
						rw "["&oHomeplus.FItemList(i).Fitemid&"]"&"ǰ��ó��"
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
						Response.Write "<script language=javascript>alert('�ٹ����� HomeplusOneItemView ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oHomeplus.FItemList(i).Fitemid & "]');</script>"
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
ElseIf (cmdparam = "EditImgSelect") Then			'�̹��� ����
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbCTget.Close: Response.End
	End If
	'## ������ ��ǰ ��� ����
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

ElseIf (cmdparam = "EditSellYn") Then				'�ǸŻ��� ����
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbCTget.Close: Response.End
	End If
	'## ������ ��ǰ ��� ����
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
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbCTget.Close: Response.End
	End If
	'## ������ ��ǰ ��� ����
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
	rw "������ ["&cmdparam&"]"
End If

If Err.Number = 0 Then
	If (IsAutoScript) then
		rw "OK|"& iMessage & "<br>"& actCnt & "���� ó���Ǿ����ϴ�."
	Else
		Response.Write "<script language=javascript>alert('" & iMessage & "\n"& actCnt & "���� ó���Ǿ����ϴ�.');</script>"
	End if
Else
	If (IsAutoScript) then
		rw "S_ERR|ó�� �߿� ������ �߻��߽��ϴ�"
	Else
		Response.Write "<script language=javascript>alert('�ڷ� ���� �߿� ������ �߻��߽��ϴ�.\n�����ڿ��� �������ּ���.');</script>"
	End if
End If
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->