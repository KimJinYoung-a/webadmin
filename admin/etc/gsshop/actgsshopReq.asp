<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/admin/etc/gsshop/gsshopItemcls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
Response.CharSet = "euc-kr"
''GSSHOP ��ǰ ���
Function GsshopOneItemReg(iitemid, strParam, byRef iErrStr, iSellCash, iGSShopSellYn, ilimityn, ilimitno, ilimiysold, iitemname)
	Dim objXML, xmlDOM, strRst
	Dim buf, strSql, AssignedRow
	Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	Dim attrPrdlist, lp, tenOptcd, gsOptcd
	Dim Tlimitno, Tlimitsold, Titemoption, Toptionname, Toptlimitno, Toptlimitsold, Toptsellyn, Toptlimityn, Toptaddprice, Tlimityn, Tsellyn, Titemsu, Tsellcash

	On Error Resume Next
	GsshopOneItemReg = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
'rw gsshopAPIURL "?"&strparam
'response.end
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
		    buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf
				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					rw buf
				End If

				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�
				attrPrdCd	= Split(buf, "|")(5)	'�̼��Ӽ���ǰ�ڵ�^���»�Ӽ���ǰ�ڵ�,�̼��Ӽ���ǰ�ڵ�^���»�Ӽ���ǰ�ڵ�	'�Ӽ��Ķ��Ÿ �������� ������ ���� �� ����

				If resultcode = "S" Then	'����(S)
					'��ǰ���翩�� Ȯ��
					strSql = "Select count(itemid) From db_item.dbo.tbl_gsshop_regitem Where itemid='" & iitemid & "'"
					rsget.Open strSql,dbget,1
					If rsget(0) > 0 Then
						'// ���� -> ����
						strSql = ""
						strSql = strSql & " UPDATE R" & VbCRLF
						strSql = strSql & "	Set GSShopLastUpdate = getdate() "  & VbCRLF
						strSql = strSql & "	, GSShopGoodNo = '" & prdCd & "'"  & VbCRLF
						strSql = strSql & "	, GSShopPrice = " &iSellCash& VbCRLF
						strSql = strSql & "	, accFailCnt = 0"& VbCRLF
						strSql = strSql & "	, GSShopRegdate = isNULL(GSShopRegdate, getdate())"
						If (prdCd <> "") Then
						    strSql = strSql & "	, GSShopstatCD = '3'"& VbCRLF					'��ϿϷ�(�ӽ�)
						Else
							strSql = strSql & "	, GSShopstatCD = '1'"& VbCRLF					'���۽õ�
						End If
						strSql = strSql & "	From db_item.dbo.tbl_gsshop_regItem R"& VbCRLF
						strSql = strSql & " Where R.itemid = '" & iitemid & "'"
						dbget.Execute(strSql)
					Else
						'// ���� -> �űԵ��
						strSql = ""
						strSql = strSql & " INSERT INTO db_item.dbo.tbl_gsshop_regItem "
						strSql = strSql & " (itemid, regitemname, reguserid, GSShopRegdate, GSShopLastUpdate, GSShopGoodNo, GSShopPrice, GSShopSellYn, GSShopStatCd) VALUES " & VbCRLF
						strSql = strSql & " ('" & iitemid & "'" & VBCRLF
						strSql = strSql & " , '" & iitemname & "'" &_
						strSql = strSql & " , '" & session("ssBctId") & "'" &_
						strSql = strSql & " , getdate(), getdate()" & VBCRLF
						strSql = strSql & " , '" & prdCd & "'" & VBCRLF
						strSql = strSql & " , '" & iSellCash & "'" & VBCRLF
						strSql = strSql & " , '" & iGSShopSellYn & "'" & VBCRLF
						If (prdCd <> "") Then
						    strSql = strSql & ",'3'"											'��ϿϷ�(�ӽ�)
						Else
						    strSql = strSql & ",'1'"											'���۽õ�
						End If
						strSql = strSql & ")"
						dbget.Execute(strSql)
						actCnt = actCnt + 1
					End If
					rsget.Close

					attrPrdlist = split(attrPrdCd,",")
					If Ubound(attrPrdlist) = 0 Then
						strSql = ""
						strSql = strSql & " SELECT COUNT(*) FROM db_item.dbo.tbl_item_option WHERE itemid = '"&iitemid&"' "
						rsget.Open strSql,dbget,1
						If rsget(0) = 0 Then
							tenOptcd	= "0000"
						End If
						rsget.Close
					End If

					If (Ubound(attrPrdlist)) = 0 AND (tenOptcd = "0000")  Then	'���� ��ǰ�̶��
						gsOptcd			= split(attrPrdCd,"^")(0)
						Toptionname		= "���ϻ�ǰ"
						Tlimitno		= ilimitno
						Tlimitsold		= ilimiysold
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
						sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
						sqlStr = sqlStr & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
						sqlStr = sqlStr & " VALUES " & VBCRLF
						sqlStr = sqlStr & " ('"&iitemid&"',  '"&tenOptcd&"', 'gsshop', '"&gsOptcd&"', '"&html2db(Toptionname)&"', 'Y', '"&Tlimityn&"', '"&Titemsu&"', '0', getdate()) "
						dbget.Execute sqlStr
					Else														'�ɼ� ���� ��ǰ�̶��
						For lp = Lbound(attrPrdlist) to Ubound(attrPrdlist)
							gsOptcd		= split(attrPrdlist(lp),"^")(0)
			                tenOptcd	= split(attrPrdlist(lp),"^")(1)
							sqlStr = ""
							sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
							sqlStr = sqlStr & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
							sqlStr = sqlStr & " SELECT itemid, itemoption, 'gsshop', '"&gsOptcd&"', optionname, optsellyn, optlimityn, " & VBCRLF
							sqlStr = sqlStr & " Case WHEN optlimityn = 'Y' AND optlimitno - optlimitsold <= 5 THEN '0' " & VBCRLF
							sqlStr = sqlStr & " 	 WHEN optlimityn = 'Y' AND optlimitno - optlimitsold > 5 THEN optlimitno - optlimitsold - 5 " & VBCRLF
							sqlStr = sqlStr & " 	 WHEN optlimityn = 'N' THEN '999' End " & VBCRLF
							sqlStr = sqlStr & " , '0', getdate() " & VBCRLF
							sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_option " & VBCRLF
							sqlStr = sqlStr & " WHERE itemid= '"&iitemid&"' " & VBCRLF
							sqlStr = sqlStr & " and itemoption = '"& tenOptcd &"' "
							dbget.Execute sqlStr
						Next
					End If
					strSql = ""
					strSql = strSql & " UPDATE R " & VBCRLF
					strSql = strSql & " SET regedOptCnt = isNULL(T.regedOptCnt,0) " & VBCRLF
					strSql = strSql & " FROM db_item.dbo.tbl_gsshop_regItem R " & VBCRLF
					strSql = strSql & " Join ( " & VBCRLF
					strSql = strSql & " 	SELECT R.itemid, count(*) as CNT, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt "
					strSql = strSql & " 	FROM db_item.dbo.tbl_gsshop_regItem R " & VBCRLF
					strSql = strSql & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro on R.itemid = Ro.itemid and Ro.mallid = 'gsshop' and Ro.itemid = " &iitemid & VBCRLF
					strSql = strSql & " 	GROUP BY R.itemid " & VBCRLF
					strSql = strSql & " ) T on R.itemid = T.itemid " & VBCRLF
					dbget.Execute strSql
				Else						'����(E)
				    iErrStr =  "��ǰ ����� ���� [" & iitemid & "]:"&resultmsg
			        Set objXML = Nothing
			        Set xmlDOM = Nothing
				    Exit Function
				End If
			Set xmlDOM = Nothing
			GsshopOneItemReg= true
		Else
			iErrStr = "GSSHOP�� ����߿� ������ �߻��߽��ϴ�..[ERR-REG-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'���û�ǰ �ǸŰ� ����
Function GSShopOnItemPriceEdit(iitemid,strParam,byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	GSShopOnItemPriceEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					rw buf
				End If

				If Err <> 0 Then
					iErrStr = "GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-PRICEEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'��� �ڵ�
				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�

				If resultcode = "S" Then
					iErrStr = "["&iitemid & "]:"&resultmsg
					GSShopOnItemPriceEdit = True
				Else
					iMessage = resultmsg
					GSShopOnItemPriceEdit = False
				End If

				If Err <> 0 Then
					If (IsAutoScript) Then
						rw "["&iitemid & "]:"&iMessage
					Else
						iErrStr =  "["&iitemid & "]:"&iMessage
					End If
					Set objXML = Nothing
				    Set xmlDOM = Nothing
				    On Error Goto 0
				    Exit Function
			    End If
			Set xmlDOM = Nothing
		Else
			GSShopOnItemPriceEdit = False
			iErrStr = "GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-PRICEEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'���û�ǰ �̹��� ����
Function GSShopOneItemImageEdit(iitemid,strParam,byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	GSShopOneItemImageEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					rw buf
				End If

				If Err <> 0 Then
					iErrStr = "GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-IMGEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'��� �ڵ�
				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�

				If resultcode = "S" Then
					iErrStr = "["&iitemid & "]:"&resultmsg
				Else
					iMessage = resultmsg
				End If

				If Err <> 0 Then
					If (IsAutoScript) Then
						rw "["&iitemid & "]:"&iMessage
					Else
						iErrStr =  "["&iitemid & "]:"&iMessage
					End If
					Set objXML = Nothing
				    Set xmlDOM = Nothing
				    On Error Goto 0
				    Exit Function
			    End If
			Set xmlDOM = Nothing
			GSShopOneItemImageEdit = True
		Else
			iErrStr = "GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-IMGEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'���û�ǰ ���� ����
Function GSShopOneItemContentsEdit(iitemid,strParam,byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	GSShopOneItemContentsEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					rw buf
				End If

				If Err <> 0 Then
					iErrStr = "GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-CONTEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'��� �ڵ�
				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�

				If resultcode = "S" Then
					iErrStr = "["&iitemid & "]:"&resultmsg
				Else
					iMessage = resultmsg
				End If

				If Err <> 0 Then
					If (IsAutoScript) Then
						rw "["&iitemid & "]:"&iMessage
					Else
						iErrStr =  "["&iitemid & "]:"&iMessage
					End If
					Set objXML = Nothing
				    Set xmlDOM = Nothing
				    On Error Goto 0
				    Exit Function
			    End If
			Set xmlDOM = Nothing
			GSShopOneItemContentsEdit = True
		Else
			iErrStr = "GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-CONTEDIT-002]"
	        Set objXML = Nothing
	        Set xmlDOM = Nothing
	        On Error Goto 0
		    Exit Function
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'��ǰ �ɼ� �߰� �� ���� ����
Function GSShopOPTSuEdit(iitemid,strParam,byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, tenOptcd, lp, gsOptcd
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd, attrPrdlist
	On Error Resume Next
	GSShopOPTSuEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					rw buf
				End If

				If Err <> 0 Then
					iErrStr = "GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-OPTSuEdit-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If

				'��� �ڵ�
				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�
				attrPrdCd	= Split(buf, "|")(5)	'�̼��Ӽ���ǰ�ڵ�^���»�Ӽ���ǰ�ڵ�,�̼��Ӽ���ǰ�ڵ�^���»�Ӽ���ǰ�ڵ�	'�Ӽ��Ķ��Ÿ �������� ������ ���� �� ����

				If resultcode = "S" Then	'����(S)
					iErrStr = "["&iitemid & "]:"&resultmsg
					attrPrdlist = split(attrPrdCd,",")
					For lp = Lbound(attrPrdlist) to Ubound(attrPrdlist)
						gsOptcd		= split(attrPrdlist(lp),"^")(0)
		                tenOptcd	= split(attrPrdlist(lp),"^")(1)
						If Ubound(attrPrdlist) = 0 AND tenOptcd = "0000" Then	'��ǰ�̶��
							sqlStr = ""
							sqlStr = sqlStr & "UPDATE db_item.dbo.tbl_OutMall_regedoption SET "
							sqlStr = sqlStr & "outmalllimitno =  "
							sqlStr = sqlStr & "Case WHEN B.limityn = 'Y' and B.limitno - B.limitsold <= 5 THEN '0'  "
							sqlStr = sqlStr & "	 WHEN B.limityn = 'Y' and B.limitno - B.limitsold > 5 THEN B.limitno - B.limitsold - 5 "
							sqlStr = sqlStr & "	 WHEN B.limityn = 'N' THEN '999' END "
							sqlStr = sqlStr & "FROM db_item.dbo.tbl_OutMall_regedoption A  "
							sqlStr = sqlStr & "JOIN db_item.dbo.tbl_item B on A.itemid = B.itemid "
							sqlStr = sqlStr & "WHERE A.itemid = '"&iitemid&"' and A.itemoption = '"&tenOptcd&"' and A.mallid = 'gsshop' "
							dbget.Execute sqlStr
						Else
							sqlStr = ""
							sqlStr = sqlStr & " IF Exists(SELECT * FROM db_item.dbo.tbl_OutMall_regedoption where itemid='"&iitemid&"' and itemoption = '"&tenOptcd&"' and mallid = 'gsshop') "
							sqlStr = sqlStr & " BEGIN"& VbCRLF
							sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_OutMall_regedoption " & VbCRLF
							sqlStr = sqlStr & " SET outmalllimitno = " & VbCRLF
							sqlStr = sqlStr & " Case WHEN optlimityn = 'Y' AND optlimitno - optlimitsold <= 5 THEN '0' " & VbCRLF
							sqlStr = sqlStr & " 	 WHEN optlimityn = 'Y' AND optlimitno - optlimitsold > 5 THEN optlimitno - optlimitsold - 5" & VbCRLF
							sqlStr = sqlStr & " 	 WHEN optlimityn = 'N' THEN '999' End" & VbCRLF
							sqlStr = sqlStr & " ,outmalllimityn = B.optlimityn " & VbCRLF
							sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption A  " & VbCRLF
							sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_option B on A.itemid = B.itemid and A.itemoption = B.itemoption " & VbCRLF
							sqlStr = sqlStr & " WHERE B.itemid = '"&iitemid&"' and B.itemoption = '"&tenOptcd&"' and A.mallid = 'gsshop' "
							sqlStr = sqlStr & " END ELSE "
							sqlStr = sqlStr & " BEGIN"& VbCRLF
							sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_OutMall_regedoption " & VBCRLF
							sqlStr = sqlStr & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastUpdate) " & VBCRLF
							sqlStr = sqlStr & " SELECT itemid, itemoption, 'gsshop', '"&gsOptcd&"', optionname, optsellyn, optlimityn, " & VBCRLF
							sqlStr = sqlStr & " Case WHEN optlimityn = 'Y' AND optlimitno - optlimitsold <= 5 THEN '0' " & VBCRLF
							sqlStr = sqlStr & " 	 WHEN optlimityn = 'Y' AND optlimitno - optlimitsold > 5 THEN optlimitno - optlimitsold - 5 " & VBCRLF
							sqlStr = sqlStr & " 	 WHEN optlimityn = 'N' THEN '999' End " & VBCRLF
							sqlStr = sqlStr & " , '0', getdate() " & VBCRLF
							sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_option " & VBCRLF
							sqlStr = sqlStr & " WHERE itemid= '"&iitemid&"' " & VBCRLF
							sqlStr = sqlStr & " and itemoption = '"& tenOptcd &"' "
							sqlStr = sqlStr & " END "
						    dbget.Execute sqlStr
						End If
					Next
				Else						'����(E)
				    iErrStr =  "��ǰ �ɼ� ���� ���� ���� [" & iitemid & "]:"&resultmsg
			        Set objXML = Nothing
			        Set xmlDOM = Nothing
			        On Error Goto 0
				    Exit Function
				End If
			Set xmlDOM = Nothing
			GSShopOPTSuEdit = True
		Else
			iErrStr = "GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-OPTSuEdit-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'��ǰ �ɼ� �߰� �� ���� ����
Function GSShopOPTSellEdit(iitemid,strParam,byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf, tenOptcd, lp, gsOptcd
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd, attrPrdlist
'	On Error Resume Next
	GSShopOPTSellEdit = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					rw buf
				End If

				If Err <> 0 Then
					iErrStr = "GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-OPTSellEdit-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If

				'��� �ڵ�
				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�
				''���ߵ� �� S->P�� ���� ���� �ִ�.(�ɼǻ�ǰ���� ��ǰ��ǰ���� ���ϴ� ��쿡�� ���� ������ ���� �����ؾߵǴ� �� �ݵ�� ó���ؾ���..

				If resultcode = "S" Then	'����(S)
					iErrStr = "["&iitemid & "]:"&resultmsg

	                sqlStr = ""
					sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_OutMall_regedoption " & VbCRLF
					sqlStr = sqlStr & " SET outmallsellyn = " & VbCRLF
					sqlStr = sqlStr & " Case WHEN (B.isusing = 'N' OR B.optsellyn = 'N') THEN 'N' " & VbCRLF
					sqlStr = sqlStr & " 	 WHEN (B.optlimityn = 'Y' AND B.optlimitno - B.optlimitsold <= 5) THEN 'N'  " & VbCRLF
					sqlStr = sqlStr & " 	 WHEN (A.outmallOptName <> B.optionname) THEN 'N'  " & VbCRLF
					sqlStr = sqlStr & " ELSE 'Y' END " & VbCRLF
					sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption A  " & VbCRLF
					sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_option B on A.itemid = B.itemid and A.itemoption = B.itemoption " & VbCRLF
					sqlStr = sqlStr & " WHERE B.itemid = '"&iitemid&"' and A.mallid = 'gsshop' "
				    dbget.Execute sqlStr
				Else						'����(E)
				    iErrStr =  "�ɼ� �ǸŻ��� ���� ���� [" & iitemid & "]:"&resultmsg
			        Set objXML = Nothing
			        Set xmlDOM = Nothing
				    Exit Function
				End If
			Set xmlDOM = Nothing
			GSShopOPTSellEdit = True
		Else
			iErrStr = "GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-OPTSellEdit-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'���û�ǰ ��ǰ�� ����
Function GSShopOnItemItemnameEdit(iitemid,strParam,byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	GSShopOnItemItemnameEdit = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					rw buf
				End If

				If Err <> 0 Then
					iErrStr = "GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-NMEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'��� �ڵ�
				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�

				If resultcode = "S" Then
					iErrStr = "["&iitemid & "]:"&resultmsg
				Else
					iMessage = resultmsg
				End If

				If Err <> 0 Then
					If (IsAutoScript) Then
						rw "["&iitemid & "]:"&iMessage
					Else
						iErrStr =  "["&iitemid & "]:"&iMessage
					End If
					Set objXML = Nothing
				    Set xmlDOM = Nothing
				    On Error Goto 0
				    Exit Function
			    End If
			Set xmlDOM = Nothing
			GSShopOnItemItemnameEdit = True
		Else
			iErrStr = "GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-NMEDIT-002]"
	        Set objXML = Nothing
	        Set xmlDOM = Nothing
		    Exit Function
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'���ΰ���׸� ����
Function GSShopInfodivEdit(iitemid,strParam,byRef iErrStr)
    Dim objXML,xmlDOM,strRst,iMessage
    Dim buf
    Dim resultmsg, resultcode, supPrdCd, supCd, prdCd, attrPrdCd
	On Error Resume Next
	GSShopInfodivEdit = False
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strParam)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					rw buf
				End If

				If Err <> 0 Then
					iErrStr = "GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-DIVEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'��� �ڵ�
				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�

				If resultcode = "S" Then
					iErrStr = "["&iitemid & "]:"&resultmsg
				Else
					iMessage = resultmsg
				End If

				If Err <> 0 Then
					If (IsAutoScript) Then
						rw "["&iitemid & "]:"&iMessage
					Else
						iErrStr =  "["&iitemid & "]:"&iMessage
					End If
					Set objXML = Nothing
				    Set xmlDOM = Nothing
				    On Error Goto 0
				    Exit Function
			    End If
			Set xmlDOM = Nothing
			GSShopInfodivEdit = True
		Else
			iErrStr = "GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-DIVEDIT-002]"
	        Set objXML = Nothing
	        Set xmlDOM = Nothing
		    Exit Function
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'���û�ǰ ���� ����
Function getGSShopOnItemSellYNEditParameter(iitemid, ichgSellYn, byRef iErrStr)
    Dim strParam, resultcode, resultmsg, supPrdCd, supCd, prdCd
    Dim objXML, xmlDOM
    Dim strRst, strSql, notitemId
    getGSShopOnItemSellYNEditParameter = False
	strRst = ""
	strRst = strRst & "regGbn=U"														'(*)��ϱ��� U : ����
	strRst = strRst & "&modGbn=S"														'(*)�������� S : �ǸŻ��� ����
	strRst = strRst & "&regId="&COurRedId												'(*)�����
	'��ǰ�⺻(prdBaseInfo)
	strRst = strRst & "&supPrdCd="&iitemid												'(*)���»��ǰ�ڵ�
	strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
	'��ǰ����(prdPrc)
	strSql = ""
	strSql = "select count(*) as cnt from db_temp.dbo.tbl_jaehyumall_not_in_itemid where mallgubun = 'gsshop' and itemid =" & iitemid
	rsget.Open strSql, dbget
	If Not(rsget.EOF or rsget.BOF) Then
		notitemId = rsget("cnt")
	End If
	rsget.close

	If ichgSellYn = "Y" Then
		strRst = strRst & "&saleEndDtm=29991231235959"									'(*)�Ǹ������Ͻ� | ��ǰ�� �ߴ�(�Ǹ�����)�Ϸ��� �ߴܽ����� �Ǹ������Ͻø� �Է��մϴ�.
	ElseIf (ichgSellYn = "N") OR (notitemId > 0) Then
		strRst = strRst & "&saleEndDtm="&FormatDate(now(), "00000000000000")			'(*)�Ǹ������Ͻ� | ��ǰ�� �ߴ�(�Ǹ�����)�Ϸ��� �ߴܽ����� �Ǹ������Ͻø� �Է��մϴ�.
	End If
	strRst = strRst & "&attrSaleEndStModYn=N"											'(*)�Ӽ��Ǹ�������¼������� | �Ӽ�����(S) ��ǰ�ǸŻ��¸� ������ �� ����ϴ� �׸�����, ��ǰ������ ���� �� ���� �� �Ӽ���ǰ�� ���µ� �Բ� ���� �� �����Ϸ��� Y, ��ǰ�����Ϳ� �Ӽ� ������ ���º��� ���� �ÿ� N

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", gsshopAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		objXML.Send(strRst)
		If objXML.Status = "200" Then
			buf = BinaryToText(objXML.ResponseBody, "euc-kr")
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML buf

				If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
					rw buf
				End If

				If Err <> 0 Then
					iErrStr = "GSShop ��� �м� �߿� ������ �߻��߽��ϴ�.[ERR-SELLEDIT-001]"
				    Set objXML = Nothing
				    Set xmlDOM = Nothing
				    Exit Function
				End If
				'��� �ڵ�
				resultcode	= Split(buf, "|")(0)	'��ϰ���ڵ�
				resultmsg	= Split(buf, "|")(1)	'��ϰ���޽���
				supPrdCd	= Split(buf, "|")(2)	'���»��ǰ�ڵ�
				supCd		= Split(buf, "|")(3)	'���»��ڵ�
				prdCd		= Split(buf, "|")(4)	'��ǰ�ڵ�

				If resultcode <> "S" Then
					iMessage = resultmsg
				Else
					strSql = ""
					strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regItem " & VbCRLF
					strSql = strSql & " SET GSShopLastUpdate = getdate() " & VbCRLF
					strSql = strSql & " ,GSShopSellYn = '" & ichgSellYn & "'" & VbCRLF
					strSql = strSql & " WHERE itemid='" & iitemid & "'"
					dbget.Execute(strSql)
		        End If

				If Err <> 0 Then
					If (IsAutoScript) Then
						rw "["&iitemid & "]:"&iMessage
					Else
						iErrStr =  "["&iitemid & "]:"&iMessage
					End If
					Set objXML = Nothing
				    Set xmlDOM = Nothing
				    On Error Goto 0
				    Exit Function
			    End If
			Set xmlDOM = Nothing
			getGSShopOnItemSellYNEditParameter = True
		Else
			iErrStr = "GSShop�� ����߿� ������ �߻��߽��ϴ�..[ERR-SELLEDIT-002]"
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function
					''///////////////						�������� ���, �Ʒ����� ���μ��� 				//////////////////
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),20)
Dim arrItemid : arrItemid = request("cksel")
Dim oGSShop, i, strParam, iErrStr, ret1
Dim sqlStr, strSql, AssignedRow, SubNodes
Dim chgSellYn, actCnt, retErrStr
Dim buf, buf2, CNT10, CNT20, CNT30, iitemid
Dim ArrRows
Dim retFlag
Dim iMessage
dim iItemName, pregitemname

retFlag   = request("retFlag")
chgSellYn = request("chgSellYn")
If (cmdparam = "RegSelect") Then				''���û�ǰ ���� ���
    arrItemid = Trim(arrItemid)
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbget.Close: Response.End
	End If

	'## ���û�ǰ ��� ����
	Set oGSShop = new CGSShop
		oGSShop.FPageSize	= 20
		oGSShop.FRectItemID	= arrItemid
		oGSShop.getGSShopNotRegItemList
	    If (oGSShop.FResultCount < 1) Then
	        arrItemid = split(arrItemid,",")
	        For i = LBound(arrItemid) to UBound(arrItemid)
	            CALL Fn_AcctFailTouch("gsshop",arrItemid(i),"��ϰ��ɻ�ǰ ���� :������� Ȯ��: �Ǹ�Y, �ɼ��߰���...")
	        Next

	        If (IsAutoScript) Then
	            rw "S_ERR|��ϰ��ɻ�ǰ ���� :������� Ȯ��: �Ǹ�Y, ����..."
	            dbget.Close: Response.End
	        Else
	            Response.Write "<script language=javascript>alert('��ϰ��ɻ�ǰ ���� :������� Ȯ��: �Ǹ�Y, �ɼ��߰���...');</script>"
				dbget.Close: Response.End
			End If
		End If

		For i = 0 to (oGSShop.FResultCount - 1)
			If oGSShop.FItemList(i).FDivcode = "" Then		'���� ��ǰ�з� ��Ī�� �� �� ī�װ� ��ǰ�̶��..
				Response.Write "<script language=javascript>alert('��ǰ�з� ��Ī�� ���� ���� ��ǰ��ȣ: [" & oGSShop.FItemList(i).Fitemid & "]');</script>"
				dbget.Close: Response.End
			End If

			If oGSShop.FItemList(i).FDeliveryType = "9" OR oGSShop.FItemList(i).FDeliveryType = "7" OR oGSShop.FItemList(i).FDeliveryType = "2" Then
				If oGSShop.FItemList(i).FDeliveryCd = "" OR oGSShop.FItemList(i).FDeliveryAddrCd = "" Then
					Response.Write "<script language=javascript>alert('�ù��/�ּ��� ��Ī�� ���� ���� ��ǰ��ȣ: [" & oGSShop.FItemList(i).Fitemid & "]');</script>"
					dbget.Close: Response.End
				End If
			End If

			If oGSShop.FItemList(i).FBrandcd = "" Then
				Response.Write "<script language=javascript>alert('�귣���ڵ� ��Ī�� ���� ���� ��ǰ��ȣ: [" & oGSShop.FItemList(i).Fitemid & "]');</script>"
				dbget.Close: Response.End
			End If

			If (oGSShop.FItemList(i).FItemdiv = "06") OR (oGSShop.FItemList(i).FItemdiv = "16") Then
				Response.Write "<script language=javascript>alert('�ֹ����ۻ�ǰ ��ǰ��ȣ: [" & oGSShop.FItemList(i).Fitemid & "]');</script>"
				dbget.Close: Response.End
			End If

			sqlStr = ""
			sqlStr = sqlStr & " IF NOT Exists(SELECT * FROM db_item.dbo.tbl_gsshop_regitem where itemid="&oGSShop.FItemList(i).Fitemid&")"
			sqlStr = sqlStr & " BEGIN"& VbCRLF
			sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_gsshop_regitem "
	        sqlStr = sqlStr & " (itemid, regdate, reguserid, gsshopstatCD, regitemname)"
	        sqlStr = sqlStr & " VALUES ("&oGSShop.FItemList(i).Fitemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oGSShop.FItemList(i).FItemName)&"')"
			sqlStr = sqlStr & " END "
			dbget.Execute sqlStr
			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oGSShop.FItemList(i).checkTenItemOptionValid Then
			    On Error Resume Next
				'//��ǰ��� �Ķ����
				strParam = oGSShop.FItemList(i).getGSShopItemRegParameter()
				If (session("ssBctID") = "icommang") or (session("ssBctID") = "kjy8517") Then
					rw gsshopAPIURL &"?"& strParam
				End If
				If Err <> 0 Then
				    rw Err.Description
					Response.Write "<script language=javascript>alert('�ٹ����� ��ǰ���� ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oGSShop.FItemList(i).Fitemid & "]');</script>"
					dbget.Close: Response.End
				End If
	            On Error Goto 0

	            iErrStr = ""
				ret1 = GsshopOneItemReg(oGSShop.FItemList(i).FItemid, strParam, iErrStr, oGSShop.FItemList(i).FSellCash, oGSShop.FItemList(i).getGSShopSellYn, oGSShop.FItemList(i).FLimityn, oGSShop.FItemList(i).FLimitNo, oGSShop.FItemList(i).FLimitSold, html2db(oGSShop.FItemList(i).FItemName))
	            If (ret1) Then
	                actCnt = actCnt+1
	            Else
	                CALL Fn_AcctFailTouch("gsshop", oGSShop.FItemList(i).Fitemid, iErrStr)
	                retErrStr = retErrStr & iErrStr
	            End If
			Else
				CALL Fn_AcctFailTouch("gsshop", oGSShop.FItemList(i).Fitemid, iErrStr)
				iErrStr = "["&oGSShop.FItemList(i).Fitemid&"] �ɼǰ˻� ����"
				retErrStr = retErrStr & iErrStr
			End If
		Next
	Set oGSShop = Nothing

    If (retErrStr <> "") Then
        Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
    End If
ElseIf (cmdparam = "EditPrice") Then				''���û�ǰ ���� ����
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbget.Close: Response.End
	End If

	'## ������ ��ǰ ��� ����
	Set oGSShop = new CGSShop
		oGSShop.FPageSize	= 20
		oGSShop.FRectItemID	= arrItemid
		oGSShop.getGSShopEditedItemList

		For i = 0 to (oGSShop.FResultCount - 1)
			On Error Resume Next
'			If (oGSShop.FItemList(i).FSellcash <> oGSShop.FItemList(i).FGSShopPrice) Then
	            strParam = oGSShop.FItemList(i).getGSShopOnItemPriceEditParameter()
				If Err <> 0 Then
					Response.Write "<script language=javascript>alert('�ٹ����� PriceEditParameter ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oGSShop.FItemList(i).Fitemid & "]');</script>"
					dbget.Close: Response.End
				End If
				On Error Goto 0

	            ret1 = false
	            ret1 = GSShopOnItemPriceEdit(oGSShop.FItemList(i).Fitemid,strParam,iErrStr)

				If (ret1) Then
				    '// ��ǰ�������� ����
				    strSql = ""
	    			strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regItem  " & VbCRLF
	    			strSql = strSql & "	SET GSShopLastUpdate=getdate() " & VbCRLF
	    			strSql = strSql & "	, GSShopPrice = " & oGSShop.FItemList(i).MustPrice & VbCRLF
	    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
	    			strSql = strSql & " Where itemid='" & oGSShop.FItemList(i).Fitemid & "'"& VbCRLF
	    			dbget.Execute(strSql)
	    			actCnt = actCnt+1
	    			rw iErrStr
	            Else
	                CALL Fn_AcctFailTouch("gsshop",oGSShop.FItemList(i).Fitemid,iErrStr)
	                rw "[���ݼ�������]"&iErrStr
				End If
'			Else
'				rw "["&oGSShop.FItemList(i).Fitemid&"] : GSShop���ݰ� �ٹ����� ������ �����Ƿ� �������� �ʽ��ϴ�."
'			End If
		Next
	Set oGSShop = Nothing

	IF (session("ssBctID")="icommang") then
		response.end
	End If
ElseIf (cmdparam = "EditImage") Then				''���û�ǰ �̹���(��ǥ �� �����) ����
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbget.Close: Response.End
	End If

	'## ������ ��ǰ ��� ����
	Set oGSShop = new CGSShop
		oGSShop.FPageSize	= 20
		oGSShop.FRectItemID	= arrItemid
		oGSShop.getGSShopEditedItemList

		For i = 0 to (oGSShop.FResultCount - 1)
			On Error Resume Next
            strParam = oGSShop.FItemList(i).getGSShopOneItemImageEditParameter()

			If Err <> 0 Then
				Response.Write "<script language=javascript>alert('�ٹ����� PriceEditParameter ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oGSShop.FItemList(i).Fitemid & "]');</script>"
				dbget.Close: Response.End
			End If
			On Error Goto 0

            ret1 = false
            ret1 = GSShopOneItemImageEdit(oGSShop.FItemList(i).Fitemid,strParam,iErrStr)

			If (ret1) Then
    			actCnt = actCnt+1
    			rw iErrStr
            Else
                CALL Fn_AcctFailTouch("gsshop",oGSShop.FItemList(i).Fitemid,iErrStr)
                rw "[�̹�����������]"&iErrStr
			End If
		Next
	Set oGSShop = Nothing

	IF (session("ssBctID")="icommang") then
		response.end
	End If
ElseIf (cmdparam = "EditContents") Then				''��ǰ���� ����
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbget.Close: Response.End
	End If

	'## ������ ��ǰ ��� ����
	Set oGSShop = new CGSShop
		oGSShop.FPageSize	= 20
		oGSShop.FRectItemID	= arrItemid
		oGSShop.getGSShopEditedItemList

		For i = 0 to (oGSShop.FResultCount - 1)
			On Error Resume Next
            strParam = oGSShop.FItemList(i).getGSShopOneItemContentsEditParameter()

			If Err <> 0 Then
				Response.Write "<script language=javascript>alert('�ٹ����� PriceEditParameter ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oGSShop.FItemList(i).Fitemid & "]');</script>"
				dbget.Close: Response.End
			End If
			On Error Goto 0

            ret1 = false
            ret1 = GSShopOneItemContentsEdit(oGSShop.FItemList(i).Fitemid,strParam,iErrStr)

			If (ret1) Then
    			actCnt = actCnt+1
    			rw iErrStr
            Else
                CALL Fn_AcctFailTouch("gsshop",oGSShop.FItemList(i).Fitemid,iErrStr)
                rw "[��ǰ ���� ��������]"&iErrStr
			End If
		Next
	Set oGSShop = Nothing

	IF (session("ssBctID")="icommang") then
		response.end
	End If
ElseIf (cmdparam = "EditOPT") Then					''����&���&�ɼ�&���� ����
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbget.Close: Response.End
	End If

	'## ������ ��ǰ ��� ����
	Set oGSShop = new CGSShop
		oGSShop.FPageSize	= 20
		oGSShop.FRectItemID	= arrItemid
		oGSShop.getGSShopEditedItemList

		For i = 0 to (oGSShop.FResultCount - 1)
			On Error Resume Next
			'2014-12-02 18:49:00 ������ �߰�
			'1.�ɼ� �߰� �ݾ��ִ� ��ǰ�� ǰ���� ������ ����
			'2.�ɼ��ִ� ��ǰ���� ��ǰ���� ����Ǿ��� �� ǰ���� ����
			strSql = ""
			strSql = strSql & " SELECT COUNT(*) as cnt FROM db_item.dbo.tbl_item_option WHERE itemid = '"& oGSShop.FItemList(i).Fitemid &"' and isusing='Y' and optAddPrice > 0 "
			rsget.Open strSql,dbget,1
			If rsget("cnt") > 0 Then
				oGSShop.FItemList(i).FmaySoldOut = "Y"
			ElseIf oGSShop.FItemList(i).FOptionCnt = 0 and oGSShop.FItemList(i).FregedOptCnt > 0 Then
				oGSShop.FItemList(i).FmaySoldOut = "Y"
			End If
			rsget.Close
			'2014-12-02 18:49:00 ������ �߰� ��

			If (oGSShop.FItemList(i).FmaySoldOut = "Y") OR (oGSShop.FItemList(i).IsSoldOutLimit5Sell) Then
				iErrStr = ""
				chgSellYn = "N"
				If (getGSShopOnItemSellYNEditParameter(oGSShop.FItemList(i).Fitemid, chgSellYn, iErrStr)) Then
					actCnt = actCnt+1
					rw "["&oGSShop.FItemList(i).Fitemid&"]"&"ǰ��ó��"
				Else
					rw "["&oGSShop.FItemList(i).Fitemid&"]"&iErrStr
				End if
				strSql = ""
				strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regItem SET lastStatCheckDate=getdate(),GSShopLastUpdate=getdate() WHERE itemid = '"&oGSShop.FItemList(i).Fitemid&"' "
				dbget.Execute strSql
			Else
				If (oGSShop.FItemList(i).FGsshopSellYn = "N" AND oGSShop.FItemList(i).IsSoldOut = False) Then
					iErrStr = ""
					chgSellYn = "Y"
					If (getGSShopOnItemSellYNEditParameter(oGSShop.FItemList(i).Fitemid, chgSellYn, iErrStr)) Then
						rw "["&oGSShop.FItemList(i).Fitemid&"]"&"�Ǹ������� ����"
					End if
					strSql = ""
					strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regItem SET lastStatCheckDate=getdate() WHERE itemid = '"&oGSShop.FItemList(i).Fitemid&"' "
					dbget.Execute strSql
				End If

				If (oGSShop.FItemList(i).FSellcash <> oGSShop.FItemList(i).FGSShopPrice) Then
					strParam = ""
		            strParam = oGSShop.FItemList(i).getGSShopOnItemPriceEditParameter()
					If Err <> 0 Then
						Response.Write "<script language=javascript>alert('�ٹ����� PriceEditParameter ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oGSShop.FItemList(i).Fitemid & "]');</script>"
						dbget.Close: Response.End
					End If
					On Error Goto 0
	
		            ret1 = false
		            ret1 = GSShopOnItemPriceEdit(oGSShop.FItemList(i).Fitemid,strParam,iErrStr)
	
					If (ret1) Then
					    '// ��ǰ�������� ����
					    strSql = ""
		    			strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regItem  " & VbCRLF
		    			strSql = strSql & "	SET GSShopLastUpdate=getdate() " & VbCRLF
		    			strSql = strSql & "	, GSShopPrice = " & oGSShop.FItemList(i).MustPrice & VbCRLF
		    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
		    			strSql = strSql & " Where itemid='" & oGSShop.FItemList(i).Fitemid & "'"& VbCRLF
		    			dbget.Execute(strSql)
		    			rw iErrStr
		            Else
		                CALL Fn_AcctFailTouch("gsshop",oGSShop.FItemList(i).Fitemid,iErrStr)
		                rw "[���ݼ�������]"&iErrStr
					End If
				End If

				strParam = ""
	            strParam = oGSShop.FItemList(i).getGSShopOptParameter()
				If Err <> 0 Then
					Response.Write "<script language=javascript>alert('�ٹ����� PriceEditParameter ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oGSShop.FItemList(i).Fitemid & "]');</script>"
					dbget.Close: Response.End
				End If
				On Error Goto 0

				'������ ���� �ɼǼ� 219, 181 ���ϰ�� timeout
				''rw strParam
			    ''response.end
			
	            ret1 = false
	            ret1 = GSShopOPTSuEdit(oGSShop.FItemList(i).Fitemid,strParam,iErrStr)
	            
				If (ret1) Then
				    rw "�ɼ� GSShopOPTSuEdit"&ret1
				    
					strSql = ""
					strSql = strSql & " UPDATE R " & VBCRLF
					strSql = strSql & " SET regedOptCnt = isNULL(T.regedOptCnt,0), accFailCnt = 0 " & VBCRLF
					strSql = strSql & " , GSShopLastUpdate = getdate() " & VBCRLF
					strSql = strSql & " FROM db_item.dbo.tbl_gsshop_regItem R " & VBCRLF
					strSql = strSql & " Join ( " & VBCRLF
					strSql = strSql & " 	SELECT R.itemid, count(*) as CNT, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt "
					strSql = strSql & " 	FROM db_item.dbo.tbl_gsshop_regItem R " & VBCRLF
					strSql = strSql & " 	JOIN db_item.dbo.tbl_OutMall_regedoption Ro on R.itemid = Ro.itemid and Ro.mallid = 'gsshop' and Ro.itemid = '"&oGSShop.FItemList(i).Fitemid&"' " & VBCRLF
					strSql = strSql & " 	GROUP BY R.itemid " & VBCRLF
					strSql = strSql & " ) T on R.itemid = T.itemid " & VBCRLF
					dbget.Execute strSql  ''',Assignedrow 
					
					''�α׸� �װ�, GSShopLastUpdate�� �ľ���
					rw iErrStr
				Else
					rw "["&oGSShop.FItemList(i).Fitemid&"]"&iErrStr
				End If
				strParam = ""
	            strParam = oGSShop.FItemList(i).getGSShopOptSellParameter()

				'�ɼ� �ǸŻ��� ����
	            ret1 = false
	            ret1 = GSShopOPTSellEdit(oGSShop.FItemList(i).Fitemid,strParam,iErrStr)
                
				If (ret1) Then
					rw "�ɼ� �ǸŻ��� ����"&iErrStr
	    			actCnt = actCnt+1
	            Else
	                CALL Fn_AcctFailTouch("gsshop",oGSShop.FItemList(i).Fitemid,iErrStr)
	                rw "[��ǰ �ɼ� �ǸŻ��� ��������]"&iErrStr
				End If
			End If
		Next
	Set oGSShop = Nothing

	IF (session("ssBctID")="icommang") then
		response.end
	End If
ElseIf (cmdparam = "EditItemname") Then				''���û�ǰ�� ����
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbget.Close: Response.End
	End If

	'## ������ ��ǰ ��� ����
	Set oGSShop = new CGSShop
		oGSShop.FPageSize	= 20
		oGSShop.FRectItemID	= arrItemid
		oGSShop.getGSShopEditedItemList

		For i = 0 to (oGSShop.FResultCount - 1)
			'//��ǰ��� �Ķ����
			On Error Resume Next
            strParam = oGSShop.FItemList(i).getGSShopOneItemnameEditParameter()

			If Err <> 0 Then
				Response.Write "<script language=javascript>alert('�ٹ����� EditItemnameParameter ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oGSShop.FItemList(i).Fitemid & "]');</script>"
				dbget.Close: Response.End
			End If
			On Error Goto 0

            ret1 = false
            ret1 = GSShopOnItemItemnameEdit(oGSShop.FItemList(i).Fitemid,strParam,iErrStr)

			If (ret1) Then
				'// ��ǰ�������� ����
				strSql = ""
				strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regItem " & VbCRLF
				strSql = strSql & " SET regitemname = B.itemname "& VbCRLF
				strSql = strSql & " FROM db_item.dbo.tbl_gsshop_regItem A "& VbCRLF
				strSql = strSql & " JOIN db_item.dbo.tbl_item B on A.itemid = B.itemid "& VbCRLF
				strSql = strSql & " WHERE A.itemid='" & oGSShop.FItemList(i).Fitemid & "'"& VbCRLF
				dbget.Execute(strSql)
				
				actCnt = actCnt+1
				rw iErrStr
			Else
                CALL Fn_AcctFailTouch("gsshop",oGSShop.FItemList(i).Fitemid,iErrStr)
                rw "[��ǰ�� ��������]"&iErrStr
			End If
		Next
	Set oGSShop = Nothing

	IF (session("ssBctID")="icommang") then
		response.end
	End If
ElseIf (cmdparam = "CheckItemNmAuto") Then				'������ ��ǰ�� ����
	Dim xitemid, xGSShopGoodNo, xItemName
	buf = ""
	CNT10 = 0
	strSql = ""
	strSql = strSql & " SELECT TOP 10 r.itemid, r.GSShopGoodNo, i.ItemName "
	strSql = strSql & "	FROM db_item.dbo.tbl_gsshop_regItem r "
	strSql = strSql & "	JOIN db_item.dbo.tbl_item i on r.itemid = i.itemid "
	strSql = strSql & "	WHERE r.regitemname is Not NULL "
	strSql = strSql & "	and (r.GSShopStatCd=3 OR r.GSShopStatCd=7) "
	strSql = strSql & "	and r.GSShopGoodNo is Not Null "
	strSql = strSql & "	and r.regitemname <> i.itemname"
	strSql = strSql & "	ORDER BY r.regdate DESC"
	rsget.Open strSql,dbget,1
	If not rsget.Eof Then
		ArrRows = rsget.getRows()
	End If
	rsget.close

	If isArray(ArrRows) Then
	    For i =0 To UBound(ArrRows,2)
	        iErrStr = ""
	        xitemid			= CStr(ArrRows(0,i))
	        xGSShopGoodNo	= CStr(ArrRows(1,i))
	        xItemName		= CStr(ArrRows(2,i))
	        buf = buf & xitemid & ","
	        If (xitemid <> "") Then
	        	On Error Resume Next
	            strParam = fnGetGSShopOneItemnameEditParameter(xitemid, xItemName)
				If Err <> 0 Then
					Response.Write "<script language=javascript>alert('�ٹ����� CheckItemNmAutoParameter ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oGSShop.FItemList(i).Fitemid & "]');</script>"
					dbget.Close: Response.End
				End If
				On Error Goto 0

				ret1 = false
				ret1 = GSShopOnItemItemnameEdit(xitemid, strParam, iErrStr)
				IF (ret1) THEN
				    pregitemname = ""
					strSql = ""
				    strSql = strSql & " SELECT isNULL(regitemname,'') as regitemname from db_item.dbo.tbl_gsshop_regItem "& VbCRLF
				    strSql = strSql & "	WHERE itemid='" & xitemid & "'"& VbCRLF
				    rsget.Open strSql,dbget,1
	                If not rsget.Eof Then
	                    pregitemname = rsget("regitemname")
	                End If
	                rsget.close
	
	    			strSql = ""
				    strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regItem  " & VbCRLF
	    			strSql = strSql & "	SET regitemname='" & html2db(xItemName) &"'"& VbCRLF
	    			strSql = strSql & " WHERE itemid='" & xitemid & "'"& VbCRLF
	    			dbget.Execute(strSql)
	    			CNT10 = CNT10+1
	
	    			If (pregitemname <> xItemName) Then
	    			    buf2 = buf2 & pregitemname & "::" & xItemName &"<br>"
	    			End If
	            Else
					CALL Fn_AcctFailTouch("gsshop", xitemid, iErrStr)
					rw "[��ǰ�� ��������]"&iErrStr
				End If
	        End If
	    Next
	End If

	If buf <> "" Then
		rw buf
	End If

	If buf2 <> "" Then
		rw buf2
	End If
	rw CNT10&"�� ��ǰ����� ����"
	response.end
ElseIf (cmdparam = "EditInfoDiv") Then				'���ΰ���׸� ����
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbget.Close: Response.End
	End If

	'## ������ ��ǰ ��� ����
	Set oGSShop = new CGSShop
		oGSShop.FPageSize	= 20
		oGSShop.FRectItemID	= arrItemid
		oGSShop.getGSShopEditedItemList

		For i = 0 to (oGSShop.FResultCount - 1)
			On Error Resume Next
            strParam = oGSShop.FItemList(i).getGSShopInfodivEditParameter()

			If Err <> 0 Then
				Response.Write "<script language=javascript>alert('�ٹ����� InfodivEditParameter ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oGSShop.FItemList(i).Fitemid & "]');</script>"
				dbget.Close: Response.End
			End If
			On Error Goto 0

			ret1 = false
			ret1 = GSShopInfodivEdit(oGSShop.FItemList(i).Fitemid,strParam,iErrStr)

			If (ret1) Then
				actCnt = actCnt+1
				rw iErrStr
			Else
				rw "[���ΰ���׸� ��������]"&iErrStr
			End If
		Next
	Set oGSShop = Nothing

	IF (session("ssBctID")="icommang") then
		response.end
	End If
ElseIf (cmdparam = "EditSellYn") Then				''�ǸŻ��� ����
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbget.Close: Response.End
	End If
	'## ������ ��ǰ ��� ����
	Set oGSShop = new CGSShop
		oGSShop.FPageSize	= 20
		oGSShop.FRectItemID	= arrItemid
		oGSShop.getGSShopEditedItemList

		If (chgSellYn="N") and (oGSShop.FResultCount < 1) and (arrItemid = "") Then
		    oGSShop.getGSShopreqExpireItemList
		End If

		For i = 0 to (oGSShop.FResultCount - 1)
		    iErrStr = ""
			If (getGSShopOnItemSellYNEditParameter(oGSShop.FItemList(i).Fitemid, chgSellYn, iErrStr)) Then
				actCnt = actCnt+1
			Else
				rw "["&iitemid&"]"&iErrStr
			End If
			retErrStr = retErrStr & iErrStr
		Next
		Set oGSShop = Nothing
		If (retErrStr<>"") Then
			Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
		End If
ElseIf (cmdparam = "sugiRegedoption") Then			''Ÿ�Ӿƿ����� ���� ����� Regedoption���
	Dim ckLimit, arrGSShopInfo
	ckLimit = request("ckLimit")
	If ckLimit = "" Then 
		Response.Write "<script language=javascript>alert('���� ���� ���� �� �����ϼ���');</script>"
		dbget.Close: Response.End
	End If
	
	strSql = ""
	strSql = strSql & " SELECT itemid, gsshopgoodno FROM db_item.dbo.tbl_gsshop_regitem WHERE itemid in ("&arrItemid&") " 
	rsget.Open strSql,dbget,1
		arrGSShopInfo = rsget.getrows()
	rsget.Close
	
	rw "--���� ����Ǵ� ������ �ƴմϴ�~!"
	For i = 0 To Ubound(arrGSShopInfo,2)
		If ckLimit = "N" Then
			rw "insert into db_item.dbo.tbl_outmall_regedoption (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastupdate) values "
			rw "('"&arrGSShopInfo(0,i)&"', '0000', 'gsshop', '"&arrGSShopInfo(1,i)&"001', '���ϻ�ǰ', 'Y', 'N', '999', '0', getdate())"&"<br>"
		ElseIf ckLimit = "Y" Then
			rw "insert into db_item.dbo.tbl_outmall_regedoption (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastupdate) values "
			rw "('"&arrGSShopInfo(0,i)&"', '0000', 'gsshop', '"&arrGSShopInfo(1,i)&"001', '���ϻ�ǰ', 'Y', 'Y', '220', '0', getdate())"&"<br>"'
		End If
	Next
	response.end
ElseIf (cmdparam = "EditStatCd") Then				''���δ�� -> ���οϷ� ���μ���
	Dim chgStatItemCode
	chgStatItemCode = request("chgStatItemCode")
	strSql = ""
	strSql = strSql & " UPDATE db_item.dbo.tbl_gsshop_regItem SET "
	strSql = strSql & " GSShopStatCd = '7' "
	strSql = strSql & " WHERE itemid = '"& chgStatItemCode &"' "
	dbget.Execute(strSql)
	actCnt = 1
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
<!-- #include virtual="/lib/db/dbclose.asp" -->