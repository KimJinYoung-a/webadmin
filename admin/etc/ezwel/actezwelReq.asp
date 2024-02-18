<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 600 %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/ezwel/ezwelcls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'#################################### ��������� �⺻ ���� Setting ####################################
Public ezwelAPIURL
Public postParam
Const dataHead = "<?xml version=""1.0"" encoding=""euc-kr"" standalone=""yes""?>"

IF application("Svr_Info") = "Dev" THEN
	ezwelAPIURL = "http://api.dev.ezwel.com/if/api/goodsInfoAPI.ez"
Else
	ezwelAPIURL = "http://api.ezwel.com/if/api/goodsInfoAPI.ez"
End if
postParam	= "cspCd="&cspCd&"&crtCd="&crtCd&"&dataSet="
'#####################################################################################################

'################################### ���� Function Setting  ##########################################
Function EzwelOneItemReg(iitemid, strParam, byRef iErrStr, iSellCash, iezwelSellYn, ilimityn, ilimitno, ilimitsold, iitemname, iimageNm)
	Dim xmlStr : xmlStr = strParam
	Dim objXML, xmlDOM, strSql, tenOptCnt
	Dim retCode, goodsCd, iMessage, AssignedRow
	If (xmlStr = "") Then
		EzwelOneItemReg = false
		Exit Function
    End If

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", ezwelAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
		objXML.send(postParam&xmlStr)

	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
'			response.write objXML.ResponseText
'			response.end
		If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
'			response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		End If

		goodsCd		= xmlDOM.getElementsByTagName("goodsCd").item(0).text
		retCode		= xmlDOM.getElementsByTagName("resultCode").item(0).text
		iMessage	= xmlDOM.getElementsByTagName("resultMsg").item(0).text

		If retCode = "200" Then		'����(200)
			strSql = "SELECT COUNT(itemid) FROM db_outmall.dbo.tbl_ezwel_regItem WHERE itemid='" & iitemid & "' and ezwelgoodno = '"&goodsCd&"'"
			rsCTget.Open strSql, dbCTget, 1
			If rsCTget(0) = 0 Then
				strSql = ""
				strSql = strSql & " UPDATE R" & VbCRLF
				strSql = strSql & "	Set ezwelLastUpdate = getdate() "  & VbCRLF
				strSql = strSql & "	, ezwelGoodNo = '" & goodsCd & "'"  & VbCRLF
				strSql = strSql & "	, ezwelPrice = " &iSellCash& VbCRLF
				strSql = strSql & "	, accFailCnt = 0"& VbCRLF
				strSql = strSql & "	, ezwelRegdate = isNULL(ezwelRegdate, getdate())"
				If (goodsCd <> "") Then
				    strSql = strSql & "	, ezwelstatCD = '7'"& VbCRLF					'��ϿϷ�(�ӽ�)
				Else
					strSql = strSql & "	, ezwelstatCD = '1'"& VbCRLF					'���۽õ�
				End If
				strSql = strSql & "	From db_outmall.dbo.tbl_ezwel_regItem R"& VbCRLF
				strSql = strSql & " Where R.itemid = '" & iitemid & "'"
				dbCTget.Execute(strSql)
			Else
				'// ���� -> �űԵ��
				strSql = ""
				strSql = strSql & " INSERT INTO db_outmall.dbo.tbl_ezwel_regItem "
				strSql = strSql & " (itemid, regitemname, reguserid, ezwelRegdate, ezwelLastUpdate, ezwelGoodNo, ezwelPrice, ezwelSellYn, ezwelStatCd, regImageName) VALUES " & VbCRLF
				strSql = strSql & " ('" & iitemid & "'" & VBCRLF
				strSql = strSql & " , '" & iitemname & "'" &_
				strSql = strSql & " , '" & session("ssBctId") & "'" &_
				strSql = strSql & " , getdate(), getdate()" & VBCRLF
				strSql = strSql & " , '" & goodsCd & "'" & VBCRLF
				strSql = strSql & " , '" & iSellCash & "'" & VBCRLF
				strSql = strSql & " , '" & iezwelSellYn & "'" & VBCRLF
				If (goodsCd <> "") Then
				    strSql = strSql & ",'7'"											'��ϿϷ�(�ӽ�)
				Else
				    strSql = strSql & ",'1'"											'���۽õ�
				End If
				strSql = strSql & " , '" & iimageNm & "'" & VBCRLF
				strSql = strSql & ")"
				dbCTget.Execute(strSql)
				actCnt = actCnt + 1
			End If
			rsCTget.Close

			strSql = ""
			strSql = strSql &  "SELECT count(*) as cnt "
			strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
			strSql = strSql & " WHERE itemid=" & iitemid
			strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
			rsget.Open strSql,dbget,1
				tenOptCnt = rsget("cnt")
			rsget.Close

			strSql = ""
			strSql = strSql & " UPDATE db_outmall.dbo.tbl_ezwel_regItem SET "
			strSql = strSql & " regedOptCnt = " & tenOptCnt
			strSql = strSql & " WHERE itemid = " & iitemid
			dbCTget.Execute strSql
			EzwelOneItemReg = true
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

Function EzwelOneItemEdit(iitemid, iEzwelGoodNo, byRef iErrStr, strParam, imustprice, isellyn, optMust)
	Dim xmlStr : xmlStr = strParam
	Dim objXML, xmlDOM, strSql, tenOptCnt
	Dim retCode, goodsCd, iMessage, AssignedRow, oMsg, ocount
	If (xmlStr = "") Then
		EzwelOneItemEdit = false
		Exit Function
    End If
'rw oEzwel.FItemList(i).isImageChanged
'rw oEzwel.FItemList(i).getBasicImage
'rw "=-=-="
'response.end
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", ezwelAPIURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
		objXML.send(postParam&xmlStr)

	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
'			response.write objXML.ResponseText
'			response.end
		If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
			'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		End If

		goodsCd		= xmlDOM.getElementsByTagName("goodsCd").item(0).text
		retCode		= xmlDOM.getElementsByTagName("resultCode").item(0).text
		iMessage	= xmlDOM.getElementsByTagName("resultMsg").item(0).text

		If retCode = "200" Then		'����(200)
			strSql = ""
			strSql = strSql & " UPDATE db_outmall.dbo.tbl_ezwel_regItem " & VbCRLF
			strSql = strSql & " SET accFailCnt = 0 " & VbCRLF
			strSql = strSql & " ,ezwelLastUpdate = getdate() " & VbCRLF
			strSql = strSql & " ,ezwelprice = '"&imustprice&"' " & VbCRLF
			strSql = strSql & " ,ezwelsellyn = '"&isellyn&"' " & VbCRLF
			If oEzwel.FItemList(i).isImageChanged Then
				strSql = strSql & " ,regImageName = '"&oEzwel.FItemList(i).getBasicImage&"' " & VbCRLF	
			End If
			strSql = strSql & " WHERE itemid='" & iitemid & "'"
			dbCTget.Execute(strSql)
			EzwelOneItemEdit = true
			If optMust = "all" Then
				strSql = ""
				strSql = strSql &  "SELECT count(*) as cnt "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " WHERE itemid=" & oEzwel.FItemList(i).Fitemid
				strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
				rsget.Open strSql,dbget,1
					ocount = rsget("cnt")
				rsget.Close
	
				strSql = ""
				strSql = strSql & " UPDATE db_outmall.dbo.tbl_ezwel_regItem SET "
				strSql = strSql & " regedOptCnt = " & ocount
				strSql = strSql & " WHERE itemid = " & oEzwel.FItemList(i).Fitemid
				dbCTget.Execute strSql
			End If

			Set objXML = Nothing
			Set xmlDOM = Nothing
			If isellyn = "N" Then
				oMsg = " | ǰ��ó��"
			End If
			rw "[" & iitemid & "]:" & iMessage & oMsg
		Else						'����(E)
			iErrStr =  "��ǰ ���� �� ���� [" & iitemid & "]:"&iMessage
			Set objXML = Nothing
			Set xmlDOM = Nothing
		    Exit Function
		End If
		On Error Goto 0
	End If
End Function
'#####################################################################################################
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),25)
Dim arrItemid : arrItemid = request("cksel")
Dim alertMsg, iMessage, actCnt, sqlStr, retErrStr
Dim oEzwel, i, strParam, iErrStr, ret1, chgSellYn, iitemid
Dim getMustprice, sellgubun, chkparam
chgSellYn = request("chgSellYn")
actCnt = 0

If (cmdparam = "RegSelect") Then				'���û�ǰ ���� ���
    arrItemid = Trim(arrItemid)
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbCTget.Close: Response.End
	End If

	'## ���û�ǰ ��� ����
	Set oEzwel = new CEzwel
		oEzwel.FPageSize	= 20
		oEzwel.FRectItemID	= arrItemid
		oEzwel.getEzwelNotRegItemList

	    If (oEzwel.FResultCount < 1) Then
	        arrItemid = split(arrItemid,",")
	        For i = LBound(arrItemid) to UBound(arrItemid)
	            CALL Fn_AcctFailTouch("ezwel",arrItemid(i),"��ϰ��ɻ�ǰ ���� :������� Ȯ��: �Ǹ�Y, �ɼ��߰���...")
	        Next

	        If (IsAutoScript) Then
	            rw "S_ERR|��ϰ��ɻ�ǰ ���� :������� Ȯ��: �Ǹ�Y, ����..."
	            dbCTget.Close: Response.End
	        Else
	            Response.Write "<script language=javascript>alert('��ϰ��ɻ�ǰ ���� :������� Ȯ��: �Ǹ�Y, �ɼ��߰���...');</script>"
				dbCTget.Close: Response.End
			End If
		End If

		For i = 0 to (oEzwel.FResultCount - 1)
			If (oEzwel.FItemList(i).FDepthCode = "") OR (oEzwel.FItemList(i).FdepthCode = "0") Then
				Response.Write "<script language=javascript>alert('ī�װ� ��Ī�� ���� ���� ��ǰ��ȣ: [" & oEzwel.FItemList(i).Fitemid & "]');</script>"
				dbCTget.Close: Response.End
			End If
			sqlStr = ""
			sqlStr = sqlStr & " IF NOT Exists(SELECT * FROM db_outmall.dbo.tbl_ezwel_regItem where itemid="&oEzwel.FItemList(i).Fitemid&")"
			sqlStr = sqlStr & " BEGIN"& VbCRLF
			sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_ezwel_regItem "
	        sqlStr = sqlStr & " (itemid, regdate, reguserid, ezwelstatCD, regitemname)"
	        sqlStr = sqlStr & " VALUES ("&oEzwel.FItemList(i).Fitemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oEzwel.FItemList(i).FItemName)&"')"
			sqlStr = sqlStr & " END "
			dbCTget.Execute sqlStr
			'##��ǰ�ɼ� �˻�(�ɼǼ��� ���� �ʰų� ��� ��ü ���ܿɼ��� ��� Pass)
			If oEzwel.FItemList(i).checkTenItemOptionValid Then
			    On Error Resume Next
				'//��ǰ��� �Ķ����
				strParam = oEzwel.FItemList(i).getEzwelItemRegXML("Reg")

				If Err <> 0 Then
				    rw Err.Description
					Response.Write "<script language=javascript>alert('�ٹ����� ��ǰ���� ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oEzwel.FItemList(i).Fitemid & "]');</script>"
					dbCTget.Close: Response.End
				End If

				On Error Goto 0
				iErrStr = ""
				ret1 = EzwelOneItemReg(oEzwel.FItemList(i).FItemid, strParam, iErrStr, oEzwel.FItemList(i).FSellCash, oEzwel.FItemList(i).getEzwelSellYn, oEzwel.FItemList(i).FLimityn, oEzwel.FItemList(i).FLimitNo, oEzwel.FItemList(i).FLimitSold, html2db(oEzwel.FItemList(i).FItemName), oEzwel.FItemList(i).FbasicimageNm)
				If (ret1) Then
					actCnt = actCnt+1
				Else
					CALL Fn_AcctFailTouch("ezwel", oEzwel.FItemList(i).Fitemid, iErrStr)
					retErrStr = retErrStr & iErrStr
					rw iErrStr
				End If
			Else
				CALL Fn_AcctFailTouch("ezwel", oEzwel.FItemList(i).Fitemid, iErrStr)
				iErrStr = "["&oEzwel.FItemList(i).Fitemid&"] : �ɼǰ˻� ���� | �����ɼ��� ������ 5�� ������ �� ����"
				retErrStr = retErrStr & iErrStr
			End If
		Next
	Set oEzwel = Nothing
    If (retErrStr <> "") Then
        Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
    End If
ElseIf (cmdparam = "EditSelect") Then				'���� ����
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbCTget.Close: Response.End
	End If
	'## ������ ��ǰ ��� ����
	Set oEzwel = new CEzwel
		oEzwel.FPageSize	= 20
		oEzwel.FRectItemID	= arrItemid
		oEzwel.getEzwelEditedItemList
		For i = 0 to (oEzwel.FResultCount - 1)
			On Error Resume Next
			strParam = ""
			chkparam = ""
			iErrStr = ""
			If (oEzwel.FItemList(i).FmaySoldOut = "Y") OR (oEzwel.FItemList(i).IsSoldOutLimit5Sell) Then
				strParam = oEzwel.FItemList(i).getEzwelItemRegXML("SellN")
				chgSellYn = "N"
			Else
				strParam = oEzwel.FItemList(i).getEzwelItemRegXML("SellY")
				chgSellYn = "Y"
			End If
			getMustprice = ""
			getMustprice = oEzwel.FItemList(i).fngetMustPrice()
			'*********************************************************************************************************************************************************
			'2014-11-06 ������ | dev_Comment
			'API�� ���۵Ǵ� ���� ��ǰ�ɼ��� �ν����� ���� | ��ϵ� �ɼ�ī��Ʈ�� ũ�ٸ� 10x10���� �ɼ� ������ ���� �������
			'�ᱹ �������� �ɼǻ��������� ������ �ɼ��� �ʱ�ȭ ���� �߰�
			'�ɼ� �������� API���� �� �ѹ� �� �ɼ��ִ� ������·� �����ϸ� ���ϴ� �����Ͱ� Ȯ�� ��.
			'�߰� : �ι� API���۽� ���� Ȯ���� ������ �� | �Ƹ� ��������� DB�� ��ǰ���� �����ϴ� �� ���� �ɷ��ִ� �� ��..
			'		���� �켱 �̷� ��ǰ�� ǰ���� ���ѵξ��ٰ� �������� �׸�鸸 ����� �����ϴ� �ϴ� ��ġ�� �ʿ��� ����.
			'�߰�_������(2015-03-04) 1173474 4:2 �̷������̶� If CInt(rsCTget("optioncnt")) <> CInt(rsCTget("regedoptcnt")) Then �̷� ������ ����
			sqlStr = ""
			sqlStr = sqlStr &  "SELECT top 1 r.itemid, i.optioncnt, r.regedoptcnt "
			sqlStr = sqlStr & " FROM db_Appwish.dbo.tbl_item as i "
			sqlStr = sqlStr & " join db_outmall.dbo.tbl_ezwel_regitem as r on i.itemid=r.itemid "
			sqlStr = sqlStr & " WHERE i.itemid=" & oEzwel.FItemList(i).Fitemid
			rsCTget.Open sqlStr,dbCTget,1
			If not rsCTget.EOF Then
				If CInt(rsCTget("optioncnt")) > 0 Then
					If CInt(rsCTget("optioncnt")) <> CInt(rsCTget("regedoptcnt")) Then
						chkparam = oEzwel.FItemList(i).getEzwelItemOptZeroNotScheduleXML("SellN")
						Call EzwelOneItemEdit(oEzwel.FItemList(i).Fitemid, oEzwel.FItemList(i).FEzwelGoodNo, iErrStr, chkparam, getMustprice, "N", "optMustN")
					End If
				End If
			End If
			rsCTget.Close
			'*********************************************************************************************************************************************************
			iErrStr = ""
			ret1 = EzwelOneItemEdit(oEzwel.FItemList(i).Fitemid, oEzwel.FItemList(i).FEzwelGoodNo, iErrStr, strParam, getMustprice, chgSellYn, "all")
			If (ret1) Then
				actCnt = actCnt+1
			Else
				CALL Fn_AcctFailTouch("ezwel", oEzwel.FItemList(i).Fitemid, iErrStr)
				retErrStr = retErrStr & iErrStr
				rw iErrStr
			End If

			If Err <> 0 Then
			    rw Err.Description
				Response.Write "<script language=javascript>alert('�ٹ����� ��ǰ���� ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oEzwel.FItemList(i).Fitemid & "]');</script>"
				dbCTget.Close: Response.End
			End If
			On Error Goto 0
			retErrStr = retErrStr & iErrStr
		Next
	Set oEzwel = Nothing
	If (retErrStr<>"") Then
		Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
	End If
ElseIf (cmdparam = "EditSelectNotSchedule") Then				'�����ٸ� ���N(������ ���� ��ư) | �̹���, ��ǰ������ �� ����
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbCTget.Close: Response.End
	End If
	'## ������ ��ǰ ��� ����
	Set oEzwel = new CEzwel
		oEzwel.FPageSize	= 20
		oEzwel.FRectItemID	= arrItemid
		oEzwel.getEzwelEditedItemList
		For i = 0 to (oEzwel.FResultCount - 1)
			On Error Resume Next
			strParam = ""
			chkparam = ""
			iErrStr = ""
			If (oEzwel.FItemList(i).FmaySoldOut = "Y") OR (oEzwel.FItemList(i).IsSoldOutLimit5Sell) Then
				strParam = oEzwel.FItemList(i).getEzwelItemEditNotScheduleXML("SellN")
				chgSellYn = "N"
			Else
				strParam = oEzwel.FItemList(i).getEzwelItemEditNotScheduleXML("SellY")
				chgSellYn = "Y"
			End If

			getMustprice = ""
			getMustprice = oEzwel.FItemList(i).fngetMustPrice()
			'*********************************************************************************************************************************************************
			'2014-11-06 ������ | dev_Comment
			'API�� ���۵Ǵ� ���� ��ǰ�ɼ��� �ν����� ���� | ��ϵ� �ɼ�ī��Ʈ�� ũ�ٸ� 10x10���� �ɼ� ������ ���� �������
			'�ᱹ �������� �ɼǻ��������� ������ �ɼ��� �ʱ�ȭ ���� �߰�
			'�ɼ� �������� API���� �� �ѹ� �� �ɼ��ִ� ������·� �����ϸ� ���ϴ� �����Ͱ� Ȯ�� ��.
			'�߰� : �ι� API���۽� ���� Ȯ���� ������ �� | �Ƹ� ��������� DB�� ��ǰ���� �����ϴ� �� ���� �ɷ��ִ� �� ��..
			'		���� �켱 �̷� ��ǰ�� ǰ���� ���ѵξ��ٰ� �������� �׸�鸸 ����� �����ϴ� �ϴ� ��ġ�� �ʿ��� ����.
			'�߰�_������(2015-03-04) 1173474 4:2 �̷������̶� If CInt(rsCTget("optioncnt")) <> CInt(rsCTget("regedoptcnt")) Then �̷� ������ ����
			sqlStr = ""
			sqlStr = sqlStr &  "SELECT top 1 r.itemid, i.optioncnt, r.regedoptcnt "
			sqlStr = sqlStr & " FROM db_Appwish.dbo.tbl_item as i "
			sqlStr = sqlStr & " join db_outmall.dbo.tbl_ezwel_regitem as r on i.itemid=r.itemid "
			sqlStr = sqlStr & " WHERE i.itemid=" & oEzwel.FItemList(i).Fitemid
			rsCTget.Open sqlStr,dbCTget,1
			If not rsCTget.EOF Then
				If CInt(rsCTget("optioncnt")) > 0 Then
					If CInt(rsCTget("optioncnt")) <> CInt(rsCTget("regedoptcnt")) Then
						chkparam = oEzwel.FItemList(i).getEzwelItemOptZeroNotScheduleXML("SellN")
						Call EzwelOneItemEdit(oEzwel.FItemList(i).Fitemid, oEzwel.FItemList(i).FEzwelGoodNo, iErrStr, chkparam, getMustprice, "N", "optMustN")
					End If
				End If
			End If
			rsCTget.Close
			'*********************************************************************************************************************************************************
			iErrStr = ""
			ret1 = EzwelOneItemEdit(oEzwel.FItemList(i).Fitemid, oEzwel.FItemList(i).FEzwelGoodNo, iErrStr, strParam, getMustprice, chgSellYn, "all")
			If (ret1) Then
				actCnt = actCnt+1
			Else
				CALL Fn_AcctFailTouch("ezwel", oEzwel.FItemList(i).Fitemid, iErrStr)
				retErrStr = retErrStr & iErrStr
				rw iErrStr
			End If

			If Err <> 0 Then
			    rw Err.Description
				Response.Write "<script language=javascript>alert('�ٹ����� ��ǰ���� ������ ������ �߻��߽��ϴ�.\n�����ڿ��� ���� ��Ź�帳�ϴ�.[��ǰ��ȣ:" & oEzwel.FItemList(i).Fitemid & "]');</script>"
				dbCTget.Close: Response.End
			End If
			On Error Goto 0
			retErrStr = retErrStr & iErrStr
		Next
	Set oEzwel = Nothing
	If (retErrStr<>"") Then
		Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
	End If
ElseIf (cmdparam = "EditSellYn") Then				'�ǸŻ��� ����
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');</script>"
		dbCTget.Close: Response.End
	End If
	'## ������ ��ǰ ��� ����
	Set oEzwel = new CEzwel
		If (session("ssBctID")="kjy8517") Then
			oEzwel.FPageSize	= 100
		Else
			oEzwel.FPageSize	= 20
		End If
		oEzwel.FRectItemID	= arrItemid
		oEzwel.getEzwelEditedItemList

'		If (chgSellYn="N") and (oEzwel.FResultCount < 1) and (arrItemid = "") Then
'		    oEzwel.getEzwelreqExpireItemList
'		End If

		If chgSellYn = "N" Then
			sellgubun = "SellN"
		Else
			sellgubun = "SellY"
		End If

		For i = 0 to (oEzwel.FResultCount - 1)
			strParam = oEzwel.FItemList(i).getEzwelItemRegXML(sellgubun)
			getMustprice = ""
			getMustprice = oEzwel.FItemList(i).fngetMustPrice()
		    iErrStr = ""
			If EzwelOneItemEdit(oEzwel.FItemList(i).Fitemid, oEzwel.FItemList(i).FEzwelGoodNo, iErrStr, strParam, getMustprice, chgSellYn, "all") Then
				actCnt = actCnt+1
			Else
				rw "["&iitemid&"]"&iErrStr
			End If
			retErrStr = retErrStr & iErrStr
		Next
	Set oEzwel = Nothing
	If (retErrStr<>"") Then
		Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
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