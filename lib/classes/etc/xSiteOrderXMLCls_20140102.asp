<%

'' !!!! �Ʒ� ������ ��Ŭ��� �Ǿ� �־�� �Ѵ�.
''/admin/etc/lotte/inc_dailyAuthCheck.asp
''/lib/classes/etc/lotteitemcls.asp
''/admin/etc/incOutMallCommonFunction.asp

Class CxSiteOrderXML
    public FItemList()
	public FOneItem
	public FResultCount
	public FTotalCount

	public FRectSellSite
	public FRectYYYYMMDD
	public FRectStartYYYYMMDD
	public FRectEndYYYYMMDD

	public FRectGubun

	public FRectAPIURL
	public FRectAuthNo

	public ErrMsg
	private objXML
	private xmlDOM

	private xmlURL
	private objData

	public function GetCheckStatus(byRef LastCheckDate, byRef isSuccess)
		dim strSql

        ''2013/11/20 ���� �����ʹ� �ѹ��� ���� ���°ɷ�
        strSql = " IF Exists("
        strSql = strSql + " 	select LastcheckDate"
        strSql = strSql + " 	from db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp]"
        strSql = strSql + " 	where dateDiff(d,LastUpdate,getdate())>0"
        strSql = strSql + " 	and sellsite='" + CStr(FRectSellSite) + "'"
        strSql = strSql + " )"
        strSql = strSql + " BEGIN"
        strSql = strSql + " 	Update T"
        strSql = strSql + " 	set LastcheckDate=dateadd(d,-1, LastcheckDate)"
        strSql = strSql + " 	from db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] T"
        strSql = strSql + " 	where sellsite='" + CStr(FRectSellSite) + "'"
        strSql = strSql + " END"
        dbget.Execute strSql

		strSql = " select LastCheckDate, isSuccess from db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] "
		strSql = strSql + " where sellsite = '" + CStr(FRectSellSite) + "' "

    	rsget.Open strSql,dbget,1
			LastCheckDate = rsget("LastCheckDate")
			isSuccess = rsget("isSuccess")
		rsget.Close

	end function

	public function SetCheckStatusStarting(LastCheckDate)
		dim strSql

		strSql = " update db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] "
		strSql = strSql + " set LastCheckDate = '" + CStr(LastCheckDate) + "', isSuccess = 'N' "
		strSql = strSql + " where sellsite = '" + CStr(FRectSellSite) + "' "
    	rsget.Open strSql,dbget,1

	end function

	public function SetCheckStatusEnded()
		dim strSql

		strSql = " update db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] "
		strSql = strSql + " set isSuccess = 'Y' "
		strSql = strSql + " ,LastUpdate=getdate()"
		strSql = strSql + " where sellsite = '" + CStr(FRectSellSite) + "' "
    	rsget.Open strSql,dbget,1

	end function

	public function SetCheckDate(LastCheckDate)
		dim strSql

		strSql = " update db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] "
		strSql = strSql + " set LastCheckDate = '" + CStr(LastCheckDate) + "' "
		strSql = strSql + " where sellsite = '" + CStr(FRectSellSite) + "' "
    	rsget.Open strSql,dbget,1

	end function

	public function SavexSiteOrderListtoDB()
		ErrMsg = ""

		if (ErrMsg = "") then
			xmlURL = GetXMLURL()
			response.write xmlURL

			if (xmlURL = "") and (ErrMsg = "") then
				ErrMsg = "��ϵ��� ���� ���޸��Դϴ�.[0]"
			end if
		end if
'dbget.close() : response.end
		if (ErrMsg = "") then
			Call GetXmlFromWeb()

			if (objData = "") and (ErrMsg = "") then
				ErrMsg = "�Ե����İ� ����߿� ������ �߻��߽��ϴ�."
			end if
		end if

		if (ErrMsg = "") then
			Call ActSavexSiteOrderListtoDB()
		end if

    end function

	function ActSavexSiteOrderListtoDB()
		dim i, j
		dim objMasterListXML, objMasterOneXML, objDetailListXML, objDetailOneXML
		dim masterCnt, detailCnt
		dim SellSite, OutMallOrderSerial, SellDate, PayType, Paydate, partnerItemID, partnerItemName, partnerOption, partnerOptionName, OrderUserID, OrderName, OrderEmail, OrderTelNo, OrderHpNo
		dim ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey, deliverymemo, requireDetail, orderDlvPay, etc1, countryCode
		dim matchItemID,matchItemOption, orderCsGbn, outMallGoodsNo
		dim strSql
		dim errCode, errStr

		dim tmpStr, isCSOrder
        dim retVal, succCNT,failCNT

        succCNT=0
        failCNT=0
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		xmlDOM.LoadXML replace(objData,"&","��")

		if (FRectSellSite = "lotteimall") then
			'// lotteimall

			'// ī��Ʈ
			masterCnt = (xmlDOM.selectNodes("/Response/Result/OrderCount").item(0).text * 1)

			response.write "(" & masterCnt & "��)"

			if (masterCnt > 0) then
				set objMasterListXML = xmlDOM.selectNodes("/Response/Result/OrderInfo")
				masterCnt = objMasterListXML.length

				for i = 0 to masterCnt - 1
					set objMasterOneXML = objMasterListXML.item(i)

					SellSite 			= FRectSellSite
					OutMallOrderSerial	= objMasterOneXML.selectSingleNode("OrdNo").text
					SellDate			= objMasterOneXML.selectSingleNode("TrdDate").text
					PayType				= "50"
					Paydate				= objMasterOneXML.selectSingleNode("TrdDate").text
					OrderUserID			= ""
					OrderName			= objMasterOneXML.selectSingleNode("OrderName").text
					OrderEmail			= ""
					OrderTelNo			= objMasterOneXML.selectSingleNode("OrderTelNo").text
					OrderHpNo			= objMasterOneXML.selectSingleNode("OrderHpNo").text

					ReceiveName			= objMasterOneXML.selectSingleNode("DelvInfo/recvName").text
					ReceiveTelNo		= objMasterOneXML.selectSingleNode("DelvInfo/recvTel").text
					ReceiveHpNo			= objMasterOneXML.selectSingleNode("DelvInfo/recvHp").text
					ReceiveZipCode		= objMasterOneXML.selectSingleNode("DelvInfo/recvPostCode").text
					if Len(ReceiveZipCode) = 6 then
						'
					end if
					ReceiveAddr1		= objMasterOneXML.selectSingleNode("DelvInfo/recvAddr1").text
					ReceiveAddr2		= objMasterOneXML.selectSingleNode("DelvInfo/recvAddr2").text

					deliverymemo		= objMasterOneXML.selectSingleNode("DlvMemoCont").text
					if (deliverymemo = "null") then
						deliverymemo = ""
					end if

					etc1				= ""
					countryCode			= "KR"

					'// ������
					set objDetailListXML = objMasterOneXML.selectNodes("ProdInfo")
					detailCnt = objDetailListXML.length
					for j = 0 to detailCnt - 1
						set objDetailOneXML = objDetailListXML.item(j)

						OrgDetailKey		= objDetailOneXML.selectSingleNode("ProdSeq").text

						SellPrice			= objDetailOneXML.selectSingleNode("ordPrice").text

						''RealSellPrice		= objDetailOneXML.selectSingleNode("buyRealPrice").text
						RealSellPrice		= objDetailOneXML.selectSingleNode("ordPrice").text ''2013/07/01 ����

						ItemOrderCount		= objDetailOneXML.selectSingleNode("ordQty").text

						requireDetail		= objDetailOneXML.selectSingleNode("GoodsChocDesc").text
						if (requireDetail = "null") then
							requireDetail = ""
						end if

						orderDlvPay			= 0

						tmpStr = Split(objDetailOneXML.selectSingleNode("CorpItemNo").text, "_")
						matchItemID = ""        ''�ʱ�ȭ 2013/08/19 �߰�
						matchItemOption = ""    ''�ʱ�ȭ 2013/08/19 �߰�
						IF IsArray(tmpStr) then

						    if (Ubound(tmpStr)>=0) then
        						matchItemID			= tmpStr(0)
        					end if

        					if (Ubound(tmpStr)>=1) then
    						    matchItemOption		= tmpStr(1)
    						end if
                            'rw Ubound(tmpStr)
                            'rw matchItemID
                            'rw matchItemOption
    						if (matchItemID="null") then matchItemID=""
                        end if

                        if (matchItemID="") then
                            matchItemID =  objDetailOneXML.selectSingleNode("EntrProdNo").text

                        end if


						partnerItemName		= objDetailOneXML.selectSingleNode("ProdName").text
						if (matchItemID="791471") then
						    partnerItemName     = "[�������] �̲��� ���� Ʃ���Ʈ Ư�� 100 x 120cm 6�� ��1"'replace(partnerItemName,"&nbsp;","")
						elseif (matchItemID="635526") then
						    partnerItemName = replace(partnerItemName,"&nbsp;"," ")
						    partnerItemName = replace(partnerItemName,"��nbsp;"," ")
						    partnerItemName = replace(partnerItemName,"   "," ")
						    partnerItemName = replace(partnerItemName,"   "," ")
                            partnerItemName = replace(partnerItemName,"  "," ")

					    end if

						partnerOptionName	= Trim(objDetailOneXML.selectSingleNode("prodOption").text)

						outMallGoodsNo		= objDetailOneXML.selectSingleNode("ProdCode").text

						if (partnerOptionName = "null") then
							partnerOptionName = ""
						end if

						partnerItemID		= objDetailOneXML.selectSingleNode("ProdCode").text
						partnerOption		= ""

						if (matchItemOption="") and (partnerOptionName="") then ''2013/07/30 �߰�
						    matchItemOption="0000"
						end if

                        if (matchItemOption="") then
                            rw "<font color=red>["&matchItemID&"]�ɼǸ�Ī ����:["&partnerOptionName&"]</font>"
                            matchItemOption="0000"
                        end if
'                rw "partnerItemID"&partnerItemID
'                rw "partnerOption"&partnerOption
'                rw "matchItemID"&matchItemID
'                rw "matchItemOption"&matchItemOption

						'// CS�������
						tmpStr = objDetailOneXML.selectSingleNode("Exchange").text
						isCSOrder = (tmpStr <> "�Ϲ�")

                        orderCsGbn=""               ''�ʱ�ȭ 2013/08/19 �߰�
						if (Not isCSOrder) then
							orderCsGbn = "0"
						else
							orderCsGbn = "3"
						end if

						''response.write "<br>" & matchItemID & partnerItemName

						retVal= saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName,outMallGoodsNo _
								, OrderName, OrderTelNo, OrderHpNo _
								, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
								, SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
								, deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
								, errCode, errStr )

                        if (retVal) then
                            succCNT=succCNT+1
                        else
                            failCNT=failCNT+1
                        end if
						set objDetailOneXML = Nothing
					next

					set objDetailListXML = Nothing
					set objMasterOneXML = Nothing
				next
			end if

			set objMasterListXML = Nothing

			strSql = " update c "
			strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
			strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
			strSql = strSql + " from "
			strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
			strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
			strSql = strSql + " on "
			strSql = strSql + " 	1 = 1 "
			strSql = strSql + " 	and c.SellSite = o.SellSite and c.sellsite='"&FRectSellSite&"'"
			strSql = strSql + " 	and c.OutMallOrderSerial = Replace(o.OutMallOrderSerial, '-', '') "
			strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
			strSql = strSql + " where "
			strSql = strSql + " 	1 = 1 "
			strSql = strSql + " 	and c.orderserial is NULL "
			strSql = strSql + " 	and o.orderserial is not NULL "
			''rw strSql
			''rsget.Open strSql, dbget, 1
		ElseIf (FRectSellSite = "lotteCom") then
			'// ī��Ʈ
			on Error resume next
			masterCnt = (xmlDOM.selectNodes("/Response/Result/OrderCount").item(0).text * 1)
			on Error Goto 0
			If masterCnt = "" Then
				rw "�ش�Ⱓ�� �´� �����Ͱ� �����ϴ�"
				Exit Function
			Else
				rw "(" & masterCnt & "��)"
			End If

			If (masterCnt > 0) Then
				set objMasterListXML = xmlDOM.selectNodes("/Response/Result/OrderInfo")
				masterCnt = objMasterListXML.length

				For i = 0 to masterCnt - 1
					Set objMasterOneXML = objMasterListXML.item(i)
					SellSite 			= FRectSellSite
					OutMallOrderSerial	= objMasterOneXML.selectSingleNode("OrdNo").text
					SellDate			= objMasterOneXML.selectSingleNode("TrdDate").text
					PayType				= "50"
					Paydate				= objMasterOneXML.selectSingleNode("TrdDate").text
					OrderUserID			= ""
					OrderName			= objMasterOneXML.selectSingleNode("OrderName").text
					OrderEmail			= ""
					OrderTelNo			= objMasterOneXML.selectSingleNode("OrderTelNo").text
					OrderHpNo			= objMasterOneXML.selectSingleNode("OrderHpNo").text
					ReceiveName			= objMasterOneXML.selectSingleNode("DelvInfo/recvName").text
					ReceiveTelNo		= objMasterOneXML.selectSingleNode("DelvInfo/recvTel").text
					ReceiveHpNo			= objMasterOneXML.selectSingleNode("DelvInfo/recvHp").text
					ReceiveZipCode		= objMasterOneXML.selectSingleNode("DelvInfo/recvPostCode").text
					If Len(ReceiveZipCode) = 6 Then
						'
					End If
					ReceiveAddr1		= objMasterOneXML.selectSingleNode("DelvInfo/recvAddr1").text
					ReceiveAddr2		= objMasterOneXML.selectSingleNode("DelvInfo/recvAddr2").text
					deliverymemo		= objMasterOneXML.selectSingleNode("DlvMemoCont").text
					If (deliverymemo = "null") Then
						deliverymemo = ""
					End If

					etc1				= ""
					countryCode			= "KR"

					'// ������
					set objDetailListXML = objMasterOneXML.selectNodes("ProdInfo")
					detailCnt = objDetailListXML.length
					For j = 0 to detailCnt - 1
						Set objDetailOneXML = objDetailListXML.item(j)

						OrgDetailKey		= objDetailOneXML.selectSingleNode("ProdSeq").text
						SellPrice			= objDetailOneXML.selectSingleNode("ordPrice").text
						''RealSellPrice		= objDetailOneXML.selectSingleNode("buyRealPrice").text
						RealSellPrice		= objDetailOneXML.selectSingleNode("ordPrice").text
						ItemOrderCount		= objDetailOneXML.selectSingleNode("ordQty").text
						requireDetail		= objDetailOneXML.selectSingleNode("GoodsChocDesc").text
						If (requireDetail = "null") Then
							requireDetail = ""
						End If

						orderDlvPay			= 0
						tmpStr = Split(objDetailOneXML.selectSingleNode("EntrProdNo").text, "_")
						matchItemID = ""
						matchItemOption = ""

                        If (matchItemID="") Then
                            matchItemID =  objDetailOneXML.selectSingleNode("EntrProdNo").text
                            matchItemOption = getOptionCodByOptionNameLotte(objDetailOneXML.selectSingleNode("EntrProdNo").text, Trim(objDetailOneXML.selectSingleNode("prodOption").text))
                        End If

						partnerItemName		= objDetailOneXML.selectSingleNode("ProdName").text
						if (matchItemID="791471") then
						    partnerItemName     = "[�������] �̲��� ���� Ʃ���Ʈ Ư�� 100 x 120cm 6�� ��1"'replace(partnerItemName,"&nbsp;","")
						elseif (matchItemID="635526") then
						    partnerItemName = replace(partnerItemName,"&nbsp;"," ")
						    partnerItemName = replace(partnerItemName,"��nbsp;"," ")
						    partnerItemName = replace(partnerItemName,"   "," ")
						    partnerItemName = replace(partnerItemName,"   "," ")
                            partnerItemName = replace(partnerItemName,"  "," ")
					    end if

						partnerOptionName	= Trim(objDetailOneXML.selectSingleNode("prodOption").text)
						outMallGoodsNo		= objDetailOneXML.selectSingleNode("ProdCode").text

						if (partnerOptionName = "null") then
							partnerOptionName = ""
						end if

						partnerItemID		= objDetailOneXML.selectSingleNode("ProdCode").text
						partnerOption		= ""

						if (matchItemOption="") and (partnerOptionName="") then
						    matchItemOption="0000"
						end if

                        if (matchItemOption="") then
                            rw "<font color=red>["&matchItemID&"]�ɼǸ�Ī ����:["&partnerOptionName&"]</font>"
                            matchItemOption="0000"
                        end if
'                rw "partnerItemID"&partnerItemID
'                rw "partnerOption"&partnerOption
'                rw "matchItemID"&matchItemID
'                rw "matchItemOption"&matchItemOption

						'// CS�������
						tmpStr = objDetailOneXML.selectSingleNode("Exchange").text
						isCSOrder = (tmpStr <> "�Ϲ�")

                        orderCsGbn=""
						if (Not isCSOrder) then
							orderCsGbn = "0"
						else
							orderCsGbn = "3"
						end if

						''response.write "<br>" & matchItemID & partnerItemName

						retVal = saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName,outMallGoodsNo _
								, OrderName, OrderTelNo, OrderHpNo _
								, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
								, SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
								, deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
								, errCode, errStr )

                        if (retVal) then
                            succCNT=succCNT+1
                        else
                            failCNT=failCNT+1
                        end if

						set objDetailOneXML = Nothing
					next

					set objDetailListXML = Nothing
					set objMasterOneXML = Nothing
				next
			end if

			set objMasterListXML = Nothing

			strSql = " update c "
			strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
			strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
			strSql = strSql + " from "
			strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
			strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
			strSql = strSql + " on "
			strSql = strSql + " 	1 = 1 "
			strSql = strSql + " 	and c.SellSite = o.SellSite and c.sellsite='"&FRectSellSite&"'"
			strSql = strSql + " 	and c.OutMallOrderSerial = Replace(o.OutMallOrderSerial, '-', '') "
			strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
			strSql = strSql + " where "
			strSql = strSql + " 	1 = 1 "
			strSql = strSql + " 	and c.orderserial is NULL "
			strSql = strSql + " 	and o.orderserial is not NULL "
			''rw strSql
			''rsget.Open strSql, dbget, 1
		else
			ErrMsg = "�Ľ̿� �����߽��ϴ�."
		end if
		Set xmlDOM = Nothing

		if (failCNT<>0) then
		    rw "["&failCNT&"] �� ����"
		end if

		if (succCNT<>0) then
		    rw "["&succCNT&"] �� ����"
		end if

	end function

	public function GetXmlFromWeb()
		objData = ""

		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		if objXML.Status = "200" then
			objData = BinaryToText(objXML.ResponseBody, "euc-kr")
		end if

		Set objXML  = Nothing
	end function

	public function GetXMLURL()
		dim tmp

		tmp = GetxSiteDateFormat(FRectStartYYYYMMDD)

		if (tmp = "") then
			GetXMLURL = ""
			ErrMsg = "���������� �������� �ʾҽ��ϴ�."
			exit function
		end if

		if (sellsite = "lotteimall") then
			if (FRectGubun = "new") then
				'// �ű�(01)
				GetXMLURL = FRectAPIURL + "/openapi/searchNewOrdLstOpenApi.lotte?subscriptionId=" + CStr(FRectAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD)) + "&SelOption=01"
			elseif (FRectGubun = "all") then
				'// ��ü(02)
				GetXMLURL = FRectAPIURL + "/openapi/searchNewOrdLstOpenApi.lotte?subscriptionId=" + CStr(FRectAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD)) + "&SelOption=02"
			else
				GetXMLURL = ""
				ErrMsg = "��ϵ��� ���� ���޸��Դϴ�.[1]"
			end if
		Elseif (sellsite = "lotteCom") Then
			If (FRectGubun = "new") Then
				'// �ű�(01)
				GetXMLURL = FRectAPIURL + "/openapi/searchNewOrdLstOpenApi.lotte?subscriptionId=" + CStr(FRectAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD)) + "&SelOption=01"
			ElseIf (FRectGubun = "all") Then
				'// ��ü(02)
				GetXMLURL = FRectAPIURL + "/openapi/searchNewOrdLstOpenApi.lotte?subscriptionId=" + CStr(FRectAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD)) + "&SelOption=02"
			Else
				GetXMLURL = ""
				ErrMsg = "��ϵ��� ���� ���޸��Դϴ�.[1]"
			End If
		else
			GetXMLURL = ""
			ErrMsg = "��ϵ��� ���� ���޸��Դϴ�.[2]"
		end if
	end function

	public function GetxSiteDateFormat(dt)
		if (FRectSellSite = "lotteimall") then
			GetxSiteDateFormat = Replace(dt, "-", "")
		elseif (FRectSellSite = "lotteCom") then
			GetxSiteDateFormat = Replace(dt, "-", "")
		else
			GetxSiteDateFormat = ""
		end if
	end function

	public function ResetXML()
		Set objXML = Nothing
		Set xmlDOM = Nothing
    end function

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FResultCount = 0
		FTotalCount = 0

		Call ResetXML()
	End Sub

	Private Sub Class_Terminate()
		Call ResetXML()
	End Sub

End Class

function saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName,outMallGoodsNo _
        , OrderName, OrderTelNo, OrderHpNo _
        , ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
        , SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
        , deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
        , byref ierrCode, byref ierrStr )
    dim paramInfo, retParamInfo
    dim PayType  : PayType  = "50"
    dim sqlStr
	dim countryCode

	if countryCode="" then countryCode="KR"

    saveOrderOneToTmpTable =false

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
		,Array("@outMallGoodsNo"	,adVarchar, adParamInput,16, outMallGoodsNo) _
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
		,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
	)

    if (matchItemOption<>"") and (matchItemID<>"-1") and (matchItemID<>"") then
        sqlStr = "db_temp.dbo.sp_TEN_xSite_TmpOrder_Insert_FromXML"
        retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

        ierrCode = GetValue(retParamInfo, "@RETURN_VALUE") ' �����ڵ�
        ierrStr  = GetValue(retParamInfo, "@retErrStr")   ' �����޼���
    else
        ierrCode = -999
        ierrStr = "��ǰ�ڵ� �Ǵ� �ɼ��ڵ�  ��Ī ����" & OrgDetailKey & " ��ǰ�ڵ� =" & matchItemID&" �ɼǸ� = "&partnerOptionName
        rw "["&ierrCode&"]"&ierrStr
        dbget.close() : response.end
    end if

    saveOrderOneToTmpTable = (ierrCode=0)
    if (ierrCode<>0) then
        rw "["&ierrCode&"]"&ierrStr
    end if
end function

public function getChrCount(orgStr, delim)
    dim retCNT : retCNT = 0
    dim buf
    buf = split(orgStr,delim)

    if IsArray(buf) then
        retCNT = UBound(buf)
    end if
    getChrCount = retCNT
end function

public function getOptionCodByOptionNameLotte(iitemid,ioptionname)
    dim retStr, sqlStr : retStr=""
    dim IsDoubleOption, IsTreepleOption
    IF (getChrCount(ioptionname,":")=2) THEN
        IF (getChrCount(ioptionname,",")=1) THEN
            IsDoubleOption = TRUE
        END IF
    ELSEIF (getChrCount(ioptionname,":")=3) THEN  '''������:c21,��Ʈ����:��Ʈ2,������ũ�߰� ����:�߰�����
        IF (getChrCount(ioptionname,",")=2) THEN
            IsTreepleOption = TRUE
        END IF
    ENd IF

    ioptionname= replace(ioptionname,"'","''")   '' like this CASE : �𵨸�:SMN-204 you're in
    IF (IsDoubleOption) THEN
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        sqlStr = sqlStr & " and optionTypename='���տɼ�'"
        sqlStr = sqlStr & " and replace(optionname,'*','')='"&SplitValue(SplitValue(ioptionname,",",0),":",1)&","&SplitValue(SplitValue(ioptionname,",",1),":",1)&"'"
    ELSEIF (IsTreepleOption) THEN
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        sqlStr = sqlStr & " and optionTypename='���տɼ�'"
        sqlStr = sqlStr & " and replace(optionname,'*','')='"&SplitValue(SplitValue(ioptionname,",",0),":",1)&","&SplitValue(SplitValue(ioptionname,",",1),":",1)&","&SplitValue(SplitValue(ioptionname,",",2),":",1)&"'"
    ELSE
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        ''sqlStr = sqlStr & " and optionTypename='"&SplitValue(ioptionname,":",0)&"'"
        sqlStr = sqlStr & " and Replace(Replace(replace(optionname,'*',''),',',''),'#','')=Replace('"&SplitValue(ioptionname,":",1)&"','#','')"
    END IF

	'response.write sqlstr & "<Br>"
    rsget.Open sqlStr,dbget,1
    if (Not rsget.EOF) then
	    retStr = rsget("itemoption")
	end if
    rsget.Close

    If (retStr="") THEN
       ''�ɼ� ��Ī�� �ȵǾ�����. �����Ī���� ����
        sqlStr = "select count(*) as CNT "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        rsget.Open sqlStr,dbget,1
        if (Not rsget.EOF) then
    	    if (rsget("CNT")>0) THEN retStr = "0000"
    	end if
        rsget.Close

    END IF
    getOptionCodByOptionNameLotte = retStr

'	if retStr="" then
'	    rw sqlStr
'	end if
end function
%>
