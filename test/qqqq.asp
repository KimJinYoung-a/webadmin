<%

function GetOrderFrom_interpark(selldate)
	dim sellsite : sellsite = "interpark"
	dim xmlURL, xmlSelldate
	dim objXML, xmlDOM, objData
	dim masterCnt, resultcode, obj
	dim objMasterListXML, objMasterOneXML
	dim oMaster, oDetail

	GetOrderFrom_interpark = False

	'// =======================================================================
	'// 날짜형식
	''selldate = "2017-10-21"
	xmlSelldate = Replace(selldate, "-", "")

	'// API URL(기간동안의 전체내역 가져오기)
	xmlURL = "https://joinapi.interpark.com"
	xmlURL = xmlURL + "/order/OrderClmAPI.do?_method=orderListDelvForSingle&sc.entrId=10X10&sc.supplyEntrNo=3000010614&sc.supplyCtrtSeq=2&sc.strDate=" + xmlSelldate + "000000" + "&sc.endDate=" + xmlSelldate + "235959"
	''response.write xmlURL


	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.setTimeouts 5000,80000,80000,80000
	objXML.Send()

	if objXML.Status = "200" then
		objData = BinaryToText(objXML.ResponseBody, "euc-kr")
		''response.write objData
	else
		response.write "ERROR : 통신오류"
		response.end
	end if


	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML replace(objData,"&","＆")

	Set obj = xmlDOM.selectNodes("/ORDER_LIST/ORDER")

	if obj is Nothing then
		response.write "내역없음 : 종료"

		GetOrderFrom_interpark = True
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	masterCnt = (xmlDOM.selectNodes("/ORDER_LIST/ORDER").length)
	''response.write masterCnt

	if masterCnt = 0 then
		response.write "내역없음 : 종료"

		GetOrderFrom_interpark = True
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	set objMasterListXML = xmlDOM.selectNodes("/ORDER_LIST/ORDER")
	masterCnt = objMasterListXML.length

	for i = 0 to masterCnt - 1
		set objMasterOneXML = objMasterListXML.item(i)
		Set oMaster = new COrderMasterItem

		oMaster.FSellSite = sellsite
		oMaster.FOutMallOrderSerial = objMasterOneXML.selectSingleNode("ORD_NO").text
		oMaster.FSellDate			= objMasterOneXML.selectSingleNode("ORDER_DT").text
		oMaster.FPayType			= "50"
		oMaster.FPaydate			= Left(objMasterOneXML.selectSingleNode("PAY_DTS").text,8)
		oMaster.FOrderUserID		= ""
		oMaster.FOrderName			= objMasterOneXML.selectSingleNode("ORD_NM").text
		oMaster.FOrderEmail			= ""
		oMaster.FOrderTelNo			= objMasterOneXML.selectSingleNode("TEL").text
		oMaster.FOrderHpNo			= objMasterOneXML.selectSingleNode("MOBILE_TEL").text
		oMaster.FReceiveName		= objMasterOneXML.selectSingleNode("RCVR_NM").text
		oMaster.FReceiveTelNo		= objMasterOneXML.selectSingleNode("DELI_TEL").text
		oMaster.FReceiveHpNo		= objMasterOneXML.selectSingleNode("DELI_MOBILE").text
		oMaster.FReceiveZipCode		= objMasterOneXML.selectSingleNode("DEL_ZIP").text
		oMaster.FReceiveAddr1		= objMasterOneXML.selectSingleNode("DELI_ADDR1").text
		oMaster.FReceiveAddr2		= objMasterOneXML.selectSingleNode("DELI_ADDR2").text
		oMaster.Fdeliverymemo		= objMasterOneXML.selectSingleNode("DELI_COMMENT").text
		oMaster.FDEL_AMT 			= objMasterOneXML.selectNodes("DELIVERY/DELV").item(0).selectSingleNode("DEL_AMT").text




		response.write oMaster.FOutMallOrderSerial & "<br />"

		Set oMaster = Nothing
		Set objMasterOneXML = Nothing
	next


	Set xmlDOM = Nothing
	Set objXML = Nothing

	GetOrderFrom_interpark = True
end function


class COrderDetail
	public FdetailSeq
	public FItemID
	public FItemOption
	public FOutMallItemID
	public FOutMallItemOption
	public FOutMallItemOptionName
	public Fitemcost
	public FReducedPrice
	public FItemNo
	public FOutMallCouponPrice
	public FTenCouponPrice
end class

class COrderMasterItem
	public FSellSite
	public FOutMallOrderSerial
	public FSellDate
	public FPayType
	public FPaydate
	public FOrderUserID
	public FOrderName
	public FOrderEmail
	public FOrderTelNo
	public FOrderHpNo
	public FReceiveName
	public FReceiveTelNo
	public FReceiveHpNo
	public FReceiveZipCode
	public FReceiveAddr1
	public FReceiveAddr2
	public Fdeliverymemo
	public FdeliverPay
end class

%>
