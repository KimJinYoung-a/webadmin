<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 XML 주문처리
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteOrderXMLCls.asp"-->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/admin/etc/gsshop/gsshopItemcls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%

response.end

Dim refIP : refIP = Request.ServerVariables ("REMOTE_ADDR")
'if (refIP <> "211.44.122.231") then
'	dbget.close()
'	response.end
'end if

Dim sqlStr, buf
Dim i, j, k
Dim sItem
Dim NowHMS
NowHMS = Hour(Time())&Minute(Time())&Second(Time())&"0"

'For Each sItem In Request.Form
'	buf = buf & "["&sItem&"]" & Request.Form(sItem) & VbCRLF
'	''Response.Write(sItem)
'	''Response.Write(" - [" & Request.Form(sItem) & "]" & strLineBreak)
'Next

sqlStr = "insert into db_temp.dbo.tbl_tmp_gsOrder"
sqlStr = sqlStr&" (regdate,refip,xmlData)"
sqlStr = sqlStr&" values(getdate(),'"&refIP&"','"&replace(Server.HTMLEncode(Request.Form),"'","''")&"')"
dbget.Execute sqlStr
response.write "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
response.write "<PurchaseOrder_V01_00>" + vbCrLf
response.write "<MessageHeader>" + vbCrLf
response.write "	<Sender>10X10</Sender>" + vbCrLf
response.write "	<Receiver>GS SHOP</Receiver>" + vbCrLf
response.write "	<MessageID>"&Request.Form("MessageID")(1)&NowHMS&"</MessageID>" + vbCrLf
response.write "	<DateTime>"&Request.Form("DateTime")(1)&NowHMS&"</DateTime>" + vbCrLf
response.write "	<ProcessType>S</ProcessType>" + vbCrLf
response.write "	<DocumentID>"&Request.Form("DocumentID")(1)&"</DocumentID>" + vbCrLf
response.write "	<UniqueID>"&Request.Form("UniqueID")(1)&NowHMS&"</UniqueID>" + vbCrLf
response.write "	<ErrorOccur></ErrorOccur>" + vbCrLf
response.write "	<ErrorMessage></ErrorMessage>" + vbCrLf
response.write "</MessageHeader>" + vbCrLf
response.write "<MessageBody>" + vbCrLf
response.write "	<PurchaseOrders>" + vbCrLf
response.write "		<ordItemNo>"&Request.Form("ordItemNo")(1)&"</ordItemNo>" + vbCrLf
response.write "		<ordNo>"&Request.Form("ordNo")(1)&"</ordNo>" + vbCrLf
response.write "		<OrderGenerationDate>" & Left(Now(), 10) & "</OrderGenerationDate>" + vbCrLf
response.write "		<ProductLineItem>" + vbCrLf
response.write "			<ConfirmedDeliveryDate>" & Left(Now(), 10) & "</ConfirmedDeliveryDate>" + vbCrLf
response.write "			<sendFg>S</sendFg>" + vbCrLf
response.write "		</ProductLineItem>" + vbCrLf
response.write "	</PurchaseOrders>" + vbCrLf
response.write "</MessageBody>" + vbCrLf
response.write "</PurchaseOrder_V01_00>" + vbCrLf
response.end





dim mode
dim sellsite
dim reguserid
Dim AssignedRow
Dim ErrMsg
dim LastCheckDate, isSuccess
dim maxCheckCount : maxCheckCount = 10

dim resultCount

dim divcd, yyyymmdd

mode = "getxsiteorderlist"
sellsite = "gseshop"

dim oCxSiteOrderXML
Set oCxSiteOrderXML = new CxSiteOrderXML

Dim xmlDoc
if (sellsite="gseshop") then
	oCxSiteOrderXML.FRectSellSite = sellsite
	oCxSiteOrderXML.ActSavexSiteOrderListtoDB
end if










if (mode = "getxsiteorderlist") then

	oCxSiteOrderXML.FRectSellSite = sellsite

    IF (sellsite="gseshop") then
    	ErrMsg = ""

		for i = 0 to maxCheckCount - 1
			'// ================================================================
			Call oCxSiteOrderXML.GetCheckStatus(LastCheckDate, isSuccess)

			'// aaaaaaaaaaaa
			LastCheckDate = "20140807"

			oCxSiteOrderXML.FRectStartYYYYMMDD = LastCheckDate
			oCxSiteOrderXML.FRectEndYYYYMMDD = LastCheckDate

			'// tnsType : 주문구분(주문/반품 : S, 취소 : C)
			'// 개발 : test1 운영 : ecb2b
			oCxSiteOrderXML.FRectAPIURL = "http://test1.gsshop.com/SupSendOrderInfo.gs?supCd=" + CStr(COurCompanyCode) + "&sdDt=" + CStr(LastCheckDate) + "&tnsType=S"

			if (isSuccess = "Y") then
				oCxSiteOrderXML.FRectGubun = "new" ''"new"

				if Not IsAutoScript then
					response.write "<br>" & LastCheckDate & " : 주문(신규) 요청 "
				end if
			else
				oCxSiteOrderXML.FRectGubun = "all"

				if Not IsAutoScript then
					response.write "<br>" & LastCheckDate & " : 주문(전체) 요청 "
				end if
			end if

			Call oCxSiteOrderXML.SetCheckStatusStarting(LastCheckDate)

			'// 신규주문 전송요청만 한다.(XML 수신X)
			Call oCxSiteOrderXML.RequestxSiteOrderListOnly

			response.write oCxSiteOrderXML.ErrMsg

			Call oCxSiteOrderXML.SetCheckStatusEnded()

			if Not IsAutoScript then
				response.write "OK"
			end if

			if (CStr(LastCheckDate) >= CStr(Left(now, 10))) then
				exit for
			end if

			LastCheckDate = Left(DateAdd("d", 1, CDate(LastCheckDate)), 10)

			Call oCxSiteOrderXML.SetCheckDate(LastCheckDate)
		next
    else
        rw "미지정 sellsite:"&sellsite
        dbget.Close : response.end
    end if
else

end if

%>
<% if (FALSE) then %>
    <% if  (IsAutoScript) then  %>
    <% rw "OK" %>
    <% else %>
    <script>alert('저장되었습니다.');</script>
    <% end if %>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
