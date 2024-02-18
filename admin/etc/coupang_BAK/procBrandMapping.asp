<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
Dim makerid, phoneNumber2, deliveryCode, returnZipCode, returnAddress, returnAddressDetail, jeju, notJeju, maeipdiv
Dim strSql, gubun
makerid				= requestCheckvar(Request("makerid"),32)
gubun				= request("gubun")
phoneNumber2		= request("phoneNumber2")
deliveryCode		= request("deliveryCode")
returnZipCode		= request("returnZipCode")
returnAddress		= request("returnAddress")
returnAddressDetail	= request("returnAddressDetail")
jeju				= request("jeju")
notJeju				= request("notJeju")
maeipdiv			= request("maeipdiv")

If gubun = "popup" Then
	If maeipdiv = "U" Then
		strSql = ""
		strSql = strSql & " IF Exists(SELECT * FROM db_etcmall.dbo.tbl_coupang_branddelivery_mapping WHERE makerid='"&makerid&"' )"
		strSql = strSql & " BEGIN "
		strSql = strSql & " UPDATE db_etcmall.dbo.tbl_coupang_branddelivery_mapping SET "
		strSql = strSql & " companyContactNumber='"&phoneNumber2&"'"
		strSql = strSql & " , returnZipCode='"&returnZipCode&"'"
		strSql = strSql & " , returnAddress='"&returnAddress&"'"
		strSql = strSql & " , returnAddressDetail='"&returnAddressDetail&"'"
		strSql = strSql & " , deliveryCode='"&deliveryCode&"'"
		strSql = strSql & " , jeju='"&jeju&"'"
		strSql = strSql & " , notJeju='"&notJeju&"'"
		strSql = strSql & " WHERE makerid = '"&makerid&"' "
		strSql = strSql & " END ELSE "
		strSql = strSql & " BEGIN "
		strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_coupang_branddelivery_mapping "
		strSql = strSql & " (makerid, vendorId, shippingPlaceName, global, addressType, countryCode, companyContactNumber, phoneNumber2, returnZipCode, returnAddress, returnAddressDetail, deliveryCode, jeju, NotJeju, regdate ) VALUES "
		strSql = strSql & " ('"&makerid&"', '', '"&makerid&" shipping place', 'false', 'JIBUN', 'KR', '1644-6035', '"&phoneNumber2&"', '"&returnZipCode&"', '"&returnAddress&"', '"&returnAddressDetail&"', '"&deliveryCode&"', '"&jeju&"', '"&NotJeju&"', getdate()) END "
		dbget.Execute strSql
	Else
		strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_coupang_branddelivery_mapping WHERE makerid='"&makerid&"' )"
		strSql = strSql & " BEGIN "
		strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_coupang_branddelivery_mapping "
		strSql = strSql & " (makerid, vendorId, outboundShippingPlaceCode, regdate ) VALUES "
		strSql = strSql & " ('"&makerid&"', '', '122412', getdate()) END "
		dbget.Execute strSql
	End If
End If
%>
<script language="javascript">
alert("정상적으로 처리되었습니다.");
parent.opener.history.go(0);
parent.self.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->