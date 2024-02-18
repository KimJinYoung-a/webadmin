<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 매장 환율 관리
' History : 2010.08.07 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">

<% if session("sslgnMethod")<>"S" then %>
	<!-- USB키 처리 시작 (2008.06.23;허진원) -->
	<OBJECT ID='MaGerAuth' WIDTH='' HEIGHT=''	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'></OBJECT>
	<script language="javascript" src="/js/check_USBToken.js"></script>
	<!-- USB키 처리 끝 -->
<% end if %>
</head>
<body bgcolor="#F4F4F4" onload="checkUSBKey()">

<%
dim idx, sitename, currencyUnit, currencyChar, exchangeRate, basedate ,sqlStr, userid, menupos, mode
	idx   = request("idx")
	sitename   = request("sitename")
	currencyUnit   = request("currencyUnit")
	currencyChar   = request("currencyChar")
	exchangeRate   = request("exchangeRate")
	basedate   = request("basedate")
	userid = session("ssBctId")
	menupos = request("menupos")
	mode = request("mode")
	
dim referer
	referer = request.ServerVariables("HTTP_REFERER")
		
if mode = "" then
	response.write "<script language='javascript'>"
	response.write "	alert(MODE 구분자가 지정되지 않았습니다.');"
	response.write "	location.replace('" & referer & "');"
	response.write "</script>"
end if

if mode = "exchangeRateedit" then

	sqlStr = "if exists(" + VbCrlf
	sqlStr = sqlStr & "		select top 1 * from db_shop.dbo.tbl_shop_exchangeRate where idx='"&idx&"'" + VbCrlf
	sqlStr = sqlStr & " )" + VbCrlf
    sqlStr = sqlStr & " 	update db_shop.dbo.tbl_shop_exchangeRate" + VbCrlf
    sqlStr = sqlStr & " 	set currencyChar=N'" + currencyChar + "'" + VbCrlf
    sqlStr = sqlStr & " 	,exchangeRate=N'" + exchangeRate + "'" + VbCrlf
    sqlStr = sqlStr & " 	,basedate=N'" + basedate + "'" + VbCrlf
    sqlStr = sqlStr & " 	,lastupdate=getdate()" + VbCrlf
    sqlStr = sqlStr & " 	,lastuserid=N'" + userid + "'" + VbCrlf
    sqlStr = sqlStr & " 	where sitename=N'" + sitename + "'" + VbCrlf
    sqlStr = sqlStr & " 	and currencyUnit=N'" + currencyUnit + "'" + VbCrlf
	sqlStr = sqlStr & " else" + VbCrlf
	sqlStr = sqlStr & " 	insert into db_shop.dbo.tbl_shop_exchangeRate (" + VbCrlf
    sqlStr = sqlStr & " 	sitename, currencyUnit ,currencyChar ,exchangeRate ,basedate ,regdate, lastupdate, reguserid, lastuserid"+ VbCrlf
    sqlStr = sqlStr & " 	) values("
    sqlStr = sqlStr & " 	N'" + sitename + "'" + VbCrlf    
    sqlStr = sqlStr & " 	,N'" + currencyUnit + "'" + VbCrlf
    sqlStr = sqlStr & " 	,N'" + currencyChar + "'" + VbCrlf    
    sqlStr = sqlStr & " 	,N'" + exchangeRate + "'" + VbCrlf
    sqlStr = sqlStr & " 	,N'" + basedate + "'" + VbCrlf
    sqlStr = sqlStr & " 	,getdate()" + VbCrlf
    sqlStr = sqlStr & " 	,getdate()" + VbCrlf
    sqlStr = sqlStr & " 	,N'" + userid + "'" + VbCrlf
    sqlStr = sqlStr & " 	,N'" + userid + "'" + VbCrlf                
    sqlStr = sqlStr & " 	)" + VbCrlf

	'response.write sqlStr &"<Br>"   
    dbget.Execute sqlStr

end if	
%>

</body>
</html>

<script language='javascript'>
	alert('OK');
	location.replace('<%=referer%>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
