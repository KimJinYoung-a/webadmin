<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'==========================================================================
'	Description: 매장날씨 관리
'	History: 2012.06.04 강준구 생성
'			 2012.06.12 한용민 수정
'==========================================================================
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/weather/weather_cls.asp"-->

<%
Dim vIdx, vQuery, vWDate, vShopID, vShopName, vWeather, vComment, shopid, vIsExist, i, vExistShop, menupos
	menupos = requestCheckVar(Request("menupos"),10)
	vIsExist = "x"
	vIdx = requestCheckVar(Request("idx"),10)
	vWDate = requestCheckVar(Request("wdate"),30)
	vShopID = requestCheckVar(Request("shopid"),32)
	vWeather = requestCheckVar(Request("weather"),32)
	vComment = html2db(Request("comment"))

If vIdx <> "" Then
	if vComment <> "" then
		if checkNotValidHTML(vComment) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		dbget.close()	:	response.End
		end if
	end if

	vQuery = "UPDATE [db_shop].[dbo].[tbl_shop_weather] SET "
	vQuery = vQuery & "wdate = '" & vWDate & "', shopid = '" & vShopID & "', weather = '" & vWeather & "', comment = '" & vComment & "' "
	vQuery = vQuery & "WHERE idx = '" & vIdx & "'"
	
	'response.write vQuery & "<Br>"
	dbget.execute vQuery
	
ElseIf vIdx = "" Then
	if vComment <> "" then
		if checkNotValidHTML(vComment) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		dbget.close()	:	response.End
		end if
	end if

	For i = LBound(Split(vShopID,",")) To UBound(Split(vShopID,","))
		vQuery = "SELECT COUNT(idx)"
		vQuery = vQuery & " FROM [db_shop].[dbo].[tbl_shop_weather]"
		vQuery = vQuery & " WHERE wdate = '" & vWDate & "' AND shopid = '" & requestCheckVar(Trim(Split(vShopID,",")(i)),32) & "'"
		
		'response.write vQuery & "<Br>"
		rsget.Open vQuery,dbget,1
		If rsget(0) > 0 Then
			vIsExist = "o"
		Else
			vIsExist = "x"
		End IF
		rsget.close()
		
		If vIsExist = "x" Then
			vQuery = "INSERT INTO [db_shop].[dbo].[tbl_shop_weather](wdate, shopid, weather, comment) "
			vQuery = vQuery & "VALUES('" & vWDate & "', '" & requestCheckVar(Trim(Split(vShopID,",")(i)),32) & "', '" & vWeather & "', '" & vComment & "')"
			
			'response.write vQuery & "<Br>"
			dbget.execute vQuery
		Else
			vExistShop = vExistShop & Trim(Split(vShopID,",")(i)) & ", "
		End IF
	Next
End If
%>

<script language="javascript">

<% If vExistShop <> "" Then %>
	alert("<%=vWDate%>에 <%=vExistShop%>의 날씨정보가\n이미 등록되어있습니다.\n이를 제외한 SHOP은 등록이 되었습니다.");
	opener.document.location.reload();
	history.back();
<% Else %>
	opener.document.location.reload();
	window.close();
<% End If %>

</script>

<!-- #include virtual="/admin/lib/poptail.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->