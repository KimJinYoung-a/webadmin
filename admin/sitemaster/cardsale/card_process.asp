<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim mode, idx, sqlStr
Dim startDate, startDateTime, endDate, endDateTime, cardCode, saleType, salePrice, minPrice, maxPrice, isUsing
dim bannerTitle, bannerView, bgcolor, blnWeb, blnMobile, blnApp
mode			= request("mode")
idx				= request("idx")
startDate   	= request("startDate")
startDateTime	= request("startDateTime")
endDate			= request("endDate")
endDateTime		= request("endDateTime")
startDate = startDate & " " & startDateTime
endDate = endDate & " " & endDateTime

cardCode		= request("cardCode")
saleType		= request("saleType")
salePrice		= request("salePrice")
minPrice		= request("minPrice")
maxPrice		= request("maxPrice")
isUsing			= request("isUsing")
bannerTitle		= request("bannerTitle")
bgcolor			= request("bgcolor")
blnWeb			= request("blnWeb")
blnMobile		= request("blnMobile")
blnApp			= request("blnApp")

if bannerView = "on" then
	bannerView="N"
else
	if bannerTitle="" then
		bannerView="N"
	else
		bannerView="Y"
	end if
end if

if blnWeb<>"Y" then blnWeb="N"
if blnMobile<>"Y" then blnMobile="N"
if blnApp<>"Y" then blnApp="N"

If mode = "I" Then
    sqlStr = ""
    sqlStr = sqlStr & " INSERT INTO [db_item].[dbo].[tbl_card_sale] (startdate, enddate, cardCode, saleType, salePrice, minPrice, maxPrice, isUsing, regdate, regUserid, bannerTitle, bannerView, bgcolor, blnWeb, blnMobile, blnApp) " & VBCRLF
    sqlStr = sqlStr & " VALUES ('"& startDate &"', '"& enddate &"', '"& cardCode &"', '"& saleType &"', '"& salePrice &"', '"& minPrice &"' " & VBCRLF
	If saleType = "1" Then
		sqlStr = sqlStr & " , NULL "
	Else
		sqlStr = sqlStr & " , '"& maxPrice &"' "
	End If
	sqlStr = sqlStr & " , '" & isUsing & "', GETDATE(), '" & session("ssBctID") & "','" & bannerTitle & "','" & bannerView & "','" & bgcolor & "','" & blnWeb & "','" & blnMobile & "','" & blnApp & "') " & VBCRLF
	dbget.Execute(sqlStr)
    Response.Write "<script language=javascript>alert('저장 하였습니다.');opener.location.reload();self.close();</script>"
    dbget.close()	:	response.End
ElseIf mode = "U" Then
	sqlStr = ""
	sqlStr = sqlStr & " UPDATE [db_item].[dbo].[tbl_card_sale] SET " & VBCRLF
	sqlStr = sqlStr & " startdate = '"& startdate &"' " & VBCRLF
	sqlStr = sqlStr & " ,enddate = '"& enddate &"' " & VBCRLF
	sqlStr = sqlStr & " ,cardCode = '"& cardCode &"' " & VBCRLF
	sqlStr = sqlStr & " ,saleType = '"& saleType &"' " & VBCRLF
	sqlStr = sqlStr & " ,salePrice = '"& salePrice &"' " & VBCRLF
	sqlStr = sqlStr & " ,minPrice = '"& minPrice &"' " & VBCRLF
	If saleType = "1" Then
		sqlStr = sqlStr & " ,maxPrice = NULL " & VBCRLF
	Else
		sqlStr = sqlStr & " ,maxPrice = '"& maxPrice &"' " & VBCRLF
	End If
	sqlStr = sqlStr & " ,isUsing = '"& isUsing &"' " & VBCRLF
	sqlStr = sqlStr & " ,bannerTitle = '"& bannerTitle &"' " & VBCRLF
	sqlStr = sqlStr & " ,bannerView = '"& bannerView &"' " & VBCRLF
	sqlStr = sqlStr & " ,lastupdate = GETDATE() " & VBCRLF
	sqlStr = sqlStr & " ,lastUpdateUserId = '"& session("ssBctID") &"' " & VBCRLF
	sqlStr = sqlStr & " ,bgcolor = '"& bgcolor &"' " & VBCRLF
	sqlStr = sqlStr & " ,blnWeb = '"& blnWeb &"' " & VBCRLF
	sqlStr = sqlStr & " ,blnMobile = '"& blnMobile &"' " & VBCRLF
	sqlStr = sqlStr & " ,blnApp = '"& blnApp &"' " & VBCRLF
	sqlStr = sqlStr & " WHERE idx = '"& idx &"' " & VBCRLF
	dbget.Execute(sqlStr)
	Response.Write "<script language=javascript>alert('수정 하였습니다.');opener.location.reload();self.close();</script>"
	dbget.close()	:	response.End
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->