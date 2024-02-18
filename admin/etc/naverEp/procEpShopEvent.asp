<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<%
Dim eventName, gubun, mode, strSql, mallid, idx
Dim startDate, endDate, isUsing, defaultCnt, existCnt
idx				= request("idx")
mode			= request("mode")
eventName		= request("eventName")
gubun			= request("gubun")
mallid			= request("mallid")

startDate		= request("startDate")
endDate			= request("endDate")
isUsing			= request("isUsing")

If mode = "I" Then
	If gubun = "1" Then	'기본
		strSql = ""
		strSql = " SELECT COUNT(*) as cnt FROM [db_outMall].[dbo].[tbl_EpShop_Event] WHERE gubun = '1' AND mallid = '"& mallid &"' "
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open strSql, dbCTget, adOpenForwardOnly, adLockReadOnly
			defaultCnt = rsCTget("cnt")
		rsCTget.Close

		If defaultCnt > 0 Then
			Response.Write "<script language=javascript>alert('기본문구는 한개만 등록가능합니다.');parent.location.reload();</script>"
			dbCTget.close()	:	response.End
		Else
			strSql = ""
			strSql = strSql & " INSERT INTO [db_outMall].[dbo].[tbl_EpShop_Event] " & vbCrLf
			strSql = strSql & " (mallid, eventName, gubun, startDate, endDate, regdate) VALUES " & vbCrLf
			strSql = strSql & " ('"& mallid &"', '"& eventName &"', '1', '1910-01-01 00:00:00', '2999-12-31 23:59:59', getdate()) " & vbCrLf
			dbCTget.execute strSql
		End If
	Else
		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as cnt "
		strSql = strSql & " FROM [db_outMall].[dbo].[tbl_EpShop_Event] "
		strSql = strSql & " WHERE gubun= '2' "
		strSql = strSql & " and startDate < '"& endDate &" 23:59:59' "
		strSql = strSql & " and endDate > '"& startDate &" 00:00:00' "
		strSql = strSql & " and isUsing = 'Y' "
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open strSql, dbCTget, adOpenForwardOnly, adLockReadOnly
			existCnt = rsCTget("cnt")
		rsCTget.Close

		If existCnt > 0 Then
			Response.Write "<script language=javascript>alert('등록 하는 날짜사이의 데이터가 존재합니다.');parent.location.reload();</script>"
			dbCTget.close()	:	response.End
		Else
			strSql = ""
			strSql = strSql & " INSERT INTO [db_outMall].[dbo].[tbl_EpShop_Event] " & vbCrLf
			strSql = strSql & " (mallid, eventName, gubun, startDate, endDate, isUsing, regdate) VALUES " & vbCrLf
			strSql = strSql & " ('"& mallid &"', '"& eventName &"', '"& gubun &"', '"& LEFT(startDate, 10) &" 00:00:00', '"& LEFT(endDate, 10) &" 23:59:59', '"& isUsing &"', getdate()) "
			dbCTget.execute strSql
		End If
	End If
	Response.Write "<script language=javascript>parent.location.reload();</script>"
	dbCTget.close()	:	response.End
Else
	If gubun = "1" Then	'기본
		strSql = ""
		strSql = strSql & " UPDATE [db_outMall].[dbo].[tbl_EpShop_Event] " & vbCrLf
		strSql = strSql & " SET eventName = '"& eventName &"' "
		strSql = strSql & " WHERE mallid = '"& mallid &"' " & vbCrLf
		strSql = strSql & " and gubun = '1' "
		strSql = strSql & " and idx = '"& idx &"' "
		dbCTget.execute strSql
	Else
		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as cnt "
		strSql = strSql & " FROM [db_outMall].[dbo].[tbl_EpShop_Event] "
		strSql = strSql & " WHERE gubun= '2' "
		strSql = strSql & " and startDate < '"& endDate &" 23:59:59' "
		strSql = strSql & " and endDate > '"& startDate &" 00:00:00' "
		strSql = strSql & " and isUsing = 'Y' "
		strSql = strSql & " and idx <> '"& idx &"' "
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open strSql, dbCTget, adOpenForwardOnly, adLockReadOnly
			existCnt = rsCTget("cnt")
		rsCTget.Close

		If existCnt > 0 Then
			Response.Write "<script language=javascript>alert('등록 하는 날짜사이의 데이터가 존재합니다.');parent.location.reload();</script>"
			dbCTget.close()	:	response.End
		Else
			strSql = ""
			strSql = strSql & " UPDATE [db_outMall].[dbo].[tbl_EpShop_Event] " & vbCrLf
			strSql = strSql & " SET eventName = '"& eventName &"' "
			strSql = strSql & " ,startDate = '"& startDate &" 00:00:00' "
			strSql = strSql & " ,endDate = '"& endDate &" 23:59:59' "
			strSql = strSql & " ,isUsing = '"& isUsing &"' "
			strSql = strSql & " WHERE mallid = '"& mallid &"' " & vbCrLf
			strSql = strSql & " and gubun = '2' "
			strSql = strSql & " and idx = '"& idx &"' "
			dbCTget.execute strSql
		End If
	End If
	Response.Write "<script language=javascript>parent.location.reload();</script>"
	dbCTget.close()	:	response.End
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->