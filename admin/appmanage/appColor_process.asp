<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim mode, sqlStr
Dim idx, colorCode, colorName, iconImage1, iconImage2, color_str, word_rgbCode, isusing, sortNo, menupos
Dim uidx
mode		= request("mode")
colorCode	= request("colorCode")
colorName	= request("colorName")
iconImage1	= request("iconImage1")
iconImage2	= request("iconImage2")
color_str	= request("color_str")
word_rgbCode= request("word_rgbCode")
isusing		= request("isusing")
sortNo		= request("sortNo")
menupos		= request("menupos")
uidx		= request("idx")			'수정시 쓰이는 idx

If colorCode = "New" Then
	sqlStr = ""
	sqlStr = sqlStr & " SELECT MAX(colorcode) as colorcode FROM db_contents.dbo.tbl_app_color_list "
	rsget.Open sqlStr, dbget, 1
	If not rsget.EOF Then
		If rsget("colorcode") < 1000 OR IsNull(rsget("colorcode")) Then
			colorcode = 1000
		Else
			colorcode = rsget("colorcode") + 1
		End If
	Else
		colorcode = 1000
	End If
	rsget.Close
End If

If mode = "I" Then
	sqlStr = ""
	sqlStr = sqlStr & " INSERT INTO db_contents.dbo.tbl_app_color_list ( "
	sqlStr = sqlStr & " colorCode, colorName, iconImageUrl1, iconImageUrl2, color_str, word_rgbCode, isusing, regdate, sortNo) VALUES "
	sqlStr = sqlStr & " ('"&colorCode&"', '"&colorName&"', '"&iconImage1&"', '"&iconImage2&"', '"&color_str&"', '"&word_rgbCode&"', '"&isusing&"', getdate(), '"&sortNo&"') "
	dbget.execute sqlStr
	Response.Write "<script language='javascript'>alert('저장되었습니다');opener.location.href='/admin/appmanage/appColorList.asp?menupos="&menupos&"';window.close();</script>"
ElseIf mode = "U" Then
	sqlStr = ""
	sqlStr = sqlStr & " UPDATE db_contents.dbo.tbl_app_color_list SET "
	sqlStr = sqlStr & " colorName = '"& colorName &"' "
	sqlStr = sqlStr & " ,iconImageUrl1 = '"& iconImage1 &"'"
	sqlStr = sqlStr & " ,iconImageUrl2 = '"& iconImage2 &"'"
	sqlStr = sqlStr & " ,color_str = '"&color_str&"'"
	sqlStr = sqlStr & " ,word_rgbCode = '"&word_rgbCode&"'"
	sqlStr = sqlStr & " ,isusing = '"&isusing&"'"
	sqlStr = sqlStr & " ,sortNo = '"&sortNo&"'"
	sqlStr = sqlStr & " WHERE idx ='" & Cstr(uidx) & "'"
	dbget.execute sqlStr
	Response.Write "<script language='javascript'>alert('수정되었습니다');opener.location.href='/admin/appmanage/appColorList.asp?menupos="&menupos&"';window.close();</script>"
Else
	Response.Write "<script language='javascript'>alert('구분자가 없습니다.'); history.back(-1);</script>"
	dbget.close()	:	response.End	
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->