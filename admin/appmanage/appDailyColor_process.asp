<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim mode, sqlStr, menupos
Dim yyyymmdd, ImageUrl, ImageUrl2, color_idx

mode		= request("mode")
menupos		= request("menupos")
yyyymmdd	= request("yyyymmdd")
ImageUrl	= request("ImageUrl")
ImageUrl2	= request("ImageUrl2")
color_idx	= request("color_idx")

Dim itemidarr, existsitemcnt, i
itemidarr		= request("itemidarr")

If mode = "I" Then
	sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as cnt FROM db_contents.dbo.tbl_app_color_master "
	sqlStr = sqlStr & " WHERE yyyymmdd = '"& yyyymmdd &"' "
	rsget.Open sqlStr,dbget,1
		If rsget("cnt") > 0 Then
			Response.Write "<script language='javascript'>alert('같은 날짜에 등록된 데이터가 있습니다.');history.back(-1);</script>"
		End If
	rsget.Close

	sqlStr = ""
	sqlStr = sqlStr & " INSERT INTO db_contents.dbo.tbl_app_color_master ( "
	sqlStr = sqlStr & " yyyymmdd, imageURL, ImageUrl2, color_idx, regdate) VALUES "
	sqlStr = sqlStr & " ('"&yyyymmdd&"', '"&ImageUrl&"', '"&ImageUrl2&"', '"&color_idx&"', getdate()) "
	dbget.execute sqlStr
	Response.Write "<script language='javascript'>alert('저장되었습니다');location.href='/admin/appmanage/appDailyColorList.asp?menupos="&menupos&"';</script>"
ElseIf mode = "U" Then
	sqlStr = ""
	sqlStr = sqlStr & " UPDATE db_contents.dbo.tbl_app_color_master SET "
	sqlStr = sqlStr & " yyyymmdd = '"& yyyymmdd &"' "
	sqlStr = sqlStr & " ,imageURL = '"& ImageUrl &"'"
	sqlStr = sqlStr & " ,imageURL2 = '"& ImageUrl2 &"'"
	sqlStr = sqlStr & " ,color_idx = '"& color_idx &"'"
	sqlStr = sqlStr & " ,lastupdate = getdate()"
	sqlStr = sqlStr & " WHERE yyyymmdd ='" & Cstr(yyyymmdd) & "'"
	dbget.execute sqlStr
	Response.Write "<script language='javascript'>alert('수정되었습니다');location.href='/admin/appmanage/appDailyColorList.asp?menupos="&menupos&"';</script>"
ElseIf mode = "itemreg" Then
	If itemidarr = "" Then
		Response.Write  "<script language='javascript'>"
		Response.Write  "	alert('상품을 선택하세요.');"
		Response.Write  "</script>"
		dbget.close()	:	response.End		
	End If

	sqlStr = ""
	sqlStr = sqlStr & " SELECT Count(*) as cnt FROM db_contents.dbo.tbl_app_color_detail WHERE yyyymmdd='"& trim(yyyymmdd) &"' "
	rsget.Open sqlStr,dbget,1
	If rsget("cnt") > 50 Then
		Response.Write "<script language='javascript'>alert('최대 등록가능 한 상품50개를 초과하였습니다');self.close();</script>"
	End If
	rsget.Close
	
	sqlStr = ""
	sqlStr = sqlStr & " SELECT Count(*) as cnt FROM db_contents.dbo.tbl_app_color_detail WHERE yyyymmdd='"& trim(yyyymmdd) &"' AND itemid in ("& trim(itemidarr) &") "
	rsget.Open sqlStr,dbget,1
	If rsget("cnt") > 0 Then
		Response.Write "<script language='javascript'>alert('같은 날짜에 등록된 상품이 있습니다.');history.back(-1);</script>"
	End If
	rsget.Close

	sqlStr = ""
	sqlStr = sqlStr & " insert into db_contents.dbo.tbl_app_color_detail (" & VBCRLF
	sqlStr = sqlStr & " yyyymmdd, itemid, regdate, sortNo, isusing)" & VBCRLF
	sqlStr = sqlStr & " 	select top 40" & VBCRLF
	sqlStr = sqlStr & " 	'"& trim(yyyymmdd) &"', i.itemid, getdate(), '0', 'Y'" & VBCRLF
	sqlStr = sqlStr & " 	FROM db_item.dbo.tbl_item i" & VBCRLF
	sqlStr = sqlStr & " 	left join db_contents.dbo.tbl_app_color_detail d" & VBCRLF
	sqlStr = sqlStr & " 		on i.itemid=d.itemid" & VBCRLF
	sqlStr = sqlStr & " 		and d.yyyymmdd='"& trim(yyyymmdd) &"'" & VBCRLF
	sqlStr = sqlStr & " 		and d.isusing='Y'" & VBCRLF
	sqlStr = sqlStr & " 	where i.itemid in ("& trim(itemidarr) &")"
	dbget.execute sqlStr

	response.write "<script language='javascript'>"
	response.write "	alert('저장되었습니다');"
	response.write "	opener.document.location.reload();"
	response.write "	self.close();"
	response.write "</script>"
ElseIf mode = "sortisusingedit" Then
	Dim detailitemarr, sortnoarr, isusingarr, cnt, tmpSort, tmpIsusing
	detailitemarr	= Request("detailitemarr")
	sortnoarr		= Request("sortnoarr")
	isusingarr		= Request("isusingarr")
	
	If sortnoarr= "" OR isusingarr = "" THEN
		Response.Write "<script language='javascript'>alert('순서 및 사용여부가 지정되지 않았습니다.'); history.back(-1);</script>"
		dbget.close()	:	response.End
	end if
	
	'선택상품 파악
	detailitemarr = split(detailitemarr,",")
	cnt = ubound(detailitemarr)
	
	'// 정렬순서 저장
	If detailitemarr <> "" THEN
		sortnoarr 	= split(sortnoarr,",")
		isusingarr	= split(isusingarr,",")
		For i = 0 to cnt
			IF sortnoarr(i) = "" THEN
				 tmpSort = "0"				
			ELSE	
				 tmpSort = sortnoarr(i)	
			END IF

			IF isusingarr(i) = "" THEN
				 tmpIsusing = "0"				
			ELSE	
				 tmpIsusing = isusingarr(i)	
			END IF
			sqlStr = " UPDATE db_contents.dbo.tbl_app_color_detail SET" + vbcrlf
			sqlStr = sqlStr & " sortNo = '"&tmpSort&"'" + vbcrlf
			sqlStr = sqlStr & " ,isusing = '"&tmpIsusing&"'" + vbcrlf
			sqlStr = sqlStr & " WHERE itemid =" + detailitemarr(i)
			dbget.execute sqlStr
		Next
	END IF
	response.write "<script language='javascript'>"
	response.write "	alert('저장되었습니다');"
	response.write "	location.href='/admin/appmanage/iframe_appDailyColorDetail.asp?yyyymmdd="&yyyymmdd&"&menupos="&menupos&"';</script>"
	response.write "</script>"
Else
	Response.Write "<script language='javascript'>alert('구분자가 없습니다.'); history.back(-1);</script>"
	dbget.close()	:	response.End	
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->