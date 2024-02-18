<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim sqlStr, menupos, cnt
Dim idx, rank, likeword, isusing
idx			= request("idx")
rank		= requestCheckvar(request("rank"),2)
likeword	= request("likeword")
isusing		= requestCheckvar(request("isusing"),1)

If NOT isnumeric(idx) AND idx <> "" Then
	response.write	"<script language='javascript'>" &_
					"	alert('글번호가 잘못 되었습니다');" &_
					"	window.close();" &_
					"</script>"	
End If

If idx = "" Then
	cnt = 0
	sqlStr = ""
	sqlStr = sqlStr & " SELECT count(*) as cnt FROM db_outmall.dbo.tbl_between_search_likeWord WHERE rank = '"&rank&"' and isusing = 'Y' "
	rsCTget.Open sqlStr,dbCTget,1
	If rsCTget("cnt") > 0 Then
		cnt = rsCTget("cnt")
		response.write	"<script language='javascript'>" &_
						"	alert('저장된 동일 순서가 있습니다. 체크 후 등록하세요');" &_
						"	location.replace('/admin/etc/between/search/popRegWord.asp?idx="&idx&"');" &_
						"</script>"
	End If
	rsCTget.Close

	If cnt = 0 Then
		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_between_search_likeWord (rank, likeword, isusing) VALUES  "
		sqlStr = sqlStr & " ("&rank&", '"&likeword&"', '"&isusing&"')  "
		dbCTget.execute sqlStr
		response.write	"<script language='javascript'>" &_
						"	alert('저장되었습니다');" &_
						"	top.opener.location.reload();" &_
						"	window.close();" &_
						"</script>"	
	End If
Else
	cnt = 0
	If isusing = "Y" Then
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt FROM db_outmall.dbo.tbl_between_search_likeWord WHERE rank = '"&rank&"' and isusing = 'Y' and idx <> '"&idx&"' "
		rsCTget.Open sqlStr,dbCTget,1
		If rsCTget("cnt") > 0 Then
			cnt = rsCTget("cnt")
			response.write	"<script language='javascript'>" &_
							"	alert('저장된 동일 순서가 있습니다. 체크 후 등록하세요');" &_
							"	location.replace('/admin/etc/between/search/popRegWord.asp?idx="&idx&"');" &_
							"</script>"	
		End If
		rsCTget.Close
	End If

	If cnt = 0 Then
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_outmall.dbo.tbl_between_search_likeWord SET "
		sqlStr = sqlStr & " rank = "&rank&",  "
		sqlStr = sqlStr & " likeword = '"&likeword&"', "
		sqlStr = sqlStr & " isusing = '"&isusing&"' "
		sqlStr = sqlStr & " WHERE idx = "&idx&"  "
		dbCTget.execute sqlStr
		response.write	"<script language='javascript'>" &_
						"	alert('저장되었습니다');" &_
						"	top.opener.location.reload();" &_
						"	location.replace('/admin/etc/between/search/popRegWord.asp?idx="&idx&"');" &_
						"</script>"
	End If
End If
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->