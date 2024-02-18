<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim id,sql,masterid,thispage

id = request("id")
id = replace(id,",","','")
id = left(id,len(id)-3)
'response.write id
'sdbget.close()	:	response.End

			sql = "update [db_board].[10x10].tbl_offshop_board" + vbcrlf
			sql = sql + " set masterid='31'" + vbcrlf
			sql = sql + " where id in ('" + Cstr(id) + "')"
			rsget.Open sql, dbget, 1

		response.write "<script language='javascript'>alert('삭제하였습니다.')</script>"
        response.write "<script language='javascript'>location.replace('/academy/academy_album.asp')</script>"
        dbget.close()	:	response.End

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->