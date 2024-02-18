
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<% 
dim idx,answer,mode,page,masteridx

idx=request("idx")
answer=html2db(request("answer"))
mode=request("mode")
page=request("page")
masteridx=request("masteridx")
dim sql
if mode="add" then

sql ="update [db_cs].[dbo].tbl_weekly_qna" + vbcrlf
sql = sql + " set answer='" + answer + "'" + vbcrlf
sql = sql + " where idx='" + Cstr(idx) + "'" + vbcrlf

rsget.open sql,dbget,1
response.write "<script>document.location.href='/admin/board/weekly_codi_qna_view.asp?idx=" + idx  + "';</script>;"
dbget.close()	:	response.End
elseif mode="del" then

sql ="update [db_cs].[dbo].tbl_weekly_qna" + vbcrlf
sql = sql + " set isusing='N'" + vbcrlf
sql = sql + " where idx='" + Cstr(idx) + "'" + vbcrlf

rsget.open sql,dbget,1
response.write "<script>document.location.href='/admin/board/weekly_codi_qna_list.asp?page=" & page & "&masteridx=" & masteridx & "';</script>;"
dbget.close()	:	response.End
end if



%>


<!-- #include virtual="/lib/db/dbclose.asp" -->