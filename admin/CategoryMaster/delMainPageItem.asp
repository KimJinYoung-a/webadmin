<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim idarr,mode,cdl,menupos
idarr	= request("idx")
if Not(idarr="" or isNull(idarr)) then
	if right(idarr,1)="," then
		idarr	= left(idarr,len(idarr)-1)
	end if
end if
cdl		= request("cdl")
mode	= request("mode")
menupos	= request("menupos")

dim sqlStr

Select Case mode
	Case "del"
		sqlStr= "update [db_sitemaster].[dbo].tbl_category_mainItem" + vbcrlf
		sqlStr = sqlStr + " set isusing='N'" + vbcrlf
		sqlStr = sqlStr + " where idx in (" + Cstr(idarr) + ")" +vbcrlf
end Select

'response.Write sqlStr
'dbget.close()	:	response.End
dbget.execute(sqlStr)

	response.write "<script language='javascript'>alert('적용 하였습니다.')</script>"
 	response.write "<script language='javascript'>location.replace('/admin/categorymaster/category_Main_pageItem.asp?cdl=" + cdl+ "&menupos=" + menupos + "');</script>"
	dbget.close()	:	response.End
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
