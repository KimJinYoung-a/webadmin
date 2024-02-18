<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<%
dim idx,mode,itemid,sql
dim arritemid,arrcnt,i
dim referer
referer = request.ServerVariables("HTTP_REFERER")

idx = request("idx")
mode=request("mode")
itemid=request("itemid")


if mode="write" then
	if not (itemid="" or isnull(itemid)) then
		
		arritemid=split(itemid,",")
		arrcnt=ubound(arritemid)
		
		for i=0 to arrcnt
		sql = sql + " insert into db_contents.dbo.tbl_diary_linkitem (idx,itemid) " &_
								" values(" & idx & "," & arritemid(i) & ")"
		next
		
		rsget.open sql,dbget,1
		
	end if
		
elseif mode="del" then
		
	if not (itemid="" or isnull(itemid)) then
		sql	=	" delete from db_contents.dbo.tbl_diary_linkitem " &_
					" where idx='" & idx & "'" &_
					" and itemid in (" & itemid &")" 
		
	rsget.open sql,dbget,1
	end if
end if

response.write "<script language='javascript'>alert('적용되었습니다.')</script>"
response.write "<script language='javascript'>location.replace('" + CStr(referer) + "')</script>"
dbget.close()	:	response.End


%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
