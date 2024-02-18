<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  상품고시
' History : 2013.12.11 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/shopcls.asp"-->
<%
dim Eval_excludeitemidarr, adminid, mode, itemidarr, sqlStr, i
	Eval_excludeitemidarr 	= request("Eval_excludeitemid")
	mode 	= request("mode")
	menupos 	= request("menupos")
	itemidarr 	= request("itemidarr")
	
adminid = session("ssBctId")

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

'/상품삭제
if mode="delitem" then
	Eval_excludeitemidarr = split(Eval_excludeitemidarr,",")

	for i = 0 to ubound(Eval_excludeitemidarr)
		sqlStr = "delete from" + VBCRLF
		sqlStr = sqlStr & " db_board.dbo.tbl_Item_Evaluate_exclude" + VBCRLF
		sqlStr = sqlStr & " where itemid in ("& trim(Eval_excludeitemidarr(i)) &")"
	
		'response.write sqlStr & "<BR>"	
		dbget.execute sqlStr
	next
	
	response.write "<script language='javascript'>"
	response.write "	alert('삭제되었습니다');"
	response.write "	document.location.href='"& refer &"'"
	response.write "</script>"

'/상품추가
elseif mode="regitem" then
	itemidarr = split(itemidarr,",")

	for i = 0 to ubound(itemidarr)
		
		sqlStr = "if not exists(" & VBCRLF
		sqlStr = sqlStr & " 	select top 1 *" & VBCRLF
		sqlStr = sqlStr & " 	from db_board.dbo.tbl_Item_Evaluate_exclude" & VBCRLF
		sqlStr = sqlStr & " 	where itemid in ("& trim(itemidarr(i)) &")" & VBCRLF
		sqlStr = sqlStr & " )" & VBCRLF
		sqlStr = sqlStr & " 	insert into db_board.dbo.tbl_Item_Evaluate_exclude(" & VBCRLF
		sqlStr = sqlStr & " 	itemid, regdate, lastupdate, regadminid, lastadminid)" & VBCRLF
		sqlStr = sqlStr & " 		select top 500" & VBCRLF
		sqlStr = sqlStr & " 		i.itemid, getdate(), getdate(), '"&adminid&"', '"&adminid&"'" & VBCRLF
		sqlStr = sqlStr & " 		from db_item.dbo.tbl_item i" & VBCRLF
		sqlStr = sqlStr & " 		where itemid in ("& trim(itemidarr(i)) &")"

		'response.write sqlStr & "<Br>"
		dbget.execute sqlStr
	next

	response.write "<script language='javascript'>"
	response.write "	alert('저장되었습니다');"
	response.write "	opener.document.location.reload();"
	response.write "	self.close();"
	response.write "</script>"
	
else
	Response.Write "<script language='javascript'>alert('구분자가 없습니다.'); history.back(-1);</script>"
	dbget.close()	:	response.End	
End If
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->