<%@ language=vbscript %>
<%
option Explicit
'Response.Buffer = True
Response.CharSet = "euc-kr"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	dim dispCate, depth, i
	dispCate = requestCheckVar(request("disp"),16)
	depth = cInt(len(dispCate)/3)
	if depth=0 then depth=1
	if dispCate<>"" then depth=depth+1
	if depth>4 then depth=4

	dim sqlStr, lp, vBody

	'// 선택 상자 출력
	for i=1 to depth
		sqlStr = " SELECT catecode, depth, catename, useyn, sortNo "
		sqlStr = sqlStr & " FROM [db_item].[dbo].[tbl_display_cate] "
		sqlStr = sqlStr & " WHERE depth = '" & i & "' "
		if i>1 then
			sqlStr = sqlStr & " and catecode like '" & left(dispCate,(i-1)*3) & "%'"
		end if
		sqlStr = sqlStr & " order by sortNo, catecode "
		rsget.Open sqlStr,dbget,1

		if Not(rsget.EOF or rsget.BOF) then
			vBody = vBody & "<div id=""dispcateval" & i & """ style=""float:left;"">"
			vBody = vBody & "<select name=""disp" & i & """ class=""formSlt"" onchange=""chgDispCate(this.value)"">" & vbCrLf
			vBody = vBody & "	<option value='" & chkIIF(i>1,left(dispCate,(i-1)*3),"") & "'>::Depth" & i & "::</option>" & vbCrLf
	
			For lp=0 To rsget.RecordCount -1
	
				vBody = vBody & "	<option value="""& rsget("catecode") &""""
				If CStr(rsget("catecode")) = left(dispCate,i*3) Then
					vBody = vBody & " selected"
				End If
				vBody = vBody & ">"& rsget("catename") &"</option>" & vbCrLf
	
				rsget.MoveNext
			next
	
			vBody = vBody & "</select> "& vbCrLf
			vBody = vBody & "</div> "
		end if

		rsget.close
	next

	Response.Write vBody
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->