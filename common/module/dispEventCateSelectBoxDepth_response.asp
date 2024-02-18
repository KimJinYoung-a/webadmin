<%@ language=vbscript %>
<%
option Explicit
'Response.Buffer = True
Response.CharSet = "euc-kr"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	dim dispCate, depth, i,maxdepth
	dispCate = requestCheckVar(request("disp"),16)
	maxdepth = requestCheckVar(request("maxD"),10)
	depth = cInt(len(dispCate)/3)
	if depth=0 then depth=1
	if dispCate<>"" then depth=depth+1
	if depth>4 then depth=4
	if cint(depth) > cint(maxdepth) then depth = maxdepth  '요청 최대 깊이까지만 
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
			if depth=2 then
				vBody = vBody & "<div class=""formInline lMar05"">" & vbCrLf
			end if
			vBody = vBody & "<select id=""disp" & i & """ name=""disp" & i & """ class=""formControl"" onchange=""chgDispCate(this.value,"&maxdepth&")"">" & vbCrLf
			vBody = vBody & "	<option value='" & chkIIF(i>1,left(dispCate,(i-1)*3),"") & "'>::Depth" & i & "::</option>" & vbCrLf

			For lp=0 To rsget.RecordCount -1

				vBody = vBody & "	<option value="""& rsget("catecode") &""""
				If CStr(rsget("catecode")) = left(dispCate,i*3) Then
					vBody = vBody & " selected"
				End If
				'//2015.04.13 정윤정 추가 사용안함 카테고리일 경우 회색으로 보이게
				if rsget("useyn") ="N" then
				    vBody = vBody & " style='color:gray;'"
			    end if  
				vBody = vBody & ">"& rsget("catename") &"</option>" & vbCrLf

				rsget.MoveNext
			next

			vBody = vBody & "</select>"
			if depth=2 then
				vBody = vBody & "</div>"
			end if
		end if

		rsget.close
	next

	Response.Write vBody
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->