<%@ language=vbscript %>
<%
option Explicit
'Response.Buffer = True
Response.CharSet = "euc-kr"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
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
		sqlStr = sqlStr & " FROM [db_academy].[dbo].[tbl_display_cate_Academy] "
		sqlStr = sqlStr & " WHERE depth = '" & i & "' "
		if i>1 then
			sqlStr = sqlStr & " and catecode like '" & left(dispCate,(i-1)*3) & "%'"
		end if

		sqlStr = sqlStr & " order by sortNo, catecode "
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			vBody = vBody & "<select name=""disp" & i & """ class=""formSlt"" onchange=""chgDispCate(this.value)"">" & vbCrLf
			vBody = vBody & "	<option value='" & chkIIF(i>1,left(dispCate,(i-1)*3),"") & "'>::Depth" & i & "::</option>" & vbCrLf
	
			For lp=0 To rsACADEMYget.RecordCount -1
	
				vBody = vBody & "	<option value="""& rsACADEMYget("catecode") &""""
				If CStr(rsACADEMYget("catecode")) = left(dispCate,i*3) Then
					vBody = vBody & " selected"
				End If
				'//2015.04.13 정윤정 추가 사용안함 카테고리일 경우 회색으로 보이게
				if rsACADEMYget("useyn") ="N" then
				    vBody = vBody & " style='color:gray;'"
			    end if    
				vBody = vBody & " >"& rsACADEMYget("catename") &"</option>" & vbCrLf
	
				rsACADEMYget.MoveNext
			next
	
			vBody = vBody & "</select> "
		end if

		rsACADEMYget.close
	next

	Response.Write vBody
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->