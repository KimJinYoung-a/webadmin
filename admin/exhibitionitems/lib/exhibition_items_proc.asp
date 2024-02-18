<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
'// 멀티 저장 처리 페이지 
dim i , lp
dim mastercode , detailcode , mode
dim idx , tempidx
dim itemcount
dim strSQL
dim idxStrSQL

dim arrItemid , tmpArrIid , strErr , actItemid , strRst
dim Ecnt : Ecnt = 0 
dim Scnt : Scnt = 0

mode = request("mode")
mastercode = request("mastercode")
detailcode = request("detailcode")

arrItemid = split(replace(request("chkitem"),vbCrLf,","),",")
if trim(request("itemidarr"))<>"" then
	tmpArrIid = trim(request("itemidarr"))
	if Right(tmpArrIid,1)="," then tmpArrIid=Left(tmpArrIid,Len(tmpArrIid)-1)
	arrItemid = split(tmpArrIid,",")
end if

'// 검수
for lp=0 to ubound(arrItemid)
	if isNumeric(arrItemid(lp)) then
		strSQL = "SELECT itemid FROM db_event.dbo.tbl_exhibition_items with (nolock) where itemid = "& arrItemid(lp) &" and mastercode = '"& mastercode &"' and detailcode = '"& detailcode &"'"
		rsget.Open strSQL, dbget, 1
		if Not rsget.Eof then
			strErr = strErr & chkIIF(strErr<>"",",","") & arrItemid(lp)
			Ecnt=Ecnt+1
		else
			actItemid = actItemid & chkIIF(actItemid<>"",",","") & getNumeric(arrItemid(lp))
			Scnt=Scnt+1
		end if
		rsget.close
	else
		if trim(arrItemid(lp))<>"" then
			strErr = strErr & chkIIF(strErr<>"",",","") & arrItemid(lp)
			Ecnt=Ecnt+1
		end if
	end if
next

if Scnt>0 Then '// 등록할 상품 코드가 있는 경우
    strSQL = " insert into db_event.dbo.tbl_exhibition_items" & vbCrLf
    strSQL = strSQL + " (mastercode , detailcode , itemid , pickitem , adminid) " & vbCrLf
    strSQL = strSQL + " select '" & mastercode & "'" & vbCrLf
    strSQL = strSQL + " ,'" & detailcode & "'" & vbCrLf
    strSQL = strSQL + " ,itemid " & vbCrLf
    strSQL = strSQL + " , 0 " & vbCrLf
    strSQL = strSQL + " ,'" & session("ssBctId") & "'" & vbCrLf
    strSQL = strSQL + " from db_item.dbo.tbl_item" & vbCrLf
    strSQL = strSQL + " where itemid in (" & actItemid & ")" & vbCrLf
    strSQL = strSQL + " 	and isusing='Y'" & vbCrLf
    strSQL = strSQL + " 	and itemid not in (" & vbCrLf
    strSQL = strSQL + " 		select itemid" & vbCrLf
    strSQL = strSQL + " 		from db_event.dbo.tbl_exhibition_items" & vbCrLf
    strSQL = strSQL + " 		where mastercode='" & mastercode & "' and detailcode = '"& detailcode &"' " & vbCrLf
    strSQL = strSQL + " 	)"

    dbget.Execute strSQL
end if 

strRst = "[" & Scnt & "]건 성공"
if Ecnt>0 then strRst = strRst & "\n[" & Ecnt & "]건 실패\n※중복상품코드: " & strErr

Response.Write "<script>" & vbCrLf
Response.Write "alert('" & strRst & "');"& vbCrLf
	if trim(request("itemidarr"))="" then
		Response.Write "opener.location.reload();" & vbCrLf
		Response.Write "window.close();"& vbCrLf
	end if
Response.Write "</script>"
response.End
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->