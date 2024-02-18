<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
'// 멀티 저장 처리 페이지 
dim i , lp , p
dim mastercode , detailcode , mode
dim idx , tempidx
dim itemcount
dim strSQL
dim idxStrSQL

dim arrItemid , tmpArrIid , strErr , strRst , arrItemopt , tmpArrIopt , dupleArrayID
dim arrItemGubun , tmpArrIgubun
dim actItemid , actItemGubun , actItemOption
dim Ecnt : Ecnt = 0 
dim Scnt : Scnt = 0
dim arrActItemid , arrActItemGubun , arrActItemOption

mode = request("mode")
mastercode = request("mastercode")
detailcode = request("detailcode")

'// array 중복 제거
Function DuplValRemove(ByVal varArr)
 	   Dim dic, items, rtnVal

 	   Set dic = CreateObject("Scripting.Dictionary")
 	   dic.removeall
 	   dic.CompareMode = 0

 	   For Each items In varArr
 	   	   If not dic.Exists(items) Then dic.Add items, items
 	   Next

 	   rtnVal = dic.keys
 	   Set dic = Nothing
 	   DuplValRemove = rtnVal
End Function

if trim(request("itemidarr")) = "" then 
	response.write "<script>등록할 상품이 없습니다.</script>"
	response.end
end if 

if trim(request("itemidarr"))<>"" then
	tmpArrIid = trim(request("itemidarr"))
	if Right(tmpArrIid,1)="," then tmpArrIid=Left(tmpArrIid,Len(tmpArrIid)-1)
	arrItemid = split(tmpArrIid,",")
	dupleArrayID = DuplValRemove(arrItemid)
end if

if trim(request("itemoptarr"))<>"" then
	tmpArrIopt = trim(request("itemoptarr"))
	if Right(tmpArrIopt,1)="," then tmpArrIopt=Left(tmpArrIopt,Len(tmpArrIopt)-1)
	arrItemopt = split(tmpArrIopt,",")
end if

if trim(request("itemgubunarr"))<>"" then
	tmpArrIgubun = trim(request("itemgubunarr"))
	if Right(tmpArrIgubun,1)="," then tmpArrIgubun=Left(tmpArrIgubun,Len(tmpArrIgubun)-1)
	arrItemGubun = split(tmpArrIgubun,",")
end if

'// 중복 입력 검증
for p = 0 to ubound(arrItemid)
	if isNumeric(arrItemid(lp)) then
		strSQL = "SELECT itemid FROM db_item.dbo.tbl_exhibition_item_detail with (nolock) where itemid = "& arrItemid(p) &" and mastercode = '"& mastercode &"'" & vbCrLf
		strSQL = strSQL & "and gubuncode = "& arrItemGubun(p) &" and optioncode = '"& arrItemopt(p) &"' and detailcode = "& detailcode
		rsget.Open strSQL, dbget, 1
		if Not rsget.Eof then
			strErr = strErr & chkIIF(strErr<>"",",","") & arrItemid(p)
			Ecnt=Ecnt+1
		else
			actItemid = actItemid & chkIIF(actItemid<>"",",","") & getNumeric(arrItemid(p))
			actItemGubun = actItemGubun & chkIIF(actItemGubun<>"",",","") & getNumeric(arrItemGubun(p))
			actItemOption = actItemOption & chkIIF(actItemOption<>"",",","") & arrItemopt(p)
			Scnt=Scnt+1
		end if
		rsget.close
	end if
next


'// master 입력
for lp=0 to ubound(dupleArrayID)
	if isNumeric(dupleArrayID(lp)) then
		strSQL = "IF NOT EXISTS(SELECT itemid FROM db_item.dbo.tbl_exhibition_item_master with (nolock) where itemid = "& dupleArrayID(lp) &" and mastercode = '"& mastercode &"')" & vbCrLf
		strSQL = strSQL & "	BEGIN " & vbCrLf
		strSQL = strSQL & "		INSERT INTO db_item.dbo.tbl_exhibition_item_master(mastercode , itemid , itemscore) " & vbCrLf
		strSQL = strSQL & "		SELECT "& mastercode &" , itemid , itemscore FROM db_item.dbo.tbl_item WHERE itemid = "& dupleArrayID(lp) & vbCrLf
		strSQL = strSQL & "	END " & vbCrLf

		dbget.execute strSQL
	end if
next

'// detail 입력
if Scnt > 0 then 
	arrActItemid = split(actItemid,",")
	arrActItemGubun = split(actItemGubun,",")
	arrActItemOption = split(actItemOption,",")

	strSQL = ""
	strSQL = "INSERT INTO db_item.dbo.tbl_exhibition_item_detail" & VbCrlf
	strSQL = strSQL & " (mastercode , gubuncode , itemid , optioncode , detailcode) " & VbCrlf
	strSQL = strSQL & " VALUES " + VbCrlf
	FOR i = 0 TO ubound(arrActItemid)
		strSQL = strSQL & chkiif(i > 0 ,",","")  &" ( "& mastercode &" , "& arrActItemGubun(i) &" , "& arrActItemid(i) &" , '"& arrActItemOption(i) &"' , "& detailcode &" ) " & VbCrlf
	NEXT 
	dbget.Execute strSQL
end if 

strRst = "[" & Scnt & "]건 입력"
if Ecnt > 0 then strRst = strRst & "\n[" & Ecnt & "]건 실패\n※중복상품코드: " & strErr

Response.Write "<script>" & vbCrLf
Response.Write "alert('" & strRst & "');"& vbCrLf
Response.Write "parent.location.reload();" & vbCrLf
'Response.Write "parent.window.close();"& vbCrLf
Response.Write "</script>"
response.End
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->