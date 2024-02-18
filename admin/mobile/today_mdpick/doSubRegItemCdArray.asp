<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
Dim sqlStr, lp, strRst, strErr, Scnt, Ecnt
Dim listidx, arrItemid, actItemid, subSortNo, subIsUsing , gubun , topview
dim tmpArrIid, dealyn

listidx = request("listidx")
gubun = request("gubun")
topview = request("topview")
dealyn = request("dealyn")

If gubun = topview Then
	topview = 1
Else
	topview = 0
End If 

arrItemid = split(replace(request("subItemidArray"),vbCrLf,","),",")
if trim(request("itemidarr"))<>"" then
	tmpArrIid = trim(request("itemidarr"))
	if Right(tmpArrIid,1)="," then tmpArrIid=Left(tmpArrIid,Len(tmpArrIid)-1)
	arrItemid = split(tmpArrIid,",")
end if
subSortNo = request("subSortNo")
subIsUsing = request("subIsUsing")

if listidx="" then
	Call Alert_Return("템플릿 정보가 없습니다.")
	dbget.close(): response.End
end if

if Not(isArray(arrItemid)) then
	Call Alert_Return("상품코드 정보가 잘못되었습니다.")
	dbget.close(): response.End
end if

if subSortNo="" then subSortNo="0"
if subIsUsing="" then subIsUsing="Y"
Scnt=0: Ecnt=0

for lp=0 to ubound(arrItemid)
	if isNumeric(arrItemid(lp)) then
		sqlStr = "SELECT itemid FROM [db_sitemaster].[dbo].tbl_mobile_main_mdpick_item with (nolock) where itemid = "& arrItemid(lp) &" and listidx='" & listidx & "' and gubun='"& gubun &"'"
		rsget.Open sqlStr, dbget, 1
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

if Scnt>0 Then
	If dealyn = "Y" Then
		sqlStr = " insert into [db_sitemaster].[dbo].tbl_mobile_main_mdpick_item" + VbCrlf
		sqlStr = sqlStr + " (listidx, itemid, isusing, sortnum , gubun , topview) " + VbCrlf
		sqlStr = sqlStr + " select '" & listidx & "'" + VbCrlf
		sqlStr = sqlStr + " ,itemid " + VbCrlf
		sqlStr = sqlStr + " ,'" & subIsUsing & "'" + VbCrlf
		sqlStr = sqlStr + " ,'" & subSortNo & "'" + VbCrlf
		sqlStr = sqlStr + " ,'" & gubun & "'" + VbCrlf
		sqlStr = sqlStr + " ,'" & topview & "'" + VbCrlf
		sqlStr = sqlStr + " from db_item.dbo.tbl_item" + VbCrlf
		sqlStr = sqlStr + " where itemid in (" & actItemid & ")" + VbCrlf
		sqlStr = sqlStr + " 	and itemdiv=21" + VbCrlf
		sqlStr = sqlStr + " 	and itemid not in (" + VbCrlf
		sqlStr = sqlStr + " 		select itemid" + VbCrlf
		sqlStr = sqlStr + " 		from [db_sitemaster].[dbo].tbl_mobile_main_mdpick_item" + VbCrlf
		sqlStr = sqlStr + " 		where listidx='" & listidx & "'" + VbCrlf
		sqlStr = sqlStr + " 	)" + VbCrlf
	Else
		sqlStr = " insert into [db_sitemaster].[dbo].tbl_mobile_main_mdpick_item" + VbCrlf
		sqlStr = sqlStr + " (listidx, itemid, isusing, sortnum , gubun , topview) " + VbCrlf
		sqlStr = sqlStr + " select '" & listidx & "'" + VbCrlf
		sqlStr = sqlStr + " ,itemid " + VbCrlf
		sqlStr = sqlStr + " ,'" & subIsUsing & "'" + VbCrlf
		sqlStr = sqlStr + " ,'" & subSortNo & "'" + VbCrlf
		sqlStr = sqlStr + " ,'" & gubun & "'" + VbCrlf
		sqlStr = sqlStr + " ,'" & topview & "'" + VbCrlf
		sqlStr = sqlStr + " from db_item.dbo.tbl_item" + VbCrlf
		sqlStr = sqlStr + " where itemid in (" & actItemid & ")" + VbCrlf
		sqlStr = sqlStr + " 	and isusing='Y'" + VbCrlf
		sqlStr = sqlStr + " 	and itemid not in (" + VbCrlf
		sqlStr = sqlStr + " 		select itemid" + VbCrlf
		sqlStr = sqlStr + " 		from [db_sitemaster].[dbo].tbl_mobile_main_mdpick_item" + VbCrlf
		sqlStr = sqlStr + " 		where listidx='" & listidx & "'" + VbCrlf
		sqlStr = sqlStr + " 	)" + VbCrlf
	End If
	dbget.Execute sqlStr

	'// 페이지정보 최종 수정자 업데이트
	sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_mdpick_list " + VbCrlf
	sqlStr = sqlStr + " Set lastadminid='" & session("ssBctId") & "'" + VbCrlf
	sqlStr = sqlStr + " ,lastupdate=getdate() " + VbCrlf
	sqlStr = sqlStr + " where idx=" + cstr(listidx)
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

<!-- #include virtual="/lib/db/dbclose.asp" -->