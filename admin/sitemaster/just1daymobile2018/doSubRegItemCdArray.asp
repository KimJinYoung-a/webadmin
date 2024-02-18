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
Dim listidx, arrItemid, actItemid, subSortNo, subIsUsing
dim tmpArrIid, ptype, itemcount, itemdiv

listidx = request("listidx")
arrItemid = split(replace(request("subItemidArray"),vbCrLf,","),",")
if trim(request("itemidarr"))<>"" then
	tmpArrIid = trim(request("itemidarr"))
	if Right(tmpArrIid,1)="," then tmpArrIid=Left(tmpArrIid,Len(tmpArrIid)-1)
	arrItemid = split(tmpArrIid,",")
end if
subSortNo = request("subSortNo")
subIsUsing = request("subIsUsing")
ptype = request("ptype")
itemcount = request("itemcount")

if listidx="" then
	Call Alert_Return("템플릿 정보가 없습니다.")
	dbget.close(): response.End
end if

if Not(isArray(arrItemid)) then
	Call Alert_Return("상품코드 정보가 잘못되었습니다.")
	dbget.close(): response.End
end If

If ubound(arrItemid)+1 > 1 Then
	Call Alert_Return("상품은 한번에 하나씩만 입력하실 수 있습니다.")
	dbget.close(): response.End
End If

if subSortNo="" then subSortNo="0"
if subIsUsing="" then subIsUsing="Y"
Scnt=0: Ecnt=0

for lp=0 to ubound(arrItemid)
	if isNumeric(arrItemid(lp)) then
		actItemid = actItemid & chkIIF(actItemid<>"",",","") & getNumeric(arrItemid(lp))
		Scnt=Scnt+1
	else
		if trim(arrItemid(lp))<>"" then
			strErr = strErr & chkIIF(strErr<>"",",","") & arrItemid(lp)
			Ecnt=Ecnt+1
		end if
	end if
Next

'// 딜 상품인지 확인한다.
sqlStr = " select itemdiv from db_item.dbo.tbl_item Where itemid='"&arrItemid(0)&"' "
rsget.Open sqlStr, dbget, 1
	itemdiv = rsget("itemdiv")
rsget.close

'if Scnt>0 then
'    sqlStr = " insert into [db_sitemaster].[dbo].tbl_pc_main_just1day_item" + VbCrlf
'    sqlStr = sqlStr + " (listidx, itemid, itemname, isusing, sortnum ) " + VbCrlf
'    sqlStr = sqlStr + " select '" & listidx & "'" + VbCrlf
'    sqlStr = sqlStr + " ,itemid, itemname " + VbCrlf
'    sqlStr = sqlStr + " ,'" & subIsUsing & "'" + VbCrlf
'    sqlStr = sqlStr + " ,'" & subSortNo & "'" + VbCrlf
'    sqlStr = sqlStr + " from db_item.dbo.tbl_item" + VbCrlf
'    sqlStr = sqlStr + " where itemid in (" & actItemid & ")" + VbCrlf
'	If ptype<>"just1day" Then
'    sqlStr = sqlStr + " 	and isusing='Y'" + VbCrlf
'	End If
'    sqlStr = sqlStr + " 	and itemid not in (" + VbCrlf
'    sqlStr = sqlStr + " 		select itemid" + VbCrlf
'    sqlStr = sqlStr + " 		from [db_sitemaster].[dbo].tbl_pc_main_just1day_item" + VbCrlf
'    sqlStr = sqlStr + " 		where listidx='" & listidx & "'" + VbCrlf
'    sqlStr = sqlStr + " 	)" + VbCrlf

'	dbget.Execute sqlStr

	'// 페이지정보 최종 수정자 업데이트
'	sqlStr = "Update [db_sitemaster].[dbo].tbl_pc_main_just1day_list " + VbCrlf
'	sqlStr = sqlStr + " Set lastadminid='" & session("ssBctId") & "'" + VbCrlf
'	sqlStr = sqlStr + " ,lastupdate=getdate() " + VbCrlf
'	sqlStr = sqlStr + " where idx=" + cstr(listidx)
'end if

strRst = "[" & Scnt & "]건 성공"
if Ecnt>0 then strRst = strRst & "\n[" & Ecnt & "]건 실패\n※실패건: " & strErr

Response.Write "<script>" & vbCrLf
'Response.Write "alert('" & strRst & "\n저장되었습니다.');"& vbCrLf
	Response.Write "parent.opener.location.href='/admin/sitemaster/just1daymobile2018/popSubItemEdit.asp?listidx="&listidx&"&itemdiv="&itemdiv&"&itemid="&arrItemid(0)&"';" & vbCrLf
	Response.Write "parent.window.close();"& vbCrLf
	if trim(request("itemidarr"))="" then
		Response.Write "opener.location.reload();" & vbCrLf
		Response.Write "window.close();"& vbCrLf
	end if
Response.Write "</script>"
response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->