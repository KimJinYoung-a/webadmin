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
Dim mainIdx, arrItemid, actItemid, subSortNo, subIsUsing
dim tmpArrIid

mainIdx = request("mainIdx")
arrItemid = split(replace(request("subItemidArray"),vbCrLf,","),",")
if trim(request("itemidarr"))<>"" then
	tmpArrIid = trim(request("itemidarr"))
	if Right(tmpArrIid,1)="," then tmpArrIid=Left(tmpArrIid,Len(tmpArrIid)-1)
	arrItemid = split(tmpArrIid,",")
end if
subSortNo = request("subSortNo")
subIsUsing = request("subIsUsing")

if mainIdx="" then
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
		actItemid = actItemid & chkIIF(actItemid<>"",",","") & getNumeric(arrItemid(lp))
		Scnt=Scnt+1
	else
		if trim(arrItemid(lp))<>"" then
			strErr = strErr & chkIIF(strErr<>"",",","") & arrItemid(lp)
			Ecnt=Ecnt+1
		end if
	end if
next

if Scnt>0 then
    sqlStr = " insert into [db_sitemaster].[dbo].tbl_cms_subInfo" + VbCrlf
    sqlStr = sqlStr + " (mainIdx, subLinkUrl, subText1, subText2, subItemid, subVideoUrl, subBGColor, subImageDesc " + VbCrlf
    sqlStr = sqlStr + " , subSortNo, subRegUserid, subRegdate, subIsUsing) " + VbCrlf
    sqlStr = sqlStr + " select '" & mainIdx & "'" + VbCrlf
    sqlStr = sqlStr + " ,'','',''" + VbCrlf
    sqlStr = sqlStr + " ,itemid " + VbCrlf
    sqlStr = sqlStr + " ,'','',''" + VbCrlf
    sqlStr = sqlStr + " ,'" & subSortNo & "'" + VbCrlf
    sqlStr = sqlStr + " ,'" & session("ssBctId") & "', getdate() " + VbCrlf
    sqlStr = sqlStr + " ,'" & subIsUsing & "'" + VbCrlf
    sqlStr = sqlStr + " from db_item.dbo.tbl_item" + VbCrlf
    sqlStr = sqlStr + " where itemid in (" & actItemid & ")" + VbCrlf
    sqlStr = sqlStr + " 	and isusing='Y'" + VbCrlf
    sqlStr = sqlStr + " 	and itemid not in (" + VbCrlf
    sqlStr = sqlStr + " 		select subItemid" + VbCrlf
    sqlStr = sqlStr + " 		from [db_sitemaster].[dbo].tbl_cms_subInfo" + VbCrlf
    sqlStr = sqlStr + " 		where mainIdx='" & mainIdx & "'" + VbCrlf
    sqlStr = sqlStr + " 	)" + VbCrlf

	dbget.Execute sqlStr

	'// 페이지정보 최종 수정자 업데이트
	sqlStr = "Update [db_sitemaster].[dbo].tbl_cms_mainInfo " + VbCrlf
	sqlStr = sqlStr + " Set mainLastModiUserid='" & session("ssBctId") & "'" + VbCrlf
	sqlStr = sqlStr + " ,mainLastModiDate=getdate() " + VbCrlf
	sqlStr = sqlStr + " where mainIdx=" + cstr(mainIdx)
	dbget.Execute sqlStr
end if

strRst = "[" & Scnt & "]건 성공"
if Ecnt>0 then strRst = strRst & "\n[" & Ecnt & "]건 실패\n※실패건: " & strErr

Response.Write "<script language='javascript'>" & vbCrLf
Response.Write "alert('" & strRst & "\n저장되었습니다.');"& vbCrLf
	if trim(request("itemidarr"))="" then
		Response.Write "opener.location.reload();" & vbCrLf
		Response.Write "window.close();"& vbCrLf
	end if
Response.Write "</script>"
response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->