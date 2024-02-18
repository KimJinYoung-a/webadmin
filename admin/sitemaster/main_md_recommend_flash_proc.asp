<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
    Session.CodePage  = 949
    Response.CharSet  = "euc-kr"
    Response.AddHeader "Pragma","no-cache"
    Response.AddHeader "cache-control", "no-staff"
    Response.Expires  = -1
%>
<%
Dim sqlStr, lp, strRst, strErr, Scnt, Ecnt
Dim arrLinkitemid , startdate , startdatetime , enddate , enddatetime , disporder , isusing
Dim arrItemid , actItemid

arrLinkitemid   = request("linkitemid")
startdate       = request("startdate")
startdatetime   = request("startdatetime")
enddate         = request("enddate")
enddatetime     = request("enddatetime")
disporder       = request("disporder")
isusing         = request("isusing")

if isusing = "" then isusing = "Y"

arrItemid = split(replace(arrLinkitemid,vbCrLf,","),",")

if Not(isArray(arrItemid)) then
	Call Alert_Return("상품코드 정보가 잘못되었습니다.")
	dbget.close(): response.End
end if

Scnt=0: Ecnt=0

for lp=0 to ubound(arrItemid)
	if isNumeric(arrItemid(lp)) then
		sqlStr = "SELECT linkitemid FROM [db_sitemaster].[dbo].tbl_main_mdchoice_flash with (nolock) where linkitemid = "& arrItemid(lp) &" and convert(varchar(10),startdate,120) = '"& startdate &"' and convert(varchar(10),enddate,120) = '"& enddate &"' "
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

if (startdatetime <> "") then startdate = startdate & " " & startdatetime &":00:00"
if (enddatetime <> "") then enddate = enddate & " " & enddatetime &":00:00"

 if Scnt>0 Then
    sqlStr = " insert into [db_sitemaster].[dbo].tbl_main_mdchoice_flash" & VbCrlf
    sqlStr = sqlStr & " (textinfo, linkitemid, linkinfo,  isusing, disporder , startdate , enddate) " & VbCrlf
    sqlStr = sqlStr & " select itemname" & VbCrlf
    sqlStr = sqlStr & " ,itemid " & VbCrlf
    sqlStr = sqlStr & " ,'/shopping/category_prd.asp?itemid='+convert(varchar,itemid)" & VbCrlf
    sqlStr = sqlStr & " ,'" & isusing & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & disporder & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & startdate & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & enddate & "'" & VbCrlf
    sqlStr = sqlStr & " from db_item.dbo.tbl_item" & VbCrlf
    sqlStr = sqlStr & " where itemid in (" & actItemid & ")" & VbCrlf
    sqlStr = sqlStr & " 	and isusing='Y'" & VbCrlf
	dbget.Execute sqlStr
 end if

strRst = "[" & Scnt & "]건 성공"
if Ecnt>0 then strRst = strRst & "\n[" & Ecnt & "]건 실패\n※중복상품코드: " & strErr

Response.Write "<script>alert('" & strRst & "');opener.location.reload();window.close();</script>"
response.End
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->