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
Dim sqlStr
Dim idx, mode, bannerTitle, startDate, expireDate, closeType, bannerType, bannerImg, linkType, linkTitle, linkURL, targetOS, targetType, importance, isUsing, status
Dim startDateSecond, expireDateSecond

idx        	= requestCheckVar(request("idx"),10)
mode       	= requestCheckVar(request("mode"),4)
bannerTitle	= requestCheckVar(request("bannerTitle"),200)
startDate  	= requestCheckVar(request("startDate"),10)
startDateSecond = requestCheckVar(request("StartDateSecond"), 30)
expireDate 	= requestCheckVar(request("expireDate"),10)
expireDateSecond = requestCheckVar(request("expireDateSecond"), 30)
closeType  	= requestCheckVar(request("closeType"),1)
bannerType 	= requestCheckVar(request("bannerType"),1)
bannerImg  	= requestCheckVar(request("bannerImg"),200)
linkType   	= requestCheckVar(request("linkType"),8)
linkTitle  	= requestCheckVar(request("linkTitle"),60)
linkURL    	= requestCheckVar(request("linkURL"),200)
targetOS   	= requestCheckVar(request("targetOS"),10)
targetType 	= requestCheckVar(request("targetType"),2)
importance 	= requestCheckVar(request("importance"),8)
isUsing    	= requestCheckVar(request("isUsing"),1)
status     	= requestCheckVar(request("status"),8)

'Banner Link URL 간소화 처리
linkURL = trim(Lcase(linkURL))
if left(linkURL,1)<>"/" then
	if left(linkURL,4)="http" then
		linkURL = replace(linkURL,"https","http")
		linkURL = replace(linkURL,"http://testm.10x10.co.kr/apps/appcom/wish/web2014","")
		linkURL = replace(linkURL,"http://m.10x10.co.kr/apps/appcom/wish/web2014","")
		linkURL = replace(linkURL,"http://testm.10x10.co.kr","")
		linkURL = replace(linkURL,"http://m.10x10.co.kr","")
		linkURL = replace(linkURL,"http://2015www.10x10.co.kr","")
		linkURL = replace(linkURL,"http://www.10x10.co.kr","")
		linkURL = replace(linkURL,"http://10x10.co.kr","")
	end if
end if

'// 처리 분기
Select Case mode
	Case "add"
	'신규 등록
	sqlStr = "Insert into [db_sitemaster].[dbo].tbl_app_startupBanner "
	sqlStr = sqlStr & "(bannerTitle, startDate, expireDate, closeType, bannerType, bannerImg, linkType, linkTitle, linkURL, targetOS, targetType, importance, isUsing, status) "
	sqlStr = sqlStr & "values ("
	sqlStr = sqlStr & "N'" & bannerTitle & "'"
	sqlStr = sqlStr & ",'" & startDate &" "& startDateSecond &"'"
	sqlStr = sqlStr & ",'" & expireDate &" "& expireDateSecond &"'"
	sqlStr = sqlStr & ",'" & closeType & "'"
	sqlStr = sqlStr & ",'" & bannerType & "'"
	sqlStr = sqlStr & ",'" & bannerImg & "'"
	sqlStr = sqlStr & ",'" & linkType & "'"
	sqlStr = sqlStr & ",N'" & linkTitle & "'"
	sqlStr = sqlStr & ",'" & linkURL & "'"
	sqlStr = sqlStr & ",'" & targetOS & "'"
	sqlStr = sqlStr & ",'" & targetType & "'"
	sqlStr = sqlStr & "," & importance
	sqlStr = sqlStr & ",'" & isUsing & "'"
	sqlStr = sqlStr & "," & status
	sqlStr = sqlStr & ")"

	dbget.Execute(sqlStr)

	Case "modi"
	'수정
	if Not(idx="" or isNull(idx)) then

		sqlStr = "Update [db_sitemaster].[dbo].tbl_app_startupBanner Set "
		sqlStr = sqlStr & "bannerTitle =N'" & bannerTitle & "'"
		sqlStr = sqlStr & ", startDate ='" & startDate&" "&startDateSecond &"'"
		sqlStr = sqlStr & ", expireDate='" & expireDate&" "&expireDateSecond &"'"
		sqlStr = sqlStr & ", closeType ='" & closeType & "'"
		sqlStr = sqlStr & ", bannerType='" & bannerType & "'"
		sqlStr = sqlStr & ", bannerImg ='" & bannerImg & "'"
		sqlStr = sqlStr & ", linkType  ='" & linkType & "'"
		sqlStr = sqlStr & ", linkTitle =N'" & linkTitle & "'"
		sqlStr = sqlStr & ", linkURL   ='" & linkURL & "'"
		sqlStr = sqlStr & ", targetOS  ='" & targetOS & "'"
		sqlStr = sqlStr & ", targetType='" & targetType & "'"
		sqlStr = sqlStr & ", importance=" & importance
		sqlStr = sqlStr & ", isUsing   ='" & isUsing & "'"
		sqlStr = sqlStr & ", status    =" & status
		sqlStr = sqlStr & " Where idx=" & idx

		dbget.Execute(sqlStr)

	end if
end Select

''response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>opener.history.go(0); self.close();</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->