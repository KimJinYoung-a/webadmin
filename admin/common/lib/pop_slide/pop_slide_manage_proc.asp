<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기획전 슬라이드 처리 페이지
' History : 2019-02-19 이종화
'###########################################################
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
Dim sqlStr
Dim idx, mode
Dim menu , mastercode , detailcode , device , titlename , bannerImg , lcolor , rcolor , isvideo
dim videohtml , linkurl , eventid , startdate , enddate , sorting , isusing
dim i
dim sIdx , sSortNo , sIsUsing , subtitlename , titlecolor

idx        	= requestCheckVar(request("idx"),10)
mode       	= requestCheckVar(request("mode"),4)
menu       	= requestCheckVar(request("menu"),15)
mastercode 	= requestCheckVar(request("mastercode"),10)
detailcode 	= requestCheckVar(request("detailcode"),10)
device  	= requestCheckVar(request("device"),1)
titlename	= requestCheckVar(request("titlename"),100)
bannerImg  	= requestCheckVar(request("bannerImg"),200)
lcolor  	= requestCheckVar(request("lcolor"),6)
rcolor  	= requestCheckVar(request("rcolor"),6)
isvideo     = requestCheckVar(request("isvideo"),1)
videohtml   = html2db(request("videohtml"))
linkurl     = requestCheckVar(request("linkurl"),125)
eventid     = requestCheckVar(request("evt_code"),10)
startdate   = requestCheckVar(request("StartDate"),10)
enddate     = requestCheckVar(request("EndDate"),10)
sorting     = requestCheckVar(request("sorting"),3)
isusing     = requestCheckVar(request("isusing"),1)
subtitlename= requestCheckVar(request("subtitlename"),100)
titlecolor= requestCheckVar(request("titlecolor"),6)


'// 처리 분기
Select Case mode
	Case "add"
	'신규 등록
	sqlStr = "INSERT INTO db_event.dbo.tbl_slide_list "
	sqlStr = sqlStr & "(menu , device , mastercode , detailcode , titlename , lcolor , rcolor , imageurl , isvideo , videohtml , linkurl , eventid , isusing ,  sorting , startdate , enddate , subtitlename,titlecolor) "
	sqlStr = sqlStr & "values ("
    sqlStr = sqlStr & "'"& menu &"'"
    sqlStr = sqlStr & ",'"& device &"'"
    sqlStr = sqlStr & ",'"& mastercode &"'"
    sqlStr = sqlStr & ",'"& detailcode &"'"
    sqlStr = sqlStr & ",N'"& titlename &"'"
    sqlStr = sqlStr & ",'"& lcolor &"'"
    sqlStr = sqlStr & ",'"& rcolor &"'"
    sqlStr = sqlStr & ",'"& bannerImg &"'"
    sqlStr = sqlStr & ",'"& isvideo &"'"
    sqlStr = sqlStr & ",'"& videohtml &"'"
    sqlStr = sqlStr & ",'"& linkurl &"'"
    sqlStr = sqlStr & ",'"& eventid &"'"
    sqlStr = sqlStr & ",'"& isusing &"'"
    sqlStr = sqlStr & ",'"& sorting &"'"
    sqlStr = sqlStr & ",'"& startdate &" 00:00:000'"
    sqlStr = sqlStr & ",'"& enddate &" 00:00:000'"
	sqlStr = sqlStr & ",N'"& subtitlename &"'"
	sqlStr = sqlStr & ",'"& titlecolor &"'"
	sqlStr = sqlStr & ")"

	dbget.Execute(sqlStr)

	Case "modi"
	'수정
	if Not(idx="" or isNull(idx)) then

		sqlStr = "UPDATE db_event.dbo.tbl_slide_list SET "
		sqlStr = sqlStr & "menu 		='" & menu & "'"
		sqlStr = sqlStr & ", device 	='" & device &"'"
		sqlStr = sqlStr & ", mastercode	='" & mastercode &"'"
		sqlStr = sqlStr & ", detailcode ='" & detailcode & "'"
		sqlStr = sqlStr & ", titlename	=N'"& titlename &"'"
		sqlStr = sqlStr & ", lcolor 	='" & lcolor & "'"
		sqlStr = sqlStr & ", rcolor  	='" & rcolor & "'"
		sqlStr = sqlStr & ", imageurl   ='" & bannerImg & "'"
		sqlStr = sqlStr & ", isvideo    ='" & isvideo & "'"
		sqlStr = sqlStr & ", videohtml  ='" & videohtml & "'"
		sqlStr = sqlStr & ", linkurl	='" & linkurl & "'"
		sqlStr = sqlStr & ", eventid	='" & eventid & "'"
		sqlStr = sqlStr & ", isusing    ='" & isusing & "'"
		sqlStr = sqlStr & ", sorting    ='" & sorting & "'"
		sqlStr = sqlStr & ", startdate  ='" & startdate & "'"
		sqlStr = sqlStr & ", enddate    ='" & enddate & "'"
		sqlStr = sqlStr & ", subtitlename =N'"& subtitlename &"'"
		sqlStr = sqlStr & ", titlecolor = '"& titlecolor &"'"
		sqlStr = sqlStr & ", lastupdate = getdate()"
		sqlStr = sqlStr & " Where idx=" & idx

		dbget.Execute(sqlStr)

	end if

	Case "sort"
		'// 리스트에서수정
		for i=1 to request.form("chkIdx").count
			sIdx = request.form("chkIdx")(i)
			sSortNo = request.form("sort"&sIdx)
			sIsUsing = request.form("use"&sIdx)
			if sSortNo="" then sSortNo="99"
			if sIsUsing="" then sIsUsing="0"

			sqlStr = sqlStr & " UPDATE db_event.dbo.tbl_slide_list SET "
			sqlStr = sqlStr & " sorting=" & sSortNo & ""
			sqlStr = sqlStr & " ,isusing='" & sIsUsing & "'"
			sqlStr = sqlStr & " Where idx='" & sIdx & "';" & vbCrLf
		Next

		If sqlStr <> "" then
			dbget.Execute sqlStr
		Else
			Call Alert_return("저장할 내용이 없습니다.")
			dbget.Close: Response.End
		End If 
	
	Case "idel"
		sIdx = request.form("chkIdx")

		sqlStr = "DELETE FROM db_event.dbo.tbl_slide_list WHERE idx='"& sIdx &"'"
		dbget.Execute sqlStr

end Select

dim returnurl
response.write "<script>alert('저장되었습니다.');</script>"
Select Case mode
	Case "add" , "modi"
	returnurl = "opener.history.go(0);self.close();"
	case "sort" , "idel"
	returnurl = "location.href='"&requestCheckVar(request("backurl"),300)&"';"
end select 
response.write "<script> "& returnurl &" </script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->