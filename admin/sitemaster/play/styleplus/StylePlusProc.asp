<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description :  play
' History : 2013.09.03 이종화 생성
'			2014.10.21 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim idx, listimg , viewtitle , reservationdate , state , mode , lastupdate, viewno , titleimg , worktext
Dim viewimg1 ,viewimg2 ,viewimg3 ,viewimg4 ,viewimg5, partwdid , partmdid , lastadminid, style_html_m
dim sqlStr
	lastadminid = session("ssBctId")
	idx		= RequestCheckVar(request("idx"),10)
	listimg	= RequestCheckVar(request("stylelistimg"),120)
	viewtitle = RequestCheckVar(request("viewtitle"),200)
	reservationdate = RequestCheckVar(request("reservationdate"),10)
	state = RequestCheckVar(request("state"),120)
	viewno = RequestCheckVar(request("viewno"),120)
	titleimg = RequestCheckVar(request("styletitleimg"),120)
	worktext = RequestCheckVar(request("worktext"),800)
	viewimg1 = RequestCheckVar(request("styleviewimg1"),120)
	viewimg2 = RequestCheckVar(request("styleviewimg2"),120)
	viewimg3 = RequestCheckVar(request("styleviewimg3"),120)
	viewimg4 = RequestCheckVar(request("styleviewimg4"),120)
	viewimg5 = RequestCheckVar(request("styleviewimg5"),120)
	partmdid = RequestCheckVar(request("partmdid"),32)
	partwdid = RequestCheckVar(request("partwdid"),32)
	mode = RequestCheckVar(request("mode"),10)
	style_html_m = request("style_html_m")

if idx = "" then
	idx = 0
end If

if idx = 0 then
	mode = "add"
else
	mode = "edit"
end if

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

if (mode = "add") then
    sqlStr = " insert into db_sitemaster.dbo.tbl_play_style_list " + VbCrlf
    sqlStr = sqlStr + " (listimg , textimg , viewimg1 , viewimg2 , viewimg3 , viewimg4 , viewimg5 " + VbCrlf
	sqlStr = sqlStr + " ,viewtitle , reservationdate , state , viewno, worktext , partmdid , partwdid )" + VbCrlf
    sqlStr = sqlStr + " values(" + VbCrlf
    sqlStr = sqlStr + " '" + listimg + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + titleimg + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewimg1 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewimg2 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewimg3 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewimg4 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewimg5 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewtitle + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + reservationdate + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + state + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewno + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + worktext + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + partmdid + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + partwdid + "'" + VbCrlf
    sqlStr = sqlStr + " )"

	'response.write sqlStr & "<br>"
    dbget.Execute sqlStr

	sqlStr = "select IDENT_CURRENT('db_sitemaster.dbo.tbl_play_style_list') as idx"
	
	'response.write sqlStr & "<br>"
	rsget.Open sqlStr, dbget, 1
	If Not Rsget.Eof then
		idx = rsget("idx")
	end if
	rsget.close

elseif mode = "edit" Then
	sqlStr = " update  db_sitemaster.dbo.tbl_play_style_list " + VbCrlf
	sqlStr = sqlStr + " set " + VbCrlf
	sqlStr = sqlStr + " listimg='" + listimg + "'" + VbCrlf
	sqlStr = sqlStr + " ,textimg='" + titleimg + "'" + VbCrlf
	sqlStr = sqlStr + " ,viewimg1='" + viewimg1 + "'" + VbCrlf
	sqlStr = sqlStr + " ,viewimg2='" + viewimg2 + "'" + VbCrlf
	sqlStr = sqlStr + " ,viewimg3='" + viewimg3 + "'" + VbCrlf
	sqlStr = sqlStr + " ,viewimg4='" + viewimg4 + "'" + VbCrlf
	sqlStr = sqlStr + " ,viewimg5='" + viewimg5 + "'" + VbCrlf
	sqlStr = sqlStr + " ,viewtitle='" + viewtitle + "'" + VbCrlf
	sqlStr = sqlStr + " ,reservationdate='" + reservationdate + "'" + VbCrlf
	sqlStr = sqlStr + " ,state='" + state + "'" + VbCrlf
	sqlStr = sqlStr + " ,viewno='" + viewno + "'" + VbCrlf
	sqlStr = sqlStr + " ,worktext='" + worktext + "'" + VbCrlf
	sqlStr = sqlStr + " ,partmdid='" + partmdid + "'" + VbCrlf
	sqlStr = sqlStr + " ,partwdid='" + partwdid + "'" + VbCrlf
	sqlStr = sqlStr + " ,lastadminid='" + lastadminid + "'" + VbCrlf
	sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
	sqlStr = sqlStr + " ,style_html_m='" + html2db(style_html_m) + "'" + VbCrlf
	sqlStr = sqlStr + " where styleidx=" + CStr(idx)

   'response.write sqlStr & "<br>"
   dbget.Execute sqlStr
end if

response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>location.href='" & manageUrl & "/admin/sitemaster/play/styleplus/popstyleplusEdit.asp?idx=" + Cstr(idx) + "&reload=on'</script>"
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->	
