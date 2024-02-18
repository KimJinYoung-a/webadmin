<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim idx, listimg , viewtitle , viewtext , reservationdate , state , mode , lastupdate, videourlM
Dim viewno , worktext , videourl , partwdid

idx		= RequestCheckVar(request("idx"),10)
listimg	= RequestCheckVar(request("pdlistimg"),150)
viewtitle = RequestCheckVar(request("viewtitle"),50)
viewtext = RequestCheckVar(request("viewtext"),800)
reservationdate = RequestCheckVar(request("reservationdate"),10)
state = RequestCheckVar(request("state"),120)
viewno = RequestCheckVar(request("viewno"),120)
worktext = RequestCheckVar(request("worktext"),500)
videourl = request("videourl")
partwdid = RequestCheckVar(request("partwdid"),32)
videourlM = RequestCheckVar(request("videourlM"),500)


mode = RequestCheckVar(request("mode"),10)

if idx = "" then
	idx = 0
end If

if idx = 0 then
	mode = "add"
else
	mode = "edit"
end if

dim sqlStr

if (mode = "add") then

    sqlStr = " insert into [db_sitemaster].[dbo].tbl_play_video_clip" + VbCrlf
    sqlStr = sqlStr + " (listimg,viewtitle,viewtext,reservationdate,state,viewno, worktext , videourl , partwdid, videourlM)" + VbCrlf
    sqlStr = sqlStr + " values(" + VbCrlf
    sqlStr = sqlStr + " '" + listimg + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewtitle + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewtext + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + reservationdate + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + state + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + viewno + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + worktext + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + videourl + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + partwdid + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + videourlM + "'" + VbCrlf
    sqlStr = sqlStr + " )"

	'response.write sqlStr
    dbget.Execute sqlStr

	sqlStr = "select IDENT_CURRENT('[db_sitemaster].[dbo].tbl_play_video_clip') as idx"
	rsget.Open sqlStr, dbget, 1
	If Not Rsget.Eof then
		idx = rsget("idx")
	end if
	rsget.close

elseif mode = "edit" Then

   sqlStr = " update  db_sitemaster.dbo.tbl_play_video_clip " + VbCrlf
   sqlStr = sqlStr + " set " + VbCrlf
   sqlStr = sqlStr + " listimg='" + listimg + "'" + VbCrlf
   sqlStr = sqlStr + " ,viewtitle='" + viewtitle + "'" + VbCrlf
   sqlStr = sqlStr + " ,viewtext='" + viewtext + "'" + VbCrlf
   sqlStr = sqlStr + " ,reservationdate='" + reservationdate + "'" + VbCrlf
   sqlStr = sqlStr + " ,state='" + state + "'" + VbCrlf
   sqlStr = sqlStr + " ,viewno='" + viewno + "'" + VbCrlf
   sqlStr = sqlStr + " ,worktext='" + worktext + "'" + VbCrlf
   sqlStr = sqlStr + " ,videourl='" + videourl + "'" + VbCrlf
   sqlStr = sqlStr + " ,partwdid='" + partwdid + "'" + VbCrlf
   sqlStr = sqlStr + " ,videourlM='" + videourlM + "'" + VbCrlf
   sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf

   sqlStr = sqlStr + " where vidx=" + CStr(idx)
   dbget.Execute sqlStr

end if


dim referer
referer = request.ServerVariables("HTTP_REFERER")
response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>location.href='" & manageUrl & "/admin/sitemaster/play/videoclip/popvideoclipEdit.asp?idx=" + Cstr(idx) + "&reload=on'</script>"
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->	
