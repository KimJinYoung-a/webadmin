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
dim gidx , idx, viewtitle , reservationdate , state , mode , position
Dim viewno , titleimg , worktext
Dim partwdid , partmkid
dim sqlStr
dim referer
Dim playmainimg , viewthumbimg1 , viewthumbimg2 , viewbgimg , downsideimg1 , downsideimg2 , subBGColor , viewcontents , mainTopBGColor , mo_contents , mo_exec_check , exec_check , exec_filepath
Dim myplayimg

referer = request.ServerVariables("HTTP_REFERER")

idx		= RequestCheckVar(request("idx"),10)
gidx		= RequestCheckVar(request("gidx"),10)
titleimg	= RequestCheckVar(request("groundtitleimg"),150)
viewtitle = RequestCheckVar(request("viewtitle"),50)
reservationdate = RequestCheckVar(request("reservationdate"),10)
state = RequestCheckVar(request("state"),120)
viewno = RequestCheckVar(request("viewno"),120)
worktext = RequestCheckVar(request("worktext"),500)
viewcontents = html2db(request("viewcontents"))
mo_contents = html2db(Trim(request("mo_contents")))
mo_exec_check = RequestCheckVar(request("mo_exec_check"),1)
exec_check = RequestCheckVar(request("exec_check"),1)
exec_filepath = RequestCheckVar(request("exec_filepath"),50)

partmkid = RequestCheckVar(request("selMId"),32)
partwdid = RequestCheckVar(request("partwdid"),32)

mode = RequestCheckVar(request("mode"),10)
position = RequestCheckVar(request("position"),10)

playmainimg		= RequestCheckVar(request("playmainimg"),150)
viewthumbimg1	= RequestCheckVar(request("beforeimg"),150)
viewthumbimg2	= RequestCheckVar(request("afterimg"),150)
viewbgimg		= RequestCheckVar(request("topbgimg"),150)
downsideimg1	= RequestCheckVar(request("sideltimg"),150)
downsideimg2	= RequestCheckVar(request("sidertimg"),150)

myplayimg	= RequestCheckVar(request("myplayimg"),150)

subBGColor	= RequestCheckVar(request("subBGColor"),150)
mainTopBGColor	= RequestCheckVar(request("mainTopBGColor"),150)

if idx = "" then
	idx = 0
end If

if idx = 0 then
	mode = "add"
else
	mode = "edit"
end if

If position  = "main" Then 'main등록
	if (mode = "add") then

		sqlStr = " insert into db_sitemaster.dbo.tbl_play_ground_main " + VbCrlf
		sqlStr = sqlStr + " (viewno , viewtitle , titleimg , mainimg , reservationdate , partmkid , partwdid , state , worktext ) " + VbCrlf
		sqlStr = sqlStr + " values(" + VbCrlf
		sqlStr = sqlStr + " '" + viewno + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + viewtitle + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + titleimg + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + playmainimg + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + reservationdate + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + partmkid + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + partwdid + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + state + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + worktext + "'" + VbCrlf
		sqlStr = sqlStr + " )"

		'response.write sqlStr
		dbget.Execute sqlStr

		sqlStr = "select IDENT_CURRENT('db_sitemaster.dbo.tbl_play_ground_main') as idx"
		rsget.Open sqlStr, dbget, 1
		If Not Rsget.Eof then
			idx = rsget("idx")
		end if
		rsget.close

	elseif mode = "edit" Then

	   sqlStr = " update  db_sitemaster.dbo.tbl_play_ground_main " + VbCrlf
	   sqlStr = sqlStr + " set " + VbCrlf
	   sqlStr = sqlStr + " titleimg='" + titleimg + "'" + VbCrlf
	   sqlStr = sqlStr + " ,mainimg='" + playmainimg + "'" + VbCrlf
	   sqlStr = sqlStr + " ,viewtitle='" + viewtitle + "'" + VbCrlf
	   sqlStr = sqlStr + " ,reservationdate='" + reservationdate + "'" + VbCrlf
	   sqlStr = sqlStr + " ,state='" + state + "'" + VbCrlf
	   sqlStr = sqlStr + " ,viewno='" + viewno + "'" + VbCrlf
	   sqlStr = sqlStr + " ,worktext='" + worktext + "'" + VbCrlf
	   sqlStr = sqlStr + " ,partmkid='" + partmkid + "'" + VbCrlf
	   sqlStr = sqlStr + " ,partwdid='" + partwdid + "'" + VbCrlf

	   sqlStr = sqlStr + " where gidx=" + CStr(idx)
	   dbget.Execute sqlStr

	end If

	response.write "<script>alert('저장되었습니다.');</script>"
	response.write "<script>location.href='" & manageUrl & "/admin/sitemaster/play/ground/groundEdit.asp?idx=" + Cstr(idx) + "&reload=on'</script>"
Else  'sub등록
	if (mode = "add") then

		sqlStr = " insert into db_sitemaster.dbo.tbl_play_ground_sub " + VbCrlf
		sqlStr = sqlStr + " (gidx , viewno , viewtitle , state , reservationdate , partMKid , partWDid , worktext)" + VbCrlf
		sqlStr = sqlStr + " values(" + VbCrlf
		sqlStr = sqlStr + " '" + gidx + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + viewno + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + viewtitle + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + state + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + reservationdate + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + partmkid + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + partwdid + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + worktext + "'" + VbCrlf
		sqlStr = sqlStr + " )"

		'response.write sqlStr
		dbget.Execute sqlStr

		sqlStr = "select IDENT_CURRENT('db_sitemaster.dbo.tbl_play_ground_sub') as idx"
		rsget.Open sqlStr, dbget, 1
		If Not Rsget.Eof then
			idx = rsget("idx")
		end if
		rsget.close

	elseif mode = "edit" Then

	   sqlStr = " update  db_sitemaster.dbo.tbl_play_ground_sub " + VbCrlf
	   sqlStr = sqlStr + " set " + VbCrlf
	   sqlStr = sqlStr + " gidx=" + gidx + "" + VbCrlf
	   sqlStr = sqlStr + " ,viewno='" + viewno + "'" + VbCrlf
	   sqlStr = sqlStr + " ,viewtitle='" + viewtitle + "'" + VbCrlf
	   sqlStr = sqlStr + " ,state='" + state + "'" + VbCrlf
	   sqlStr = sqlStr + " ,reservationdate='" + reservationdate + "'" + VbCrlf
	   sqlStr = sqlStr + " ,partmkid='" + partmkid + "'" + VbCrlf
	   sqlStr = sqlStr + " ,partwdid='" + partwdid + "'" + VbCrlf
	   sqlStr = sqlStr + " ,worktext='" + worktext + "'" + VbCrlf
	   sqlStr = sqlStr + " ,playmainimg='" + playmainimg + "'" + VbCrlf
	   sqlStr = sqlStr + " ,viewthumbimg1='" + viewthumbimg1 + "'" + VbCrlf
	   sqlStr = sqlStr + " ,viewthumbimg2='" + viewthumbimg2 + "'" + VbCrlf
	   sqlStr = sqlStr + " ,viewbgimg='" + viewbgimg + "'" + VbCrlf
	   sqlStr = sqlStr + " ,downsideimg1='" + downsideimg1 + "'" + VbCrlf
	   sqlStr = sqlStr + " ,downbgcolor='" + subBGColor + "'" + VbCrlf
	   sqlStr = sqlStr + " ,viewcontents='" + viewcontents + "'" + VbCrlf
	   sqlStr = sqlStr + " ,mainbgcolor='" + mainTopBGColor + "'" + VbCrlf
   	   sqlStr = sqlStr + " ,myplayimg='" + myplayimg + "'" + VbCrlf
	   sqlStr = sqlStr + " ,mo_contents='" + mo_contents + "'" + VbCrlf
	   sqlStr = sqlStr + " ,mo_exec_check='" + mo_exec_check + "'" + VbCrlf
   	   sqlStr = sqlStr + " ,exec_check='" + exec_check + "'" + VbCrlf
	   sqlStr = sqlStr + " ,exec_filepath='" + exec_filepath + "'" + VbCrlf
	   sqlStr = sqlStr + " where gcidx=" + CStr(idx)
	   dbget.Execute sqlStr

	end If

	response.write "<script>alert('저장되었습니다.');</script>"
	response.write "<script>location.href='" & manageUrl & "/admin/sitemaster/play/ground/groundweekEdit.asp?idx=" + Cstr(idx)+ "&gidx="+ gidx +"&reload=on'</script>"
End If
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
