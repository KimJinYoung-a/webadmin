<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###############################################
' PageName : pop_top_slide_proc.asp
' Discription : TOP slide process
' History : 2019-02-12 ������
'			2019-10-02 ������	���ø� �������� ����
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	Dim eventid , mode , device , idx, saveafter
	Dim slideimg , linkurl , sorting , isusing '�����̵� �̹���
	Dim sqlStr, contentsAlign
	Dim sIdx, sSortNo, sIsUsing, i , slinkurl, bgleft, bgright '//�����̵�
	dim checkwide, menuidx

	idx	= requestCheckVar(Request.form("idx"),10)
	eventid	= requestCheckVar(Request.form("eventid"),10)
	mode = requestCheckVar(Request.form("mode"),6)
	device = requestCheckVar(Request.form("device"),1)
	slideimg = requestCheckVar(Request.form("slideimg"),200)
	menuidx = requestCheckvar(request("menuidx"),16)
	if menuidx="" or isnull(menuidx) then menuidx=0
'Response.write mode & "<br>"
Select Case mode
	 Case "SI"
		'slide�̹��� �ű� ���
		sqlStr = "Insert Into db_event.dbo.tbl_event_top_slide_addimage " &_
					" (evt_code, device , slideimg, contentsAlign, menuidx) values " &_
					" ('" & eventid  & "'" &_
					" ,'" & device &"'" &_
					" ,'" & slideimg &"'" &_
					" ,'" & contentsAlign & "'" &_
					" ," & menuidx &")"
		dbget.Execute(sqlStr)
		sqlStr = "IF NOT EXISTS(SELECT idx FROM db_event.dbo.tbl_event_multi_contents WHERE menuidx=" & menuidx  & " and device='"& device &"')" & vbCrLf
		sqlStr = sqlStr & "	BEGIN" & vbCrLf
		sqlStr = sqlStr & "		Insert Into db_event.dbo.tbl_event_multi_contents(menuidx, device , imgurl)" & vbCrLf
		sqlStr = sqlStr & "  	values('" & menuidx  & "','" & device &"','" & slideimg &"')" & vbCrLf
		sqlStr = sqlStr & "	END"
		dbget.Execute(sqlStr)
		saveafter="SI"
	Case "SU"
		'//����Ʈ��������
		for i=1 to request.form("chkIdx").count
			sIdx = request.form("chkIdx")(i)
			sSortNo = request.form("sort"&sIdx)
			sIsUsing = request.form("use"&sIdx)
			slinkurl = request.form("linkurl"&sIdx)
			bgleft = request.form("bgleft"&sIdx)
			bgright = request.form("bgright"&sIdx)
			contentsAlign = request.form("contentsAlign"&sIdx)
			if sSortNo="" then sSortNo="0"
			if sIsUsing="" then sIsUsing="N"
			if contentsAlign="2" then checkwide="Y"
			sqlStr = sqlStr & " Update db_event.dbo.tbl_event_top_slide_addimage Set " & vbCrLf
			sqlStr = sqlStr & " sorting=" & sSortNo & "" & vbCrLf
			sqlStr = sqlStr & " ,isusing='" & sIsUsing & "'" & vbCrLf
			sqlStr = sqlStr & " ,linkurl='" & slinkurl & "'" & vbCrLf
			sqlStr = sqlStr & " ,bgleft='" & bgleft & "'" & vbCrLf
			sqlStr = sqlStr & " ,bgright='" & bgright & "'" & vbCrLf
			sqlStr = sqlStr & " ,contentsAlign='" & contentsAlign & "'" & vbCrLf
			sqlStr = sqlStr & " Where idx='" & sIdx & "';"
		Next
		If sqlStr <> "" then
			dbget.Execute sqlStr
		Else
			Call Alert_return("������ ������ �����ϴ�.")
			dbget.Close: Response.End
		End If 
		'���̵��� ��� ���̵� �÷��� ������Ʈ
		if checkwide="Y" then
			sqlStr = "update [db_event].[dbo].[tbl_event_display]"& vbCrLf
			sqlStr = sqlStr & " set evt_wideyn=1, evt_fullyn=1" & vbCrLf
			sqlStr = sqlStr & " Where evt_code="& eventid
			dbget.Execute sqlStr
		else
			sqlStr = "update [db_event].[dbo].[tbl_event_display]"& vbCrLf
			sqlStr = sqlStr & " set evt_wideyn=0, evt_fullyn=0" & vbCrLf
			sqlStr = sqlStr & " Where evt_code="& eventid
			dbget.Execute sqlStr
		end if
		saveafter="SU"
	Case "SD" '����
		sIdx = request.form("chkIdx")

		sqlStr = "delete from db_event.dbo.tbl_event_top_slide_addimage Where idx='"& sIdx &"' and device = '"& device &"'"
		dbget.Execute sqlStr
		sqlStr = "IF NOT EXISTS(SELECT top 1 idx FROM db_event.dbo.tbl_event_top_slide_addimage WHERE menuidx=" & menuidx  & " AND device='"& device &"')" & vbCrLf
		sqlStr = sqlStr & "	BEGIN" & vbCrLf
		sqlStr = sqlStr & "		DELETE FROM db_event.dbo.tbl_event_multi_contents WHERE menuidx=" & menuidx  & " AND device='"& device &"'" & vbCrLf
		sqlStr = sqlStr & "	END"
		dbget.Execute(sqlStr)
	Case "SV"
		dim hvideotype, videoLink, videoFullLink
		hvideotype = requestCheckVar(Request.Form("videotype"),1)
		videoLink = requestCheckVar(Request.Form("videoLink"),256)
		'--------------------------------------------------------
		'// 2016.2.16 �ű��߰� ��ǰ�󼼼��� ������ �߰� - ������
		'// 2016-12-13  iframe ���� ��� - iframe ���� ����
		'// ������ ������ �� ���Խ����� src, width, height�� �̾Ƴ�
		If Trim(videoLink) <> "" Then
			Dim itemvideo, RetStr, RetSrc, RetWidth, RetHeight, regEx, Matches, Match, VideoTempSrc, VideoTempWidth, VideoTempHeight, videoType
			itemvideo = videoLink
			itemvideo = itemvideo & "?rel=0"
			if hvideotype="1" then
				RetWidth="720"
				RetHeight="405"
			elseif hvideotype="2" then
				RetWidth="720"
				RetHeight="540"
			else
				RetWidth="720"
				RetHeight="720"
			end if
			'// 2016-12-13 �߰� iframe ���� �ּҸ� �Ѿ� �ð��
			If InStr(itemvideo ,"iframe") > 0 Then
			else
				if device="W" then
					'// ���� ��ȯ �� �⺻�� (���������� ��޿�����)
					If InStr(itemvideo , "youtube")>0 Then
						itemvideo = "<iframe src="""& Trim(Replace(itemvideo,"https://youtu.be/","https://www.youtube.com/embed/")) &""" width=""" & RetWidth & """  height=""" & RetHeight & """ frameborder=""0"" allowfullscreen></iframe>"
					ElseIf InStr(itemvideo , "youtu.be")>0 Then
						itemvideo = "<iframe src="""& Trim(Replace(itemvideo,"https://youtu.be/","https://www.youtube.com/embed/")) &""" width=""" & RetWidth & """  height=""" & RetHeight & """ frameborder=""0"" allowfullscreen></iframe>"
					ElseIf InStr(itemvideo, "vimeo")>0 Then
						itemvideo = "<iframe src="""& Trim(Replace(itemvideo,"https://vimeo.com/","https://player.vimeo.com/video/")) &""" width=""" & RetWidth & """  height=""" & RetHeight & """ frameborder=""0"" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>"
					End If
				else
					'// ���� ��ȯ �� �⺻�� (���������� ��޿�����)
					If InStr(itemvideo , "youtube")>0 Then
						itemvideo = "<iframe src="""& Trim(Replace(itemvideo,"https://youtu.be/","https://www.youtube.com/embed/")) &""" frameborder=""0"" allowfullscreen></iframe>"
					ElseIf InStr(itemvideo , "youtu.be")>0 Then
						itemvideo = "<iframe src="""& Trim(Replace(itemvideo,"https://youtu.be/","https://www.youtube.com/embed/")) &""" frameborder=""0"" allowfullscreen></iframe>"
					ElseIf InStr(itemvideo, "vimeo")>0 Then
						itemvideo = "<iframe src="""& Trim(Replace(itemvideo,"https://vimeo.com/","https://player.vimeo.com/video/")) &""" frameborder=""0"" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>"
					End If
				End If
			End If 

			itemvideo = Trim(Replace(itemvideo,"""","'"))
			'// iframe �̿��� �ڵ�� �߶����
			itemvideo = Left(itemvideo, InStrRev(itemvideo, "</iframe>")+9)

			'// ���� Ÿ������(���������� ��޿�����)
			If InStr(itemvideo, "youtube")>0 Then
				videoType = "youtube"
			ElseIf InStr(itemvideo, "vimeo")>0 Then
				videoType = "vimeo"
			Else
				videoType = "etc"
			End If

			Set regEx = New RegExp
			regEx.IgnoreCase = True
			regEx.Global = True

			regEx.pattern = "<iframe [^<>]*>"
			Set Matches = regEx.execute(itemvideo)
			For Each Match In Matches
				VideoTempSrc =  Mid(Match.Value, InStrRev(Match.Value,"src='")+5)
				RetSrc = Left(VideoTempSrc, InStr(VideoTempSrc, "'")-1)

				VideoTempWidth =  Mid(Match.Value, InStrRev(Match.Value,"width='")+7)
				RetWidth = Left(VideoTempWidth, InStr(VideoTempWidth, "'")-1)

				VideoTempHeight =  Mid(Match.Value, InStrRev(Match.Value,"height='")+8)
				RetHeight = Left(VideoTempHeight, InStr(VideoTempHeight, "'")-1)
			Next
			Set regEx = Nothing
			Set Matches = Nothing

			videoFullLink=chrbyte(html2db(itemvideo),255,"")
		End If
		'--------------------------------------------------------

		'slide�̹��� �ű� ���
		sqlStr = "update db_event.dbo.tbl_event_display" + vbcrlf
		sqlStr = sqlStr + "set videoLink='" + videoLink + "'" + vbcrlf
		sqlStr = sqlStr + " , videoFullLink='" + videoFullLink + "'" + vbcrlf
		sqlStr = sqlStr + " , videotype='" + hvideotype + "'" + vbcrlf
		sqlStr = sqlStr + " where evt_code=" + Cstr(eventid)
		dbget.Execute(sqlStr)
		saveafter="SV"
	Case "VD" '����
		sIdx = request.form("chkIdx")

		sqlStr = "update db_event.dbo.tbl_event_display set videoLink=null, videoFullLink=null, videotype=null Where evt_code="& eventid
		dbget.Execute sqlStr
		saveafter="VD"
End Select
%>
<script language="javascript">
<!--
	// ������� ����
	alert("<%=chkiif(mode="SD","���� �Ϸ�.","����/���� �Ϸ�.")%>");
	<% if device="W" then %>
	self.location = "pop_<%=chkiif(device="W","pcweb","mobile")%>_top_slide.asp?eC=<%=eventid%>&smode=<%=mode%>&saveafter=<%=saveafter%>&menuidx=<%=menuidx%>";
	<% else %>
	self.location = "pop_<%=chkiif(device="M","mobile","pcweb")%>_top_slide.asp?eC=<%=eventid%>&smode=<%=mode%>&saveafter=<%=saveafter%>&menuidx=<%=menuidx%>";
	<% end if %>
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
