<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###############################################
' PageName : pop_themeslide_proc.asp
' Discription : ����� slide process
' History : 2019-02-11 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	Dim menuidx , mode , device , idx, saveafter
	Dim slideimg , linkurl , sorting , isusing '�����̵� �̹���
	Dim sqlStr, videoLink, hvideotype, videoFullLink, eventid
	Dim sIdx, sSortNo, sIsUsing, i , slinkurl, bgleft, bgright '//�����̵�

	idx	= requestCheckVar(Request.form("idx"),10)
	menuidx	= requestCheckVar(Request.form("menuidx"),10)
	mode = requestCheckVar(Request.form("mode"),6)
	device = requestCheckVar(Request.form("device"),1)
	slideimg = requestCheckVar(Request.form("slideimg"),200)
	eventid	= requestCheckVar(Request.form("eventid"),10)
'Response.write mode & "<br>"
Select Case mode
	Case "SI"
		'slide�̹��� �ű� ���
		sqlStr = "Insert Into db_event.dbo.tbl_event_multi_contents " &_
					" (menuidx, device , imgurl) values " &_
					" ('" & menuidx  & "'" &_
					" ,'" & device &"'" &_
					" ,'" & slideimg &"')"
		dbget.Execute(sqlStr)
		saveafter="SI"
	Case "SV"
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
		sqlStr = "Insert Into db_event.dbo.tbl_event_multi_contents " &_
					" (menuidx, device , videoLink, videoFullLink, videotype) values " &_
					" ('" & menuidx  & "'" &_
					" ,'" & device &"'" &_
					" ,'" & videoLink &"'" &_
					" ,'" & videoFullLink &"'" &_
					" ,'" & hvideotype &"')"
		dbget.Execute(sqlStr)
		saveafter="SV"
	Case "SU"
		'//����Ʈ��������
		for i=1 to request.form("chkIdx").count
			sIdx = request.form("chkIdx")(i)
			sSortNo = request.form("viewidx"&sIdx)
			sIsUsing = request.form("use"&sIdx)
			if sSortNo="" then sSortNo="0"
			if sIsUsing="" then sIsUsing="N"

			sqlStr = sqlStr & " Update db_event.dbo.tbl_event_multi_contents Set " & vbCrLf
			sqlStr = sqlStr & " viewidx=" & sSortNo & "" & vbCrLf
			sqlStr = sqlStr & " ,isusing='" & sIsUsing & "'" & vbCrLf
			sqlStr = sqlStr & " Where idx='" & sIdx & "';"
		Next

		If sqlStr <> "" then
			dbget.Execute sqlStr
		Else
			Call Alert_return("������ ������ �����ϴ�.")
			dbget.Close: Response.End
		End If 
		saveafter="SU"
	Case "SD" '����
		sIdx = request.form("chkIdx")

		sqlStr = "delete from db_event.dbo.tbl_event_multi_contents Where idx='"& sIdx &"' and device = '"& device &"'"
		dbget.Execute sqlStr
	Case "SAD" '����
		sIdx = request.form("chkIdx")

		sqlStr = "delete from db_event.dbo.tbl_event_multi_contents Where idx in (" & sIdx & ") and device = '"& device &"'"
		dbget.Execute sqlStr
End Select
%>
<script language="javascript">
<!--
	// ������� ����
	alert("<%=chkiif(mode="SD","���� �Ϸ�.","����/���� �Ϸ�.")%>");
	<% if device="W" then %>
	self.location = "pop_<%=chkiif(device="W","pcweb","mobile")%>_themeslide.asp?eC=<%=eventid%>&menuidx=<%=menuidx%>&smode=<%=mode%>&saveafter=<%=saveafter%>";
	<% else %>
	self.location = "pop_<%=chkiif(device="M","mobile","pcweb")%>_themeslide.asp?eC=<%=eventid%>&menuidx=<%=menuidx%>&smode=<%=mode%>&saveafter=<%=saveafter%>";
	<% end if %>
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->