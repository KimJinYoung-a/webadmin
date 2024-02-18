<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : muticontentsvideo_process.asp
' Discription : I��(������) �̺�Ʈ ��Ƽ ������ ���� ���μ���
' History : 2019.02.25 ������ 
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
dim eCode, eMode, sqlStr, device, idx, videoFullLink
dim sIDX, sSortNo, ix, hvideotype, videoLink, menuidx
dim refer, BGImage, BGColorLeft, BGColorRight, contentsAlign, Margin
refer = request.ServerVariables("HTTP_REFERER")
eCode = requestCheckVar(Request.Form("eventid"),10)
menuidx	= requestCheckVar(Request.form("menuidx"),10)
eMode = requestCheckVar(Request.form("mode"),6)
device = requestCheckVar(Request.form("device"),1)
idx	= requestCheckVar(Request.form("idx"),10)
hvideotype = requestCheckVar(Request.Form("videotype"),1)
videoLink = Request.Form("videoLink")
BGImage	= requestCheckVar(Request.form("BGImage"),128)
BGColorLeft	= requestCheckVar(Request.form("BGColorLeft"),8)
BGColorRight	= requestCheckVar(Request.form("BGColorRight"),8)
contentsAlign	= requestCheckVar(Request.form("contentsAlign"),1)
Margin	= requestCheckVar(Request.form("Margin"),10)

if BGColorLeft="" then BGColorLeft="#FFFFFF"

    if eCode="" then
        response.write "<script type='text/javascript'>"
        response.write "	alert('��ȿ���� ���� ������ �Դϴ�. �ٽ� �õ��� �ּ���.');history.back();"
        response.write "</script>"
        response.End
    end if

    if videoLink <> "" then
        if checkNotValidHTML(videoLink) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
        response.write "</script>"
        response.End
        end if
    end If

    if BGImage <> "" then
        if checkNotValidHTML(BGImage) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
        response.write "</script>"
        response.End
        end if
    end If

    if not isNumeric(Margin) then Margin=0
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
                ElseIf InStr(itemvideo, "naver")>0 Then
                    itemvideo = "<iframe src="""& Trim(Replace(itemvideo,"?autoPlay=true","")) &""" width=""" & RetWidth & """  height=""" & RetHeight & """ frameborder=""no"" scrolling=""no"" marginwidth=""0"" marginheight=""0"" allow=""autoplay"" allowfullscreen></iframe>"
                End If
            else
                '// ���� ��ȯ �� �⺻�� (���������� ��޿�����)
                If InStr(itemvideo , "youtube")>0 Then
                    itemvideo = "<iframe src="""& Trim(Replace(itemvideo,"https://youtu.be/","https://www.youtube.com/embed/")) &""" frameborder=""0"" allowfullscreen></iframe>"
                ElseIf InStr(itemvideo , "youtu.be")>0 Then
                    itemvideo = "<iframe src="""& Trim(Replace(itemvideo,"https://youtu.be/","https://www.youtube.com/embed/")) &""" frameborder=""0"" allowfullscreen></iframe>"
                ElseIf InStr(itemvideo, "vimeo")>0 Then
                    itemvideo = "<iframe src="""& Trim(Replace(itemvideo,"https://vimeo.com/","https://player.vimeo.com/video/")) &""" frameborder=""0"" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>"
                ElseIf InStr(itemvideo, "naver")>0 Then
                    itemvideo = "<iframe src="""& Trim(Replace(itemvideo,"?autoPlay=true","")) &""" frameborder=""no"" scrolling=""no"" marginwidth=""0"" marginheight=""0"" allow=""autoplay"" allowfullscreen></iframe>"
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
        ElseIf InStr(itemvideo, "naver")>0 Then
            videoType = "naver"
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

    '// ��Ƽ������ ������ ���� �Է�
    sqlStr = "Update db_event.dbo.tbl_event_multi_contents_master" & vbCrLf
    sqlStr = sqlStr & " Set BGImage='" & BGImage & "'" & vbCrLf
    sqlStr = sqlStr & " ,BGColorLeft='" & BGColorLeft & "'" & vbCrLf
	sqlStr = sqlStr & " ,BGColorRight='" & BGColorRight & "'" & vbCrLf
    sqlStr = sqlStr & " ,contentsAlign='" & contentsAlign & "'" & vbCrLf
    sqlStr = sqlStr & " ,Margin='" & Margin & "'" & vbCrLf
    sqlStr = sqlStr & " Where idx='" & menuidx & "'"
    dbget.Execute sqlStr

    '--3.theme ����
    sqlStr = "UPDATE [db_event].[dbo].[tbl_event_md_theme]" & vbCrlf
    sqlStr = sqlStr + " SET contentsAlign='" & contentsAlign & "'" & vbCrlf
    sqlStr = sqlStr + " where evt_code=" & eCode
    dbget.execute sqlStr

Select Case eMode
Case "VI"
	dbget.beginTrans
		sqlStr = "Insert Into db_event.dbo.tbl_event_multi_contents " &_
					" (menuidx, device , videoLink, videoFullLink, videotype) values " &_
					" ('" & menuidx  & "'" &_
					" ,'" & device &"'" &_
					" ,'" & videoLink &"'" &_
					" ,'" & videoFullLink &"'" &_
					" ,'" & hvideotype &"')"
		dbget.Execute(sqlStr)

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.", "back", "")
            response.End 
        end if
    '===========================================================
	dbget.CommitTrans
    
	response.write "<script type='text/javascript'>"
    response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=5&viewset='+opener.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
    response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
Case "VU"
	dbget.beginTrans
        sqlStr = sqlStr & " Update db_event.dbo.tbl_event_multi_contents" & vbCrLf
        sqlStr = sqlStr & " Set videoLink='" & videoLink & "'" & vbCrLf
        sqlStr = sqlStr & " ,videoFullLink='" & videoFullLink & "'" & vbCrLf
        sqlStr = sqlStr & " ,videotype='" & hvideotype & "'" & vbCrLf
        sqlStr = sqlStr & " Where idx='" & idx & "'"
        dbget.Execute sqlStr

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.", "back", "")
            response.End 
        end if
    '===========================================================
	dbget.CommitTrans
    
	response.write "<script type='text/javascript'>"
    response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=5&viewset='+opener.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
    response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->