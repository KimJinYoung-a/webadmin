<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : muticontentsvideo_process.asp
' Discription : I형(통합형) 이벤트 멀티 컨텐츠 영상 프로세스
' History : 2019.02.25 정태훈 
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
        response.write "	alert('유효하지 않은 데이터 입니다. 다시 시도해 주세요.');history.back();"
        response.write "</script>"
        response.End
    end if

    if videoLink <> "" then
        if checkNotValidHTML(videoLink) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
        response.write "</script>"
        response.End
        end if
    end If

    if BGImage <> "" then
        if checkNotValidHTML(BGImage) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
        response.write "</script>"
        response.End
        end if
    end If

    if not isNumeric(Margin) then Margin=0
    '--------------------------------------------------------
    '// 2016.2.16 신규추가 상품상세설명 동영상 추가 - 원승현
    '// 2016-12-13  iframe 없는 경우 - iframe 생성 삽입
    '// 아이템 동영상 값 정규식으로 src, width, height값 뽑아냄
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
        '// 2016-12-13 추가 iframe 없이 주소만 넘어 올경우
        If InStr(itemvideo ,"iframe") > 0 Then
        else
            if device="W" then
                '// 비디오 변환 및 기본형 (유투브인지 비메오인지)
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
                '// 비디오 변환 및 기본형 (유투브인지 비메오인지)
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
        '// iframe 이외의 코드는 잘라버림
        itemvideo = Left(itemvideo, InStrRev(itemvideo, "</iframe>")+9)

        '// 비디오 타입지정(유투브인지 비메오인지)
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

    '// 멀티컨텐츠 마스터 정보 입력
    sqlStr = "Update db_event.dbo.tbl_event_multi_contents_master" & vbCrLf
    sqlStr = sqlStr & " Set BGImage='" & BGImage & "'" & vbCrLf
    sqlStr = sqlStr & " ,BGColorLeft='" & BGColorLeft & "'" & vbCrLf
	sqlStr = sqlStr & " ,BGColorRight='" & BGColorRight & "'" & vbCrLf
    sqlStr = sqlStr & " ,contentsAlign='" & contentsAlign & "'" & vbCrLf
    sqlStr = sqlStr & " ,Margin='" & Margin & "'" & vbCrLf
    sqlStr = sqlStr & " Where idx='" & menuidx & "'"
    dbget.Execute sqlStr

    '--3.theme 수정
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
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
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
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
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