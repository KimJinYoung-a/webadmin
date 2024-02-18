<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : muticontentinfo_process.asp
' Discription : I형(통합형) 이벤트 멀티 컨텐츠 정보 프로세스
' History : 2019.12.20 정태훈 
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim eCode, eMode, sqlStr, menuidx
dim refer, BGImage, BGColorLeft, BGColorRight, contentsAlign, Margin
refer = request.ServerVariables("HTTP_REFERER")
eCode = requestCheckVar(Request.Form("eventid"),10)
menuidx	= requestCheckVar(Request.form("menuidx"),10)
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

    if BGImage <> "" then
        if checkNotValidHTML(BGImage) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
        response.write "</script>"
        response.End
        end if
    end If

    if not isNumeric(Margin) then Margin=0
    '// 멀티컨텐츠 마스터 정보 입력
    sqlStr = sqlStr & " Update db_event.dbo.tbl_event_multi_contents_master" & vbCrLf
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

	response.write "<script type='text/javascript'>"
    response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->