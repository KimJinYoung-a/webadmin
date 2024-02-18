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
Dim idx, title, Link_Image, Isusing, mode, menuidx, eCode, device
dim refer, BGImage, BGColorLeft, BGColorRight, contentsAlign, Margin
	
	idx = requestCheckVar(request("idx"),10)
    menuidx = RequestCheckVar(request("menuidx"),256)
    Link_Image = requestCheckVar(request("Link_Image"),128)
    Isusing = requestCheckVar(request("Isusing"),1)
    BGImage	= requestCheckVar(Request.form("BGImage"),128)
    BGColorLeft	= requestCheckVar(Request.form("BGColorLeft"),8)
    BGColorRight	= requestCheckVar(Request.form("BGColorRight"),8)
    contentsAlign	= requestCheckVar(Request.form("contentsAlign"),1)
    Margin	= requestCheckVar(Request.form("Margin"),10)
    eCode = requestCheckVar(request("eventid"),10)
    device = requestCheckVar(Request.Form("device"),1)

    if BGColorLeft="" then BGColorLeft="#FFFFFF"

    if Link_Image <> "" then
        if checkNotValidHTML(Link_Image) then
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
	if idx="" then idx=0
	If idx=0 Then
	mode = "add"
	Else
	mode = "edit"
	End If

dim sqlStr

    '// 멀티컨텐츠 마스터 정보 입력
    sqlStr = " Update db_event.dbo.tbl_event_multi_contents_master" & vbCrLf
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

if (mode = "add") then

    sqlStr = " insert into [db_event].[dbo].[tbl_ImageLink_Master]" & VbCrlf
    sqlStr = sqlStr & " (menuidx, Image, RegUser, ModifyUser, Isusing, RegDate, LastUpDate, device)" & VbCrlf
    sqlStr = sqlStr & " values(" & VbCrlf
    sqlStr = sqlStr & " '" & menuidx & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & Link_Image & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & Session("ssBctId") & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & Session("ssBctId") & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & Isusing & "'" & VbCrlf
    sqlStr = sqlStr & " ,getdate()" & VbCrlf
    sqlStr = sqlStr & " ,getdate()" & VbCrlf
    sqlStr = sqlStr & " ,'" & device & "'" & VbCrlf
    sqlStr = sqlStr & " )"
    dbget.Execute sqlStr

    sqlStr = "IF NOT EXISTS(SELECT idx FROM db_event.dbo.tbl_event_multi_contents WHERE menuidx=" & menuidx  & " and device='"& device &"')" & vbCrLf
    sqlStr = sqlStr & "	BEGIN" & vbCrLf
    sqlStr = sqlStr & "		Insert Into db_event.dbo.tbl_event_multi_contents(menuidx, device , imgurl)" & vbCrLf
    sqlStr = sqlStr & "  	values('" & menuidx  & "','" & device &"','" & Link_Image &"')" & vbCrLf
    sqlStr = sqlStr & "	END"
    dbget.Execute(sqlStr)

elseif mode = "edit" then

   sqlStr = " update [db_event].[dbo].[tbl_ImageLink_Master]" & VbCrlf
   sqlStr = sqlStr & " set Image='" & Link_Image & "'" & VbCrlf
   sqlStr = sqlStr & " ,ModifyUser='" & Session("ssBctId") & "'" & VbCrlf
   sqlStr = sqlStr & " ,Isusing='" & Isusing & "'" & VbCrlf
   sqlStr = sqlStr & " ,LastUpDate=getdate()" & VbCrlf
   sqlStr = sqlStr & " where idx=" & cstr(idx)
   dbget.Execute sqlStr

end if

	response.write "<script type='text/javascript'>"
    response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" & Cstr(eCode) & "&togglediv=5&viewset='+opener.document.frmEvt.viewset.value);"
    response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->