<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : brandstoryinfo_process.asp
' Discription : I형(통합형) 이벤트 브랜드 스토리 등록 프로세스
' History : 2019.02.12 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<%

dim eCode, eMode, sqlStr, brandname, menuidx, dgncomment, ebrand
dim refer, BGImage, BGColorLeft, BGColorRight, contentsAlign, Margin
refer = request.ServerVariables("HTTP_REFERER")
eCode = requestCheckVar(Request.Form("evt_code"),10)
eMode = requestCheckVar(Request.Form("imod"),2)
menuidx = requestCheckVar(Request.Form("menuidx"),10)
brandname = requestCheckVar(Request.Form("brandname"),16)
ebrand = requestCheckVar(Request.Form("ebrand"),16)
dgncomment = html2db(Request.Form("dgncomment"))
BGImage	= requestCheckVar(Request.form("BGImage"),128)
BGColorLeft	= requestCheckVar(Request.form("BGColorLeft"),8)
BGColorRight	= requestCheckVar(Request.form("BGColorRight"),8)
contentsAlign	= requestCheckVar(Request.form("contentsAlign"),1)
Margin	= requestCheckVar(Request.form("Margin"),10)

if BGColorLeft="" then BGColorLeft="#FFFFFF"

if dgncomment <> "" then
	if checkNotValidHTML(dgncomment) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

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

select case eMode
case "BI"
	dbget.beginTrans
    '===========================================================
    '-- 추가
		sqlStr = "Insert Into db_event.dbo.tbl_event_multi_contents " & vbcrlf
		sqlStr = sqlStr + " (menuidx, BrandName, BrandContents, makerid) values "  & vbcrlf
		sqlStr = sqlStr + " ('" & menuidx  & "'"  & vbcrlf
		sqlStr = sqlStr + " ,'" & brandname &"'"  & vbcrlf
		sqlStr = sqlStr + " ,'" & dgncomment &"'"
		sqlStr = sqlStr + " ,'" & ebrand &"')"
		dbget.Execute(sqlStr)

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
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
case "BU"
	dbget.beginTrans
    '===========================================================
    '-- 수정
        sqlStr = "UPDATE [db_event].[dbo].[tbl_event_multi_contents]" & vbCrlf
        sqlStr = sqlStr + " SET BrandName='" & brandname & "'" & vbCrlf
        sqlStr = sqlStr + " , BrandContents='" & dgncomment & "'" & vbCrlf
		sqlStr = sqlStr + " , makerid='" & ebrand & "'" & vbCrlf
        sqlStr = sqlStr + " where menuidx=" & menuidx
        dbget.execute sqlStr

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
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
case else
	Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
end select
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->