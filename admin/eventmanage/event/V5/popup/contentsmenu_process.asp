<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : contentsmenu_process.asp
' Discription : I형(통합형) 이벤트 멀티 컨텐츠 메뉴 등록 프로세스
' History : 2019.02.07 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<%

dim eCode, eMode, strSql
dim menudiv, GroupItemType, GroupItemCheck, GroupItemPriceView, isusing, menuidx

eCode = requestCheckVar(Request.Form("evt_code"),10)
eMode = requestCheckVar(Request.Form("imod"),2)
menudiv = requestCheckVar(Request.Form("menudiv"),2)
isusing = requestCheckVar(Request.Form("isusing"),1)
GroupItemType = requestCheckVar(Request.Form("GroupItemType"),1)
GroupItemCheck = requestCheckVar(Request.Form("GroupItemCheck"),1)
GroupItemPriceView = requestCheckVar(Request.Form("GroupItemPriceView"),1)
menuidx = requestCheckVar(Request.Form("menuidx"),10)

dim refer
refer = request.ServerVariables("HTTP_REFERER")

if menudiv="3" then
    GroupItemType="T"
    GroupItemCheck="I"
    GroupItemPriceView="Y"
end if

select case eMode
case "MI"
	dbget.beginTrans
        '===========================================================
        '--2.disply 수정
        strSql = "INSERT INTO [db_event].[dbo].[tbl_event_multi_contents_master]"
        if menudiv="3" then
        strSql = strSql + "(evt_code, menudiv, isusing, GroupItemPriceView, GroupItemCheck, GroupItemType)" & vbCrlf
        strSql = strSql + " values(" & eCode & "," & menudiv & ",'Y','" & GroupItemPriceView & "','" & GroupItemCheck & "','" & GroupItemType & "')" & vbCrlf
        elseif menudiv="12" then '// 상품 가격 연동
        strSql = strSql + "(evt_code, menudiv, isusing, GroupItemViewType, GroupItemTitleName, GroupItemBrandName)" & vbCrlf
        strSql = strSql + " values(" & eCode & "," & menudiv & ",'Y','C','Y','Y')" & vbCrlf
        else
        strSql = strSql + "(evt_code, menudiv, isusing)" & vbCrlf
        strSql = strSql + " values(" & eCode & "," & menudiv & ",'Y')" & vbCrlf
        end if
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
            response.End 
        end if

        strSql = "SELECT SCOPE_IDENTITY()"
        rsget.Open strSql, dbget, 0
        menuidx = rsget(0)
        rsget.Close

        If menudiv = "13" Then
            strSql = "INSERT INTO [db_event].[dbo].[tbl_event_multi_contents_tabbar]" & vbCrlf
            strSql = strSql & " (master_idx, device) VALUES (" & menuidx & ", 'M'), (" & menuidx & ", 'W') "
            dbget.execute strSql

            if Err.Number <> 0 then
                dbget.RollBackTrans 
                Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
                response.End 
            end if
        End If
    '===========================================================
	dbget.CommitTrans
    
	response.write "<script type='text/javascript'>"
    response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	parent.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=5&viewset='+parent.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
    'response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
case "MU"
    if menuidx="" then
        response.write "<script type='text/javascript'>"
        response.write "	alert('유효하지 않은 데이터 입니다. 다시 시도해 주세요.');history.back();"
        response.write "</script>"
        response.End
    end if
	dbget.beginTrans
        '===========================================================
        '--2.disply 수정
        strSql = "UPDATE [db_event].[dbo].[tbl_event_multi_contents_master]" & vbCrlf
        strSql = strSql + " SET menudiv='" & menudiv & "'" & vbCrlf
        if menudiv="3" then
        strSql = strSql + ", GroupItemPriceView='" & GroupItemPriceView & "'" & vbCrlf
        strSql = strSql + ", GroupItemCheck='" & GroupItemCheck & "'" & vbCrlf
        strSql = strSql + ", GroupItemType='" & GroupItemType & "'" & vbCrlf
        end if
        strSql = strSql + " where evt_code=" & eCode
        strSql = strSql + " and idx=" & menuidx
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
            response.End 
        end if
    '===========================================================
	dbget.CommitTrans
    
	response.write "<script type='text/javascript'>"
    response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	parent.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=5&viewset='+parent.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
    'response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
case "MD"
    dbget.beginTrans
    strSql = "DELETE FROM [db_event].[dbo].[tbl_event_multi_contents_master]" & vbCrlf
    strSql = strSql + " where evt_code=" & eCode
    strSql = strSql + " and idx=" & menuidx
    dbget.execute strSql
    if Err.Number <> 0 then
        dbget.RollBackTrans 
        Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
        response.End 
    end if
    dbget.CommitTrans
	response.write "<script type='text/javascript'>"
    response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	parent.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=5&viewset='+parent.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
    'response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
case else
	Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
end select
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->