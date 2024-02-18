<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'dim baljuid, evt_code
'baljuid     = request.Form("baljuid")
'evt_code    = request.Form("evt_code")
''if (evt_code="") or Not IsNumeric(evt_code) then evt_code=0

dim gift_code
gift_code   = request.Form("gift_code")
dim ref
ref = request.ServerVariables("HTTP_REFERER")


if Len(gift_code)<1 then
    response.write "<script>alert('not valid gift_code');</script>"
    response.write "<script>history.back();</script>"
    dbget.close()	:	response.End
end if


''전체사은 이벤트인경우 생성불가하게 변경..
''------------------------------------------------------------------------------
dim ievt_code : ievt_code=0
sqlStr = " select top 1 event_code from db_event.dbo.tbl_openGift O" & VbCRLF
sqlStr = sqlStr & " Join db_event.dbo.tbl_Event E" & VbCRLF
sqlStr = sqlStr & " on O.event_code=E.evt_code" & VbCRLF
sqlStr = sqlStr & " Join db_event.dbo.tbl_Gift G" & VbCRLF
sqlStr = sqlStr & " on E.evt_code=G.evt_code" & VbCRLF
sqlStr = sqlStr & " where G.gift_code="&gift_code

rsget.Open sqlStr,dbget,1
if not rsget.Eof then
    ievt_code = rsget("event_code")
end if
rsget.Close

if (ievt_code<>0) then
    response.write "<script>alert('전체 사은이벤트 "&ievt_code&" 사은품코드 ("&gift_code&") 재작성 불가');</script>"
    response.write "전체 사은이벤트 "&ievt_code&" 사은품코드 ("&gift_code&") 재작성 불가"
    response.end
end if
''response.write  "수정중"
''response.end

''------------------------------------------------------------------------------
dim sqlStr

sqlStr = "exec [db_order].[dbo].sp_Ten_order_gift_reMake " & CStr(gift_code)
dbget.Execute sqlStr
'sqlStr = "exec [db_order].[dbo].sp_Ten_order_Gift_Maker " & CStr(baljuid) & "," & "'N'" & "," & CStr(evt_code)
'dbget.Execute sqlStr
'
'sqlStr = "exec [db_order].[dbo].sp_Ten_order_Gift_Maker " & CStr(baljuid) & "," & "'Y'" & "," & CStr(evt_code)
'dbget.Execute sqlStr



''사은품 제휴 -기타 : 다이어리 사은품  이벤트 : 야후 이벤트;;
'if (evt_code="8752") or (evt_code="8842") or (evt_code="9098") or (CStr(evt_code)="0") then
'    sqlStr = " exec [db_order].[dbo].ten_order_Gift_Maker_Etc " & CStr(baljuid) & ",'N'," & CStr(evt_code) & ",'10x10'"
'    dbget.execute sqlStr
'end if
 
response.write "<script>alert('OK');</script>"
response.write "<script>location.replace('" + ref + "');</script>"
dbget.close()	:	response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->