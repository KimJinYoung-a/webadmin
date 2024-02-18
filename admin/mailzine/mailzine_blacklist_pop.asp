<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 메일진
' History : 2018.04.27 이상구 생성(메일러 연동 생성 메일러로 발송 내역 전송. 메일 가져오기 생성.)
'			2019.06.24 정태훈 수정(템플릿 기능 신규 추가)
'			2020.05.28 한용민 수정(TMS 메일러 추가)
'###########################################################
%>
<%
response.write "<font color='red'>처리중입니다. 잠시만 기다려 주세요.</font>"
%>
<iframe id="mail" src="/admin/mailzine/mailzine_blacklist.asp" width="100%" frameborder="0" scrolling="no" height=50></iframe>
