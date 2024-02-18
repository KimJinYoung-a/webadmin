<%
'###########################################################
' Description : 위하고 Api 공용함수
' History : 2022.10.27 한용민 생성
'###########################################################
%>
<%
dim USER_IP, wehagoAppKey, wehagoServiceCode, AES256Key
    USER_IP = "110.93.128.93"

' 위하고 API
IF application("Svr_Info")="Dev" THEN
    wehagoAppKey = "98730f8cfdef4f77af17ce8ee08282fb"
    wehagoServiceCode = "10x10"
    AES256Key = "E86916E2CF3846B9BB6880CBC0447C35"
else
    wehagoAppKey = "98730f8cfdef4f77af17ce8ee08282fb"
    wehagoServiceCode = "10x10"
    AES256Key = "E86916E2CF3846B9BB6880CBC0447C35"
end if
%>