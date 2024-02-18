<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : AGV재고 처리 페이지
' Hieditor : 2020.09.11 이상구 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%

dim refer
    refer = request.ServerVariables("HTTP_REFERER")

dim mode, skuCd, skuCdArr
dim sqlStr, i, j, k

response.write "샘플파일입니다."
dbget.close : response.end

select case mode
    case "chgStockGubun":
        response.write mode
    case else
        response.write "잘못된 접근입니다."
        dbget.close : response.end
end select

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
