<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : AGV��� ó�� ������
' Hieditor : 2020.09.11 �̻� ����
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

response.write "���������Դϴ�."
dbget.close : response.end

select case mode
    case "chgStockGubun":
        response.write mode
    case else
        response.write "�߸��� �����Դϴ�."
        dbget.close : response.end
end select

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
