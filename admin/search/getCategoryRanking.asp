<%@ codepage="65001" language="VBScript" %>
<%
option Explicit
Response.CharSet = "utf-8"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
function CheckVaildIP(ref)
    CheckVaildIP = false
    dim i
    ' dim VaildIP : VaildIP = Array("13.125.145.40","13.125.12.181","52.79.73.145","61.252.133.88","192.168.1.70","61.252.133.81","192.168.1.81","192.168.1.72","110.93.128.107","61.252.133.2","61.252.133.69","61.252.133.70","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72")
    ' 
    ' for i=0 to UBound(VaildIP)
    '     if (VaildIP(i)=ref) then
    '         CheckVaildIP = true
    '         exit function
    '     end if
    ' next

    dim validToken : validToken = Array("70711546f86e45b2bb3f9b5528ded10d")
    dim authtkn : authtkn = LCASE(request("authtkn"))
    for i=0 to UBound(validToken)
        if (validToken(i)=authtkn) then
            CheckVaildIP = true
            exit function
        end if
    next

end function

dim ref : ref = request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    dbget.Close() : response.end
end if

Dim vQuery, vArr, i, dno
dno = requestCheckVar(Request("dno"),10) ''도메인 번호
if (dno="") then dno="0"


vQuery = " exec db_sitemaster.dbo.usp_ksearch_CategoryDictionary_get "&dno

rsget.CursorLocation = adUseClient
rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly

If Not rsget.Eof Then
	vArr = rsget.getRows()
End If
rsget.close
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
If isArray(vArr) Then
	For i=0 To UBound(vArr,2)
        Response.Write Trim(vArr(0,i))&vbCRLF
	Next
End If

%>