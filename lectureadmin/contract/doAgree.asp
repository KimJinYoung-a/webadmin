<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/fingersUpcheAgreeCls.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<%
dim sqlStr
dim agreeIdx
dim makerid : makerid = session("ssBctID")
dim groupid : groupid = getPartnerId2GroupID(makerid)
dim mode : mode = requestCheckvar(request("mode"),16)

agreeIdx  = requestCheckvar(request("agreeIdx"),10)

dim agreeRefIp : agreeRefIp=request.serverVariables("REMOTE_HOST")
dim AssignedRow, iagreedate

if (mode="iagree") then
    sqlStr = "update db_partner.dbo.tbl_partner_fingers_agreeHist"&vbCRLF
    sqlStr = sqlStr&" set agreedate=isNULL(agreedate,getdate())"&vbCRLF
    sqlStr = sqlStr&" , agreeRefIp='"&agreeRefIp&"'"&vbCRLF
    sqlStr = sqlStr&" where agreeIdx="&agreeIdx&vbCRLF
    sqlStr = sqlStr&" and groupid='"&groupid&"'"
    
    dbget.Execute sqlStr,AssignedRow
    
    if (AssignedRow>0) then
        sqlStr = "select agreedate from db_partner.dbo.tbl_partner_fingers_agreeHist where  agreeIdx="&agreeIdx&""
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
        if Not rsget.Eof then
            iagreedate = rsget("agreedate")
        end if
        rsget.Close
        
        response.write "<script>alert('"&iagreedate&" 동의 처리 되었습니다.');opener.location.reload();window.close()</script>"
    else
        response.write "<script>alert('작업중 오류가 발생하였습니다.');history.back()</script>"
    end if
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
