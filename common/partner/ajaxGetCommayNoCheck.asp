<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim query1, company_no, uid
company_no = requestCheckvar(request.form("company_no"),16)
uid = requestCheckvar(request.form("uid"),32)
query1 = " select top 1 company_no, groupid " & vbCrLf
query1 = query1 & " from [db_partner].[dbo].[tbl_partner]" & vbCrLf
query1 = query1 & " where id='" & uid & "'" & vbCrLf
rsget.Open query1,dbget,1

if not rsget.EOF  then
    if trim(replace(company_no,"-","")) = replace(rsget("company_no"),"-","") then
        if rsget("groupid") <> "" then
            response.write("C")
        else
            response.write("T")
        end if
    else
        response.write("F")
    end if
end if
rsget.close
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->