<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
Dim s_thread
s_thread = request("s_thread")
if (s_thread = "") then
            response.write("<script>window.alert('s_thread값이 넘어오지 않았습니다.');</script>")
            response.write("<script>history.go(-1);</script>")
            dbget.close()	:	response.End
end if
dim table_name
table_name = request("table_name")
if (table_name = "") then
            response.write("<script>window.alert('table 구분자가 넘어오지 않았습니다.');</script>")
            response.write("<script>history.go(-1);</script>")
            dbget.close()	:	response.End
end if
dim gotopage
gotopage = request("gotopage")

dim query1
query1 = "update "+table_name+" set check_flag = 'Y' where thread = "&s_thread&"  "
dbget.Execute query1

if Err.number <> 0 then
%>
    <script LANGUAGE="JavaScript">
    <!--
     alert("에러가 발생했습니다.");
     history.back();
    //-->
    </script>
<%
else
    response.redirect "boardlist.asp"
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
