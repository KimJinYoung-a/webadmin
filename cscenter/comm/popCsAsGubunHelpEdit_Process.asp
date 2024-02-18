<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim comm_cd
dim infoHtml_B001, infoHtml_B007
comm_cd = requestCheckVar(request("comm_cd"),4)
infoHtml_B001 = request("infoHtml_B001")
infoHtml_B007 = request("infoHtml_B007")

dim refer, sqlStr
refer = request.ServerVariables("HTTP_REFERER")

sqlStr = " IF Exists(select * from db_cs.dbo.tbl_cs_comm_div_info where div_comm_cd='" + comm_cd + "')"
sqlStr = sqlStr + " BEGIN"
sqlStr = sqlStr + "     update db_cs.dbo.tbl_cs_comm_div_info"
sqlStr = sqlStr + "     set infoHtml='" + html2Db(infoHtml_B001) + "'"
sqlStr = sqlStr + "     where div_comm_cd='" + comm_cd + "'"
sqlStr = sqlStr + "     and state_comm_cd='B001'"
sqlStr = sqlStr + "     update db_cs.dbo.tbl_cs_comm_div_info"
sqlStr = sqlStr + "     set infoHtml='" + html2Db(infoHtml_B007) + "'"
sqlStr = sqlStr + "     where div_comm_cd='" + comm_cd + "'"
sqlStr = sqlStr + "     and state_comm_cd='B007'"
sqlStr = sqlStr + " END"
sqlStr = sqlStr + " ELSE"
sqlStr = sqlStr + " BEGIN"
sqlStr = sqlStr + "     insert into db_cs.dbo.tbl_cs_comm_div_info"
sqlStr = sqlStr + "     (div_comm_cd, state_comm_cd, infoHtml)"
sqlStr = sqlStr + "     values("
sqlStr = sqlStr + "     '" + comm_cd + "'"
sqlStr = sqlStr + "     ,'B001'"
sqlStr = sqlStr + "     ,'" + html2Db(infoHtml_B001) + "'"
sqlStr = sqlStr + "     )"
sqlStr = sqlStr + "     insert into db_cs.dbo.tbl_cs_comm_div_info"
sqlStr = sqlStr + "     (div_comm_cd, state_comm_cd, infoHtml)"
sqlStr = sqlStr + "     values("
sqlStr = sqlStr + "     '" + comm_cd + "'"
sqlStr = sqlStr + "     ,'B007'"
sqlStr = sqlStr + "     ,'" + html2Db(infoHtml_B007) + "'"
sqlStr = sqlStr + "     )"
sqlStr = sqlStr + " END"

dbget.Execute sqlStr

%>

<script language='javascript'>
    alert('수정되었습니다.');
    location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->