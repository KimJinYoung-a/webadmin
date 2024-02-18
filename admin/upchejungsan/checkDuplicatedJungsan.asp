<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
dim yyyymm, mastercode
yyyymm      = request("yyyymm")
mastercode  = request("mastercode")

dim sqlStr,resultRows, resultRows2

sqlStr = " select top 4000 d.detailidx,d.mastercode,d.itemid,d.itemoption, count(detailidx) as cnt"
sqlStr = sqlStr & " from"
sqlStr = sqlStr & " [db_jungsan].[dbo].tbl_designer_jungsan_master m,"
sqlStr = sqlStr & "  [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
sqlStr = sqlStr & " where m.id=d.masteridx"
sqlStr = sqlStr & " and m.yyyymm>='" & yyyymm & "'"
sqlStr = sqlStr & " group by  d.detailidx,d.mastercode,d.itemid,d.itemoption"
sqlStr = sqlStr & " having count(detailidx)>1"

rw sqlStr
'if (yyyymm<>"") then
'    rsget.Open sqlStr,dbget,1
'    if Not rsget.Eof then 
'        resultRows = rsget.getRows()   
'    end if 
'    rsget.close
'end if


dim i,cnt, cnt2, cnt3
if IsArray(resultRows) then
    cnt = Ubound(resultRows,2)
else
    cnt = 0
end if

if (mastercode<>"") then
    sqlStr = " select top 10 * from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
    sqlStr = sqlStr & " where mastercode='" & mastercode & "'"
    
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then 
        resultRows2 = rsget.getRows()   
    end if 
    rsget.close
end if


if IsArray(resultRows2) then
    cnt2 = Ubound(resultRows2,2)
else
    cnt2 = 0
end if

'response.write Ubound(resultRows,1)
'response.write ","
'response.write Ubound(resultRows,2)
'response.write ","
%>
<script language='javascript'>
function detailSearch(mastercode){
    frmResearch.mastercode.value = mastercode;
    frmResearch.submit();
}
</script>
<form name="frmResearch">
<input type="hidden" name="yyyymm" value="<%= yyyymm %>">
<input type="hidden" name="mastercode" value="">
</form>
<%= yyyymm %> 이후 중복정산내역
<% if IsArray(resultRows) then %>
<table width="400" border="0" cellspacing="1" cellpadding="2" class="a" bgcolor="#CCCCCC">
<tr>
    <td>IDX </td>
    <td>M Code</td>
    <td>상품코드</td>
    <td>옵션코드</td>
    <td>갯수</td>
</tr>
<% for i=0 to cnt %>
<tr bgcolor="#FFFFFF">
    <td><%= resultRows(0,i) %></td>
    <td><a href="javascript:detailSearch('<%= resultRows(1,i) %>');"><%= resultRows(1,i) %></a></td>
    <td><%= resultRows(2,i) %></td>
    <td><%= resultRows(3,i) %></td>
    <td><%= resultRows(4,i) %></td>
</tr>
<% next %>
</table>
<% end if %>

<br><br>

<% if IsArray(resultRows2) then %>
<table width="100%" border="0" cellspacing="1" cellpadding="2" class="a" bgcolor="#CCCCCC">
<tr>
    <td>.</td>
    <td>.</td>
    <td>.</td>
    <td>.</td>
    <td>.</td>
</tr>
<% for i=0 to cnt2 %>
<tr bgcolor="#FFFFFF">
    <td><%= resultRows2(0,i) %></td>
    <td><%= resultRows2(1,i) %></td>
    <td><%= resultRows2(2,i) %></td>
    <td><%= resultRows2(3,i) %></td>
    <td><%= resultRows2(4,i) %></td>
    <td><%= resultRows2(5,i) %></td>
    <td><%= resultRows2(6,i) %></td>
    <td><%= resultRows2(7,i) %></td>
    <td><%= resultRows2(8,i) %></td>
    <td><%= resultRows2(9,i) %></td>
    <td><%= resultRows2(10,i) %></td>
    <td><%= resultRows2(11,i) %></td>
    <td><%= resultRows2(12,i) %></td>
    <td><%= resultRows2(13,i) %></td>
</tr>
<% next %>
</table>
<% end if %>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->