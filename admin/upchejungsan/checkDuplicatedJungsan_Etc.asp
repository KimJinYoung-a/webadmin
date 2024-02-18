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


sqlStr = " select T.designerid, D.detailidx, d.mastercode, d.itemname, d.sellcash, d.suplycash from ("
sqlStr = sqlStr + " select d.mastercode, m.designerid, count(*) as CNT from db_jungsan.dbo.tbl_designer_jungsan_master m"
sqlStr = sqlStr + " 	Join db_jungsan.dbo.tbl_designer_jungsan_detail d"
sqlStr = sqlStr + " 	on m.id=d.masteridx"
sqlStr = sqlStr + " where m.yyyymm>='"&yyyymm&"'"
sqlStr = sqlStr + " and d.gubuncd='witakchulgo'"
sqlStr = sqlStr + " and d.buyname is Not NULL"
sqlStr = sqlStr + " group by d.mastercode, m.designerid"
sqlStr = sqlStr + " having count(*)>1"
sqlStr = sqlStr + " ) T Join db_jungsan.dbo.tbl_designer_jungsan_detail D"
sqlStr = sqlStr + " 	on T.mastercode=D.mastercode"
sqlStr = sqlStr + " 	and D.gubuncd='witakchulgo'"
sqlStr = sqlStr + " order by D.Mastercode"


if (yyyymm<>"") then
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then 
        resultRows = rsget.getRows()   
    end if 
    rsget.close
end if



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
<%= yyyymm %> 이후 추가정산검토

<p>
<% dim flip : flip =false %>
<% if IsArray(resultRows) then %>
<table width="100%" border="0" cellspacing="1" cellpadding="2" class="a" bgcolor="#CCCCCC">
<tr>
    <td>브랜드</td>
    <td>.</td>
    <td>주문번호</td>
    <td>내역</td>
    <td> 판매가</td>
    <td>매입가</td>
</tr>
<% for i=0 to cnt %>
<tr bgcolor="<%= chkiif(flip,"#AAAAEE","#EEDDEE") %>">
    <td><%= resultRows(0,i) %></td>
    <td><%= resultRows(1,i) %></td>
    <td><%= resultRows(2,i) %></td>
    <td><%= resultRows(3,i) %></td>
    <td><%= resultRows(4,i) %></td>
    <td><%= resultRows(5,i) %></td>
</tr>
<% 
    if i<cnt then 
        if resultRows(2,i)<>resultRows(2,i+1) then
            flip = Not flip
        end if
    end if
%>
<% next %>
</table>
<% end if %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->