<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim bardata, bardataArr
bardata = request("bardata")



bardataArr = split(bardata,VbCRLF)

Dim CNT : CNT = UBOUND(bardataArr)

dim isBarExists : isBarExists=CNT>-1
dim i,j
dim Cols : cols=2
dim Rows : Rows=CNT\cols+1

'rw "Cols="&Cols&","&"Rows="&Rows

%>
<% if (isBarExists) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% for i=0 to Rows-1 %>
<tr bgcolor="#FFFFFF">
    <% for j=0 to Cols-1 %>
    <td>
        <% if CNT>=i*cols+j then %>
            <% if bardataArr(i*cols+j)<>"" then %>
                <img src="http://company.10x10.co.kr/barcode/barcode.asp?image=3&type=23&data=<%=TRIM(bardataArr(i*cols+j))%>&height=30&barwidth=1&Size=7" border=0>
            <% end if %>
        <% end if %>
    </td>
    <% next %>
</tr>    
<% next %>
</table>
<% else %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">
<tr bgcolor="#FFFFFF">
    <td>
        <textarea cols="80" rows="10" name="bardata"></textarea>
    </td>
    <td><input type="button" value="»ý¼º" onClick="document.frm.submit();"></td>
</tr>    
</form>
</table>
<% end if %>
<!-- #include virtual="/common/lib/poptail.asp"-->