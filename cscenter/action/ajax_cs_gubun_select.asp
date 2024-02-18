<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %>
<%
'###########################################################
' Description : csºæ≈Õ
' History : ¿ÃªÛ±∏ ª˝º∫
'           2023.11.07 «—øÎπŒ ºˆ¡§(ƒı∏Æ∆©¥◊)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<%
dim sqlStr, ArrData, gubun01, gubun02, name_gubun01, name_gubun02, name_gubun01name, name_gubun02name
dim name_frm, targetDiv, i
gubun01         = request("gubun01")
gubun02         = request("gubun02")
name_gubun01    = request("name_gubun01")
name_gubun02    = request("name_gubun02")
name_gubun01name= request("name_gubun01name")
name_gubun02name= request("name_gubun02name")
name_frm        = request("name_frm")
targetDiv       = request("targetDiv")

'response.write "gubun01=" + gubun01 + "<br>"
'response.write "gubun02=" + gubun02 + "<br>"
'response.write "name_gubun01=" + name_gubun01 + "<br>"
'response.write "name_gubun02=" + name_gubun02 + "<br>"
'response.write "name_gubun01name=" + name_gubun01name + "<br>"
'response.write "name_gubun02name=" + name_gubun02name + "<br>"

sqlStr = " select top 300"
sqlStr = sqlStr + " c1.comm_cd, c2.comm_cd, c1.comm_name, c2.comm_name"
sqlStr = sqlStr + " from [db_cs].[dbo].tbl_cs_comm_code c1 with (nolock)"
sqlStr = sqlStr + " join [db_cs].[dbo].tbl_cs_comm_code c2 with (nolock)"
sqlStr = sqlStr + "     on c1.comm_cd=c2.comm_group"
sqlStr = sqlStr + " where c1.comm_group='Z020'"
sqlStr = sqlStr + " and c2.comm_group='" + gubun01 + "'"
sqlStr = sqlStr + " and c1.comm_isDel='N'"
sqlStr = sqlStr + " and c2.comm_isDel='N'"
sqlStr = sqlStr + " order by c2.comm_cd asc"

if gubun01<>"" then
    'response.write sqlStr & "<br>"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if not rsget.Eof Then
    	ArrData = rsget.getRows()
    end if
    rsget.Close
end if
%>
<table width="240" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
    <td align="right" colspan="2">
        <input class="csbutton" type="button" value=" ªË¡¶ " onClick="delGubun('<%= name_gubun01 %>','<%= name_gubun02 %>','<%= name_gubun01name %>','<%= name_gubun02name %>','<%= name_frm %>','<%= targetDiv %>');">
        <input class="csbutton" type="button" value=" x " onClick="colseCausepop('<%= targetDiv %>');">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="left" width="100" valign="top">
        <% call drawSelectBoxCSCommComboOnChange("combogubun01",gubun01,"Z020","divCsAsGubunSelect(this.value,'','" + name_gubun01 + "','" + name_gubun02 + "','" + name_gubun01name + "','" + name_gubun02name + "','" + name_frm +"','" + targetDiv + "');")  %>
    </td>
    <td valign="top">
        <table width="140" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" class="a">
        <% if (isArray(ArrData)) then %>
            <% for i = 0 To UBound(ArrData,2) %>
            <tr bgcolor="#FFFFFF">
                <% if (ArrData(1,i)=gubun02) then %>
                <td><font color="red"><%=ArrData(3,i)%></font></td>
                <% else %>
                <td><a href="javascript:selectGubun('<%=ArrData(0,i)%>','<%=ArrData(1,i)%>','<%=ArrData(2,i)%>','<%=ArrData(3,i)%>','<%= name_gubun01 %>','<%= name_gubun02 %>','<%= name_gubun01name %>','<%= name_gubun02name %>','<%= name_frm %>','<%= targetDiv %>');"><%=ArrData(3,i)%></a></td>
                <% end if %>
            </tr>
            <% next %>
        <% end if %>
        </table>
    </td>
</table>
<%
function drawSelectBoxCSCommComboOnChange(selectBoxName,selectedId,groupCode,onChangefunction)
   dim tmp_str,sqlStr
   %>
     <select name="<%=selectBoxName%>" onChange="<%= onChangefunction %>" size="10">
     <option value='' <%if selectedId="" then response.write " selected" %> >º±≈√</option>
   <%
       sqlStr = " select comm_cd,comm_name "
       sqlStr = sqlStr + " from [db_cs].[dbo].tbl_cs_comm_code with (nolock)"
       sqlStr = sqlStr + " where comm_group='" + groupCode + "' "
       sqlStr = sqlStr + " and comm_isDel='N' "
       sqlStr = sqlStr + " order by comm_cd "

        'response.write sqlStr & "<br>"
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    
       if  not rsget.EOF  then
           do until rsget.EOF
               if LCase(selectedId) = LCase(rsget("comm_cd")) then
                   tmp_str = " selected"
               end if
               response.write("<option value='" & rsget("comm_cd") & "' " & tmp_str & ">" + db2html(rsget("comm_name")) + " </option>")
               tmp_str = ""
               rsget.MoveNext
           loop
       end if
       rsget.close
   %>
       </select>
   <%
End function
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->