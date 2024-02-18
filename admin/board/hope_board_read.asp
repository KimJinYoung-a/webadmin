<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/hope_boardcls.asp"-->
  
<%
dim ohope,idx
idx = request("idx")

set ohope = new CHopeBoardDetail
ohope.read idx


%>
  <table width="580" border="0" align="center" cellpadding="0" cellspacing="3">
    <tr> 
      <td background="/admin/images/topbar_bg.gif" height="25" valign="middle"> 
          <table width="520" border="0" cellpadding="0" cellspacing="0" class="a">
            <tr> 
              <td> 
                <div align="left"><span class="a"><b>☞ <%=ohope.FTitle %></b></span></div>
              </td>
            </tr>
          </table>
      </td>
    </tr>
    <tr> 
      <td class="a" height="5">글쓴이: <%=ohope.Fusername %>&nbsp;| 날짜: <% =(ohope.Fregdate) %></td>
    </tr>
    <tr> 
      <td><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
    </tr>
     <tr> 
      <td valign="top" class="a"> 
        내용 :<br>
         <%=ohope.FContents %>
          <br>
      </td>
    </tr>
    <tr> 
    <td height="2"><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
    </tr>
 </table>
 </td>
</tr>
</table>
<%
set ohope = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->