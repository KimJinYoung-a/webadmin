<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/boardcls.asp"-->
  
<%
	Dim ix,i, page, pgsize
	Dim TotalPage, TotalCount
	Dim prepage, nextpage
	Dim mode
	Dim nIndent, strtitle
	Dim nInstr,searchmode,search,searchString
    Dim nboard

	if Request("pgsize")="" then
		pgsize = 10
	else
		pgsize = Request("pgsize")
	end if
	
	if Request("page") = "" then
		page = 1
	else
		page = cInt(Request("page")) 
	end if

set nboard = new CBoard
nboard.FTableName = "[db_board].[10x10].tbl_partner_notice"
nboard.design_notice_read request("idx")

%>
<script language="JavaScript">
<!--
function gotolist(){
location.href = "partner_notice.asp?idx=<%= request("idx") %>&page=<% =page %>&menupos=89"
}
function gotomodify(){
location.href = "partner_notice_modify.asp?idx=<%= request("idx") %>&page=<% =page %>&pgsize=<% =pgsize %>&menupos=89"
}
//-->
</script>
  <table width="580" border="0" align="center" cellpadding="0" cellspacing="3">
    <tr> 
      <td background="/admin/images/topbar_bg.gif" height="25" valign="middle"> 
          <table width="520" border="0" cellpadding="0" cellspacing="0" class="a">
            <tr> 
              <td> 
                <div align="left"><span class="a"><b>☞ <%=nboard.FRectTitle %></b></span></div>
              </td>
            </tr>
          </table>
      </td>
    </tr>
    <tr> 
      <td class="a" height="5"> 아이디: <span class="id"><% =nboard.FRectID %></span> &nbsp;|
      글쓴이: <span class="id"><%=nboard.FRectName %></span>&nbsp;| 날짜: <% =(nboard.Fregdate) %></td>
    </tr>
    <tr> 
      <td><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
    </tr>
     <tr> 
      <td valign="top" class="a"> 
        내용 :<br>
         <%=nboard.FRectContents %>
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
<table width="580" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
	<td><input type="button" value="글 수정" onclick="gotomodify();">&nbsp;<input type="button" value="List" onclick="gotolist();"></td>
</tr>
</table>
<%
set nboard = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->