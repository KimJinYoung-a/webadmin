<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/lib/companybodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/boardcls.asp"-->
  
<%
	Dim ix,i, page, pgsize
	Dim TotalPage, TotalCount
	Dim prepage, nextpage
	Dim mode,idx
	Dim nIndent, strtitle
	Dim nInstr,searchmode,search,searchString
    Dim nboard

	idx = request("idx")

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

if Request("delmode") = "delete" then
nboard.FTableName = "[db_board].[dbo].tbl_partner_board"
nboard.FRectIdx = idx
nboard.FRectDesignerID = session("ssBctID")
nboard.design_board_del

else
nboard.FTableName = "[db_board].[dbo].tbl_partner_board"
nboard.design_board_read request("idx")
end if


%>
<script language="JavaScript">
<!--

function gotoreply(){
location.href = "partner_board_write.asp?replymode=reply&idx=<%= request("idx") %>&page=<% =page %>&pgsize=<% =pgsize %>&menupos=94"
}

function gotolist(){
location.href = "partner_board.asp?idx=<%= request("idx") %>&page=<% =page %>&menupos=94"
}

function gotomodify(){
location.href = "partner_board_modify.asp?idx=<%= request("idx") %>&page=<% =page %>&pgsize=<% =pgsize %>&menupos=94"
}

function gotodelete(){
//	if (CheckMember() == true){

location = "partner_board_read.asp?menupos=94&delmode=delete&page=<%=page%>&idx=<%= request("idx") %>";

}
//-->
</script>
<input type="hidden" name="menupos" value="<%= menupos %>">
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
      글쓴이: <span class="id"><a href="mailto:<%=nboard.FRectEmail %>"><%=nboard.FRectName %></a></span>&nbsp;| 날짜: <% =(nboard.Fregdate) %></td>
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
	<td><input type="button" value="글 삭제" onclick="gotodelete();">&nbsp;<input type="button" value="글 수정" onclick="gotomodify();">&nbsp;<input type="button" value="답변" onclick="gotoreply();">&nbsp;<input type="button" value="List" onclick="gotolist();"></td>
</tr>
</table>
<%
set nboard = Nothing
%>
<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->