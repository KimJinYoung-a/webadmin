<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/boardcls.asp"-->

<%
	Dim ix,i, page, pgsize
	Dim TotalPage, TotalCount
	Dim prepage, nextpage
	Dim mode
	Dim nIndent ,idx,ref,ref_level,ref_serial,ref_userid
	Dim name,email,title,contents
    Dim nboard

set nboard = new CBoard

if Request.Form("writemode") = "write" then

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


ref = Request("ref")
ref_level = Request("ref_level")
ref_serial = Request("ref_serial")
ref_userid = Request("ref_userid")
name = Request("name")
email = Request("email")
title = Request("title")
contents = Request("contents")
title = html2db(title)
contents = html2db(contents)

nboard.FTableName = "[db_board].[10x10].tbl_designer_board"
nboard.FRectIdx = request("idx")
nboard.FRectDesignerID = session("ssBctId")
nboard.FPageSize = pgsize
nboard.FRectName = name
nboard.FRectEmail = email
nboard.FRectTitle = title
nboard.FRectContents = contents
nboard.FRectRef = ref
nboard.FRectLevel = ref_level
nboard.FRectSerial = ref_serial
nboard.FRectRefuserid = ref_userid
nboard.FCurrPage = page
nboard.design_board_write

response.redirect "designer_board.asp?menupos=90"
end if

if request("replymode") = "reply" then
nboard.FTableName = "[db_board].[10x10].tbl_designer_board"
nboard.FRectIdx = request("idx")
nboard.design_board_reply
end if

%>

<script language="javascript">
	<!--
	function checkform()
	{
		if (document.boardform.title.value == "") {
			alert("제목을 입력해 주십시요...");
			document.boardform.title.focus();
			return ;
		}
		else if (document.boardform.contents.value == "") {
			alert("내용을 입력해 주십시요");
			document.boardform.contents.focus();
			return ;
		}
		else if (document.boardform.ref_userid.value == "") {
			alert("업체ID를 선택 해 주세요");
			document.boardform.ref_userid.focus();
			return ;
		}
	    else {
			document.boardform.submit();
		}

	}

	//-->
</script>

<form method="POST" name="boardform" action="designer_board_write.asp" name="boardform">
		<input type=hidden name="idx" value=<%=nboard.FRectIdx%>>
		<input type=hidden name="ref" value=<%=nboard.FRectRef%>>
		<input type=hidden name="ref_level" value=<%=nboard.FRectLevel%>>
		<input type=hidden name="ref_serial" value=<%=nboard.FRectSerial%>>
        <input type="hidden" name="writemode" value="write">

	<table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td class="a" width="409"><b><img src="/admin/images/mini_icon.gif" width="17" height="17">
              디자이너 게시판 쓰기</b></td>
            <td class="a">
              <div align="right"> </div>
            </td>
          </tr>
        </table>
<br>
        <table width="750" border="0" cellpadding="3" cellspacing="1">
		  <tr>
            <td width="100" bgcolor="#eeeeee" height="7">
              <div align="right"><font color="#CCCCCC" class="a">아이디 : </font></div>
            </td>
            <td width="407" height="7"  class="a">
			  <% = session("ssBctId") %>
            </td>
          </tr>
          <tr>
            <td width="100" bgcolor="#eeeeee" height="7">
              <div align="right"><font color="#CCCCCC" class="a">글쓴이 : </font></div>
            </td>
            <td width="407" height="7">
			  <input type="text" name="name" maxlength="32" value='<%=session("ssBctCname")%>'>
            </td>
          </tr>
          <tr>
            <td width="100" bgcolor="#eeeeee" height="7">
              <div align="right"><font color="#CCCCCC" class="a">업체 : </font></div>
            </td>
            <td width="407" height="7">
			  <% drawSelectBoxDesigner "ref_userid",ref_userid %>
            </td>
          </tr>
          <tr>
            <td width="100" bgcolor="#eeeeee" height="6">
              <div align="right"><font color="#CCCCCC" class="a">메일 : </font></div>
            </td>
            <td width="407" height="6">
              <input type="text" name="email" size="54" maxlength="128">
            </td>
          </tr>
          <tr>
            <td width="100" bgcolor="#eeeeee" height="6">
              <div align="right"><font color="#CCCCCC" class="a">제목 : </font></div>
            </td>
            <td width="407" height="6">
              <input type="text" name="title" size="54" maxlength="128">
            </td>
          </tr>
          <tr>
            <td width="100" bgcolor="#eeeeee">
              <div align="right" class="a"><font color="#CCCCCC" class="a">게시판
                내용 : </font></div>
            </td>
            <td>
              <textarea name="contents" cols="53" rows="15"><% =nboard.FRectContents  %></textarea>
            </td>
          </tr>
        </table>
        <table border="0" align="center" cellpadding="0" cellspacing="5">
          <tr>
            <td height="2">
              <div align="right">
                <p><a href="javascript:checkform()"><img src="/admin/images/write_butten.gif" width="55" border="0"></a></p>
              </div>
            </td>
            <td valign="top" height="2">
              <div align="center"><a href="javascript:history.back()"><img src="/admin/images/cancle_butten.gif" width="55" border="0"></a></div>
            </td>
          </tr>
        </table>
       </form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->