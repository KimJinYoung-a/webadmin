<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/sitemasterClass/main_reviewCls.asp"-->
<!-- #include virtual="/lib/classes/sitemasterClass/main_TopReviewCls.asp"-->
<%
dim seachbox,eventbox
seachbox = request("seachbox")

Dim page
page = Request("page")
If page="" Then page = 1

Dim idx

idx = Request("idx")

dim oeventuserlist , i

	set oeventuserlist = new Ceventuserlist
		oeventuserlist.FPagesize = 20
		oeventuserlist.FCurrPage = page
		oeventuserlist.frectseachbox = seachbox
		oeventuserlist.Feventuserlist99()


	dim omainreview
	Set omainreview = new CSearchKeyWord
	omainreview.FRectidx = idx

	if idx<>"" then
		omainreview.GetSearchreview
	end if


%>
<script language='javascript'>
	// 페이지 이동
	function goPage(pg)
	{
		document.searchfrm.page.value=pg;
		document.searchfrm.action="main_topkeyword_write_1.asp";
		document.searchfrm.submit();
	}
	function choice(uid,cmt,iid,Lcate,Mcate)
	{
		document.frm.userid.value= uid;
		document.frm.comment.value=cmt;
		document.frm.itemid.value=iid;
		document.frm.cate_large.value=Lcate;
		document.frm.cate_mid.value=Mcate;
	}
	function goSubmit()
	{
		// id 입력여부 검사
		if(!document.frm.userid.value) {
			alert("관련 키워드를 입력해주세요.");
			document.frm.userid.focus();
			return;
		}
		// 코멘트 입력여부 검사
		if(!document.frm.comment.value) {
			alert("키워드 클릭시 이동할 링크를 입력해주세요.");
			document.frm.comment.focus();
			return;
		}

		// 순서 입력여부 검사
		if(!document.frm.sortNo.value) {
			alert("표시 순서를 입력해주세요.\n※ 순서는 숫자이며 적을수록 순번이 높습니다.");
			document.frm.sortNo.focus();
			return;
		}

		<% if idx="" then %>
		if(confirm("작성하신 내용을 등록하시겠습니까?")) {
			document.frm.mode.value="add";
			document.frm.action="doMainReview.asp";
			document.frm.submit();
		}
		<% else %>
		if(confirm("수정하신 내용을 저장하시겠습니까?")) {
			document.frm.mode.value="modify";
			document.frm.action="doMainReview.asp";
			document.frm.submit();
		}
		<% end if %>
	}

</script>
<!-- 검색 시작 -->
<form name="searchfrm" method="post" >
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

	<tr align="center" bgcolor="#FFFFFF" >
		<td width="100" bgcolor="<%= adminColor("gray") %>">검색조건</td>
		<td align="left">
			&nbsp;ItemID: <input type="text" name="seachbox" value="<%= seachbox %>" size="10">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.searchfrm.submit();">
		</td>
	</tr>
</table>
</form>
<!-- 검색 끝 -->

<% if seachbox <> "" or idx <> ""  then %>
<!-- 리스트 시작 -->
<form name="frm" method="post" action="doMainReview.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="cate_large" value="">
<input type="hidden" name="cate_mid" value="">
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<% if idx="" then %>
		<font color="red"><b>코멘트 등록</b></font>
		<% else %>
		<font color="red"><b>코멘트 수정</b></font>
		<% end if%>
	</td>
</tr>
<% if idx<>"" then %>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">일련번호</td>
	<td align="left"><input type="text" name="idx" value="<%=idx%>" readonly size="10" class="text_ro"></td>
</tr>
<% end if %>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">User ID</td>
	<td align="left"><input type="text" name="userid" value="<% if idx<> "" then Response.Write omainreview.FitemList(0).fuserid %>" size="32" readonly maxlength="32" class="text"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">코멘트</td>
	<td align="left">
		<table cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td colspan="2"><input type="text" bgcolor="#707080" name="comment" value="<% if idx<>"" then Response.Write omainreview.FitemList(0).fcomment%>" size="200" readonly class="text"></td>
		<tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">표시순서</td>
	<td align="left"><input type="text" name="sortNo" value="<% if idx<>"" then Response.Write omainreview.FitemList(0).FsortNo: else Response.Write "0" %>" size="3" class="text"></td></td>
</tr>
	<% if oeventuserlist.ftotalcount > 0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="17">
			검색결과 : <b><%= oeventuserlist.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center">ID</td>
		<td align="center">Comment</td>


    </tr>

	<% for i=0 to oeventuserlist.FResultCount-1 %>
    	<tr align="center" bgcolor="#FFFFFF">
			<td align="center"><a href="javascript:choice('<%= oeventuserlist.flist(i).fuserid %>','<%= chrbyte(oeventuserlist.flist(i).fcontents,300,"Y") %>','<%= oeventuserlist.flist(i).fitemid %>','<%= oeventuserlist.flist(i).fcate_large %>','<%= oeventuserlist.flist(i).fcate_mid %>')"><%= oeventuserlist.flist(i).fuserid %></a></td>
			<td align="center"><a href="javascript:choice('<%= oeventuserlist.flist(i).fuserid %>','<%= chrbyte(oeventuserlist.flist(i).fcontents,300,"Y") %>','<%= oeventuserlist.flist(i).fitemid %>','<%= oeventuserlist.flist(i).fcate_large %>','<%= oeventuserlist.flist(i).fcate_mid %>')"><%= oeventuserlist.flist(i).fcontents %></a></td>


    	</tr>
	<% next %>

	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="7" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>

</table>
<% end if %>
</form>

<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr valign="bottom" bgcolor="FFFFFF">
			<td align="center">
			<!-- 페이지 시작 -->
			<%
				if oeventuserlist.HasPreScroll then
					Response.Write "<a href='javascript:goPage(" & oeventuserlist.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
				else
					Response.Write "[pre] &nbsp;"
				end if

				for i=0 + oeventuserlist.StartScrollPage to oeventuserlist.FScrollCount + oeventuserlist.StartScrollPage - 1

					if i>oeventuserlist.FTotalpage then Exit for

					if CStr(page)=CStr(i) then
						Response.Write " <font color='red'>[" & i & "]</font> "
					else
						Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
					end if

				next

				if oeventuserlist.HasNextScroll then
					Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
				else
					Response.Write "&nbsp; [next]"
				end if
			%>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" colspan="2" bgcolor="#FFFFFF">
		<input type="button" class="button" value="저장" onClick="goSubmit()"> &nbsp;
		<input type="button" class="button" value="취소" onClick="self.history.back()">
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->