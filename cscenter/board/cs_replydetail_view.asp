<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_replycls.asp"-->
<%

dim idx, currmode
dim masterIdx, gubunCode, masterUseYN

idx = request("idx")
masterIdx = request("masterIdx")
gubunCode = request("gubunCode")
masterUseYN = request("masterUseYN")


dim oCReply
Set oCReply = new CReply

if (idx <> "") then
	currmode = "modiDetail"
	oCReply.FRectDetailIDX = idx
	oCReply.GetReplyDetailOne()
else
	currmode = "insDetail"
	oCReply.GetReplyDetailEmptyOne()

	oCReply.FOneItem.Fmasteridx = masterIdx
	oCReply.FOneItem.FgubunCode = gubunCode
	oCReply.FOneItem.FmasterUseYN = masterUseYN
end if

%>
<script language='javascript'>

function fnSaveReplyMaster() {
	var frm = document.frm;

	if (frm.masterIdx.value == "") {
		alert("기본 카테고리를 선택하세요.");
		frm.masterIdx.focus();
		return;
	}

	if (frm.subtitle.value == "") {
		alert("상세 카테고리명을 입력하세요.");
		frm.subtitle.focus();
		return;
	}

	/*
	if (frm.contents.value == "") {
		alert("안내문구를 입력하세요.");
		frm.contents.focus();
		return;
	}
	*/

	if (frm.dispOrderNo.value == "") {
		alert("표시순서를 입력하세요.");
		frm.dispOrderNo.focus();
		return;
	}

	if (frm.dispOrderNo.value*0 != 0) {
		alert("표시순서는 숫자만 가능합니다.");
		frm.dispOrderNo.focus();
		return;
	}

	if (confirm("저장하시겠습니까?") == true) {
		frm.submit();
	}
}

function fnGotoList() {
	document.location.href = "cs_replydetail_list.asp?menupos=<%= menupos %>";
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td>
	        	<img src="/images/icon_arrow_down.gif" align="absbottom">
	        	<font color="red"><strong>상세 카테고리 <% if (currmode = "insMaster") then %>작성<% else %>수정<% end if %></strong></font>
	        </td>
	        <td align="right">

	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="cs_reply_process.asp">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="mode" value="<%= currmode %>">
	<input type="hidden" name="detailidx" value="<%= oCReply.FOneItem.Fidx %>">
	<input type="hidden" name="gubunCode" value="<%= oCReply.FOneItem.FgubunCode %>">
	<tr bgcolor="#FFFFFF">
		<td width=100 bgcolor="<%= adminColor("tabletop") %>" height="30">기본 카테고리명</td>
		<td>
			<% Call drawSelectBoxReplyMaster("masterIdx", oCReply.FOneItem.Fmasteridx, oCReply.FOneItem.FgubunCode, oCReply.FOneItem.FmasterUseYN) %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width=100 bgcolor="<%= adminColor("tabletop") %>" height="30">상세 카테고리명</td>
		<td>
			<input type="text" class="text" name="subtitle" value="<%= oCReply.FOneItem.Fsubtitle %>" size="40">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" height="30">안내문구</td>
		<td>
			<textarea class="textarea" name="contents" cols="100" rows="10"><%= oCReply.FOneItem.Fcontents %></textarea>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" height="30">표시순서</td>
		<td>
			<input type="text" class="text" name="dispOrderNo" value="<%= oCReply.FOneItem.FdispOrderNo %>" size="4">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" height="30">사용구분</td>
		<td>
			<select class="select" name="useYN">
				<option value="Y" <% if (oCReply.FOneItem.FuseYN = "Y") then %>selected<% end if %> >사용함</option>
				<option value="N" <% if (oCReply.FOneItem.FuseYN = "N") then %>selected<% end if %> >사용안함</option>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>" height="30">등록자</td>
		<td>
			<%= oCReply.FOneItem.Freguserid %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">최종수정</td>
		<td>
			<%= oCReply.FOneItem.Flastupdate %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2" height="35" align="center">
			<input type="button" class="button" value="저장하기" onclick="fnSaveReplyMaster()">
			&nbsp;
			<input type="button" class="button" value="목록으로" onclick="fnGotoList()">
		</td>
	</tr>
	</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
