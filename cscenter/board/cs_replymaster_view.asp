<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 고객센터 [안내문구] 기본 카테고리
' History : 이상구 생성
'			2021.09.10 한용민 수정(이문재이사님요청 자사몰 필드추가, 소스표준화, 보안강화)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_replycls.asp"-->
<%
dim idx, gubunCode, currmode
	idx = requestcheckvar(getNumeric(request("idx")),10)
	gubunCode = requestcheckvar(request("gubunCode"),4)

dim oCReply
Set oCReply = new CReply
if (idx <> "") then
	currmode = "modiMaster"
	oCReply.FRectMasterIDX = idx
	oCReply.GetReplyMasterOne()
else
	currmode = "insMaster"
	oCReply.GetReplyMasterEmptyOne()
	oCReply.FOneItem.FgubunCode = gubunCode
end if

%>
<script type="text/javascript">

function fnSaveReplyMaster() {
	var frm = document.frm;

	if (frm.sitename.value == "") {
		alert("구분을 입력하세요.");
		frm.sitename.focus();
		return;
	}
	if (frm.title.value == "") {
		alert("카테고리명을 입력하세요.");
		frm.title.focus();
		return;
	}

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
	document.location.href = "/cscenter/board/cs_replymaster_list.asp?menupos=<%= menupos %>";
}

</script>

<form name="frm" method="post" action="/cscenter/board/cs_reply_process.asp" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="<%= currmode %>">
<input type="hidden" name="masteridx" value="<%= oCReply.FOneItem.Fidx %>">
<input type="hidden" name="gubunCode" value="<%= oCReply.FOneItem.FgubunCode %>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFFFFF" height="30" colspan=2>※ 기본 카테고리 <% if (currmode = "insMaster") then %>작성<% else %>수정<% end if %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width=100 bgcolor="<%= adminColor("tabletop") %>" height="30">구분</td>
	<td>
		<% Drawreplysitename "sitename", oCReply.FOneItem.fsitename, "" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width=100 bgcolor="<%= adminColor("tabletop") %>" height="30">기본 카테고리명</td>
	<td>
		<input type="text" class="text" name="title" value="<%= oCReply.FOneItem.Ftitle %>" size="40">
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
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
