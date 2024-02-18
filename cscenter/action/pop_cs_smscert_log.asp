<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_smscertcls.asp" -->
<%

dim i
dim userid, usercell, usermail

userid = requestCheckVar(request("userid"), 32)
usercell = requestCheckVar(request("usercell"), 32)
usermail = requestCheckVar(request("usermail"), 128)


dim occssmscert
set occssmscert = New CCSSMSCert

occssmscert.FCurrPage = 1
occssmscert.FPageSize = 100
occssmscert.FRectUserID = userid
occssmscert.FRectUserCell = usercell
occssmscert.FRectUserMail = usermail

if (userid<>"") or (usercell<>"") or (usermail<>"") then
    occssmscert.GetCSSMSCertLogList
end if

%>
<script language='javascript'>

function jsReSendSMSCert(idx) {
	if (confirm("인증번호 재전송 하시겠습니까?") == false) {
		return;
	}

	var frm = document.frmAct;
	frm.idx.value = idx;
	frm.submit();
}

</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			고객 아이디 : <input type="text" class="text" name="userid" value="<%= userid %>" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
			&nbsp;
			핸드폰 : <input type="text" class="text" name="usercell" value="<%= usercell %>" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
          	<input type="button" class="button_s" value="검색" onclick="document.frm.submit()">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			이메일 : <input type="text" class="text" name="usermail" value="<%= usermail %>" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
		</td>
	</tr>
	</form>
</table>

<p>

* <font color="red">직원 SMS인증</font>은 검색되지 않습니다.

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
        <td width="50">IDX</td>
		<td>아이디</td>
		<td>휴대폰</td>
		<td>이메일</td>
		<td width="50">인증번호</td>
		<td width="150">전송일시</td>
		<td width="150">확인일시</td>
        <td>비고</td>
    </tr>
<% for i = 0 to (occssmscert.FResultCount - 1) %>
    <tr align="center" bgcolor="#FFFFFF" height="25">
        <td><%= occssmscert.FItemList(i).Fidx %></td>
		<td><%= occssmscert.FItemList(i).Fuserid %></td>
		<td><%= occssmscert.FItemList(i).Fusercell %></td>
		<td><%= occssmscert.FItemList(i).Fusermail %></td>
		<td><%= occssmscert.FItemList(i).FsmsCD %></td>
		<td><%= occssmscert.FItemList(i).Fregdate %></td>
		<td><%= occssmscert.FItemList(i).FconfDate %></td>
		<td>
			<% if IsNull(occssmscert.FItemList(i).FconfDate) and (i = 0) and (occssmscert.FItemList(i).FconfDiv = "S") then %>
			<input type="button" class="button_s" value="재전송" onclick="jsReSendSMSCert(<%= occssmscert.FItemList(i).Fidx %>);">
			<% end if %>
		</td>
    </tr>
<% next %>
<% if (occssmscert.FResultCount < 1) then %>
    <tr bgcolor="#FFFFFF" align="center">
        <td height="25" colspan="11">검색결과가 없습니다.</td>
    </tr>
<% end if %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
		</td>
	</tr>
</table>

<form name="frmAct" method="post" action="pop_cs_smscert_log_process.asp">
	<input type="hidden" name="mode" value="resendcert">
	<input type="hidden" name="idx" value="">
</form>

<%

set occssmscert = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
