<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 고객블랙리스트
' Hieditor : 2014.03.06 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/member/tenbyten/invalid/invalid_user_cls.asp"-->

<%
dim adminuserid, i, menupos, oinvalid
dim idx, gubun, invaliduserid, isusing, regdate, lastupdate, reguserid, lastuserid, comment
	idx = requestcheckvar(request("idx"),10)
	menupos = requestcheckvar(request("menupos"),10)

adminuserid=session("ssBctId")

set oinvalid = new cinvalid_list
	oinvalid.frectidx = idx
	
	if idx <> "" then
		oinvalid.getinvalid_oneitem()
		
		if oinvalid.ftotalcount > 0 then
			idx = oinvalid.foneitem.fidx
			gubun = oinvalid.foneitem.fgubun
			invaliduserid = oinvalid.foneitem.finvaliduserid
			isusing = oinvalid.foneitem.fisusing
			regdate = oinvalid.foneitem.fregdate
			lastupdate = oinvalid.foneitem.flastupdate
			reguserid = oinvalid.foneitem.freguserid
			lastuserid = oinvalid.foneitem.flastuserid
			comment = oinvalid.foneitem.fcomment
		end if
	end if

if gubun = "" then gubun = "ONEVT"
if isusing = "" then isusing = "Y"
%>

<script type="text/javascript">

function reg_invalid(){
	if (frm.gubun.value==""){
		alert('특별고객구분을 선택해주세요.');
		frm.gubun.focus();
		return;
	}
	if (frm.invaliduserid.value==""){
		alert('아이디를 입력해 주세요.');
		frm.invaliduserid.focus();
		return;
	}
	if (frm.isusing.value==""){
		alert('사용구분을 선택해주세요.');
		frm.isusing.focus();
		return;
	}
	
	frm.action="/admin/member/tenbyten/invalid/invalid_user_process.asp";
	frm.mode.value="edit";
	frm.submit();
}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">		
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<form name="frm" method="post" style="margin:0px;">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<input type="hidden" name="mode">

<tr bgcolor="#FFFFFF">
	<td align="center"><b>IDX</b><br></td>
	<td>
		<%=idx%>
		<input type="hidden" name="idx" value="<%=idx%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>특별고객구분</b><br></td>
	<td>
		<% Drawinvalidgubun "gubun", gubun, "" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>아이디</b><br></td>
	<td>
		<input type="text" name="invaliduserid" value="<%= invaliduserid %>" size=32 maxlength=32>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>코맨트</b><br></td>
	<td>
		<textarea name="comment" cols=80 rows=5><%= comment %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>사용여부</b><br></td>
	<td>
		<% drawSelectBoxisusingYN "isusing", isusing, "" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>최근수정</b><br></td>
	<td>
		<% if lastupdate<>"" then %>
			<%= lastupdate %>
		<% end if %>

		<% if lastuserid<>"" then %>
			<Br>(<%= lastuserid %>)
		<% end if %>
	</td>
</tr>	
<tr bgcolor="#FFFFFF">
	<td align="center" colspan="2">
		<input type="button" value="저장" onclick="reg_invalid();" class="button">
	</td>
</tr>
</table>
</form>

<%
set oinvalid = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->
