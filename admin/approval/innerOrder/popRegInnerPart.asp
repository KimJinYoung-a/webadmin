<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %>
<%
'###########################################################
' Description : 결제요청서 등록
' History : 2011.03.14 정윤정  생성
' 0 요청/1 진행중/ 5 반려/7 승인/ 9 완료
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/innerPartcls.asp"-->
<%

dim idx, mode

idx = requestCheckvar(Request("idx"),32)

if (idx = "") then
	idx = 0
	mode = "regnewpart"
else
	mode = "modifypart"
end if

'==============================================================================
dim oinnerpart
set oinnerpart = New CInnerPart

oinnerpart.FCurrPage = 1
oinnerpart.FPageSize = 1

oinnerpart.FRectIdx = idx

oinnerpart.GetInnerPartOne

if (mode = "modifypart") and (oinnerpart.FOneItem.Fidx = "") then
	response.write "잘못된 접속입니다."
	response.end
end if

%>
<script language="javascript">

function jsReg(frm) {
	if (frm.divcd.value == "") {
		alert("내부부서 구분을 지정하세요.");
		return;
	}

	if (frm.BIZSECTION_CD.value == "") {
		alert("ERP부서코드를 지정하세요.");
		return;
	}

	if (frm.scmid.value == "") {
		alert("어드민부서코드를 지정하세요.");
		return;
	}



	if (confirm("내부부서를 등록 하시겠습니까?") == true) {
		frm.submit();
	}
}

function jsDel(frm) {
	if (confirm("정말로 삭제하시겠습니까?") == true) {
		frm.mode.value = "delpart";
		frm.submit();
	}
}

function jsRegInsertUpcheShopWitak(frm) {
	if (confirm("일괄생성 하시겠습니까?\n\n생성에 시간이 소요됩니다.(5~10초)") == true) {
		frm.mode.value = "reginsertupcheshopwitak";
		frm.submit();
	}
}

function jsRegInsertPartToOnline(frm) {
	if (confirm("일괄생성 하시겠습니까?") == true) {
		frm.mode.value = "reginsertparttoonline";
		frm.submit();
	}
}

function jsRegInsertPartToOffline(frm) {
	if (confirm("일괄생성 하시겠습니까?") == true) {
		frm.mode.value = "reginsertparttooffline";
		frm.submit();
	}
}

</script>
<table width="100%" cellpadding="5" cellspacing="1" class="a"  style="padding-bottom:50px;" >
<tr>
	<td>
		<table width="100%" align="left" cellpadding="1" cellspacing="1" class="a"   border="0" >
		<form name="frm" method="post" action="innerpart_process.asp">
		<input type="hidden" name="mode" value="<%= mode %>">
		<tr>
			<td>
				<table width="100%" cellpadding="1" cellspacing="1" class="a">
				<tr>
					<td height=30><b>내부부서 등록</b></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="30" width="100">
						IDX
					</td>
					<input type="hidden" name="idx" value="<%= oinnerpart.FOneItem.Fidx %>">
					<td bgcolor="#FFFFFF">
						<%= oinnerpart.FOneItem.Fidx %>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="30" width="100">
						구분
					</td>
					<td bgcolor="#FFFFFF">
						<select name="divcd">
						<option value="">--선택--</option>
						<option value="S" <% if (oinnerpart.FOneItem.Fdivcd = "S") then %>selected<% end if %>>매장</option>
						<option value="M" <% if (oinnerpart.FOneItem.Fdivcd = "M") then %>selected<% end if %>>매입부서</option>
						</select>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="30" width="100">
						ERP부서명
					</td>
					<td bgcolor="#FFFFFF">
						<%= oinnerpart.FOneItem.FBIZSECTION_NM %>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="30" width="100">
						ERP부서코드
					</td>
					<td bgcolor="#FFFFFF">
						<input type="text" class="text" name="BIZSECTION_CD" value="<%= oinnerpart.FOneItem.FBIZSECTION_CD %>">
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="30" width="100">
						어드민부서코드
					</td>
					<td bgcolor="#FFFFFF">
						<input type="text" class="text" name="scmid" value="<%= oinnerpart.FOneItem.Fscmid %>">
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFFFFF" colspan="2" align=center height="40">

						<% if (mode = "regnewpart") then %>
						<input type="button" class="button" value="등록" onClick="jsReg(frm)">
						<% else %>
						<!--
						<input type="button" class="button" value="수정">
						&nbsp;
						-->
						<input type="button" class="button" value="삭제" onClick="jsDel(frm)">
						<% end if %>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
