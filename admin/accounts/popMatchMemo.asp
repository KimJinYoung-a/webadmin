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
<!-- #include virtual="/lib/classes/jungsan/ipkumlistcls.asp"-->
<%

dim inoutidx

dim i, j

inoutidx = requestCheckvar(Request("inoutidx"),32)

if (inoutidx = "") then
	inoutidx = -1
end if

dim ipkum
set ipkum = new IpkumChecklist
	ipkum.FCurrpage=1
	ipkum.FPagesize=1
	ipkum.FScrollCount = 10
	ipkum.FRectShowDismatch = "Y"

	ipkum.FRectInOutIDX = inoutidx

	ipkum.GetipkumlistAccounts

if ipkum.FResultCount = 0 then
	response.write "잘못된 접근입니다."
	response.end
end if

dim IsMemoInserted : IsMemoInserted = Not IsNull(ipkum.Fipkumitem(0).Fmatchmemo)

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">

function jsModifyInnerOrderPercentage(frm) {
	if (frm.innerorderpercentage.value == "") {
		alert("분배비율을 입력하세요.");
		return;
	}

	if (frm.innerorderpercentage.value*0 != 0) {
		alert("분배비율은 숫자만 가능합니다.");
		return;
	}

	if (confirm("수정하시겠습니까?") == true) {
		frm.mode.value = "modifyinnerorderpercentage";
		frm.submit();
	}
}

function jsModifyInnerOrderOne(frm) {
	if (confirm("과세/면세 내역 모두 재작성됩니다.\n\n재작성하시겠습니까?") == true) {
		frm.mode.value = "updateOneDetail";
		frm.submit();
	}
}

function jsSelectChanged(obj) {
	if (obj.value == "직접입력") {
		$("tr#idmemodetail").show();
	} else {
		$("tr#idmemodetail").hide();
	}
}

function jsUpdateMatchMemo() {
	var frm = document.frm;

	if (frm.matchMemoTMP.value == "") {
		alert("매칭메모를 선택하세요.");
		return;
	}

	if ((frm.matchMemoTMP.value == "직접입력") && (frm.matchMemo.value == "")) {
		alert("매칭메모를 입력하세요.");
		return;
	}

	if (frm.matchMemo.value.length > 100) {
		alert("메모는 100글자까지 가능합니다.");
		return;
	}

	if (confirm("저장하시겠습니까?") != true) {
		return;
	}

	<% if IsMemoInserted then %>
		frm.mode.value = "modMatchMemo";
	<% else %>
		frm.mode.value = "insMatchMemo";
	<% end if %>

	if (frm.matchMemoTMP.value != "직접입력") {
		frm.matchMemo.value = frm.matchMemoTMP.value;
	}

	frm.submit();
}

function jsDelMatchMemo() {
	var frm = document.frm;

	if (confirm("매칭을 삭제하시겠습니까?") != true) {
		return;
	}

	frm.mode.value = "delMatchMemo";
	frm.submit();
}




</script>
<table width="100%" cellpadding="5" cellspacing="1" class="a"  style="padding-bottom:50px;" >
<form name="frm" method="post" action="matchMemo_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="inoutidx" value="<%= inoutidx %>">
<tr>
	<td>
		<table width="100%" align="left" cellpadding="1" cellspacing="1" class="a"   border="0" >
		<tr>
			<td>
				<table width="100%" cellpadding="1" cellspacing="1" class="a">
				<tr>
					<td height=30 colspan="2"><b>메모등록</b></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="1" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="25" width="15%" align=center>
						은행명
					</td>
					<td bgcolor="#FFFFFF" align="center" width="35%">
						<%= ipkum.Fipkumitem(0).Fbkname %>
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="15%" align=center>
						계좌번호
					</td>
					<td bgcolor="#FFFFFF" align="center">
						<%= ipkum.Fipkumitem(0).Fbkacctno %>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="25" align=center>
						입출금일
					</td>
					<td bgcolor="#FFFFFF" align="center">
						<%= mid(ipkum.Fipkumitem(0).Fbkdate,1,4) %>-<%= mid(ipkum.Fipkumitem(0).Fbkdate,5,2) %>-<%= mid(ipkum.Fipkumitem(0).Fbkdate,7,2) %>
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						적요
					</td>
					<td bgcolor="#FFFFFF" align="center">
						<%= ipkum.Fipkumitem(0).Fbkjukyo %>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="25" align=center>
						입금금액
					</td>
					<td align="center" bgcolor="#FFFFFF">
						<% if ipkum.Fipkumitem(0).finout_gubun = "2" then %>
							<b><%= FormatNumber(ipkum.Fipkumitem(0).Fbkinput,0) %></b>
						<% end if %>
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						출금금액
					</td>
					<td align="center" bgcolor="#FFFFFF">
						<% if ipkum.Fipkumitem(0).finout_gubun = "1" then %>
							<b><%= FormatNumber(ipkum.Fipkumitem(0).Fbkinput,0) %></b>
						<% end if %>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="25" align=center>
						매칭상태
					</td>
					<td align="left" bgcolor="#FFFFFF" colspan="3">
						&nbsp;
						<% if Not IsNull(ipkum.Fipkumitem(i).Fmatchstate) and (ipkum.Fipkumitem(i).Fmatchstate <> "X") then %>
							입력불가
						<% else %>
						<input type="radio" name="matchstate" value="X" <% if (ipkum.Fipkumitem(i).Fmatchstate = "X") then %>checked<% end if %> > 매칭제외
						<input type="radio" name="matchstate" value="D" <% if IsNull(ipkum.Fipkumitem(i).Fmatchstate) then %>disabled<% end if %> > 매칭제외 취소
						<input type="radio" name="matchstate" value="" <% if IsNull(ipkum.Fipkumitem(i).Fmatchstate) then %>checked<% end if %> > 입력않함
						<% end if %>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="25" align=center>
						메모
					</td>
					<td align="left" bgcolor="#FFFFFF" colspan="3">
						&nbsp;
						<select class="select" name="matchMemoTMP" onChange="jsSelectChanged(this);">
						<option value=""></option>
						<option value="카드사 입금(고객결제)" <% if (ipkum.Fipkumitem(i).Fmatchmemo = "카드사 입금(고객결제)") then %>selected<% end if %> >카드사 입금(고객결제)</option>
						<option value="PG사 입금(고객결제)" <% if (ipkum.Fipkumitem(i).Fmatchmemo = "PG사 입금(고객결제)") then %>selected<% end if %> >PG사 입금(고객결제)</option>
						<option value="직접입력" <% if IsMemoInserted and (InStr("카드사 입금(고객결제),PG사 입금(고객결제)", ipkum.Fipkumitem(i).Fmatchmemo) = 0) then %>selected<% end if %> >직접입력</option>
						</select>
					</td>
				</tr>
				<tr id="idmemodetail" style="display:<% if IsMemoInserted and (InStr("카드사 입금(고객결제),PG사 입금(고객결제)", ipkum.Fipkumitem(i).Fmatchmemo) = 0) then %>inline<% else %>none<% end if %>">
					<td bgcolor="<%= adminColor("tabletop") %>" height="25" align=center>
						메모상세
					</td>
					<td align="left" bgcolor="#FFFFFF" colspan="3">
						&nbsp;
						<textarea class="textarea" name="matchMemo" cols="50" rows="4"><%= ipkum.Fipkumitem(i).Fmatchmemo %></textarea>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" cellpadding="1" cellspacing="1" class="a">
				<tr>
					<td height=30 colspan="2" align="center">
						<input type="button" class="button" value="메모<% if IsMemoInserted then %>수정<% else %>등록<% end if %>" onClick="jsUpdateMatchMemo();">
						<% if IsMemoInserted then %>
						&nbsp;
						<input type="button" class="button" value="메모[삭제]" onClick="jsDelMatchMemo();">
						<% end if %>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
</body>
</html>
<%
set ipkum = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
