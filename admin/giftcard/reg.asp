<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/giftcard/giftcard_cls.asp"-->
<%
Dim i, idx, mode, oGiftCard
idx = request("idx")
If idx ="" Then
	mode = "I"
Else
	mode = "U"
End If

Set oGiftCard = new cGiftCard
	oGiftCard.FRectIdx = idx
	oGiftCard.getGiftCardOneItem
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function targetOpt(v){
	if (v == "0000"){
		$("#sugiTarget").show();
		$("#sugiPrice").val("");
	}else{
		$("#sugiTarget").hide();
		$("#sugiPrice").val("");
	}
}
function regGift(){
	<% 'if not(C_ADMIN_AUTH or C_PSMngPart) then %>
	<% if not(C_ADMIN_AUTH) then %>
	if ($("#eappidx").val()== ""){
		alert('품의서IDX를 입력하세요');
		$("#eappidx").focus();
		return;
	}
	<% end if %>

	if ($("#reqTitle").val()== ""){
		alert('제목을 입력하세요');
		$("#reqTitle").focus();
		return;
	}

	if ($("#reqContent").val()== ""){
		alert('내용을 입력하세요');
		$("#reqContent").focus();
		return;
	}


	if ($("#userid").val()== ""){
		alert('텐바이텐ID를 입력하세요');
		$("#userid").focus();
		return;
	}

	if ($("#makeCnt").val()== ""){
		alert('발급할 카드 수량을 입력하세요');
		$("#makeCnt").focus();
		return;
	}

	if ($("#opt").val()== ""){
		alert('옵션을 선택하세요');
		$("#opt").focus();
		return;
	}

	if ($("#MMSTitle").val()== ""){
		alert('MMS 제목을 입력하세요');
		$("#MMSTitle").focus();
		return;
	}

	if ($("#MMSContent").val()== ""){
		alert('MMS 내용을 입력하세요');
		$("#MMSContent").focus();
		return;
	}

	if (confirm("저장 하시겠습니까?")){
		var frm = document.frm;
		frm.action = "/admin/giftcard/giftcardProc.asp";
		frm.submit();
	}
}
function pop_checkId(){
	if ($("#userid").val()== ""){
		alert('텐바이텐ID를 입력해야 확인 가능합니다.');
		$("#userid").focus();
		return;
	} else {
		var str = $("#userid").val();
		str = str.replace(/(?:\r\n|\r|\n)/g, ',');

		var popwin = window.open("/admin/giftcard/pop_checkId.asp?userid="+str,"popcheckId","width=1200,height=600,scrollbars=yes,resizable=yes");
		popwin.focus();
	}
}
function pop_eappView(){
	if ($("#eappidx").val()== ""){
		alert('품의서IDX를 입력해야 확인 가능합니다.');
		$("#eappidx").focus();
		return;
	} else {
		var iridx = $("#eappidx").val();
		var popwin = window.open("/admin/approval/eapp/vieweapp.asp?iridx="+iridx,"popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
		popwin.focus();
	}
}
</script>
<form name="frm" method="post">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="idx" value="<%= idx %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="15%">품의서IDX</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" id="eappidx" name="eappidx" class="text" size="10" value="<%= oGiftCard.FOneItem.FEappIdx %>">
		&nbsp;<input type="button" class="button" value="품의서보기" onclick="pop_eappView();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="15%">제목</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" id="reqTitle" name="reqTitle" class="text" size="100" value="<%= oGiftCard.FOneItem.FReqTitle %>">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="15%">내용</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<textarea class="textarea" id="reqContent" name="reqContent" cols="150" rows="20"><%= oGiftCard.FOneItem.FReqContent %></textarea>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="15%">텐바이텐ID</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<textarea name="userid" id="userid" class="textarea" cols="32" rows="5"><%= Chkiif(mode="U", getUserids(idx), "") %></textarea>
		<strong>(※ 반드시 엔터로 구분해서 넣어주세요)</strong>
		<br />
		<input type="button" class="button" value="ID체크" onclick="pop_checkId();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="15%">발급할 카드 수량</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" id="makeCnt" name="makeCnt" class="text" size="10" value="<%= oGiftCard.FOneItem.FMakeCnt %>">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="15%">옵션</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<select id="opt" name="opt" class="select" onchange="targetOpt(this.value);">
			<option value="">-Choice-</option>
			<option value="0001" <%= Chkiif(oGiftCard.FOneItem.FOpt="0001", "selected", "") %> >1만원권</option>
			<option value="0002" <%= Chkiif(oGiftCard.FOneItem.FOpt="0002", "selected", "") %>>2만원권</option>
			<option value="0003" <%= Chkiif(oGiftCard.FOneItem.FOpt="0003", "selected", "") %>>3만원권</option>
			<option value="0004" <%= Chkiif(oGiftCard.FOneItem.FOpt="0004", "selected", "") %>>5만원권</option>
			<option value="0005" <%= Chkiif(oGiftCard.FOneItem.FOpt="0005", "selected", "") %>>8만원권</option>
			<option value="0006" <%= Chkiif(oGiftCard.FOneItem.FOpt="0006", "selected", "") %>>10만원권</option>
			<option value="0007" <%= Chkiif(oGiftCard.FOneItem.FOpt="0007", "selected", "") %>>15만원권</option>
			<option value="0008" <%= Chkiif(oGiftCard.FOneItem.FOpt="0008", "selected", "") %>>20만원권</option>
			<option value="0009" <%= Chkiif(oGiftCard.FOneItem.FOpt="0009", "selected", "") %>>30만원권</option>
			<option value="0000" <%= Chkiif(oGiftCard.FOneItem.FOpt="0000", "selected", "") %>>수기입력</option>
		</select>
		<span id="sugiTarget" <%= Chkiif(oGiftCard.FOneItem.FSugiPrice = "" , "style='display:none;'", "") %>  >
			<input type="text" id="sugiPrice" name="sugiPrice" value="<%= oGiftCard.FOneItem.FSugiPrice %>">원
		</span>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="15%">MMS 제목</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" class="text" size="100" id="MMSTitle" name="MMSTitle" value="<%= oGiftCard.FOneItem.FMMSTitle %>">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="15%">MMS 내용</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<textarea class="textarea" id="MMSContent" name="MMSContent" cols="150" rows="20"><%= oGiftCard.FOneItem.FMMSContent %></textarea>
	</td>
</tr>
<tr align="center" height="25" bgcolor="FFFFFF">
	<td colspan="3">
		<% If oGiftCard.FOneItem.FIsSend <> "Y" Then %>
			<input type="button" class="button" value="저장" onClick="regGift();">
		<% End If %>
		<input type="button" class="button" value="리스트로" onClick="location.href='/admin/giftcard/list.asp?menupos=<%= menupos %>';">
	</td>
</tr>
</table>
</form>
<% Set oGiftCard = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->