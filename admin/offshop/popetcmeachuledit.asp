<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  기타매출관리
' History : 2009.04.07 서동석 생성
'			2010.05.13 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/etcmeachulcls.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<%
dim idx
dim ofranchulgojungsan, shopid

idx = RequestCheckvar(request("idx"),10)

if idx="" then idx="0"


'// ===========================================================================
set ofranchulgojungsan = new CEtcMeachul
ofranchulgojungsan.FRectidx = idx
ofranchulgojungsan.getOneEtcMeachul

dim IsMeaipPriceEditPossible	: IsMeaipPriceEditPossible = True
if (idx <> "0") and (ofranchulgojungsan.FOneItem.Fdivcode <> "GC") and (ofranchulgojungsan.FOneItem.Fdivcode <> "ET") then
	IsMeaipPriceEditPossible = False
end if

'// ===========================================================================
'수익부서목록
Dim clsBS, arrBizList
Set clsBS = new CBizSection
	clsBS.FUSE_YN = "Y"
	clsBS.FOnlySub = "Y"
	clsBS.FSale = "N"
	arrBizList = clsBS.fnGetBizSectionList
Set clsBS = nothing
'// ===========================================================================
Dim defaultYYYY, defaultMM, defaultShopDiv
Dim i

IF idx="0" THen
    defaultYYYY = Left(DateAdd("m",-1,now()),4)
    defaultMM   = Mid(DateAdd("m",-1,now()),6,2)
    defaultShopDiv = ""
ELSE
    defaultYYYY = ""
    defaultMM   = ""
    defaultShopDiv = ""
END IF
%>
<script type='text/javascript'>

function SaveInfo(frm){
	if (frm.title.value.length<1){
		alert('Title을 입력하세요');
		frm.title.focus();
		return;
	}

	if (frm.shopdiv.value.length<1){
		alert('구분을 입력하세요');
		frm.shopdiv.focus();
		return;
	}

	if (frm.diffKey.value.length<1){
	    alert('발행 차수를 입력하세요');
		frm.diffKey.focus();
		return;
	}

	if (frm.shopdiv.value == "7") {
		if ((frm.papertype.value != "200") && (frm.papertype.value != "102")) {
			alert("구분이 수출(해외)인경우 \n\n수출신고필증 또는 영세계산서만 증빙서류로 등록할 수 있습니다.");
			frm.papertype.focus();
			return;
		}
        // 20036 => 4010005
		if (frm.selltype.value != "4010005") {
			alert("구분이 수출(해외)인경우 \n\n계정과목은 영세만 가능합니다.");
			frm.selltype.focus();
			return;
		}
	} else if (frm.shopdiv.value == "9") {
	    //if (frm.idx.value!=9861){
    		if (frm.papertype.value != "102") {
    			alert("구분이 영세인경우 \n\n영세계산서만 증빙서류로 등록할 수 있습니다.");
    			frm.papertype.focus();
    			return;
    		}
    	//}

		if (frm.selltype.value != "4010005") {
			alert("구분이 영세인경우 \n\n계정과목은 영세만 가능합니다.");
			frm.selltype.focus();
			return;
		}
	} else {
		if ((frm.papertype.value == "200") || (frm.papertype.value == "102")) {
			alert("출고처구분이 수출 또는 영세인경우만 등록 가능합니다.");
			frm.papertype.focus();
			return;
		}

		if (frm.selltype.value == "4010005") {
			alert("출고처구분이 수출 또는 영세인경우만 등록 가능합니다..");
			frm.selltype.focus();
			return;
		}
	}

<% if idx="0" then %>
	if (frm.shopid.value.length<1){
		alert('매출처를 입력하세요');
		frm.shopid.focus();
		return;
	}

	if (frm.totalbuycash.value.length<1){
		alert('총 매입가를 입력하세요');
		frm.totalbuycash.focus();
		return;
	}

	if (frm.totalsuplycash.value.length<1){
		alert('총 공급가를 입력하세요');
		frm.totalsuplycash.focus();
		return;
	}
<% elseif (IsMeaipPriceEditPossible) then %>
	if (frm.totalbuycash.value.length<1){
		alert('총 매입가를 입력하세요');
		frm.totalbuycash.focus();
		return;
	}
<% end if %>

	if (frm.totalsum.value.length > 0) {
		frm.totalsum.value = replaceAll(frm.totalsum.value, ",", "");
	}

/*
	if (frm.totalsum.value.length<1){
		alert('총 발행금액을 입력하세요');
		frm.totalsum.focus();
		return;
	}


	if ((!frm.statecd[0].checked)&&(!frm.statecd[1].checked)&&(!frm.statecd[2].checked)&&(!frm.statecd[3].checked)){
		alert('상태를 선택하세요.');
		frm.statecd[0].focus();
		return;
	}
*/

	if (confirm('저장 하시겠습니까?')){
		frm.submit();
	}
}

function escapeRegExp(string) {
    return string.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1");
}

// frm.totalsum.value = replaceAll(frm.totalsum.value, ",", "");
function replaceAll(string, find, replace) {
  return string.replace(new RegExp(escapeRegExp(find), 'g'), replace);
}

function changeState(state)
{
	var f = document.frm;

	switch (state)
	{
	case "0":
		var msg = "수정중으로 변경하시겠습니까?";
		break;
	case "1":
		var msg = "업체확인중으로 변경하시겠습니까?";
		break;
	case "3":
		var msg = "업체확인완료로 변경하시겠습니까?";
		break;
	case "7":
		var msg = "입금완료로 변경하시겠습니까?";
		if (f.ipkumdate.value.length!=10)
		{
			alert("입금일을 입력하십시오.");
			return;
		}

		if (f.taxdate.value.length!=10)
		{
			alert("매출기준일이 없습니다.\n\n증빙서류 작성 후 매출기준일을 입력하십시오.");
			return;
		}

		break;
	}

	if (confirm(msg))
	{
		f.mode.value = "changeState";
		f.stateCd.value = state;
		f.submit();
	}
}

function changeIssueState(state)
{
	var f = document.frm;
	var msg = "";

	switch (state)
	{
	case "0":
		msg = "발행신청으로 변경하시겠습니까?";
		break;
	case "9":
		msg = "발행완료로 변경하시겠습니까?";
		if (f.taxdate.value.length != 10) {
			alert("매출기준일이 없습니다.\n\n증빙서류 작성 후 매출기준일을 입력하십시오.");
			return;
		}
		break;
	case "NULL":
		msg = "증빙서류 정보를 삭제하시겠습니까?";
		break;
	}

	if ((f.paperissuetype.value == "1") && (state == "NULL")) {
		// 발행된 계산서 정보 삭제
		msg = "발행된 계산서가 국세청에 전송된 경우\n수정세금계산서를 추가로 발행해야 하고\n\n전송되지 않은 경우\nBILL36524 에서 발행된 계산서를 취소해야 합니다.\n\n" + msg;
	}

	if (msg == "") {
		alert("ERROR");
		return;
	}

	if (confirm(msg) == true) {
		f.mode.value = "changeIssueState";
		f.issueStateCd.value = state;
		f.submit();
	}
}

function changeIpkumState(state)
{
	var f = document.frm;
	var msg = "";

	switch (state)
	{
	case "0":
		msg = "입금이전으로 변경하시겠습니까?";
		break;
	case "5":
		msg = "일부입금으로 변경하시겠습니까?";
		break;
	case "9":
		msg = "입금완료로 변경하시겠습니까?";
		if (f.ipkumdate.value.length != 10) {
			alert("먼저 입금일을 입력하세요");
			return;
		}
		break;
	case "NULL":
		msg = "입금상태 정보를 삭제하시겠습니까?";
		break;
	}

	if (msg == "") {
		alert("ERROR");
		return;
	}

	if (confirm(msg) == true) {
		f.mode.value = "changeIpkumState";
		f.ipkumStateCd.value = state;
		f.submit();
	}
}

function jsGetTax(ibizNo, itotSum){
	var sSearchText = ibizNo;
	var itotSum = itotSum;

	if (sSearchText == "2118700620") {
		sSearchText = "";
	}

	var winTax = window.open("/admin/tax/popSetEseroTax.asp?sST="+sSearchText+"&totSum="+itotSum+"&tgType=NRM&iTST=1","popGetTaxInfo","width=1200, height=800, resizable=yes, scrollbars=yes");
	winTax.focus();
}

function fillTaxInfo(eTax,iDK,iVK,dID,sInm,mTP,mSP,mVP){
    var frm = document.frm;

	frm.eserotaxkey.value = eTax;
}

</script>

<form name="frm" method=post action="/admin/offshop/etc_meachul_process.asp" style="margin:0px;">
<input type=hidden name="idx" value="<%= ofranchulgojungsan.FOneItem.Fidx %>">
<% if idx="0" then %>
<input type=hidden name="mode" value="addmaster">
<% else %>
<input type=hidden name="mode" value="modimaster">
<input type="hidden" name="stateCd" value="">
<input type="hidden" name="issueStateCd" value="">
<input type="hidden" name="ipkumStateCd" value="">
<% end if %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>IDX</td>
		<td bgcolor="#FFFFFF" ><%= ofranchulgojungsan.FOneItem.Fidx %></td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">매출처</td>
		<% if idx="0" then %>
		<td bgcolor="#FFFFFF" >
			<% NewdrawSelectBoxShopAll "shopid", shopid %>
		</td>
		<% else %>
		<td bgcolor="#FFFFFF" ><%= ofranchulgojungsan.FOneItem.Fshopid %></td>
		<% end if %>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">정산대상월</td>
		<% if idx="0" then %>
			<td bgcolor="#FFFFFF" ><% call DrawYMBox(defaultYYYY,defaultMM) %></td>
		<% else %>
			<td bgcolor="#FFFFFF" >
				<% if C_ADMIN_AUTH or C_MngPart or C_PSMngPart then %>
					<% call DrawYMBox(Left(NULL2Blank(ofranchulgojungsan.FOneItem.FYYYYMM),4),Right(NULL2Blank(ofranchulgojungsan.FOneItem.FYYYYMM),2)) %>
					※ 관리자,재무팀만 수정가능
				<% else %>
					<%= Left(NULL2Blank(ofranchulgojungsan.FOneItem.FYYYYMM),4) %>-<%= Right(NULL2Blank(ofranchulgojungsan.FOneItem.FYYYYMM),2) %>
					<input type="hidden" name="yyyy1" value="<%= Left(NULL2Blank(ofranchulgojungsan.FOneItem.FYYYYMM),4) %>">
					<input type="hidden" name="mm1" value="<%= Right(NULL2Blank(ofranchulgojungsan.FOneItem.FYYYYMM),2) %>">
				<% end if %>
			</td>
		<% end if %>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">구분</td>
		<td bgcolor="#FFFFFF" >
			<% if idx="0" then %>
				<% Call DrawShopDivBox(defaultShopDiv) %>
				/
				<select class="select" name="divcode">
					<option value="GC">가맹비
					<option value="ET">기타매출
				</select>
			<% else %>
				<% Call DrawShopDivBox(ofranchulgojungsan.FOneItem.FShopDiv) %>
				/
				<font color="<%= ofranchulgojungsan.FOneItem.GetDivCodeColor %>"><%= ofranchulgojungsan.FOneItem.GetDivCodeName %></font>
			<% end if %>
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">차수</td>
	    <td bgcolor="#FFFFFF" >
	    <% if idx="0" then %>
	    <input type="text" name="diffKey" maxlength="2" class="text">
	    <% else %>
	    <input type="text" name="diffKey" value="<%= ofranchulgojungsan.FOneItem.FdiffKey %>" size="2" maxlength="2" class="text">
	    <% end if %>
	    </td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">Title</td>
		<td bgcolor="#FFFFFF" >
			<input type="text" class="text" name=title value="<%= ofranchulgojungsan.FOneItem.Ftitle %>" size="40" maxlength="40" <%If ofranchulgojungsan.FOneItem.Fstatecd>="4" Then %>readOnly<%End If %> >
			(ex) OO점 4월 1차 상품대
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">총소비자가</td>
		<td bgcolor="#FFFFFF">
			<% if idx="0" then %>

			<% else %>
			<%= FormatNumber(ofranchulgojungsan.FOneItem.Ftotalsellcash,0) %>
			<% end if %>
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>"><b>총출고가</b></td>
		<td bgcolor="#FFFFFF">
			<% if idx="0" then %>
			<input type=text name=totalsuplycash value="" size=10 maxlength=10 style="border:1px #999999 solid; text-align=right">
			<font color="#AAAAAA">(매출처로 공급한 상품가격)</font>
			<% else %>
			<%= FormatNumber(ofranchulgojungsan.FOneItem.Ftotalsuplycash,0) %>
			<font color="#AAAAAA">(매출처로 공급한 상품가격)</font>
			<% end if %>
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">총매입가</td>
		<td bgcolor="#FFFFFF">
			<% if idx="0" then %>
				<input type=text name=totalbuycash value="" size=10 maxlength=10 style="border:1px #999999 solid; text-align=right">
				<font color="#AAAAAA">(소요 비용:매입)</font>
			<% else %>
				<% if IsMeaipPriceEditPossible then %>
					<input type=text name=totalbuycash value="<%= ofranchulgojungsan.FOneItem.Ftotalbuycash %>" size=10 maxlength=10 style="border:1px #999999 solid; text-align=right">
				<% else %>
					<%= FormatNumber(ofranchulgojungsan.FOneItem.Ftotalbuycash,0) %>
					<input type="hidden" name="totalbuycash" value="<%= ofranchulgojungsan.FOneItem.Ftotalbuycash %>">
				<% end if %>
				<font color="#AAAAAA">(업체로부터 공급받은 상품가격)</font>
			<% end if %>
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">진행상태</td>
		<td bgcolor="#FFFFFF" >
		<font color="<%= ofranchulgojungsan.FOneItem.GetStateColor %>"><%= ofranchulgojungsan.FOneItem.GetStateName %></font>

		<% if (ofranchulgojungsan.FOneItem.Fstatecd="0") then %>
		==&gt; <input type="button" class="button" onclick="changeState('1');" value="업체확인중으로 변경">
		<% elseif (ofranchulgojungsan.FOneItem.Fstatecd="1") then %>
		==&gt; <input type="button" class="button" onclick="changeState('3');" value="업체확인완료로 변경">
		<% elseif (ofranchulgojungsan.FOneItem.Fstatecd="4") or (ofranchulgojungsan.FOneItem.Fstatecd="3") then %>
		==&gt; <input type="button" class="button" onclick="changeState('7');" value="완료 로 변경">
		<% else %>
		<% end if %>

		<% if (ofranchulgojungsan.FOneItem.Fstatecd="1") or (ofranchulgojungsan.FOneItem.Fstatecd="3") then %>
		<input type="button" class="button" onclick="changeState('0');" value="수정중으로 변경">
		<% elseif (ofranchulgojungsan.FOneItem.Fstatecd="4") then %>
		<input type="button" class="button" onclick="changeState('0');" value="수정중으로 변경">
		<% else %>

	    <% end if %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">기타사항</td>
		<td bgcolor="#FFFFFF" >
			<textarea name="etcstr" class="textarea" cols="86" rows="8"><%= ofranchulgojungsan.FOneItem.Fetcstr %></textarea>
		</td>
	</tr>

	<tr>
		<td height="10" bgcolor="#FFFFFF" style="padding: 1px;"></td>
		<td bgcolor="#FFFFFF" style="padding: 1px;"></td>
	</tr>

	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">매출부서</td>
		<td bgcolor="#FFFFFF">
	        <select class="select" name="bizsection_cd">
	        <option value="">--선택--</option>
	        <% For i = 0 To UBound(arrBizList,2)	%>
	    		<option value="<%=arrBizList(0,i)%>" <%IF (ofranchulgojungsan.FOneItem.Fbizsection_cd) = Cstr(arrBizList(0,i)) THEN%> selected <%END IF%>><%=arrBizList(1,i)%></option>
	    	<% Next %>
	        </select>
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">계정과목</td>
		<td bgcolor="#FFFFFF">
			<% drawPartnerCommCodeBox true,"sellacccd","selltype",ofranchulgojungsan.FOneItem.Fselltype,"" %>
		</td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">증빙서류</td>
		<td bgcolor="#FFFFFF">
			<select class="select" name="papertype">
				<option value="">선택
				<option value="100" <% if ofranchulgojungsan.FOneItem.Fpapertype="100" then response.write "selected" %> > 세금 계산서
				<option value="101" <% if ofranchulgojungsan.FOneItem.Fpapertype="101" then response.write "selected" %> > 면세 계산서
				<option value="102" <% if ofranchulgojungsan.FOneItem.Fpapertype="102" then response.write "selected" %> > 영세 계산서
				<option value="200" <% if ofranchulgojungsan.FOneItem.Fpapertype="200" then response.write "selected" %> > 수출신고필증
				<option value="999" <% if ofranchulgojungsan.FOneItem.Fpapertype="999" then response.write "selected" %> > 없음
			</select>

	        <select class="select" name="paperissuetype">
	        	<option value="">--선택--</option>
				<option value="1" <%IF (ofranchulgojungsan.FOneItem.Fpaperissuetype = "1") THEN%> selected <%END IF%>>정발행</option>
				<option value="2" <%IF (ofranchulgojungsan.FOneItem.Fpaperissuetype = "2") THEN%> selected <%END IF%>>역발행</option>
	        </select>
	        *역발행 = 매입자 발행
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">이세로</td>
		<td bgcolor="#FFFFFF">
			<% if ofranchulgojungsan.FOneItem.Fpaperissuetype = "2" then %>
				<input type="text" class="text_ro" name="eserotaxkey" value="<%= ofranchulgojungsan.FOneItem.Feserotaxkey %>" size="30" maxlength="32" readonly>
				<input type="button" class="button" value="검색" onClick="jsGetTax('<%= ofranchulgojungsan.FOneItem.FbizNo %>','<%= ofranchulgojungsan.FOneItem.Ftotalsum %>');">
		    <% else %>
		        <%= ofranchulgojungsan.FOneItem.Feserotaxkey %>
		        <% if IsNull(ofranchulgojungsan.FOneItem.Feserotaxkey) then %>매칭이전<% end if %>
			<% end if %>
		</td>
	</tr>

	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>"><b>총발행금액</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text_ro" name="totalsum" value="<%= ofranchulgojungsan.FOneItem.Ftotalsum %>" size="10" maxlength="10" style="text-align=right">
			<font color="#AAAAAA">(계산서 발행 금액)</font>
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">매출기준일</td>
		<td bgcolor="#FFFFFF">
   			<input type="text" id="termTaxDt" name="taxdate" readonly size="11" maxlength="10" value="<%= ofranchulgojungsan.FOneItem.Ftaxdate %>" class="text_ro" style="text-align:center;" />
			<%if (ofranchulgojungsan.FOneItem.Fstatecd > "0") or (ofranchulgojungsan.FOneItem.Fstatecd < "7") then %>
				<% if (Not IsNull(ofranchulgojungsan.FOneItem.Fpapertype)) and (ofranchulgojungsan.FOneItem.Fpapertype <> "100" and Not (ofranchulgojungsan.FOneItem.Fpapertype = "200" and IsNull(ofranchulgojungsan.FOneItem.Finvoiceidx))) then %>
				<img src="/images/calicon.gif" align="absmiddle" border="0" id="btnTaxDt" style="cursor:pointer;" />
				<script type="text/javascript">
					var CAL_TaxDate = new Calendar({
						inputField : "termTaxDt", trigger    : "btnTaxDt",
						bottomBar: true, dateFormat: "%Y-%m-%d",
						onSelect: function() {
							this.hide();
						}
					});
				</script>
				<% end if %>
			<% end if %>
			<font color="#AAAAAA">(계산서발행일,수출신고일자)</font>
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">진행상태</td>
		<td bgcolor="#FFFFFF" >
			<%If idx <> "0" Then %>

				<%= ofranchulgojungsan.FOneItem.GetIssueStateName %>(<%= ofranchulgojungsan.FOneItem.Fpaperissuetype %>)

				<% if (ofranchulgojungsan.FOneItem.Fpaperissuetype = "1") then %>
					<% if (C_ADMIN_AUTH) and (ofranchulgojungsan.FOneItem.FIssueStateCD="9") then %>
					<input type="button" class="button" onclick="changeIssueState('NULL');" value="증빙서류 삭제"> [관리자뷰]
					<% end if %>
				<% elseif (ofranchulgojungsan.FOneItem.Fpaperissuetype = "2") then %>

					<% if IsNull(ofranchulgojungsan.FOneItem.FIssueStateCD) then %>
						==&gt;
						<input type="button" class="button" onclick="changeIssueState('0');" value="발행신청으로 변경">
						<input type="button" class="button" onclick="changeIssueState('9');" value="발행완료로 변경">
					<% else %>
						==&gt;
						<% if (ofranchulgojungsan.FOneItem.FIssueStateCD="0") then %>
							<input type="button" class="button" onclick="changeIssueState('9');" value="발행완료로 변경">
						<% elseif (ofranchulgojungsan.FOneItem.FIssueStateCD="9") then %>

						<% else %>
							ERROR
						<% end if %>
						<input type="button" class="button" onclick="changeIssueState('NULL');" value="증빙서류 삭제" <% if (Not C_ADMIN_AUTH) then %>disabled<% end if %> > <% if (C_ADMIN_AUTH) then %>[관리자뷰]<% end if %>
					<% end if %>

				<% end if %>

	    	<% end if %>
		</td>
	</tr>
	<tr>
		<td height="10" bgcolor="#FFFFFF" style="padding: 1px;"></td>
		<td bgcolor="#FFFFFF" style="padding: 1px;"></td>
	</tr>

	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">입금일</td>
		<td bgcolor="#FFFFFF">
   			<input type="text" id="termIpkumDt" name="ipkumdate" readonly size="11" maxlength="10" value="<%= ofranchulgojungsan.FOneItem.Fipkumdate %>" class="text_ro" style="text-align:center;" />
			<%if (ofranchulgojungsan.FOneItem.Fstatecd > "0") or (ofranchulgojungsan.FOneItem.Fstatecd < "7") then %>
			<img src="/images/calicon.gif" align="absmiddle" border="0" id="btnIpkumDt" style="cursor:pointer;" />
			<script type="text/javascript">
				var CAL_IpkumDate = new Calendar({
					inputField : "termIpkumDt", trigger    : "btnIpkumDt",
					bottomBar: true, dateFormat: "%Y-%m-%d",
					onSelect: function() {
						this.hide();
					}
				});
			</script>
			<% end if %>
		</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">진행상태</td>
		<td bgcolor="#FFFFFF" >
			<%If idx <> "0" Then %>

				<%= ofranchulgojungsan.FOneItem.GetIpkumStateName %>

				<% if IsNull(ofranchulgojungsan.FOneItem.FIpkumStateCD) then %>
					==&gt;
					<!--
					<input type="button" class="button" onclick="changeIpkumState('0');" value="입금이전으로 변경">
					-->
					<input type="button" class="button" onclick="changeIpkumState('5');" value="일부입금으로 변경">
					<input type="button" class="button" onclick="changeIpkumState('9');" value="입금완료로 변경">
				<% else %>
					==&gt;
					<% if (ofranchulgojungsan.FOneItem.FIpkumStateCD="0") then %>
						<input type="button" class="button" onclick="changeIpkumState('5');" value="일부입금으로 변경">
						<input type="button" class="button" onclick="changeIpkumState('9');" value="입금완료로 변경">
					<% elseif (ofranchulgojungsan.FOneItem.FIpkumStateCD="5") then %>
						<input type="button" class="button" onclick="changeIpkumState('9');" value="입금완료로 변경">
					<% elseif (ofranchulgojungsan.FOneItem.FIpkumStateCD="9") then %>

					<% else %>
						ERROR
					<% end if %>
					<input type="button" class="button" onclick="changeIpkumState('NULL');" value="입금상태 삭제">
				<% end if %>

	    	<% end if %>
		</td>
	</tr>
	<tr>
		<td height="10" bgcolor="#FFFFFF" style="padding: 1px;"></td>
		<td bgcolor="#FFFFFF" style="padding: 1px;"></td>
	</tr>

	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">최초등록자</td>
		<td bgcolor="#FFFFFF"><%= ofranchulgojungsan.FOneItem.Fregusername %>(<%= ofranchulgojungsan.FOneItem.Freguserid %>)</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">최종처리자</td>
		<td bgcolor="#FFFFFF"><%= ofranchulgojungsan.FOneItem.Ffinishusername %>(<%= ofranchulgojungsan.FOneItem.Ffinishuserid %>)</td>
	</tr>
	<tr height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">등록일</td>
		<td bgcolor="#FFFFFF"><%= ofranchulgojungsan.FOneItem.Fregdate %></td>
	</tr>
	<tr height="30">
		<td colspan=2 align=center bgcolor="#FFFFFF">
		<%If idx="0" Then %>
			<input type="button" class="button" value="내용저장" onclick="SaveInfo(frm);">
		<% else %>
			<input type="button" class="button" value="전체수정" onclick="SaveInfo(frm);">
		<%End If %>

		</td>
	</tr>
</table>
</form>
<%
set ofranchulgojungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
