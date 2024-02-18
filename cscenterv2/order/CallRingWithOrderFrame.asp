<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs 메모 -CALL 추가
' History : 2007.10.26 이상구 생성
'           2009-01-07 서동석 수정,
'###########################################################
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/cscenterv2/lib/classes/history/cs_memocls.asp" -->
<%

Sub DrawOneDateBoxCS(byval yyyy1,mm1,dd1)
	dim buf,i
	dim startyyyy, endyyyy

	startyyyy = Year(now) - 3
	endyyyy = Year(now)

	if (CLng(mm1) = 12) then
		endyyyy = endyyyy + 1
	end if

	buf = "<select class='select' name='yyyy1'>"
    buf = buf + "<option value='" + CStr(yyyy1) +"' selected>" + CStr(yyyy1) + "</option>"
    for i = startyyyy to endyyyy
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='mm1' >"
    buf = buf + "<option value='" + CStr(mm1) + "' selected>" + CStr(mm1) + "</option>"

    for i=1 to 12
    	buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
	next

    buf = buf + "</select>"

    buf = buf + "<select class='select' name='dd1'>"
    buf = buf + "<option value='" + CStr(dd1) +"' selected>" + CStr(dd1) + "</option>"
    for i=1 to 31
        buf = buf + "<option value='" + Format00(2,i) + "' >" + Format00(2,i) + "</option>"
    next
    buf = buf + "</select>"

    response.write buf
end Sub

dim i, userid, orderserial,id
dim mode, sqlStr
dim ippbxuser, intel, phoneNumber, phoneNumberOut
dim yyyy1,mm1,dd1, hh1, retrydateyn
dim isEditMode
dim sitename

ippbxuser   = requestCheckVar(request("ippbxuser"),32)
intel       = requestCheckVar(request("intel"),32)
phoneNumber      = requestCheckVar(request("phoneNumber"),32)
phoneNumberOut  = requestCheckVar(request("phoneNumberOut"),32)

if (phoneNumber<>"") then phoneNumber=ParsingPhoneNumber(phoneNumber)

userid          = RequestCheckVar(request("userid"),32)
orderserial     = RequestCheckVar(request("orderserial"),11)
id              = RequestCheckVar(request("id"),9)
sitename        = RequestCheckVar(request("sitename"),32)

dim ocsmemo
set ocsmemo = New CCSMemo

retrydateyn = "N"

if (id <> "") then
	ocsmemo.FRectId = id
	ocsmemo.FRectUserID = userid
	ocsmemo.FRectOrderserial = orderserial
	ocsmemo.GetCSMemoDetail

	userid = ocsmemo.FOneItem.FUserID
	orderserial = ocsmemo.FOneItem.Forderserial
	phoneNumber = ocsmemo.FOneItem.FphoneNumber
	sitename = ocsmemo.FOneItem.Fsitename

	if (Not IsNull(ocsmemo.FOneItem.Fretrydate)) and (ocsmemo.FOneItem.Fretrydate <> "") then
		yyyy1 = Year(ocsmemo.FOneItem.Fretrydate)
		mm1 = Month(ocsmemo.FOneItem.Fretrydate)
		dd1 = Day(ocsmemo.FOneItem.Fretrydate)
		hh1 = Hour(ocsmemo.FOneItem.Fretrydate)

		retrydateyn = "Y"
	else
		yyyy1 = Year(Now)
		mm1 = Month(Now)
		dd1 = Day(Now)
		hh1 = Hour(Now)
	end if

	isEditMode = true
else
	ocsmemo.GetCSMemoBlankDetail
	''mayBe Inbound
	if (phoneNumber<>"") then
		ocsmemo.FOneItem.FmmGubun = "1"
	else
		ocsmemo.FOneItem.FmmGubun = "0"
	end if

	yyyy1 = Year(Now)
	mm1 = Month(Now)
	dd1 = Day(Now)
	hh1 = Hour(Now)

	isEditMode = false
end if
'=============================================================================
%>
<script language="JavaScript">
var sitename = "<%= sitename %>";
</script>
<script language="JavaScript" src="/cscenter/js/date.format.js"></script>
<script language="JavaScript" src="/cscenter/ippbxmng/ippbxClick2Call.js"></script>
<script language='javascript'>
var NowDoing = false;
<% if (phoneNumber<>"") or (orderserial<>"") or (userid<>"") then %>
    NowDoing = true;
<% end if %>
function setDoingState(){
    document.all.doingdispinfo.innerHTML = (NowDoing)?"<strong><font color=red>[처리중]</font></strong>":"[대기중]";
}

function setGubunState(){
    // do nothing
}

function checkDoing(){
    if (!NowDoing){
        NowDoing=true;
        setDoingState();
    }
}

function reInput(){
    document.location.href = '/cscenterv2/order/CallRingWithOrderFrame.asp?sitename=' + sitename;
}

function Clip2Paste(){
    var clipTxt = window.clipboardData.getData("Text");

    if (clipTxt.length<1){ return; }

    //indexOf
    var posSpliter = clipTxt.indexOf("|");
    var iorderserial ="";
    var iuserid ="";
    if (posSpliter>0){
        iorderserial = clipTxt.substring(0,posSpliter);
        iuserid      = clipTxt.substring(posSpliter+1,255);

        frm.orderserial.value = iorderserial;
        frm.userid.value = iuserid;
    }


}


function SearchOrderByPhoneNo(comp){
    var iphoneNum = comp.value;
    if (iphoneNum.length<1){
        alert('전화번호를 넣고 검색하세요.');
        if (comp.enabled) { comp.focus(); };
        return;
    }

    top.listFrame.SearchByPhoneNumber(iphoneNum);
}

function SearchOrderByOrderSerial(comp){
    var iOrderserial = comp.value;
    if (iOrderserial.length<1){
        alert('주문번호를 넣고 검색하세요.');
        if (comp.enabled) { comp.focus(); };
        return;
    }

    top.listFrame.SearchByOrderserial(iOrderserial);
}

function SearchOrderByUserID(comp){
    var iUserid = comp.value;
    if (iUserid.length<1){
        alert('아이디를 넣고 검색하세요.');
        if (comp.enabled) { comp.focus(); };
        return;
    }
    top.listFrame.SearchByUserID(iUserid);
}

function iMemoList(comp){
    var iphoneNum    = "";
    var iuserid      = "";
    var iorderserial = "";

    if ((comp.name=="phoneNumber")||(comp.name=="phoneNumberOut")){
        iphoneNum = comp.value;
        if (iphoneNum.length<1){
            alert('전화번호를 입력하세요.');
            if (!comp.disabled) { comp.focus(); };
            return;
        }
    }else if (comp.name=="userid"){
        iuserid = comp.value;
        if (iuserid.length<1){
            alert('아이디를 입력하세요.');
            comp.focus();
            return;
        }
    }else if (comp.name=="orderserial"){
        iorderserial = comp.value;
        if (iorderserial.length<1){
            alert('주문번호를  입력하세요.');
            comp.focus();
            return;
        }
    }

    document.all.i_history_memo.src = "/cscenterv2/order/iframeHistory.asp?sitename=" + sitename + "&userid=" + iuserid + "&orderserial=" + iorderserial + "&phoneNumer=" + iphoneNum;
}

function GotoHistoryMemoMidify(id,userid,orderserial)
{
    frm.action="/cscenterv2/history/history_memo_write.asp?id=" + id + "&backwindow=" + "opener" + "&userid=" + userid + "&orderserial=" + orderserial
    frm.submit();
}

function SubmitSave()
{
    if ((document.frm.orderserial.value.length<1)&&(document.frm.userid.value.length<1)&&(document.frm.phoneNumber.value.length<1)) {
	    alert("전화번호, 주문번호, 아이디 중 하나는 입력 되어야 합니다.");
		return;
	}

	if (document.frm.contents_jupsu.value == "") {
		alert("메모내용을 입력하세요.");
		document.frm.contents_jupsu.focus();
		return;
	}

	if (document.frm.mmgubun.value.length<1){
	    alert("구분을 선택 하세요.");
		document.frm.mmgubun.focus();
		return;
	}

	if (document.frm.qadiv.value.length<1){
	    alert("구분상세를 선택 하세요.");
		document.frm.qadiv.focus();
		return;
	}

	if(document.frm.id.value == "") {
    	document.frm.mode.value = "write";
    	document.frm.submit();
	}else{
    	document.frm.mode.value = "modify";
    	document.frm.submit();
	}
}

function SubmitFinish(){
	if (document.frm.contents_jupsu.value == "") {
			alert("메모내용을 입력하세요.");
			return;
	}

    if (confirm("완료처리하겠습니까?") == true) {
            document.frm.mode.value = "finish";
            document.frm.submit();
    }
}

function SubmitDelete()
{
    if (confirm("삭제하겠습니까?") == true) {
            document.frm.mode.value = "delete";
            document.frm.submit();
    }
}

function ToggleRetryDate(v)
{
	var yyyy1 = frm.yyyy1;
	var mm1 = frm.mm1;
	var dd1 = frm.dd1;
	var hh1 = frm.hh1;

	if (v.checked != true) {
		// 지정않함
		yyyy1.disabled = true;
		mm1.disabled = true;
		dd1.disabled = true;
		hh1.disabled = true;
	} else {
		// 직접설정
		yyyy1.disabled = false;
		mm1.disabled = false;
		dd1.disabled = false;
		hh1.disabled = false;
	}
}

function WriteNowDateString(v) {
	var d = new Date();
	v.focus();

	<% if (id = "") then %>
		alert("접수시에는 시간을 입력할 수 없습니다.[접수시간과 동일]");
		return;
	<% end if %>

	// /cscenter/js/date.format.js
	v.value = v.value + "\n\n+" + d.format("yyyy-mm-dd HH:MM:ss") + "\n";
}

function SetToRetryDateTommorow() {
	var d = new Date();

	var yyyy1 = frm.yyyy1;
	var mm1 = frm.mm1;
	var dd1 = frm.dd1;
	var hh1 = frm.hh1;

	d.setDate(d.getDate()+1);

	frm.retrydateyn.checked = true;
	ToggleRetryDate(frm.retrydateyn);

	yyyy1.value = d.getFullYear();
	mm1.value = d.getMonth() + 1;
	dd1.value = d.getDate();
	hh1.value = 10;
}

</script>
<!--<body topmargin=10 leftmargin=10 marginwidth=0 marginheight=0>-->

<!-- CS메모-CALL 시작-->
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="<%= adminColor("topbar") %>">
        <td colspan="2">
			<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
        		<tr>
			        <td>
			        	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>CS메모-CALL</b>
			        	<b><DIV id=pindispinfo style='display:inline;border:solid 0 gray;font-size:9pt;height:17px;text-align:absbottom'>[ ]</div></b>
			            <DIV id=doingdispinfo style='display:inline;border:solid 0 gray;font-size:9pt;height:17px;text-align:absbottom'></div>
			            <input type="button" class="button" value="신규입력" onclick="javascript:reInput();">
			        </td>
			        <td align="right">

			            <input type="button" class="button" value="<%= chkIIF(isEditMode,"수정","저장") %>" onclick="javascript:SubmitSave();">
				       	<input type="button" class="button" value="완료" <%= chkIIF((Not isEditMode) or (ocsmemo.FOneItem.Fdivcd<>"2"),"disabled","") %> onclick="javascript:SubmitFinish();">
						<!-- CS팀장님 요청으로 숨김
				        <input type="button" class="button" value="삭제" <%= chkIIF(isEditMode,"","disabled") %> onclick="javascript:SubmitDelete();">
						-->
				    </td>
				</tr>
			</table>
		</td>
	</tr>

    <form name="frm" onsubmit="return false;" method="post" action="popCallRing_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="sitename" value="<%= sitename %>">
	<input type="hidden" name="id" value="<%= ocsmemo.FOneItem.Fid %>">
    <tr>
    	<td width="50" height="30" bgcolor="<%= adminColor("tabletop") %>">고객전화</td>
    	<td bgcolor="#FFFFFF">
        	<table width="370" cellpadding="0" cellspacing="0" border="0" >
        	<tr>
        	    <td width="100"><input type="text" name="phoneNumber" class="text" value="<%= phoneNumber %>" size="20" onKeyDown="checkDoing();" onKeyPress="if (event.keyCode == 13) SearchOrderByPhoneNo(frm.phoneNumber);"></td>
        	    <td width="90" align="center"><a href="javascript:SearchOrderByPhoneNo(frm.phoneNumber);">주문검색</a></td>
        	    <td width="90" align="center"><a href="javascript:alert('사용불가!!');;">전화걸기</a></td>
        	    <td width="90" align="center"><a href="javascript:iMemoList(frm.phoneNumber);">관련메모</a></td>
        	</tr>
        	</table>
    	</td>
    </tr>
    <tr>
    	<td height="30" bgcolor="<%= adminColor("tabletop") %>">주문번호</td>
    	<td bgcolor="#FFFFFF">
    	    <table width="370" cellpadding="0" cellspacing="0" border="0" >
        	<tr>
        	    <td width="100"><input type="text" name="orderserial" class="text" value="<%= orderserial %>" size="20" onKeyDown="checkDoing();" ></td>
        	    <td width="90" align="center"><a href="javascript:SearchOrderByOrderSerial(frm.orderserial)">주문검색</a></td>
        	    <td width="90" align="center"><a href="javascript:Clip2Paste()">붙여넣기</a></td>
        	    <td width="90" align="center"><a href="javascript:iMemoList(frm.orderserial);">관련메모</a></td>
        	</tr>
        	</table>
    	</td>
    </tr>
    <tr>
    	<td height="30" bgcolor="<%= adminColor("tabletop") %>">고객ID</td>
    	<td bgcolor="#FFFFFF">
    	    <table width="370" cellpadding="0" cellspacing="0" border="0" >
        	<tr>
        	    <td width="100"><input type="text" name="userid" class="text" value="<%= userid %>" size="20" onKeyDown="checkDoing();"></td>
        	    <td width="90" align="center"><a href="javascript:SearchOrderByUserID(frm.userid)">주문검색</a></td>
        	    <td width="90" align="center"></td>
        	    <td width="90" align="center"><a href="javascript:iMemoList(frm.userid);">관련메모</a></td>
        	</tr>
        	</table>
    	</td>
    </tr>
    <% if id = "" then %>
    <% else %>
	    <tr>
	    	<td height="30" bgcolor="<%= adminColor("tabletop") %>">접수일</td>
	    	<td bgcolor="#FFFFFF">
	    		<input type="text" name="regdate" class="text_ro" value="<%= ocsmemo.FOneItem.fregdate %>" size="26" readonly>&nbsp;
	    		등록자ID : <%= ocsmemo.FOneItem.Fwriteuser %>
	    	</td>
	    </tr>

	<% end if %>

    <tr>
    	<td height="30" bgcolor="<%= adminColor("tabletop") %>">다음처리</td>
    	<td bgcolor="#FFFFFF">

    		<input type="checkbox" name="retrydateyn" <% if retrydateyn = "Y" then %>checked<% end if %> onClick="ToggleRetryDate(this)">

    		<% Call DrawOneDateBoxCS(yyyy1,mm1,dd1) %>
    		&nbsp;
    		<select name="hh1" class="select">
    			<option value=""></option>
    			<% for i = 9 to 18 %>
    				<option value="<%= i %>" <% if (hh1 = i) then %>selected<% end if %>><%= Right(("0" & i), 2) %></option>
    			<% next %>
    		</select>

    		<input type="button" class="button" value="내일오전" onClick="SetToRetryDateTommorow()">
    	</td>
    </tr>

	<% if ucase(ocsmemo.FOneItem.Ffinishyn) <> "Y" then %>

    <% else %>
	    <tr>
	    	<td height="30" bgcolor="<%= adminColor("tabletop") %>">완료일</td>
	    	<td bgcolor="#FFFFFF">
	    		<input type="text" name="regdate" class="text_ro" value="<%= ocsmemo.FOneItem.Ffinishdate %>" size="26" readonly>&nbsp;
	    		처리자ID : <%= ocsmemo.FOneItem.Ffinishuser %>
	    	</td>
	    </tr>
	<% end if %>

	<tr>
    	<td height="30" bgcolor="<%= adminColor("tabletop") %>">특이사항</td>
    	<td bgcolor="#FFFFFF">
			<input type="radio" name="specialmemo" value="" <% if (ocsmemo.FOneItem.Fspecialmemo = "") then %>checked<% end if %> >없음
			&nbsp; &nbsp; &nbsp;
    	    <input type="radio" name="specialmemo" value="제휴몰" <% if (ocsmemo.FOneItem.Fspecialmemo = "제휴몰") then %>checked<% end if %> >제휴몰
			&nbsp; &nbsp; &nbsp;
			<input type="radio" name="specialmemo" value="#" <% if (ocsmemo.FOneItem.Fspecialmemo = "#") then %>checked<% end if %> >#
			&nbsp; &nbsp; &nbsp;
			<input type="radio" name="specialmemo" value="##" <% if (ocsmemo.FOneItem.Fspecialmemo = "##") then %>checked<% end if %> >##
			&nbsp; &nbsp; &nbsp;
			<input type="radio" name="specialmemo" value="###" <% if (ocsmemo.FOneItem.Fspecialmemo = "###") then %>checked<% end if %> ><font color="red">###</font>
	    </td>
    </tr>

	<tr>
    	<td height="30" bgcolor="<%= adminColor("tabletop") %>">구분</td>
    	<td bgcolor="#FFFFFF">
    	    <!-- #include virtual="/cscenter/memo/mmgubunselectbox.asp"-->
	    </td>
    </tr>

	<tr>
    	<td height="30" bgcolor="<%= adminColor("tabletop") %>">처리유형</td>
    	<td bgcolor="#FFFFFF">
    	    <% if ocsmemo.FOneItem.Fdivcd="2" then %>
    	    <input type=hidden name="divcd" value="2">
    	    <input type="checkbox" name="dummi" checked disabled > 요청
    	    <% else %>
    	    <input type="checkbox" name="divcd" value="2" > 요청
    	    <% end if %>
	    </td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>" align="center">
    		메 모<br>내 용<br><br>
    		<input type="button" class="button" value="시간" onClick="WriteNowDateString(frm.contents_jupsu)">
    	</td>
    	<td bgcolor="#FFFFFF">
    		<textarea name="contents_jupsu" class="textarea" cols="60" rows="10" onKeyDown="checkDoing();"><%= replace(db2html(ocsmemo.FOneItem.Fcontents_jupsu),"<br>",vbCrlf) %></textarea><br>
    	</td>
    </tr>
	</form>
</table>

<!-- CS메모-CALL 끝-->

<p>

<!-- 관련메모 시작-->

<!-- CS메모-CALL 시작-->
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="<%= adminColor("topbar") %>">
        <td colspan="2">
			<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
        		<tr>
			        <td>
			        	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>관련메모</b>
			        </td>
			        <td align="right">

				    </td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="2">
			<iframe id="i_history_memo" name="i_history_memo" src="/cscenterv2/order/iframeHistory.asp?userid=<%= userid %>&orderserial=<%= orderserial %>&phoneNumer=<%= phoneNumber %>&id=<%= id %>" width="100%" height="500" scrolling="auto" frameborder="0"></iframe>
		</td>
	</tr>
</table>

<script language='javascript'>
function getOnLoad(){
    setDoingState();
    setGubunState();
    document.all.pindispinfo.innerHTML = "[" + window.name.substr(18,9) + "]";

	// /cscenter/memo/mmgubunselectbox.asp 참조
	startRequest('mmgubun','<%= ocsmemo.FOneItem.FmmGubun %>','<%= ocsmemo.FOneItem.Fqadiv %>');

	ToggleRetryDate(frm.retrydateyn);
}

window.onload = getOnLoad;
</script>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
