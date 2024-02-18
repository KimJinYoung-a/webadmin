<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs 메모 -CALL 추가
' History : 2007.10.26 한용민 수정
'           2009-01-07 서동석 수정,
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_memocls.asp" -->
<%

1 않쓴다.!!!!!!

dim i, userid, orderserial,id
dim mode, sqlStr
dim ippbxuser, intel, phoneNumber, phoneNumberOut, qadiv
dim isEditMode


ippbxuser   = requestCheckVar(request("ippbxuser"),32)
intel       = requestCheckVar(request("intel"),32)
phoneNumber      = requestCheckVar(request("phoneNumber"),32)
phoneNumberOut  = requestCheckVar(request("phoneNumberOut"),32)

if (phoneNumber<>"") then phoneNumber=ParsingPhoneNumber(phoneNumber)

userid          = RequestCheckVar(request("userid"),32)
orderserial     = RequestCheckVar(request("orderserial"),11)
id              = RequestCheckVar(request("id"),9)




dim ocsmemo
set ocsmemo = New CCSMemo

if (id <> "") then
	ocsmemo.FRectId = id
	ocsmemo.FRectUserID = userid
	ocsmemo.FRectOrderserial = orderserial
	ocsmemo.GetCSMemoDetail

	userid = ocsmemo.FOneItem.FUserID
	orderserial = ocsmemo.FOneItem.Forderserial
	phoneNumber = ocsmemo.FOneItem.FphoneNumber

	isEditMode = true
else
	ocsmemo.GetCSMemoBlankDetail
	''mayBe Inbound
	if (phoneNumber<>"") then ocsmemo.FOneItem.FmmGubun = "1"
	isEditMode = false
end if




'=============================================================================
%>
<script language='javascript'>
var NowDoing = false;
<% if (phoneNumber<>"") or (orderserial<>"") or (userid<>"") then %>
    NowDoing = true;
<% end if %>
function setDoingState(){
    document.all.doingdispinfo.innerHTML = (NowDoing)?"<strong><font color=red>[처리중]</font></strong>":"[대기중]";
}

function setGubunState(){
    var comp = frm.mmGubun;

    if (comp.value == "0") {
        //일반메모
        frm.phoneNumber.disabled = true;
        frm.phoneNumber.style.background = "#DDDDDD";

        frm.phoneNumberOut.disabled = true;
        frm.phoneNumberOut.style.background = "#DDDDDD";

    }else if(comp.value=="1"){
        //인바운드
        frm.phoneNumber.disabled = false;
        frm.phoneNumber.style.background = "#FFFFFF"; //className="text";

        frm.phoneNumberOut.disabled = true;
        frm.phoneNumberOut.style.background = "#DDDDDD"; //className="text_ro";
    }else if(comp.value=="2"){
        //아웃바운드
        frm.phoneNumber.disabled = true;
        frm.phoneNumber.style.background = "#DDDDDD";

        frm.phoneNumberOut.disabled = false;
        frm.phoneNumberOut.style.background = "#FFFFFF";

    }else if(comp.value=="3"){
        //업체통화
        frm.phoneNumber.disabled = true;
        frm.phoneNumber.style.background = "#DDDDDD";

        frm.phoneNumberOut.disabled = false;
        frm.phoneNumberOut.style.background = "#FFFFFF";
    }
}

function checkDoing(){
    if (!NowDoing){
        NowDoing=true;
        setDoingState();
    }
}

function reInput(){
    document.location.href = '/cscenter/ippbxmng/popCallRing.asp';
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

    var isNewWin = false;
    var window_width = 1280;
    var window_height = 1024;

    try{
        opener.Hndlw;
    }catch(e){
        //
        alert('창을 닫으신 후 다시 시도해 주세요.');
        return;
    }

    if (opener.Hndlw==null){
	    opener.Hndlw = window.open("/cscenter/ordermaster/ordermaster.asp","PopOrderMaster","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	    opener.Hndlw.focus();
	    isNewWin=true;
	}

	//창을 닫은경우 체크
	try{
	    opener.Hndlw.focus();
	}catch(e){
	    opener.Hndlw = window.open("/cscenter/ordermaster/ordermaster.asp","PopOrderMaster","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	    opener.Hndlw.focus();
	    isNewWin=true;
	}

	//this.focus();
	//어 씽크로나이즈 오류.. setTimeout();
	//opener.Hndlw.listFrame.SearchByPhoneNumber(iphoneNum); //스크립 오류..

	if (isNewWin){
	    setTimeout("opener.Hndlw.listFrame.SearchByPhoneNumber('" + iphoneNum + "')",1000);
	}else{
	    opener.Hndlw.listFrame.SearchByPhoneNumber(iphoneNum);
	}

}

function SearchOrderByOrderSerial(comp){
    var iOrderserial = comp.value;
    if (iOrderserial.length<1){
        alert('주문번호를 넣고 검색하세요.');
        if (comp.enabled) { comp.focus(); };
        return;
    }

    var isNewWin = false;
    var window_width = 1280;
    var window_height = 1024;

    try{
        opener.Hndlw;
    }catch(e){
        //
        alert('창을 닫으신 후 다시 시도해 주세요.');
        return;
    }

    if (opener.Hndlw==null){
	    opener.Hndlw = window.open("/cscenter/ordermaster/ordermaster.asp","PopOrderMaster","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	    opener.Hndlw.focus();
	    isNewWin=true;
	}

	//창을 닫은경우 체크
	try{
	    opener.Hndlw.focus();
	}catch(e){
	    opener.Hndlw = window.open("/cscenter/ordermaster/ordermaster.asp","PopOrderMaster","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	    opener.Hndlw.focus();
	    isNewWin=true;
	}

	//this.focus();
	//어 씽크로나이즈 오류.. setTimeout();
	//opener.Hndlw.listFrame.SearchByOrderserial(iOrderserial); //스크립 오류..

	if (isNewWin){
	    setTimeout("opener.Hndlw.listFrame.SearchByOrderserial('" + iOrderserial + "')",1000);
	}else{
	    opener.Hndlw.listFrame.SearchByOrderserial(iOrderserial);
	}

}

function SearchOrderByUserID(comp){
    var iUserid = comp.value;
    if (iUserid.length<1){
        alert('아이디를 넣고 검색하세요.');
        if (comp.enabled) { comp.focus(); };
        return;
    }

    var isNewWin = false;
    var window_width = 1280;
    var window_height = 1024;

    try{
        opener.Hndlw;
    }catch(e){
        //
        alert('창을 닫으신 후 다시 시도해 주세요.');
        return;
    }

    if (opener.Hndlw==null){
	    opener.Hndlw = window.open("/cscenter/ordermaster/ordermaster.asp","PopOrderMaster","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	    opener.Hndlw.focus();
	    isNewWin=true;
	}

	//창을 닫은경우 체크
	try{
	    opener.Hndlw.focus();
	}catch(e){
	    opener.Hndlw = window.open("/cscenter/ordermaster/ordermaster.asp","PopOrderMaster","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	    opener.Hndlw.focus();
	    isNewWin=true;
	}

	//this.focus();
	//어 씽크로나이즈 오류.. setTimeout();
	//opener.Hndlw.listFrame.SearchByUserID(iUserid); //스크립 오류..

	if (isNewWin){
	    setTimeout("opener.Hndlw.listFrame.SearchByUserID('" + iUserid + "')",1000);
	}else{
	    opener.Hndlw.listFrame.SearchByUserID(iUserid);
	}

}

function fnClick2Call(comp){
    var iphoneNum = comp.value;
    if (iphoneNum.length<1){
        alert('전화번호를 입력하세요.');
        if (!comp.disabled) { comp.focus(); };
        return;
    }

    //대시 추가 루틴 필요 js로


    if (!opener){
        alert('Err1 - 창을 다시 열어주세요..');
        return;
    }

    //권한 문제.. 부모창이 사라졌을 수 있음..
    try{
        opener.name;
    }catch(e){
        alert('창을 닫으신 후 다시 시도해 주세요.');
        return;
    }

    opener.click2call(iphoneNum);
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

    document.all.i_history_memo.src = "/cscenter/ippbxmng/iframeHistory.asp?userid=" + iuserid + "&orderserial=" + iorderserial + "&phoneNumer=" + iphoneNum;
}

function GotoHistoryMemoMidify(id,userid,orderserial)
{
    frm.action="/cscenter/history/history_memo_write.asp?id=" + id + "&backwindow=" + "opener" + "&userid=" + userid + "&orderserial=" + orderserial
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

	if (document.frm.qadiv.value.length<1){
	    alert("문의 유형을 선택 하세요.");
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


</script>
<body topmargin=10 leftmargin=10 marginwidth=0 marginheight=0>

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
        	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>CS메모 - CALL <DIV id=pindispinfo style='display:inline;border:solid 0 gray;font-size:9pt;height:20px;text-align:center'>[ ]</div></b>
        	<input type="button" class="button" value="신규입력" onclick="javascript:reInput();">
            <DIV id=doingdispinfo style='display:inline;border:solid 0 gray;font-size:9pt;height:20px;text-align:center'></div>
        </td>
        <td align="right">

            <input type="button" class="button" value="<%= chkIIF(isEditMode,"수정","저장") %>" onclick="javascript:SubmitSave();">
	       	<input type="button" class="button" value="완료" <%= chkIIF((Not isEditMode) or (ocsmemo.FOneItem.Fdivcd<>"2"),"disabled","") %> onclick="javascript:SubmitFinish();">
	        <input type="button" class="button" value="삭제" <%= chkIIF(isEditMode,"","disabled") %> onclick="javascript:SubmitDelete();">
	        <input type="button" class="button" value="닫기" onclick="javascript:window.close();">
	    </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" onsubmit="return false;" method="post" action="popCallRing_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="id" value="<%= ocsmemo.FOneItem.Fid %>">
	<tr>
    	<td width="50" bgcolor="<%= adminColor("tabletop") %>">구분</td>
    	<td bgcolor="#FFFFFF">
	        <select name="mmGubun" onChange="setGubunState(this);">
	            <option value="0" <% if ocsmemo.FOneItem.FmmGubun = "0" then %>selected<% end if %>>일반메모</option>
	            <option value="1" <% if ocsmemo.FOneItem.FmmGubun = "1" then %>selected<% end if %>>인바운드통화</option>
	            <option value="2" <% if ocsmemo.FOneItem.FmmGubun = "2" then %>selected<% end if %>>아웃바운드통화</option>
	            <option value="3" <% if ocsmemo.FOneItem.FmmGubun = "3" then %>selected<% end if %>>업체통화</option>
	            <!--
	            <option value="4" <% if ocsmemo.FOneItem.FmmGubun = "4" then %>selected<% end if %>>SMS</option>
	            <option value="5" <% if ocsmemo.FOneItem.FmmGubun = "5" then %>selected<% end if %>>EMAIL</option>
	            -->
	        </select>
        </td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">In전화</td>
    	<td bgcolor="#FFFFFF">
        	<table width="400" cellpadding="0" cellspacing="0" border="0" >
        	<tr>
        	    <td width="100"><input type="text" name="phoneNumber" class="text" value="<%= phoneNumber %>" size="20" onKeyDown="checkDoing();" onKeyPress="if (event.keyCode == 13) SearchOrderByPhoneNo(frm.phoneNumber);"></td>
        	    <td width="100" align="center"><a href="javascript:SearchOrderByPhoneNo(frm.phoneNumber);">주문검색</a></td>
        	    <td width="100" align="center"><!-- a href="javascript:fnClick2Call(frm.phoneNumber);">전화걸기</a --></td>
        	    <td width="100" align="center"><a href="javascript:iMemoList(frm.phoneNumber);">관련메모</a></td>
        	</tr>
        	</table>
    	</td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">Out전화</td>
    	<td bgcolor="#FFFFFF">
        	<table width="400" cellpadding="0" cellspacing="0" border="0" >
        	<tr>
        	    <td width="100"><input type="text" name="phoneNumberOut" class="text" value="<%= phoneNumberOut %>" size="20" onKeyDown="checkDoing();" onKeyPress="if (event.keyCode == 13) SearchOrderByPhoneNo(frm.phoneNumber);"></td>
        	    <td width="100" align="center"><a href="javascript:SearchOrderByPhoneNo(frm.phoneNumberOut);">주문검색</a></td>
        	    <td width="100" align="center"><a href="javascript:fnClick2Call(frm.phoneNumberOut);">전화걸기</a></td>
        	    <td width="100" align="center"><a href="javascript:iMemoList(frm.phoneNumberOut);">관련메모</a></td>
        	</tr>
        	</table>
    	</td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">주문번호</td>
    	<td bgcolor="#FFFFFF">
    	    <table width="400" cellpadding="0" cellspacing="0" border="0" >
        	<tr>
        	    <td width="100"><input type="text" name="orderserial" class="text" value="<%= orderserial %>" size="20" onKeyDown="checkDoing();" ></td>
        	    <td width="100" align="center"><a href="javascript:SearchOrderByOrderSerial(frm.orderserial)">주문검색</a></td>
        	    <td width="100" align="center"><a href="javascript:Clip2Paste()">붙여넣기</a></td>
        	    <td width="100" align="center"><a href="javascript:iMemoList(frm.orderserial);">관련메모</a></td>
        	</tr>
        	</table>
    	</td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">고객ID</td>
    	<td bgcolor="#FFFFFF">
    	    <table width="400" cellpadding="0" cellspacing="0" border="0" >
        	<tr>
        	    <td width="100"><input type="text" name="userid" class="text" value="<%= userid %>" size="20" onKeyDown="checkDoing();"></td>
        	    <td width="100" align="center"><a href="javascript:SearchOrderByUserID(frm.userid)">주문검색</a></td>
        	    <td width="100" align="center"></td>
        	    <td width="100" align="center"><a href="javascript:iMemoList(frm.userid);">관련메모</a></td>
        	</tr>
        	</table>
    	</td>
    </tr>
    <% if id = "" then %>
    <% else %>
	    <tr>
	    	<td bgcolor="<%= adminColor("tabletop") %>">접수<br>일</td>
	    	<td bgcolor="#FFFFFF">
	    		<input type="text" name="regdate" class="text_ro" value="<%= ocsmemo.FOneItem.fregdate %>" size="26" readonly>&nbsp;
	    		당담자ID : <%= ocsmemo.FOneItem.Fwriteuser %>
	    	</td>
	    </tr>
	<% end if %>
	<% if ucase(ocsmemo.FOneItem.Ffinishyn) <> "Y" then %>
    <% else %>
	    <tr>
	    	<td bgcolor="<%= adminColor("tabletop") %>">완료일</td>
	    	<td bgcolor="#FFFFFF">
	    		<input type="text" name="regdate" class="text_ro" value="<%= ocsmemo.FOneItem.Ffinishdate %>" size="26" readonly>&nbsp;
	    		당담자ID : <%= ocsmemo.FOneItem.Ffinishuser %>
	    	</td>
	    </tr>
	<% end if %>
	<tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">유형</td>
    	<td bgcolor="#FFFFFF">

    	    <% if ocsmemo.FOneItem.Fdivcd="2" then %>
    	    <input type=hidden name="divcd" value="2">
    	    <input type="checkbox" name="dummi" checked disabled >처리요청
    	    <% else %>
    	    <input type="checkbox" name="divcd" value="2" >처리요청
    	    <% end if %>

	        <!-- 유형 : -->
	        &nbsp;&nbsp;
  			<select class="select" name="qadiv">
                <option value="">전체</option>
                <option value="00" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="00","selected","") %> >배송문의</option>
                <option value="01" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="01","selected","") %> >주문문의</option>
                <option value="02" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="02","selected","") %> >상품문의</option>
                <option value="03" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="03","selected","") %> >재고문의</option>
                <option value="04" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="04","selected","") %> >취소문의</option>
                <option value="05" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="05","selected","") %> >환불문의</option>
                <option value="06" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="06","selected","") %> >교환문의</option>
                <option value="07" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="07","selected","") %> >AS문의</option>
                <option value="08" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="08","selected","") %> >이벤트문의</option>
                <option value="09" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="09","selected","") %> >증빙서류문의</option>
                <option value="10" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="10","selected","") %> >시스템문의</option>
                <option value="11" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="11","selected","") %> >회원제도문의</option>
                <option value="12" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="12","selected","") %> >회원정보문의</option>
                <option value="13" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="13","selected","") %> >당첨문의</option>
                <option value="14" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="14","selected","") %> >반품문의</option>
                <option value="15" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="15","selected","") %> >입금문의</option>
                <option value="16" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="16","selected","") %> >오프라인문의</option>
                <option value="17" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="17","selected","") %> >쿠폰/마일리지문의</option>
                <option value="18" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="18","selected","") %> >결제방법문의</option>
                <option value="20" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="20","selected","") %> >기타문의</option>
            </select>
	    </td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">메모<br>내용</td>
    	<td bgcolor="#FFFFFF"><textarea name="contents_jupsu" class="textarea" cols="68" rows="10" onKeyPress="checkDoing();"><%= db2html(ocsmemo.FOneItem.Fcontents_jupsu) %></textarea></td>
    </tr>

</table>

<p>

관련 메모
<br>
<iframe id="i_history_memo" name="i_history_memo" src="/cscenter/ippbxmng/iframeHistory.asp?userid=<%= userid %>&orderserial=<%= orderserial %>&phoneNumer=<%= phoneNumber %>" width="480" height="300" scrolling="auto" frameborder="1"></iframe>


<script language='javascript'>
function getOnLoad(){
    alert('사용 중지 파일입니다. 관리자 문의 요망');
    setDoingState();
    setGubunState();
    document.all.pindispinfo.innerHTML = "[" + window.name.substr(18,9) + "]";
}

window.onload = getOnLoad;
</script>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->