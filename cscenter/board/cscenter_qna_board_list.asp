<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 1:1 상담
' History : 2009.04.17 이상구 생성
'			2016.03.25 한용민 수정(문의분야 모두 DB화 시킴)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<!-- #include virtual="/lib/classes/board/myqnacls.asp" -->
<%
dim i, j, sDate ,eDate ,blnDate, page, ckReplyDate, replyDate1, replyDate2, searchDiv, searchText, tmpqadivname
dim itemqanotinclude, research, finishyn, userid, orderserial, qadiv , writeid, chargeid, replyqadiv, userlevel, evalPoint
dim isusing, sitename
dim currstate, userGubun
	qadiv               = request("qadiv")
	itemqanotinclude    = request("itemqanotinclude")
	research            = request("research")
	userid              = request("userid")
	orderserial         = request("orderserial")
	writeid             = request("writeid")
	chargeid            = request("chargeid")
	replyqadiv          = request("replyqadiv")
	userlevel			= request("userlevel")
	evalPoint			= request("evalPoint")
	searchDiv			= request("searchDiv")
	searchText			= Trim(request("searchText"))
	sDate = request("sdt")
	eDate = request("edt")
	blnDate = request("edc")
	'if (itemqanotinclude="") and (research="") then itemqanotinclude="on"
	isusing				= request("isusing")
	sitename			= request("sitename")

	if (sitename = "") and not(C_ADMIN_AUTH) then
		sitename = "10x10"
	end if

	currstate			= request("currstate")
	userGubun			= request("userGubun")

if (research = "") then
	isusing = "Y"
	sitename = "10x10"
end if

if ((userid <> "") or (orderserial <> "")) then
    qadiv = ""
    itemqanotinclude = ""
end if

page	= req("page",1)

ckReplyDate	= req("ckReplyDate",req("ckReplyDateDefault",""))
replyDate1	= req("replyDate1",LEFT(CStr(dateAdd("d",-7,now())),10))
replyDate2	= req("replyDate2",LEFT(CStr(now()),10))

if (blnDate="") and (research="") then
    blnDate = "on"
    sDate   = LEFT(CStr(dateAdd("m",-3,now())),10)
    eDate   = LEFT(CStr(now()),10)
end if

finishYN = req("finishYN","")

dim boardqna
set boardqna = New CMyQNA
	boardqna.FPageSize = 50
	boardqna.FCurrPage = page
	boardqna.RectQadiv = qadiv
	boardqna.FSearchUserID = userid
	boardqna.FSearchOrderSerial = orderserial
	boardqna.FSearchWriteId = writeId
	boardqna.FSearchChargeId = chargeid
	boardqna.FRectReplyQADiv = replyqadiv
	boardqna.FSearchUserLevel = userlevel
	boardqna.FSearchDiv = searchDiv
	boardqna.FSearchText = searchText
	boardqna.FRectEvalPoint = evalPoint
	boardqna.FRectIsUsing = isusing

	IF blnDate="on" Then
		boardqna.FSearchStartDate = sDate
		boardqna.FSearchEndDate =eDate
	End IF

	IF ckReplyDate="on" Then
		boardqna.FreplyDate1 = replyDate1
		boardqna.FreplyDate2 =replyDate2
	End IF

	boardqna.FRectItemNotInclude = itemqanotinclude
	boardqna.FRectSiteName = sitename
	boardqna.FRectCurrState = currstate
	boardqna.FRectUserGubun = userGubun

	''boardqna.list finishYN

	''old ver
	'if (finishyn = "N") then
	    boardqna.SearchNew = finishyn
	'end if

	boardqna.fqnalist

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">

$(function() {
  	$("#finishyn_A").button().children().attr("style","font-size:12px;<%=CHKIIF(finishyn="","font-weight:bold;color:red;","color:black;")%>");
  	$("#finishyn_N").button().children().attr("style","font-size:12px;<%=CHKIIF(finishyn="N","font-weight:bold;color:red;","color:black;")%>");
  	$("#finishyn_VV").button().children().attr("style","font-size:12px;<%=CHKIIF(finishyn="V","font-weight:bold;color:red;","color:black;")%>");
	$("#finishyn_VE").button().children().attr("style","font-size:12px;<%=CHKIIF(finishyn="E","font-weight:bold;color:red;","color:black;")%>");
	$("#finishyn_VD").button().children().attr("style","font-size:12px;<%=CHKIIF(finishyn="D","font-weight:bold;color:red;","color:black;")%>");
  	$("#finishyn_V").button().children().attr("style","font-size:12px;<%=CHKIIF(finishyn="V","font-weight:bold;color:red;","color:black;")%>");
	$("#finishyn_E").button().children().attr("style","font-size:12px;<%=CHKIIF(finishyn="E","font-weight:bold;color:red;","color:black;")%>");
	$("#finishyn_D").button().children().attr("style","font-size:12px;<%=CHKIIF(finishyn="D","font-weight:bold;color:red;","color:black;")%>");
});

function CloseWindow(){
	window.close();
}

function  TnSearch(frm){
	if (frm.rectuserid.length<1){
		alert('검색어를 입력하세요.');
		return;
	}
	frm.method="get";
	frm.submit();
}

function NextPage(ipage){
	document.frmSrc.page.value= ipage;
	document.frmSrc.submit();
}

function SubmitSearch() {
    document.qnaform.submit();
}

function SubmitSearchUserId(userid) {
    document.qnaform.userid.value = userid;
    document.qnaform.orderserial.value = "";
    document.qnaform.submit();
}


function jsPopCal(fName,sName){
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

function EnableDate(obj){
	var f = document.qnaform;
	if (obj.checked)
	{
		f.sdt.readOnly=false;
		f.edt.readOnly=false;
		f.sdt.className="text";
		f.edt.className="text";
	}
	else
	{
		f.sdt.readOnly=true;
		f.edt.readOnly=true;
		f.sdt.className="text_ro";
		f.edt.className="text_ro";
	}
}

function replyEnableDate(obj){
	var f = document.qnaform;
	if (obj.checked)
	{
		f.replyDate1.readOnly=false;
		f.replyDate2.readOnly=false;
		f.replyDate1.className="text";
		f.replyDate2.className="text";
	}
	else
	{
		f.replyDate1.readOnly=true;
		f.replyDate2.readOnly=true;
		f.replyDate1.className="text_ro";
		f.replyDate2.className="text_ro";
	}
}

function jsFinishYNButton(a) {
    document.qnaform.userid.value = "";
    document.qnaform.orderserial.value = "";
    document.qnaform.finishyn.value = a;
    document.qnaform.submit();
}

function jsDelQna(id) {
	if (confirm("삭제하시겠습니까?")) {
		document.delform.id.value = id;
		document.delform.submit();
	}
}

document.title = "1:1 상담리스트";

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td width="400" style="padding:5; border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>"  background="/images/menubar_1px.gif">
					<font color="#333333"><b>[CS]고객센터>>[1:1상담]게시판관리</b></font>
				</td>

				<td align="right" style="border-bottom:1px solid <%= adminColor("tablebg") %>" bgcolor="#F4F4F4">
				</td>

			</tr>
		</table>
	</td>
</tr>

<!--	설명 있으면 들어갑니다.	-->
<tr bgcolor="#FFFFFF">
	<td style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
		- 최근 100건만 검색됩니다.<br>
		- 아이디 또는 주문번호로 검색할 경우는 답변유무/질문유형에 관계없이 모두 표시됩니다.
	</td>
</tr>
</table>

<!-- 검색 시작 -->
<form method="get" name="qnaform" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">검 색<br>조 건</td>
	<td height="31" align="left">&nbsp;
		고객아이디 : <input type="text" class="text" name="userid" value="<%= userid %>" size="15" onKeyPress="if (event.keyCode == 13) document.qnaform.submit();">
  		&nbsp;&nbsp;
  		주문번호 : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" size="15" onKeyPress="if (event.keyCode == 13) document.qnaform.submit();">
  		&nbsp;&nbsp;/&nbsp;&nbsp;
  		담당자아이디 : <input type="text" class="text" name="chargeid" value="<%= chargeid %>" size="15" onKeyPress="if (event.keyCode == 13) document.qnaform.submit();">
  		&nbsp;&nbsp;
  		답변자아이디 : <input type="text" class="text" name="writeid" value="<%= writeid %>" size="15" onKeyPress="if (event.keyCode == 13) document.qnaform.submit();">
		&nbsp;&nbsp;
		쇼핑몰 :
	    <% call drawSelectBoxXSiteOrderInputPartnerCS("sitename", sitename) %>
		&nbsp;&nbsp;
		상태 :
        <select class="select" name="currstate">
        	<option value="" <%=CHKIIF(currstate="","selected","")%>>선택</option>
        	<option value="B001" <%=CHKIIF(currstate="B001","selected","")%>>답변이전 전체</option>
			<option value="B006" <%=CHKIIF(currstate="B006","selected","")%>>업체 답변완료</option>
			<option value="B007" <%=CHKIIF(currstate="B007","selected","")%>>답변완료</option>
			<option value="B008" <%=CHKIIF(currstate="B008","selected","")%>>전송이전(제휴)</option>
			<option value="B009" <%=CHKIIF(currstate="B009","selected","")%>>전송완료(제휴)</option>
        </select>
  	</td>
	<td rowspan="4" width="80" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검 색" style="width:60px;height:70px;" onClick="SubmitSearch()">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td height="31" align="left">&nbsp;
  		질문유형 :
  		<% drawSelectBoxqadiv "qadiv", qadiv, "", "Y", "N", "Y" %>
        &nbsp;&nbsp;
  		질문구분 :
		<select class="select" name="replyqadiv">
			<option value="">전체</option>
			<option value="">======</option>
            <option value="01" <% if replyqadiv = "01" then response.write "selected" %>>단순문의</option>
			<option value="">======</option>
            <option value="all" <% if replyqadiv = "all" then response.write "selected" %>>고객불만 전체</option>
			<option value="02"  <% if replyqadiv = "02" then response.write "selected" %>>업체불만</option>
            <option value="03"  <% if replyqadiv = "03" then response.write "selected" %>>배송(CJ)불만</option>
            <option value="10"  <% if replyqadiv = "10" then response.write "selected" %>>시스템개선요청</option>
            <option value="99"  <% if replyqadiv = "99" then response.write "selected" %>>기타불만</option>
        </select>
        &nbsp;&nbsp;
        회원등급 : <% DrawselectboxUserLevel "userlevel", userlevel, "" %>
        &nbsp;&nbsp;
        고객만족도 :
        <select class="select" name="evalPoint">
        	<option value="" <%=CHKIIF(evalPoint="","selected","")%>>전체</option>
        	<option value="5" <%=CHKIIF(evalPoint="5","selected","")%>>5점</option>
			<option value="4" <%=CHKIIF(evalPoint="4","selected","")%>>4점</option>
			<option value="3" <%=CHKIIF(evalPoint="3","selected","")%>>3점</option>
			<option value="2" <%=CHKIIF(evalPoint="2","selected","")%>>2점</option>
			<option value="1" <%=CHKIIF(evalPoint="1","selected","")%>>1점</option>
			<option>--------</option>
			<option value="3DN" <%=CHKIIF(evalPoint="3DN","selected","")%>>3점 이하 전체</option>
        </select>
		&nbsp;&nbsp;
		검색조건 :
        <select class="select" name="searchDiv">
        	<option value="" <%=CHKIIF(searchDiv="","selected","")%>>선택</option>
        	<option value="title" <%=CHKIIF(searchDiv="title","selected","")%>>제목</option>
			<option value="contents" <%=CHKIIF(searchDiv="contents","selected","")%>>내용</option>
			<option value="makerid" <%=CHKIIF(searchDiv="makerid","selected","")%>>브랜드</option>
			<option value="username" <%=CHKIIF(searchDiv="username","selected","")%>>고객명</option>
        </select>
		<input type="text" class="text" name="searchText" value="<%= searchText %>" size="12" onKeyPress="if (event.keyCode == 13) document.qnaform.submit();">
		&nbsp;&nbsp;
		작성자 :
        <select class="select" name="userGubun">
			<option></option>
			<option value="C" <%= CHKIIF(userGubun="C", "selected", "") %>>고객</option>
			<option value="M" <%= CHKIIF(userGubun="M", "selected", "") %>>상담사</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td height="31" align="left">&nbsp;
        <input type="checkbox" name="edc" <%IF blnDate="on" then response.write "checked" %> onclick="EnableDate(this);">
        고객작성일 : <input type="text" size="10" name="sdt" value="<%= sDate %>" onClick="jsPopCal('qnaform','sdt');" <% IF blnDate="on" Then%>class="text" <%Else%>readonly class="text_ro"<%END IF%> style="cursor:hand;">
        ~<input type="text" size="10" name="edt" value="<%= eDate %>" onClick="jsPopCal('qnaform','edt');" <% IF blnDate="on" Then%>class="text" <%Else%>readonly class="text_ro"<%END IF%> style="cursor:hand;">
        &nbsp;&nbsp;
        <input type="checkbox" name="ckReplyDate" <%IF ckReplyDate="on" then response.write "checked" %> onclick="replyEnableDate(this);">
        답변일 : <input type="text" size="10" name="replyDate1" value="<%= replyDate1 %>" onClick="jsPopCal('qnaform','replyDate1');" <% IF ckReplyDate="on" Then%>class="text" <%Else%>readonly class="text_ro"<%END IF%> style="cursor:hand;">
        ~<input type="text" size="10" name="replyDate2" value="<%= replyDate2 %>" onClick="jsPopCal('qnaform','replyDate2');" <% IF ckReplyDate="on" Then%>class="text" <%Else%>readonly class="text_ro"<%END IF%> style="cursor:hand;">
		&nbsp;&nbsp;
		<input type="checkbox" name="isusing" value="Y" <%= CHKIIF(isusing="Y", "checked", "") %> > 삭제내역 표시안함
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td height="31" align="left">&nbsp;
		<input type="hidden" name="finishyn" value="<%=finishyn%>">
		<button type="button" id="finishyn_A" onClick="jsFinishYNButton('');">전체목록</button>
		&nbsp;
		<button type="button" id="finishyn_N" onClick="jsFinishYNButton('N');">미처리목록</button>
		&nbsp;
		<button type="button" id="finishyn_VV" onClick="jsFinishYNButton('VV');">VVIP 미처리전체</button>
		<button type="button" id="finishyn_VE" onClick="jsFinishYNButton('VE');">VVIP 일반상담</button>
		<button type="button" id="finishyn_VD" onClick="jsFinishYNButton('VD');">VVIP 배송문의</button>
		&nbsp;
		<button type="button" id="finishyn_V" onClick="jsFinishYNButton('V');">VIP 미처리전체</button>
		<button type="button" id="finishyn_E" onClick="jsFinishYNButton('E');">VIP 일반상담</button>
		<button type="button" id="finishyn_D" onClick="jsFinishYNButton('D');">VIP 배송문의</button>
	</td>
</tr>
</table>
</form>

<br>
* <font color="blue">[M]</font> : 모바일 사이트에서 작성한 문의입니다.<br>
* <font color="orange">[A]</font> : 앱에서 작성한 문의입니다.
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="15" style="padding:3 0 3 5">검색결과 : <b><%=boardqna.ResultCount%></b> / <%=boardqna.TotalCount%></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="70" height="25">레벨</td>
    <td width="135">고객명(아이디)</td>
    <td width="70">사이트</td>
	<td width="70">주문번호</td>
    <td width="120">브랜드</td>
    <td width="50">문의상품</td>
	<td width="100">구분</td>
    <td>제목</td>
    <td width="30">첨부</td>
	<td width="80">작성일</td>
    <td width="60">담당자</td>
    <td width="100">업체답변</td>
	<td width="100">답변여부</td>
	<td width="30">전송</td>
    <td width="30">삭제</td>
</tr>
<% if boardqna.ResultCount>0 then %>
	<%
	for i = 0 to boardqna.ResultCount - 1

	if isarray(split(boardqna.results(i).fqadivname,"!@#")) then
		if ubound(split(boardqna.results(i).fqadivname,"!@#")) > 0 then
			tmpqadivname =  split(boardqna.results(i).fqadivname,"!@#")(1)
		end if
	end if
	%>
	<% if (boardqna.results(i).dispyn = "N") then %>
		<tr align="center" bgcolor="#EEEEEE">
	<% else %>
		<tr align="center" bgcolor="#FFFFFF">
	<% end if %>

		<td align="center" height="25">
			<% if (boardqna.results(i).Fsitename = "10x10") or (boardqna.results(i).Fsitename = "") then %>
				<font color="<%= getUserLevelColorByDate(boardqna.results(i).fUserLevel,Left(boardqna.results(i).regdate,10)) %>">
				<strong><%= getUserLevelStrByDate(boardqna.results(i).fUserLevel,Left(boardqna.results(i).regdate,10)) %></strong></font>
			<% end if %>
		</td>
	    <td align="left">
			<% if boardqna.results(i).Fuserlevel="7" then %>
				<font color="blue"><%= boardqna.results(i).username %></font>
		    	<!--<a href="javascript:SubmitSearchUserId('<%'= boardqna.results(i).userid %>');">-->
		    	<font color="blue">(<%= printUserId(boardqna.results(i).userid, 2, "*") %>)</font>
		    	<!--</a>-->
				</font>
			<% else %>
		    	<%= boardqna.results(i).username %>
		    	<!--<a href="javascript:SubmitSearchUserId('<%'= boardqna.results(i).userid %>');">-->
		    	(<%= printUserId(boardqna.results(i).userid, 2, "*") %>)
		    	<!--</a>-->
			<% end if %>
	    </td>
	    <td><%= boardqna.results(i).Fsitename %></td>
		<td><%= boardqna.results(i).orderserial %></td>
	    <td>
	    	<%= boardqna.results(i).Fmakerid %>
	    	<% if (boardqna.results(i).IsUpchebeasong) then %>
	    		<font color=red>(업배)</font>
	    	<% end if %>
	    </td>
	    <td><%= boardqna.results(i).itemid %></td>
		<td>
			<a href="cscenter_qna_board_reply.asp?id=<%= boardqna.results(i).id %>">
			<% if boardqna.results(i).qadiv="26" then %>
				<font color="blue"><%= tmpqadivname %></font>
			<% else %>
				<%= tmpqadivname %>
			<% end if %>
			</a>
		</td>
	    <td align="left">
			<% if Not IsNull(boardqna.results(i).FExtSiteName) then %>
				<% if (boardqna.results(i).FExtSiteName = "mobile") then %>
					<font color="blue">[M]</font>
				<% elseif (boardqna.results(i).FExtSiteName = "app") then %>
					<font color="orange">[A]</font>
				<% end if %>
			<% end if %>
			<% if (boardqna.results(i).FEvalPoint > 0) then %>
				<% for j = 1 to boardqna.results(i).FEvalPoint %><img src="http://fiximage.10x10.co.kr/web2009/mytenbyten/star_red.gif"><% next %>
			<% end if %>
			<a href="cscenter_qna_board_reply.asp?id=<%= boardqna.results(i).id %>">
				<%= db2html(boardqna.results(i).title) %>
				<% if (boardqna.results(i).title = "") then %>(제목없음)<% end if %>
			</a>
			<!--
			<a href="cscenter_qna_board_reply_new.asp?id=<%= boardqna.results(i).id %>">
				[새버전]
			</a>
			-->
		</td>
		<td><%= CHKIIF(boardqna.results(i).FattachFile <> "", "Y", "") %></td>
	    <td align="center">
	    	<%
			' 이문재이사님 지시. 금일 표기 하지 말고 그냥 날짜와 시간 단순하게 표기하라 하심.	' 2019.05.16 한용민
			'if (Left(boardqna.results(i).regdate, 10) < Left(now, 10)) then
			%>
			<% if boardqna.results(i).regdate<>"" and not(isnull(boardqna.results(i).regdate)) then %>
				<%= Left(boardqna.results(i).regdate,10) %>
				<br><%= mid(boardqna.results(i).regdate,11,18) %>
	    	<% end if %>
			<% 'else %>
	      	<!--금일 <%'= Right(FormatDate(boardqna.results(i).regdate, "0000.00.00 00:00:00"), 8) %>-->
	    	<% 'end if %>
	    </td>
	    <td>
	    	<% if (boardqna.results(i).chargeid<>"") then %><%= boardqna.results(i).chargeid %><% end if %>
	    </td>
	    <td>
			<% if  boardqna.results(i).FtargetMakerID<>"" and Not IsNull(boardqna.results(i).FtargetMakerID) then %>

				<% if  boardqna.results(i).Fupchereplydate<>"" and Not IsNull(boardqna.results(i).Fupchereplydate) then %>
					<% if boardqna.results(i).replyDate<>"" and not(isnull(boardqna.results(i).replyDate)) then %>
					답변완료<br />
					<% else %>
					<b>답변완료</b><br />
					<% end if %>
					<%= Left(boardqna.results(i).Fupchereplydate,10) %><br />
					<%= mid(boardqna.results(i).Fupchereplydate,11,18) %>
				<% elseif  boardqna.results(i).Fupcheviewdate<>"" and Not IsNull(boardqna.results(i).Fupcheviewdate) then %>
					업체확인중<br />
					<%= Left(boardqna.results(i).Fupcheviewdate,10) %><br />
					<%= mid(boardqna.results(i).Fupcheviewdate,11,18) %>
				<%  else %>
					<%= boardqna.results(i).FtargetMakerID %><br />
				<% end if %>
			<% end if %>
	    </td>
	    <td>
	    	<% if boardqna.results(i).replyuser<>"" and not(isnull(boardqna.results(i).replyuser)) then %>
				완료(<%= boardqna.results(i).replyuser %>)

				<% if boardqna.results(i).replyDate<>"" and not(isnull(boardqna.results(i).replyDate)) then %>
					<br>
					<%= Left(boardqna.results(i).replyDate,10) %>
					<br><%= mid(boardqna.results(i).replyDate,11,18) %>
				<% end if %>
			<% end if %>
	    </td>
	    <td>
	    	<%= boardqna.results(i).FsendYN %>
	    </td>
	    <td>
	    	<% if (boardqna.results(i).dispyn="N") then %>
			<font color="red">삭제</font>
			<% elseif (boardqna.results(i).Fsitename <> "10x10") and (boardqna.results(i).Fsitename <> "") and (boardqna.results(i).replyuser = "" or isnull(boardqna.results(i).replyuser)) then %>
				<%
					If (session("ssAdminPOsn") = "4") OR (session("ssAdminPOsn") = "5") Then
				 %>
					<a href="javascript:jsDelQna(<%= boardqna.results(i).id %>)">X</a>
				<%
					Else
						response.write "권한없음"
					End If
				%>
			<% end if %>
	    </td>
	</tr>
	<% next %>

    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% sbDisplayPaging "page="&page, boardqna.FTotalCount, boardqna.FPageSize, 10%>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center">검색결과가 없습니다.</td>
	</tr>
<% end if %>
</table>

<form method="post" name="delform" action="cscenter_qna_board_act.asp" onsubmit="return false">
<input type="hidden" name="id" value="">
<input type="hidden" name="mode" value="del">
</form>

<%
Set boardqna = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
