<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 온라인 1:1 게시판 문의 보기
' Hieditor : 2010.01.03 한용민 온라인 이동 수정/생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/myqnacls.asp" -->
<%
dim itemqanotinclude, research, finishyn, userid, orderserial, qadiv , writeid ,i, j ,sDate ,eDate ,blnDate
Dim ckReplyDate, replyDate1, replyDate2 ,page ,boardqna ,ocheck ,shopflg
	qadiv               = request("qadiv")
	itemqanotinclude    = request("itemqanotinclude")
	research            = request("research")
	userid              = request("userid")
	orderserial         = request("orderserial")
	shopflg             = request("shopflg")
	writeid             = request("writeid")
	page	= req("page",1)
	finishYN = req("finishYN","")
	'if (itemqanotinclude="") and (research="") then itemqanotinclude="on"
	
	qadiv = "16"	'오프라인문의만 고정
	
	sDate = request("sdt")
	eDate = request("edt")
	blnDate = request("edc")

	ckReplyDate	= req("ckReplyDate",req("ckReplyDateDefault",""))
	replyDate1	= req("replyDate1",LEFT(CStr(dateAdd("d",-7,now())),10))
	replyDate2	= req("replyDate2",LEFT(CStr(now()),10))
	
	if (blnDate="") and (research="") then 
	    blnDate = "on"
	    sDate   = LEFT(CStr(dateAdd("m",-3,now())),10)
	    eDate   = LEFT(CStr(now()),10)
	end if

'//권한체크
set ocheck = new CMyQNA_list
	ocheck.frectssBctId = session("ssBctId")
	ocheck.fmembercheck()
	
set boardqna = New CMyQNA
	boardqna.FPageSize = 50
	boardqna.FCurrPage = page
	boardqna.RectQadiv = qadiv
	boardqna.FSearchUserID = userid
	boardqna.FSearchOrderSerial = orderserial
	boardqna.FSearchWriteId = writeId
	
	IF blnDate="on" Then
		boardqna.FSearchStartDate = sDate
		boardqna.FSearchEndDate =eDate
	End IF
	
	IF ckReplyDate="on" Then
		boardqna.FreplyDate1 = replyDate1
		boardqna.FreplyDate2 =replyDate2
	End IF
	
	boardqna.FRectItemNotInclude = itemqanotinclude
	
	''boardqna.list finishYN
	
	''old ver
	if (finishyn = "N") then
	    boardqna.SearchNew = "Y"
	end if
	
	if ocheck.foneitem.getmemberdisp = false then	
	    shopflg = "Y" ''지정된 매장만 보게 - 미지정 매장 볼 수 없음.// 2012/06/18 eastone
		if ocheck.FOneItem.getmemberofficedisp = true then
			boardqna.frectshopid = ocheck.FOneItem.fbigo
		else
			boardqna.frectshopid = ocheck.FOneItem.fssBctId
		end if
	end if
	
	boardqna.frectshopflg = shopflg
	boardqna.list
%>

<script language='javascript'>

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


function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);
	
	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

function EnableDate(obj)
{
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

function replyEnableDate(obj)
{
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

document.title = "1:1 상담리스트";

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form method="get" name="qnaform">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		아이디 : <input type="text" class="text" name="userid" value="<%= userid %>" size="15" onKeyPress="if (event.keyCode == 13) document.qnaform.submit();">
  		&nbsp;
  		주문번호 : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" size="15" onKeyPress="if (event.keyCode == 13) document.qnaform.submit();">
  		&nbsp;
  		답변아이디 : <input type="text" class="text" name="writeid" value="<%= writeid %>" size="15" onKeyPress="if (event.keyCode == 13) document.qnaform.submit();">
  		&nbsp;
  		질문유형 :
		<select class="select" name="qadiv">
            <option value="">전체</option>
            <option value="00" <% if qadiv = "00" then response.write "selected" %>>배송문의</option>
            <option value="01" <% if qadiv = "01" then response.write "selected" %>>주문문의</option>
            <option value="02" <% if qadiv = "02" then response.write "selected" %>>상품문의</option>
            <option value="03" <% if qadiv = "03" then response.write "selected" %>>재고문의</option>
            <option value="04" <% if qadiv = "04" then response.write "selected" %>>취소문의</option>
            <option value="05" <% if qadiv = "05" then response.write "selected" %>>환불문의</option>
            <option value="06" <% if qadiv = "06" then response.write "selected" %>>교환문의</option>
            <option value="07" <% if qadiv = "07" then response.write "selected" %>>AS문의</option>
            <option value="08" <% if qadiv = "08" then response.write "selected" %>>이벤트문의</option>
            <option value="09" <% if qadiv = "09" then response.write "selected" %>>증빙서류문의</option>
            <option value="10" <% if qadiv = "10" then response.write "selected" %>>시스템문의</option>
            <option value="11" <% if qadiv = "11" then response.write "selected" %>>회원제도문의</option>
            <option value="12" <% if qadiv = "12" then response.write "selected" %>>회원정보문의</option>
            <option value="13" <% if qadiv = "13" then response.write "selected" %>>당첨문의</option>
            <option value="14" <% if qadiv = "14" then response.write "selected" %>>반품문의</option>
            <option value="15" <% if qadiv = "15" then response.write "selected" %>>입금문의</option>
            <option value="16" <% if qadiv = "16" then response.write "selected" %>>오프라인문의</option>
            <option value="17" <% if qadiv = "17" then response.write "selected" %>>쿠폰/마일리지문의</option>
            <option value="18" <% if qadiv = "18" then response.write "selected" %>>결제방법문의</option>
            <option value="20" <% if qadiv = "20" then response.write "selected" %>>기타문의</option>
            <option value="21" <% if qadiv = "21" then response.write "selected" %>>아이띵소문의</option>
            <option value="23" <% if qadiv = "23" then response.write "selected" %>>사은품문의</option>
            <option value="24" <% if qadiv = "24" then response.write "selected" %>>POINT1010문의</option>
        </select>
        <br>
        <input type="checkbox" name="edc" <%IF blnDate="on" then response.write "checked" %> onclick="EnableDate(this);">
        고객작성일 : <input type="text" size="10" name="sdt" value="<%= sDate %>" onClick="jsPopCal('qnaform','sdt');" <% IF blnDate="on" Then%>class="text" <%Else%>readonly class="text_ro"<%END IF%> style="cursor:hand;">
        ~<input type="text" size="10" name="edt" value="<%= eDate %>" onClick="jsPopCal('qnaform','edt');" <% IF blnDate="on" Then%>class="text" <%Else%>readonly class="text_ro"<%END IF%> style="cursor:hand;">
        <input type="checkbox" name="ckReplyDate" <%IF ckReplyDate="on" then response.write "checked" %> onclick="replyEnableDate(this);">
        답변일 : <input type="text" size="10" name="replyDate1" value="<%= replyDate1 %>" onClick="jsPopCal('qnaform','replyDate1');" <% IF ckReplyDate="on" Then%>class="text" <%Else%>readonly class="text_ro"<%END IF%> style="cursor:hand;">
        ~<input type="text" size="10" name="replyDate2" value="<%= replyDate2 %>" onClick="jsPopCal('qnaform','replyDate2');" <% IF ckReplyDate="on" Then%>class="text" <%Else%>readonly class="text_ro"<%END IF%> style="cursor:hand;">
        매장지정:
        <select name="shopflg">
        	<option value="" <% if shopflg = "" then response.write " selected" %>>선택</option>
        	<option value="Y" <% if shopflg = "Y" then response.write " selected" %>>매장지정완료</option>
        	<option value="N" <% if shopflg = "N" then response.write " selected" %>>매장미지정</option>
        </select>
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="SubmitSearch()">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="radio" name="finishyn" value="" <% if finishyn = "" then response.write "checked" %>> 전체
    	<input type="radio" name="finishyn" value="N" <% if finishyn = "N" then response.write "checked" %>> 미처리
	</td>
</tr>
</form>
</table>

<br>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="20" style="padding:3 0 3 5">검색결과 : <b><%=boardqna.ResultCount%></b> / <%=boardqna.TotalCount%></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>레벨</td>
    <td>고객명(아이디)</td>
    <td>주문번호</td>
    <td>문의상품</td>
    <td>제목</td>
    <td>구분</td>
    <td>작성일</td>
    <td>답변여부</td>
    <td>답변자</td>
    <td>비고</td>
</tr>
<% for i = 0 to (boardqna.ResultCount - 1) %>

<% if (boardqna.results(i).dispyn = "N") then %>
<tr align="center" bgcolor="#EEEEEE">
<% else %>
<tr align="center" bgcolor="#FFFFFF">
<% end if %>
	<td><b><%= getUserLevelStrByDate(boardqna.results(i).fUserLevel, Left(boardqna.results(i).regdate, 10)) %></b></td>
    <td>
    	<%= boardqna.results(i).username %>
    	<!--<a href="javascript:SubmitSearchUserId('<%'= boardqna.results(i).userid %>');">-->
    	(<%= printUserId(boardqna.results(i).userid, 2, "*") %>)
    	<!--</a>-->
    	</td>
    <td><%= boardqna.results(i).orderserial %></td>
    <td><%= boardqna.results(i).itemid %></td>
    <td align="left"><%= db2html(boardqna.results(i).title) %></td>
    <td>
    	<%= boardqna.code2name(boardqna.results(i).qadiv) %>
    	<% if boardqna.results(i).fshopid = "" or isnull(boardqna.results(i).fshopid) then %>
    		(매장지정안됨)
    	<% else %>
    		(<%= boardqna.results(i).fshopid %>)
    	<% end if %>
    </td>
    <td align="center">
    	<% if (Left(boardqna.results(i).regdate, 10) < Left(now, 10)) then %>
      	<%= Left(boardqna.results(i).regdate,10) %>
    	<% else %>
      	금일 <%= Right(FormatDate(boardqna.results(i).regdate, "0000.00.00 00:00:00"), 8) %>
    	<% end if %>
    </td>
    <td>
    	<% if (boardqna.results(i).replyuser<>"") then %>답변완료<% end if %>
    </td>
    <td>
    	<% if (boardqna.results(i).replyuser<>"") then %><%= boardqna.results(i).replyuser %><% end if %>
    </td>

    <td>
    	<input type="button" value="수정" class="button" onclick="location.href='online_cscenter_qna_reply.asp?id=<%= boardqna.results(i).id %>';">
    	<% if (boardqna.results(i).dispyn="N") then %><font color="red">삭제</font><% end if %>
    </td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan=20>	
		<div align="center">
			<% sbDisplayPaging "page="&page, boardqna.FTotalCount, boardqna.FPageSize, 10%>
		</div>
	</td>
</tr>
</table>

<%
Set boardqna = Nothing
set ocheck = nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
