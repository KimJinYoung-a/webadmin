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
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/myqnacls.asp" -->
<%
dim i, j ,reffrom ,boardqna ,orderinfo ,myqnalist , ocheck , id
	reffrom = request("reffrom")
	id = request("id")

'//권한체크
set ocheck = new CMyQNA_list
	ocheck.frectssBctId = session("ssBctId")
	ocheck.fmembercheck()
	
set boardqna = New CMyQNA_list
	boardqna.frectid = id
	boardqna.fqnaread()

if boardqna.foneitem.fuserid <> "" then
	set orderinfo = New CMyQNAOrderInfo
	'orderinfo.UserOrderInfo (boardqna.foneitem.fuserid)
	'orderinfo.UserMinusOrderInfo (boardqna.foneitem.fuserid)
end if

if boardqna.foneitem.fuserid <> "" or boardqna.foneitem.forderserial <> "" then
	set myqnalist = New CMyQNA
	
	if boardqna.foneitem.fuserid <> "" then
	    myqnalist.SearchUserID = boardqna.foneitem.fuserid
	end if
	if boardqna.foneitem.forderserial <> "" then
	    myqnalist.SearchOrderSerial = boardqna.foneitem.forderserial
	end if
	
    myqnalist.PageSize = 100
    myqnalist.CurrPage = 1
	myqnalist.RectQadiv = 16	'오프라인 고정    
   
	if ocheck.foneitem.getmemberdisp = false then
		myqnalist.frectshopid = ocheck.FOneItem.fssBctId
	end if    
	
    myqnalist.list
end if

%>

<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script>

function SubmitForm()
{
        if (document.frm.replytitle.value == "") {
                alert("제목을 입력하세요.");
                return;
        }
        if (document.frm.replycontents.value == "") {
                alert("내용을 입력하세요.");
                return;
        }

        if (confirm("입력이 정확합니까?") == true) { document.frm.submit(); }
}

//질문 유형수정
function updateqadiv(){
	if (confirm("질문 유형을 수정하시겠습니까?")){
		updateform.mode.value="CHG";
		updateform.submit();
	}
}

//해당매장지정
function updateshopid(){
	if (confirm("문의 내용을 해당 매장으로 지정 하시겠습니까?")){
		if (updateform.shopid.value==''){
			alert('매장을 선택해 주세요');
			return;
		}
		
		updateform.mode.value="chshopid";
		updateform.submit();
	}
}

function delqadiv(){
	if (confirm("삭제하시겠습니까?")){
		document.delform.submit();
	}
}

document.title = "1:1 상담리스트";
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<img src="/images/icon_star.gif" align="absbottom">
	    <font color="red"><strong>1:1 상담 답변</strong></font>
	</td>
</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">             
<form method=post name="updateform" action="online_cscenter_qna_act.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="id" value="<% = boardqna.foneitem.fid %>">
<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="15">
		<img src="/images/icon_arrow_down.gif" align="absbottom">
	    <font color="red"><b>문의내용</b></font>
	    <% 
		if ocheck.foneitem.getmemberdisp = true then
	    %>
		    &nbsp;&nbsp;
		    질문유형수정 :
		    <select name="qadiv" class="select">
		    <option>선택</option>
		        <option value="00" <% if boardqna.foneitem.fqadiv = "00" then response.write "selected" %>>배송문의</option>
		        <option value="01" <% if boardqna.foneitem.fqadiv = "01" then response.write "selected" %>>주문문의</option>
		        <option value="02" <% if boardqna.foneitem.fqadiv = "02" then response.write "selected" %>>상품문의</option>
		        <option value="03" <% if boardqna.foneitem.fqadiv = "03" then response.write "selected" %>>재고문의</option>
		        <option value="04" <% if boardqna.foneitem.fqadiv = "04" then response.write "selected" %>>취소문의</option>
		        <option value="05" <% if boardqna.foneitem.fqadiv = "05" then response.write "selected" %>>환불문의</option>
		        <option value="06" <% if boardqna.foneitem.fqadiv = "06" then response.write "selected" %>>교환문의</option>
		        <option value="07" <% if boardqna.foneitem.fqadiv = "07" then response.write "selected" %>>AS문의</option>    
		        <option value="08" <% if boardqna.foneitem.fqadiv = "08" then response.write "selected" %>>이벤트문의</option>
		        <option value="09" <% if boardqna.foneitem.fqadiv = "09" then response.write "selected" %>>증빙서류문의</option>    
		        <option value="10" <% if boardqna.foneitem.fqadiv = "10" then response.write "selected" %>>시스템문의</option>
		        <option value="11" <% if boardqna.foneitem.fqadiv = "11" then response.write "selected" %>>회원제도문의</option>
		        <option value="12" <% if boardqna.foneitem.fqadiv = "12" then response.write "selected" %>>회원정보문의</option>
		        <option value="13" <% if boardqna.foneitem.fqadiv = "13" then response.write "selected" %>>당첨문의</option>
		        <option value="14" <% if boardqna.foneitem.fqadiv = "14" then response.write "selected" %>>반품문의</option>
		        <option value="15" <% if boardqna.foneitem.fqadiv = "15" then response.write "selected" %>>입금문의</option>
		        <option value="16" <% if boardqna.foneitem.fqadiv = "16" then response.write "selected" %>>오프라인문의</option>
		        <option value="17" <% if boardqna.foneitem.fqadiv = "17" then response.write "selected" %>>쿠폰/마일리지문의</option>
		        <option value="18" <% if boardqna.foneitem.fqadiv = "18" then response.write "selected" %>>결제방법문의</option>
		        <option value="20" <% if boardqna.foneitem.fqadiv = "20" then response.write "selected" %>>기타문의</option>
	            <option value="21" <% if boardqna.foneitem.fqadiv = "21" then response.write "selected" %>>아이띵소문의</option>
	            <option value="23" <% if boardqna.foneitem.fqadiv = "23" then response.write "selected" %>>사은품문의</option>
	            <option value="24" <% if boardqna.foneitem.fqadiv = "24" then response.write "selected" %>>POINT1010문의</option>
		    </select>
		    <input type="button" class="button" value="저장" onclick="updateqadiv();">	    
	    	해당매장지정 : <% drawSelectBoxOffShop "shopid" , boardqna.foneitem.fshopid %>
	    	<input type="button" class="button" value="저장" onclick="updateshopid();">
		<% end if %>
	</td>
</tr>
</form>

<form method="post" name="frm" action="online_cscenter_qna_act.asp" onsubmit="return false">
<!--
<%' if boardqna.foneitem.freplyuser<>"" then %>
<input type="hidden" name="mode" value="reply">
<%' else %>
<input type="hidden" name="mode" value="firstreply">
<%' end if %>
-->
<input type="hidden" name="mode" value="REP">
<input type="hidden" name="id" value="<%= boardqna.foneitem.fid %>">
<input type="hidden" name="username" value="<%= boardqna.foneitem.fusername %>">
<input type="hidden" name="regdate" value="<%= boardqna.foneitem.fregdate %>">
<input type="hidden" name="title" value="<%= boardqna.foneitem.ftitle %>">
<input type="hidden" name="contents" value='<%= replace(html2db(boardqna.foneitem.fcontents),"'","") %>'> <!-- -.- -->
<input type="hidden" name="replydate" value="<%= boardqna.foneitem.freplydate %>"> 
<input type="hidden" name="email" value="<%= boardqna.foneitem.fusermail %>">
<input type="hidden" name="emailok" value="<%= boardqna.foneitem.femailok %>">
<input type="hidden" name="extsitename" value="<%= boardqna.foneitem.fFextsitename %>">
<input type="hidden" name="replyuser" value="<%= session("ssBctID") %>">
<input type="hidden" name="imsitxt">
<tr>
	<td width="80" align="center" bgcolor="#FFFFFF"><b>작성자</b></td>
	<td bgcolor="#FFFFFF">
	    <font color="#464646"><%= boardqna.foneitem.fusername %>(<%= boardqna.foneitem.fuserid %>/<%= boardqna.foneitem.forderserial %>)</font>
	    &nbsp;&nbsp;
	    [ <b><%= getUserLevelStrByDate(boardqna.foneitem.fUserLevel, left(boardqna.foneitem.fregdate,10)) %></b> ]
	    <%
	    	if boardqna.foneitem.fFrealnamecheck="Y" then
	    		Response.Write " / 실명확인회원"
	    	end if
	    %>
	</td>
	<td align="center" bgcolor="#FFFFFF"><b>문의주문번호</b></td>
	<td width="160" bgcolor="#FFFFFF">
	    <% if boardqna.foneitem.forderserial<>"" then %>
    	    <a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= boardqna.foneitem.forderserial %>');"><%= boardqna.foneitem.forderserial %> >>상세보기</a>
        <% end if %>
	</td>
</tr>
<tr height="25">
	<td align="center" bgcolor="#FFFFFF"><b>작성일시</b></td>
	<td bgcolor="#FFFFFF"><font color="#464646"><%= boardqna.foneitem.fregdate %></font></td>
	<td align="center" bgcolor="#FFFFFF"><b>문의상품</b></td>
	<td bgcolor="#FFFFFF">
	    <%= boardqna.foneitem.fitemid %>
	    &nbsp;&nbsp;
	    <% if boardqna.foneitem.fitemid<>"" and boardqna.foneitem.fitemid>0 then %>
	        <a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= boardqna.foneitem.fitemid %>" target="_blank">>>상품보기</a>
	    <% end if %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#FFFFFF"><b>답변 예상일</b></td>
	<td colspan="3" bgcolor="#FFFFFF" height="25"><font color="#464646"><%= boardqna.foneitem.fFExpectReplyDate %></font></td>
</tr>
<tr>
	<td align="center" bgcolor="#FFFFFF"><b>문의제목</b></td>
	<td colspan="3" bgcolor="#FFFFFF" height="25"><font color="#464646"><%= nl2br(db2html(boardqna.foneitem.ftitle)) %></font></td>
</tr>
<tr>
	<td align="center" bgcolor="#FFFFFF"><b>문의내용</b></td>
	<td colspan="3" bgcolor="#FFFFFF" height="25"><font color="#464646"><%= nl2br(db2html(boardqna.foneitem.fcontents)) %></font></td>
</tr>
<tr height="25" valign="top" bgcolor="<%= adminColor("tabletop") %>">
    <td colspan="4" valign="middle">
        <img src="/images/icon_arrow_down.gif" align="absbottom">
        <font color="red"><b>답변작성</b></font>
    </td>
</tr>
 <% if boardqna.foneitem.freplyuser<>"" then %>
<tr>
    <td align="center" bgcolor="#FFFFFF">답변제목</td>
	<td colspan="3" bgcolor="#FFFFFF"><input type="text" class="text" name="replytitle" size="65" value="<%= boardqna.foneitem.freplytitle %>"></td>
</tr>
<tr>
    <td align="center" bgcolor="#FFFFFF">답변내용</td>
	<td colspan="3" bgcolor="#FFFFFF"><textarea class="textarea" name="replycontents" cols="100" rows="10"><%= db2html(boardqna.foneitem.freplycontents) %></textarea></td>
</tr>
<% Else %>
<tr>
    <td align="center" bgcolor="#FFFFFF">답변제목</td>
	<td colspan="3" bgcolor="#FFFFFF">
		  <input type="text" class="text" name="replytitle" size="65">&nbsp;
		  <% SelectBoxQnaPreface "01" %>&nbsp;
		  <% SelectBoxQnaCompliment "" %>
	</td>
</tr>
<tr>
    <td align="center" bgcolor="#FFFFFF">답변내용</td>
	<td colspan="3" bgcolor="#FFFFFF"><textarea class="textarea" name="replycontents" cols="100" rows="10"></textarea></td>
</tr>
<% End If %>

<tr>
	<td colspan="15" align="center" bgcolor="#FFFFFF">
	    <input type="button" class="button" value=" 답변저장 " onclick="SubmitForm()">
	    <input type="button" class="button" value=" 목록으로 " onclick="location.href='online_cscenter_qna_list.asp';">
	</td>
</tr>
</form>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<font color="red"><strong>예전 상담 목록</strong></font>
	</td>
</tr>
            
<% if boardqna.foneitem.fuserid <> "" or boardqna.foneitem.forderserial <> "" then %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>레벨</td>
    <td>주문번호</td>
    <td width="50">상품</td>
    <td>제목</td>
    <td>구분</td>
    <td>작성일</td>
    <td>답변여부</td>
    <td>답변자</td>
    <td>답변일</td>
    <td>비고</td>
</tr>
	<% if myqnalist.ResultCount < 0 then %>
	
	<% else %>
		<% for i = 0 to (myqnalist.ResultCount - 1) %>
		<tr align="center" <% if (myqnalist.results(i).id <> CLng(id)) then %>bgcolor="#FFFFFF"<% else %> class="tr_select" bgcolor="#AFEEEE"<% end if %>>
		    <td><b><%= getUserLevelStrByDate(myqnalist.results(i).fUserLevel, left(myqnalist.results(i).regdate,10)) %></b></td>
		    <td><%= myqnalist.results(i).orderserial %></td>
		    <td><%= myqnalist.results(i).itemid %></td>
		    <td><a href="online_cscenter_qna_reply.asp?id=<%= myqnalist.results(i).id %>&reffrom=<%= reffrom %>"><%= myqnalist.results(i).title %></a></td>
		    <td>
		    	<a href="online_cscenter_qna_reply.asp?id=<%= myqnalist.results(i).id %>&reffrom=<%= reffrom %>">
		    	<%= myqnalist.code2name(myqnalist.results(i).qadiv) %>
		    	<% if myqnalist.results(i).fshopid = "" or isnull(myqnalist.results(i).fshopid) then %>
		    		(매장지정안됨)
		    	<% else %>
		    		(<%= myqnalist.results(i).fshopid %>)
		    	<% end if %>
		    	</a>
		    </td>
		    <td><%= FormatDate(myqnalist.results(i).regdate, "0000-00-00") %></td>
		    <td><% if (myqnalist.results(i).replyuser<>"") then %>답변완료<% end if %></td>
		    <td>
		    	<% if (myqnalist.results(i).replyuser<>"") then %>
		    		<%= myqnalist.results(i).replyuser %>		    	
		    	<% end if %>
		    </td>
		    <td><acronym title="<%= myqnalist.results(i).replydate %>"><%= Left(myqnalist.results(i).replydate,10) %></acronym></td>
		    <td><% if (myqnalist.results(i).dispyn="N") then %><font color="red">삭제</font><% end if %></td>
		</tr>
		<% next %>
	<% end if %>
<% end if %>    
</table>

<form method="post" name="delform" action="online_cscenter_qna_act.asp" onsubmit="return false">
	<input type="hidden" name="id" value="<%= boardqna.foneitem.fid %>">
	<input type="hidden" name="mode" value="del">
</form>

<iframe name="PrefaceFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>

<script language="JavaScript">

function TnChangePreface(SelectGubun){
	PrefaceFrame.location.href="/cscenter/board/preface_select.asp?gubun=" + SelectGubun + "&userid=<%= boardqna.foneitem.fuserid %>&masterid=01";
}

function TnChangeText(str){
	var basictext;
	
	basictext = "안녕하세요. <%= boardqna.foneitem.fuserid %>님\n"
	basictext = basictext + "텐바이텐 <%= session("ssBctCname") %>입니다.\n"
	basictext = basictext + "(내용)\n"
	basictext = basictext + "만족스런답변이 되셨는지요\n\n"
	
	if(str == ''){
		document.frm.replycontents.value = basictext;
	}
	else{
		//오프라인 매장 전용 매뉴이기 때문에고객행복센터 라는 말은 제낀다
		document.frm.replycontents.value = str.replace("고객행복센터","");
	}
}

<% if boardqna.foneitem.freplyuser <> "" then %>
	
<% else %>
	TnChangeText('');
<% end if %>

</script>

<iframe name="ComplimentFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script language="JavaScript">

 function TnChangeCompliment(SelectGubun){
	ComplimentFrame.location.href="/cscenter/board/compliment_select.asp?masterid=01&gubun=" + SelectGubun;
 }

 function TnChangeText2(str){

	if(str == ''){
	}
	else{
		document.frm.replycontents.value = document.frm.imsitxt.value + "\n" + str;
	}
 }

</script>

<%
set myqnalist = Nothing
set boardqna = Nothing
set orderinfo = Nothing
set ocheck = nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->