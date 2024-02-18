<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/myqnacls.asp" -->
<%

dim i, j, page, rectuserid, qadiv

rectuserid = request("rectuserid")
page = request("page")
qadiv = request("qadiv")
if page="" then page=1


'==============================================================================
'나의 1:1질문답변
dim boardqna
set boardqna = New CMyQNA

boardqna.PageSize = 20
boardqna.CurrPage = page
boardqna.ScrollCount = 10
boardqna.RectQadiv = qadiv
boardqna.SearchUserID = rectuserid


boardqna.list

%>

<script language='javascript'>
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
</script>



<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
		    <img src="/images/icon_star.gif" align="absbottom">
		    <font color="red"><strong>1:1 상담리스트(전체리스트)</strong></font>
		</td>
		<td align="right" background="/images/tbl_blue_round_06.gif">
		    <a href="cscenter_qna_board_list.asp">
		    <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">미처리리스트
		    </a>
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td></td>
		<td></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmSrc" method="get" action="">
    <input type="hidden" name="page" value="<% = page %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			아이디 : <input type="text" class="text" name="rectuserid" value="<%= rectuserid %>">&nbsp;&nbsp;
        	질문유형선택 :
    	    <select class="select" name="qadiv">
                <option value="">선택</option>
                <option value="00" <% if qadiv = "00" then response.write "selected" %>>배송문의</option>
                <option value="01" <% if qadiv = "01" then response.write "selected" %>>주문문의</option>
                <option value="02" <% if qadiv = "02" then response.write "selected" %>>상품문의</option>
                <option value="03" <% if qadiv = "03" then response.write "selected" %>>재고문의</option>
                <option value="04" <% if qadiv = "04" then response.write "selected" %>>취소,환불문의</option>
                <option value="06" <% if qadiv = "06" then response.write "selected" %>>교환문의</option>
                <option value="08" <% if qadiv = "08" then response.write "selected" %>>사은품문의</option>
                <option value="10" <% if qadiv = "10" then response.write "selected" %>>시스템문의</option>
                <option value="12" <% if qadiv = "12" then response.write "selected" %>>개인정보관련</option>
                <option value="13" <% if qadiv = "13" then response.write "selected" %>>당첨문의</option>
                <option value="14" <% if qadiv = "14" then response.write "selected" %>>반품문의</option>
                <option value="15" <% if qadiv = "15" then response.write "selected" %>>입금문의</option>
                <option value="16" <% if qadiv = "16" then response.write "selected" %>>오프라인문의</option>
                <option value="20" <% if qadiv = "20" then response.write "selected" %>>기타문의</option>
            </select>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="document.frmSrc.submit()">
		</td>
	</tr>
	</form>
</table>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<form name="frmSrc" method="get" action="">
    <input type="hidden" name="page" value="<% = page %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top">
        	아이디 : <input type="text" name="rectuserid" value="<%= rectuserid %>">&nbsp;&nbsp;
        	질문유형선택 :
        	    <select name="qadiv">
                    <option value="">선택</option>
                    <option value="00" <% if qadiv = "00" then response.write "selected" %>>배송문의</option>
                    <option value="01" <% if qadiv = "01" then response.write "selected" %>>주문문의</option>
                    <option value="02" <% if qadiv = "02" then response.write "selected" %>>상품문의</option>
                    <option value="03" <% if qadiv = "03" then response.write "selected" %>>재고문의</option>
                    <option value="04" <% if qadiv = "04" then response.write "selected" %>>취소,환불문의</option>
                    <option value="06" <% if qadiv = "06" then response.write "selected" %>>교환문의</option>
                    <option value="08" <% if qadiv = "08" then response.write "selected" %>>사은품문의</option>
                    <option value="10" <% if qadiv = "10" then response.write "selected" %>>시스템문의</option>
                    <option value="12" <% if qadiv = "12" then response.write "selected" %>>개인정보관련</option>
                    <option value="13" <% if qadiv = "13" then response.write "selected" %>>당첨문의</option>
                    <option value="14" <% if qadiv = "14" then response.write "selected" %>>반품문의</option>
                    <option value="15" <% if qadiv = "15" then response.write "selected" %>>입금문의</option>
                    <option value="16" <% if qadiv = "16" then response.write "selected" %>>오프라인문의</option>
                    <option value="20" <% if qadiv = "20" then response.write "selected" %>>기타문의</option>
                </select>
            &nbsp;
            <input type="button" value="검색" onclick="document.frmSrc.submit()">
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr bgcolor="#DDDDFF">
        <td width="200" align="center">고객명(아이디/주문번호)</td>
        <td align="center">제목</td>
        <td width="100" align="center">구분</td>
        <td width="70" align="center">답변유무</td>
        <td width="50" align="center">Site</td>
        <td width="160" align="center">작성일</td>
    </tr>
    <% for i = 0 to (boardqna.ResultCount - 1) %>
    
    <% if (boardqna.results(i).dispyn = "N") then %>
    <tr align="center" bgcolor="#EEEEEE">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
        <td align="left"><%= boardqna.results(i).username %>(<%= boardqna.results(i).userid %>/<%= boardqna.results(i).orderserial %>)</td>
        <td align="left">&nbsp;<a href="cscenter_qna_board_reply.asp?id=<%= boardqna.results(i).id %>"><%= boardqna.results(i).title %></a></td>
        <td><%= boardqna.code2name(boardqna.results(i).qadiv) %></td>
        <td>
        	<% if (boardqna.results(i).replyuser<>"") then %>답변완료<% end if %>
        </td>
        <td>
        	<% if IsNull(boardqna.results(i).Fextsitename) then %>
        	&nbsp;
        	<% else %>
        	<%= boardqna.results(i).Fextsitename %>
        	<% end if %>
        </td>
        <td align="left"><%= boardqna.results(i).regdate %></td>
    </tr>
    <% next %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="center" align="center">
		<% if boardqna.HasPreScroll then %>
			<a href="javascript:NextPage('<%= boardqna.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + boardqna.StartScrollPage to boardqna.ScrollCount + boardqna.StartScrollPage - 1 %>
			<% if i>boardqna.Totalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if boardqna.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	    </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->