<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 1:1 상담
' History : 2009.04.17 이상구 생성
'			2019.05.16 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/board/myqnacls.asp" -->
<%
dim i, userid, orderserial, page
    page	= req("page",1)
    userid = request("userid")
    orderserial = request("orderserial")


dim boardqna
set boardqna = New CMyQNA
    boardqna.FPageSize = 10
    boardqna.FCurrPage = page
    boardqna.FSearchUserID = userid
    boardqna.FSearchOrderSerial = orderserial
    boardqna.list2	' 목록 가져오기
%>
<script type="text/javascript">

function popQnaView(idx)
{
	var url = "/cscenter/board/cscenter_qna_board_reply.asp?id=" + idx;
	var popwin = window.open(url,"PopMyQnaList","width=1024, height=768, left=50, top=50, scrollbars=yes, resizable=yes, status=yes");
	popwin.focus();
}
</script>
<link rel="stylesheet" href="/cscenter/css/cs.css" type="text/css">
<style>
body {
    background-color: #FFFFFF;
}

.listSep {
	border-top:0px #CCCCCC solid; height:1px; margin:0; padding:0;
}
</style>
<table width="100%" border="0" cellspacing="0" cellpadding="2" class="a" bgcolor="FFFFFF">
<tr height="20" align="center" bgcolor="F3F3FF">
    <td width="">구분</td>
    <td width="80">주문번호</td>
    <td width="60">상품번호</td>
    <td>제목</td>
    <td width="30">첨부</td>
    <td width="80">접수일</td>
    <td width="60">답변예정자</td>
    <td width="100">답변여부</td>
    <td width="30">삭제</td>
</tr>
<tr>
    <td class="listSep" colspan="15" bgcolor="#CCCCCC" style="border-top:1px"></td>
</tr>
<% for i = 0 to (boardqna.ResultCount - 1) %>
<% if (boardqna.results(i).dispyn = "N") then %>
    <tr align="center" bgcolor="#EEEEEE">
<% else %>
    <tr align="center" bgcolor="#FFFFFF">
<% end if %>

    <td><%= boardqna.code2name(boardqna.results(i).qadiv) %></td>
    <td><%= boardqna.results(i).orderserial %></td>
    <td><%= boardqna.results(i).itemid %></td>
    <td align="left">&nbsp;&nbsp;<a href="javascript:popQnaView('<%= boardqna.results(i).id %>');"><%= db2html(boardqna.results(i).title) %></a></td>
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
    <td><%= boardqna.results(i).chargeid %></td>
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
        <% if (boardqna.results(i).dispyn="N") then %><font color="red">삭제</font><% end if %>
    </td>
</tr>
<tr>
    <td class="listSep" colspan="15" bgcolor="#CCCCCC"></td>
</tr>
<% next %>
<% If boardqna.ResultCount = 0 Then  %>
    <tr align="center" bgcolor="#FFFFFF">
        <td colspan="15">검색결과가 없습니다.</td>
    </tr>
<% end if %>
</table>
<div align="center">
    <% sbDisplayPaging "page="&page, boardqna.FTotalCount, boardqna.FPageSize, 10%>
</div>

<%
Set boardqna = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
