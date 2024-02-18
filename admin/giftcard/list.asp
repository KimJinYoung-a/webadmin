<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/giftcard/giftcard_cls.asp"-->
<%
Dim oGiftCard, page, i
page	= request("page")
If page = "" Then page = 1

Set oGiftCard = new cGiftCard
	oGiftCard.FCurrPage			= page
	oGiftCard.FPageSize			= 50
	oGiftCard.getGiftCardList
%>
<script language="javascript">
function regGift(){
	location.href = "/admin/giftcard/reg.asp?menupos=<%= menupos %>";
}

//페이지 이동
function goPage(v) {
	document.getElementById("page").value = v;
	document.frm.submit();
}
function goView(v){
	location.href= "reg.asp?menupos=<%=menupos%>&idx="+v;
}
function sendCard(idx, eappidx, makecnt, optcode){
	var frm = document.tfrm;
	if(confirm("[ 품의서 IDX "+ eappidx + " ]\n\n카드 "+makecnt+"장을 발급 하시겠습니까?")){
		frm.idx.value = idx;
		frm.opt.value = optcode;
		frm.submit();
	}
}
</script>
<form name="frm" method="get" action="">
<input type="hidden" id="page" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
</form>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="7">
		검색결과 : <b><%= FormatNumber(oGiftCard.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oGiftCard.FTotalPage,0) %></b>
	</td>
	<td align="right">
		<input type="button" class="button" value="등록하기" onClick="regGift();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" >
	<td width="10%">품의서IDX</td>
	<td width="30%">제목</td>
	<td width="7%">옵션</td>
	<td width="5%">발급수량</td>
	<td width="12%">등록자</td>
	<td width="13%">등록일</td>
	<td width="10%">발급여부</td>
	<td width="13%">발급일</td>
</tr>
<% For i=0 to oGiftCard.FResultCount - 1 %>
<tr align="center" bgcolor= "#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td><%= oGiftCard.FItemList(i).FEappIdx %></td>
	<td onclick="goView(<%= oGiftCard.FItemList(i).FIdx %>);" style="cursor:pointer;">
		<%= oGiftCard.FItemList(i).FReqTitle %>
	</td>
	<td><%= oGiftCard.FItemList(i).getCardOptName %></td>
	<td><%= oGiftCard.FItemList(i).FMakeCnt %></td>
	<td><%= oGiftCard.FItemList(i).FRegUserId %></td>
	<td><%= LEFT(oGiftCard.FItemList(i).FRegdate, 10) %></td>
	<% If oGiftCard.FItemList(i).FIsSend = "Y" Then %>
		<td><font color="red">완료</font></td>
	<% Else %>
		<td><input type="button" value="발급" class="button" onclick="sendCard('<%= oGiftCard.FItemList(i).FIdx %>', '<%= oGiftCard.FItemList(i).FEappIdx %>', '<%= oGiftCard.FItemList(i).FMakeCnt %>', '<%= oGiftCard.FItemList(i).FOpt %>');"></td>
	<% End If %>
	<td><%= LEFT(oGiftCard.FItemList(i).FIsSendDate, 10) %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oGiftCard.HasPreScroll then %>
		<a href="javascript:goPage('<%= oGiftCard.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oGiftCard.StartScrollPage to oGiftCard.FScrollCount + oGiftCard.StartScrollPage - 1 %>
    		<% if i>oGiftCard.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oGiftCard.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<form name="tfrm" action="/admin/giftcard/giftcardProc.asp" method="post">
<input type="hidden" name="idx">
<input type="hidden" name="opt">
<input type="hidden" name="mode" value="S">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->