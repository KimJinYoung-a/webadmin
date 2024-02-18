<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 입금내역
' History : 서동석 생성
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/ipkumlistcls.asp"-->
<%
dim jungsanidx
	jungsanidx 		= requestCheckVar(Request("jungsanidx"),10)

dim oipkum
set oipkum = new IpkumChecklist
	oipkum.FCurrpage=1
	oipkum.FPagesize=100
	oipkum.FScrollCount = 10

	oipkum.FOrderby = "desc"

	oipkum.FRectJungsanIDX = jungsanidx

	oipkum.GetMatchedIpkumlistAccounts

dim i
dim totmatchprice

totmatchprice = 0

%>

<script language='javascript'>

function SubmitSearch(frm) {

	document.frm.submit();
}

function SubmitDelete(frm) {

	if (confirm("정말로 삭제하시겠습니까?") == true) {
		frm.submit();
	}
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="jungsanidx" value="<%= jungsanidx %>">

	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" height="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">

		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="SubmitSearch(frm)">
		</td>
	</tr>

	</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		검색결과 : <b><%= oipkum.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td>IDX</td>
	<td width="70">은행명</td>
	<td width="100">계좌번호</td>
	<td width="70">입출금일</td>
	<td>적요</td>
  	<td width="80">입금금액</td>
  	<td width="80">출금금액</td>
  	<td width="80">매칭금액</td>
  	<td>비고</td>
</tr>
<% if oipkum.FResultCount > 0 then %>
<% for i=0 to oipkum.FResultCount-1 %>
	<% totmatchprice = totmatchprice + oipkum.Fipkumitem(i).Fmatchprice %>
<form name="frmmatch<%= i %>" method="post" action="pop_ipkum_search_process.asp">
<input type="hidden" name="mode" value="delmatch">
<input type="hidden" name="jungsanidx" value="<%= jungsanidx %>">
<input type="hidden" name="inoutidx" value="<%= oipkum.Fipkumitem(i).Finoutidx %>">
<input type="hidden" name="matchdetailidx" value="<%= oipkum.Fipkumitem(i).Fmatchdetailidx %>">
<tr align="center" bgcolor="#FFFFFF" height="25">
	<td><%= oipkum.Fipkumitem(i).Finoutidx %></td>
	<td>
		<%= oipkum.Fipkumitem(i).Fbkname %>
	</td>
	<td>
		<%= oipkum.Fipkumitem(i).Fbkacctno %>
	</td>
	<td>
		<%= mid(oipkum.Fipkumitem(i).Fbkdate,1,4) %>-<%= mid(oipkum.Fipkumitem(i).Fbkdate,5,2) %>-<%= mid(oipkum.Fipkumitem(i).Fbkdate,7,2) %>
	</td>
	<td>
		<%= oipkum.Fipkumitem(i).Fbkjukyo %>
	</td>
  	<td>
		<% if oipkum.Fipkumitem(i).finout_gubun = "2" then %>
			<%= FormatNumber(oipkum.Fipkumitem(i).Fbkinput,0) %>
		<% end if %>
  	</td>
  	<td>
		<% if oipkum.Fipkumitem(i).finout_gubun = "1" then %>
			<%= FormatNumber(oipkum.Fipkumitem(i).Fbkinput,0) %>
		<% end if %>
  	</td>
  	<td>
  		<%= FormatNumber(oipkum.Fipkumitem(i).Fmatchprice,0) %>
  	</td>
	<td>
		<input type="button" class="button_s" value="삭제하기" onClick="SubmitDelete(frmmatch<%= i %>)">
	</td>
</tr>
</form>
<% next %>
<tr align="center" bgcolor="#FFFFFF" height="25">
	<td>총액</td>
	<td colspan="6"></td>
	<td>
		<%= FormatNumber(totmatchprice, 0) %>
	</td>
  	<td></td>
</tr>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="9" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>




<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->