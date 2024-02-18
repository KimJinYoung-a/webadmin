<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim itax_no, ierrmsg
itax_no = request("itax_no")
ierrmsg = request("ierrmsg")
%>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
  <tr valign="bottom">
    <td width="10" height="10" align="right" valign="bottom" bgcolor="#F3F3FF"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
    <td height="10" valign="bottom" background="/images/tbl_blue_round_02.gif" bgcolor="#F3F3FF"></td>
    <td width="10" height="10" align="left" valign="bottom" bgcolor="#F3F3FF"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
  </tr>
  <tr valign="top">
    <td height="20" background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
    <td height="20" background="/images/tbl_blue_round_06.gif" bgcolor="#F3F3FF"><img src="/images/icon_star.gif" align="absbottom">
    <font color="red"><strong>ERROR</strong></font></td>
    <td height="20" background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
  </tr>
  <tr valign="top">
    <td background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
    <td bgcolor="#F3F3FF">
    <% if Trim(itax_no)="-2" then %>
		<br>네오포트에 등록된 사업자번호와 어드민에 등록된 사업자명이 다르거나,
		<br>회원가입이 되어 있지 않습니다.
		<br>하단 메세지 참조.
	<% elseif Trim(itax_no)="-3" then %>
		<br><b>네오포트 홈페이지에서 이용료 결제후 사용하세요 (건당 200원)</b>
		<br>하단 메세지 참조.
	<% else %>
		<br>하단 메세지 참조.
	<% end if %>
	<br>
	<b>ErrCode : [<%= itax_no %>] ErrMsg : [<%= ierrmsg %>]</b>
    </td>
    <td background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
  </tr>
  <tr>
    <td background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
    <td>
    	<br>
    	오류코드 정리<br>
    	-2, Sender is not valid Neoport user<br>
		: 네오포트 회원 가입이 안되신 경우 입니다. 네오포트에서 회원 가입후 사용하세요.<br>
		<a href="http://www.neoport.net" target="_blank"><font color="blue">>>네오포트 회원가입하기</font></a><br>
		<br>
		-3, 잔여건수없음 or No Remainder<br>
		: 세금계산서 발행시 건당 200원의 금액이 과금됩니다.<br>
		네오포트 사이트 오른쪽에 보시면 [서비스/제품 구매] 라고 버튼이 있으며<br>
		이곳에 들어가셔서 종량상품,또는 정액상품을 구매하신 후 사용하시면 됩니다.
    </td>
    <td background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
  </tr>
  <tr valign="top">
    <td background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
    <td bgcolor="#F3F3FF" align="right"><a href="javascript:history.back();"><strong>&lt;&lt;뒤로가기</strong></a></td>
    <td background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
  </tr>
  <tr valign="top" bgcolor="#F3F3FF">
    <td height="10" bgcolor="#F3F3FF"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td height="10" background="/images/tbl_blue_round_08.gif"></td>
    <td height="10"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
  </tr>
</table>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
