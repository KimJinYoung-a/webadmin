<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/upchebeasongcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<script language='javascript'>
function PopWindowRequestLecture(){
	var popwin = window.open("/cscenterv2/lecture/frame.asp?menupos=1236","PopWindowRequestLecture","width=1000 height=600 left=0 top=0 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopWindowRequestDIY(){
	var popwin = window.open("/cscenterv2/order/frame.asp?menupos=1237","PopWindowRequestLecture","width=1000 height=600 left=0 top=0 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopCSMemo(sitegubun, orderserial, userid, finishyn) {
	var popwin = window.open("/academy/cs/pop_cs_memo_list.asp?sitegubun=" + sitegubun + "&orderserial=" + orderserial + "&userid=" + userid + "&finishyn=" + finishyn,"cs_memo","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
  <tr>
    <td colspan="3" width="65%">
      <!-- 강좌신청내역 검색 -->
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
      	<tr height="10" valign="bottom">
      	    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
      	    <td width="70%" background="/images/tbl_blue_round_02.gif"></td>
      	    <td width="30%" background="/images/tbl_blue_round_02.gif"></td>
      	    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
      	</tr>
      	<tr height="25">
      	    <td background="/images/tbl_blue_round_04.gif"></td>
      	    <td background="/images/tbl_blue_round_06.gif">
      	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>강좌신청내역 검색</b>
      	    </td>
      	    <td align="right" background="/images/tbl_blue_round_06.gif">
      	        <a href="javascript:PopWindowRequestLecture();"> 강좌신청내역보기 <img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
      	    </td>
      	    <td background="/images/tbl_blue_round_05.gif"></td>
      	</tr>
      	<tr valign="bottom" height="25">
      		<td background="/images/tbl_blue_round_04.gif"></td>
      	    <td></td>
      	    <td align="right"></td>
      	    <td background="/images/tbl_blue_round_05.gif"></td>
      	</tr>
      	<tr height="10" valign="top">
      		<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
      	    <td background="/images/tbl_blue_round_08.gif"></td>
      	    <td background="/images/tbl_blue_round_08.gif"></td>
      	    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
      	</tr>
      </table>
      <!-- 강좌신청내역 검색 -->
    </td>
    <td width="2%"></td>
    <td>
      <!-- 새로고침 -->
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="DDDDFF">
      	<tr height="10" valign="bottom">
      	    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
      	    <td width="70%" background="/images/tbl_blue_round_02.gif"></td>
      	    <td width="30%" background="/images/tbl_blue_round_02.gif"></td>
      	    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
      	</tr>
      	<tr height="25">
      	    <td background="/images/tbl_blue_round_04.gif"></td>
      	    <td background="/images/tbl_blue_round_06.gif">
      	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>업데이트 시간</b>
      	    </td>
      	    <td align="right" background="/images/tbl_blue_round_06.gif">
      	        <a href="javascript:document.location.reload();">
      	        새로고침
      	        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
      	        </a>
      	    </td>
      	    <td background="/images/tbl_blue_round_05.gif"></td>
      	</tr>
      	<tr valign="bottom" height="25">
      		<td background="/images/tbl_blue_round_04.gif"></td>
      	    <td><%= now %></td>
      	    <td align="right">
      	    </td>
      	    <td background="/images/tbl_blue_round_05.gif"></td>
      	</tr>
      	<tr height="10" valign="top">
      		<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
      	    <td background="/images/tbl_blue_round_08.gif"></td>
      	    <td background="/images/tbl_blue_round_08.gif"></td>
      	    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
      	</tr>
      </table>
      <!-- 새로고침 끝 -->
    </td>
  </tr>
  <tr>
    <td height="10" colspan="5"></td>
  </tr>
  <tr>
    <td colspan="3" width="65%">
      <!-- DIY주문내역 검색 -->
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
      	<tr height="10" valign="bottom">
      	    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
      	    <td width="70%" background="/images/tbl_blue_round_02.gif"></td>
      	    <td width="30%" background="/images/tbl_blue_round_02.gif"></td>
      	    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
      	</tr>
      	<tr height="25">
      	    <td background="/images/tbl_blue_round_04.gif"></td>
      	    <td background="/images/tbl_blue_round_06.gif">

      	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>DIY주문내역 검색</b>
      	    </td>
      	    <td align="right" background="/images/tbl_blue_round_06.gif">
      	        <a href="javascript:PopWindowRequestDIY();"> DIY주문내역 검색 <img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
      	    </td>
      	    <td background="/images/tbl_blue_round_05.gif"></td>
      	</tr>
      	<tr valign="bottom" height="25">
      		<td background="/images/tbl_blue_round_04.gif"></td>
      	    <td></td>
      	    <td align="right"></td>
      	    <td background="/images/tbl_blue_round_05.gif"></td>
      	</tr>
      	<tr height="10" valign="top">
      		<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
      	    <td background="/images/tbl_blue_round_08.gif"></td>
      	    <td background="/images/tbl_blue_round_08.gif"></td>
      	    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
      	</tr>
      </table>
      <!-- DIY주문내역 검색 -->
    </td>
    <td width="2%"></td>
    <td>
			<!-- CS메모 관리 -->
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
				<tr height="10" valign="bottom">
				    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
				    <td width="70%" background="/images/tbl_blue_round_02.gif"></td>
				    <td width="30%" background="/images/tbl_blue_round_02.gif"></td>
				    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
				</tr>
				<tr height="25">
				    <td background="/images/tbl_blue_round_04.gif"></td>
				    <td background="/images/tbl_blue_round_06.gif">
				    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>CS메모 관리</b>
				    </td>
				    <td align="right" background="/images/tbl_blue_round_06.gif">
				        <a href="javascript:PopCSMemo('academy','','','');">
				        바로가기
				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
				        </a>
				    </td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr valign="bottom" height="25">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>미처리메모</td>
				    <td align="right">
				    	<b>??</b> 건
				    	<a href="javascript:PopCSMemo('academy','','','N');">
            		    <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
            		    </a>
				    </td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr valign="bottom" height="25">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr valign="bottom" height="25">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="10" valign="top">
					<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
				    <td background="/images/tbl_blue_round_08.gif"></td>
				    <td background="/images/tbl_blue_round_08.gif"></td>
				    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
				</tr>
			</table>
			<!-- CS메모 관리 -->
    </td>
  </tr>
  <tr>
    <td width="33%"></td>
    <td width="33%"></td>
    <td></td>
  </tr>
  <tr>
    <td height="10" colspan="5"></td>
  </tr>
  <tr>
    <td width="33%"></td>
    <td width="33%"></td>
    <td></td>
  </tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->