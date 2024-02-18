<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->


<body topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>

<!-- 상단메뉴 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr height="10" valign="bottom" bgcolor="F4F4F4">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
    </tr>
    <tr height="25" valign="bottom" bgcolor="F4F4F4">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="center">
        	<a href="/cscenter/history/history_memo.asp"><b>HISTORY</b></a>
        	&nbsp | &nbsp
			<a href="/m_item_search.asp">CS등록건</a>
			&nbsp | &nbsp
			<a href="/m_baljulist.asp">마일지지/쿠폰</a>
		</td>
    	<td align="right"></td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 상단메뉴 -->

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td>

            <table width="100%" border=0 cellspacing=0 cellpadding=2 class=a bgcolor="FFFFFF">
                
                <tr align="center" bgcolor="F3F3FF">
                    <td width="30">구분</td>
                	<td width="30">idx</td>
                 	<td width="60">고객ID</td>   	
                	<td width="80">주문번호</td>
                	<td>내용</td>
                    <td width="50">등록자</td>
                	<td width="70">등록일</td>
                	<td width="30">완료</td>
                </tr>
                <tr>
                    <td height="1" colspan="15" bgcolor="#CCCCCC"></td>
                </tr>
                <tr align="center" bgcolor="FFFFFF">
                    <td>요청</td>
                	<td>00000</td>
                 	<td>coolhas</td>   	
                	<td>06033011111</td>
                	<td align="left">업체와 통화 후 연락주기로 함</td>
                    <td>iroo4</td>
                	<td>2006-03-30</td>
                	<td>N</td>
                </tr>
                <tr>
                    <td height="1" colspan="15" bgcolor="#CCCCCC"></td>
                </tr>
            </table>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>



<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->