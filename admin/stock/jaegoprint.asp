<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  재고파악리스트출력
' History : 2007년 7월 13일 한용민 생성
' 			2007년 12월 4일 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/jaegostock.asp"-->

<%
dim idx, fidx, sql , i 				'변수선언
	idx = request("idx")				'인덱스값을 받아온다.
	fidx = Left(idx,Len(idx)-1)			'받아온 인덱스 값을 하나씩빼서 나열
	 
dim oip1 						'클래스선언
	set oip1 = new Cfitemlist		'변수에 토탈을 넣구
	oip1.Frectidx = fidx
	oip1.fprintlist()				'클래스실행 
%>


<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<input type="button" value="시트출력하기" onclick="javascript:window.print();">
        	<font color="red"><strong>옵션이 업는 상품의 경우 "0000"으로 표기됩니다.</strong></font>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">이미지</td>
		<td width="40">상품<br>코드</td>
		<td width="100">브랜드ID</td>
		<td>상품명[옵션]</td>
		<td width="40">현재<br>재고</td>
		<td width="40">재고<br>파악</td>
		<td width="80">작업상태<br>작업지시일시</td>
	</tr>
<% 
dim sql3
if oip1.FTotalCount > 0 then 		'레코드 수가 0보다 크면 
%>	 
	<% for i=0 to oip1.FTotalCount - 1 %>
		<form name="frmBuyPrc<%=i%>" method="get">			<!--for문 안에서 i 값을 가지고 루프-->
		<input type="hidden" name="mode">
		<tr align="center" bgcolor="#FFFFFF">
			<td><img src="<%= oip1.flist(i).fsmallimage %>" width=50 height=50><input type="hidden" name="smallimage" value="<%= oip1.flist(i).fsmallimage %>"></td>	<!--'이미지 -->
			<td><%= oip1.flist(i).fitemid %><input type="hidden" name="fitemid" value="<%= oip1.flist(i).fitemid %>"></td>				 					<!--'상품번호	 -->
			<td><%= oip1.flist(i).fmakerid %><input type="hidden" name="fmakerid" value="<%= oip1.flist(i).fmakerid %>"></td>									 <!--'브랜드id -->
		<!--상품명시작 -->
			<td align="left">
				<%= oip1.flist(i).fitemname %><input type="hidden" name="fitemname" value="<%= oip1.flist(i).fitemname %>">
				<br>
				<font color="blue">
				<%= oip1.flist(i).fitemoptionname %><input type="hidden" name="itemoptionname" value="<%= oip1.flist(i).fitemoptionname %>">
				</font>
			</td>				
		<!--상품명끝 -->									
			<td><%= oip1.flist(i).frealstock %><input type="hidden" name="frealstock" value="<%= oip1.flist(i).frealstock %>"></td>									 <!--'재고파악용재고 -->
			<td></td>															
			<td>
				<!--재고파악란시작-->
					<% if oip1.flist(i).fstatecd = 1 then %>
						 작업지시
					<% elseif oip1.flist(i).fstatecd = 5 then %>
						 재고파악완료
					<% elseif oip1.flist(i).fstatecd = 7 then %>
						 완료(반영됨)
					<% elseif oip1.flist(i).fstatecd = 8 then %>
						 완료(미반영)
					<% end if %>
				<!--재고파악란끝-->
			</td>						
		</tr>
		</form>	
		

	<% next %>
	
<% else %>
	<tr bgcolor="#FFFFFF">
	<td colspan=11 align=center>[ 검색결과가 없습니다. ]</td>
	</tr>
<% end if %>
</table>

<%
set oip1 = nothing
%>	
<!-- #include virtual="/lib/db/dbclose.asp" -->
<script language="javascript">
opener.location.reload();
</script>
