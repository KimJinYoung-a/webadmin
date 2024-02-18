<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  팝업 관리
' History : 2011.01.28 김진영 생성
'			2018.08.10 한용민 수정
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/partpersonCls.asp"-->
<%

dim board, isCsCenter

board = requestCheckvar(request("board"),1)

%>

<script language="javascript">

<% if (board = "U") then %>

function workerselect(userid, username)
{
	opener.focus();
	opener.document.frm.workername.value = username;
	opener.document.frm.workerid.value = userid;
	window.close();
}

<% end if %>

</script>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr height="10" valign="bottom" bgcolor="F4F4F4">
    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_02.gif"></td>
    <td background="/images/tbl_blue_round_02.gif"></td>
    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="20" valign="bottom" bgcolor="F4F4F4">
    <td background="/images/tbl_blue_round_04.gif"></td>
    <td valign="top" bgcolor="F4F4F4" align="center"><b>텐바이텐 파트별 담당자 연락처</b></td>
    <td valign="top" align="right" bgcolor="F4F4F4"></td>
    <td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr height="15" valign="bottom" bgcolor="F4F4F4">
    <td background="/images/tbl_blue_round_04.gif"></td>
    <td valign="top" bgcolor="F4F4F4"></td>
    <td valign="top" align="right" bgcolor="F4F4F4"></td>
    <td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" align="center" cellpadding="2" cellspacing="1" bgcolor=#bababa class="a">
<%
Dim clist,clist2, arlist, arlist2, arlist3, i, j, gubun, sabun, idx

'Partlist 클래스 생성
Set clist = new Partlist
	'Partlist 클래스의 구분이 Y일 경우엔 모든 리스트 보이고 미사용시 리스트에서 제외
	clist.FGubun = "Y"
	arlist = clist.fnGetlist

For i = 0 to Ubound(arlist,2)
	clist.idx = arlist(0,i)
%>
	<tr bgcolor="#FFDDDD" height="25">
		<td colspan=4><b><%= i+1 %>.<%= arlist(1,i) %></b></td>
	</tr>
<%
	arlist2 = clist.fnGetmolist2
	If IsArray(arlist2) = "True" Then
		For j = 0 to Ubound(arlist2,2)
            isCsCenter = (arlist2(1,j) = 220)
%>
	<tr bgcolor="#FFFFFF" height="22" <% if (board = "U") then %>style="cursor:pointer" onClick="workerselect('<%= arlist2(9,j)%>', '<%= arlist2(3,j)%>')"<% end if %> >
    	<td><%= arlist2(2,j)%></td>
    	<td>
			<% if isCsCenter then %>
			고객센터
			<% else %>
			<%= arlist2(3,j)%>
			<% end if %>
		</td>
    	<td>
    		<%
    		'/cs팀 분기. 직통전화로 안받고. 공통 번호로 받는다고함.
    		if isCsCenter then
    		%>
    			070-4868-1799 (고객주문관련문의)
    		<% else %>
    			<%= arlist2(4,j)%>&nbsp;#<%= arlist2(5,j)%>
    		<% end if %>
    	</td>
    	<td>
    		<%
    		if isCsCenter then
    		%>
    			<a href="mailto:customer@10x10.co.kr">customer@10x10.co.kr</a>
    		<% else %>
    			<a href="mailto:<%= arlist2(6,j)%>"><%= arlist2(6,j)%></a>
    		<% end if %>
		</td>
    </tr>
		<% Next %>
	<% elseif (arlist(1,i) = "고객행복센터") then %>
	<tr bgcolor="#FFFFFF" height="22">
    	<td>고객센터</td>
    	<td>고객센터</td>
    	<td>070-4868-1799</td>
    	<td></td>
    </tr>
	<%End If%>
<% Next %>
<% Set clist = nothing %>
</table>
<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr valign="top" bgcolor="F4F4F4" height="30">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left" bgcolor="F4F4F4">
    	<b>* 연락처</b> <br>&nbsp;&nbsp; 본사 : 02-554-2033 &nbsp;&nbsp; 물류센터 : 1644-1851</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" bgcolor="F4F4F4" height="30">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left" bgcolor="F4F4F4">
    	<b>* 팩스번호</b> <br>&nbsp;&nbsp; 대학로 : 02-2179-9244 (MD파트), 02-2179-9245(마케팅), 02-2179-9058(오프라인)
    	<br>&nbsp;&nbsp; 물류센터 : 02-3493-1032</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" bgcolor="F4F4F4" height="30">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left" bgcolor="F4F4F4" >
    	<b>* 주소</b> <br>
    	&nbsp;&nbsp; 대학로 : (03082) 서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐<br>

		&nbsp;&nbsp; 물류센터 : 경기도 포천시 군내면 용정경제로2길 83 텐바이텐 물류센터
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" bgcolor="F4F4F4" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
