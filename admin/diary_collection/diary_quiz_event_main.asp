<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db2open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
		dim sql,strHtml
		sql =	" select top 20 eventidx,eventname from [db_cts].[dbo].[tbl_2007_diary_event_master]" &_
					" order by eventidx desc"
		db2_rsget.open sql, db2_dbget, 1

		if not db2_rsget.eof then




				do until db2_rsget.eof
					strHtml = strHtml + "<tr bgcolor='#FFFFFF'>"
					strHtml = strHtml + "	<td align='center'>" & db2_rsget("eventidx") & "</td> "
					strHtml = strHtml + "	<td align='center'>" &_
															" <span onclick=" & chr(34) & "TnEditEvent('" & db2_rsget("eventidx") & "')" & chr(34) & " style='cursor:pointer'>" & db2_rsget("eventName") & "</span>" &_
															"	<span onclick=" & chr(34) & "TnDelEvent('" & db2_rsget("eventidx") & "')" & chr(34) & " style='cursor:pointer'>[삭제]</span></td> "
					strHtml = strHtml + "	<td align='center'><span onclick=" & chr(34) & "TnGoentrant('" & db2_rsget("eventidx") & "')" & chr(34) & " style='cursor:pointer'>보기</span></td>"
					'strHtml = strHtml + "	<td align='center'><span onclick=" & chr(34) & "TnGoWinnerList('" & db2_rsget("eventidx") & "')" & chr(34) & " style='cursor:pointer'>보기</span></td> "
					strHtml = strHtml + "</tr>"

				db2_rsget.movenext
				loop

		end if

		db2_rsget.close
%>

<script language="javascript" type="text/javascript">
//응모자 보기
function TnGoentrant(idx){
	document.location.href="diary_quiz_event_winnerList.asp?eventidx=" + idx;

}
//신규 등록
function popNewReg(){
	window.open('diary_quiz_event_Pop_new.asp?mode=add','newpop','width=500,height=500,scrollbars=yes,resizable=yes');
}
//이벤트 수정
function TnEditEvent(idx){
	window.open('diary_quiz_event_Pop_new.asp?mode=edit&eventidx=' + idx,'editpop','width=500,height=500,scrollbars=yes,resizable=yes');
}
//이벤트 삭제
function TnDelEvent(idx){
	var ret = confirm('삭제 하시겠습니까.?');
	if (ret) {
		document.delFrm.mode.value='del';
		document.delFrm.eventidx.value= idx ;
		document.delFrm.submit();
	}
}

</script>
<table width="600" border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC" class="a">
	<tr>
		<td colspan="3" align="right" bgcolor="#EDEDED"><span style="cursor:pointer" onclick="popNewReg()">신규등록</span>&nbsp;&nbsp;</td>
	</tr>
	<tr bgcolor="#EDEDED">
		<td width="70" align="center">번호</td>
		<td align="center">제목</td>
		<td width="100" align="center">응모자 보기</td>
		<!--<td width="100" align="center">당첨자 보기</td>  -->
	</tr>

	<%= strHtml %>

</table>

<form name="delFrm" method="post" action="http://imgstatic.10x10.co.kr/linkweb/doDiary_quiz_event.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="del">
<input type="hidden" name="eventidx" value="">
</form>
<table width="600" border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC" class="a">
	<tr>
		<td><a href="/admin/sitemaster/diary_collection_2007/diary_eval_event_main.asp">다이어리 상품후기이벤트 당첨자 등록</a></td>
	</tr>
	<tr>
		<td><a href="/admin/sitemaster/diary_collection_2007/diary_eval_event_winList.asp">다이어리 상품후기이벤트 당첨자 보기</a></td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db2close.asp" -->