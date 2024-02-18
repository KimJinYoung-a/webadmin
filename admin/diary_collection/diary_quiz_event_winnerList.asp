<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db2open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim eventidx
eventidx= request("eventidx")

if eventidx="" then
response.write "<h1>잘못된 접근입니다</h1>"
dbget.close()	:	response.End
end if

		dim SQL,strHtml,FResultCount

		SQL = " select win_idx,userid" &_
					" from [db_cts].[dbo].[tbl_2007_diary_event_winner] " &_
					" where  eventidx=" & eventidx &_
					" order by userid "


		db2_rsget.open SQL,db2_dbget,1

				if not db2_rsget.eof then
							FResultCount = db2_rsget.recordCount
				end if

				if not db2_rsget.eof then

						do until db2_rsget.eof
							strHtml = strHtml + "<tr bgcolor='#FFFFFF'>"
							strHtml = strHtml + "	<td align='center'>" & db2_rsget("userid") & "</td> "
							strHtml = strHtml + "	<td align='center'><span style='cursor:pointer; color:blue;' onclick=" & chr(34) & "FnDelWinner('" & db2_rsget("win_idx") & "');" & chr(34) &">[del]</span></td> "
							strHtml = strHtml + "</tr>"

						db2_rsget.movenext
						loop

				end if

		db2_rsget.close

%>

<table width="900" border="0" cellpadding="0" cellspacing="0" class="a">
<form name="winnerFrm" method="post" action="/admin/sitemaster/diary_collection_2007/doDiary_quiz_event_winner.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="eventidx" value="<%= eventidx %>">
<input type="hidden" name="win_idx" value="">
	<tr>
		<td bgcolor="#EDEDED">
			<font color="red">notice:</font><br>
			상단에 "정답자만 보기 박스"를 체크하신후 "정답입력" 란에 정답을 작성하신후 검색하세요.<br>
			아이디를 클릭하시면 우측 박스에 아이디가 임시 저장됩니다.<br>
			페이지를 넘기셔두 저장된값은 유지됩니다.<br>
			아이디 선정후 우측하단에 저장 버튼을 누르시면 당첨자로 선정 됩니다.
			</td>
	</tr>
	<tr>
		<td width="700"><iframe src="/admin/sitemaster/diary_collection_2007/diary_quiz_event_entryList.asp?eventidx=<%= eventidx %>" width="700" height="700" frameborder="0"></iframe></td>
		<td width="200" align="left" valign="top">
			<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td>
						<table width="150" border="0" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC" class="a">
							<%= strHtml %>
							<tr>
								<td colspan="2">당첨자수 : 총 <font color="red"><%= FresultCount %></font>명</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<textarea name="winnerList" value="" cols="20" rows="20"></textarea>
						<br>
						<input type="button" value="저장"  onclick="FnsaveWinner();" />
					</td>
				</tr>
			</table>
		</td>

	</tr>
</form>
</table>

<script language="javascript" type="text/javascript">
function FnsaveWinner(){
	var frm = document.winnerFrm;

	if (frm.winnerList.value.length<1){
		alert('당첨자를 선택해 주세요	');
		return false;
	}

	var conf = confirm('선정하시겠습니까?\n' + frm.winnerList.value);

	if (conf) {
		frm.mode.value="write"
		frm.submit();
	}
}

function FnDelWinner(idx){
	var frm = document.winnerFrm;
	var conf = confirm("삭제 하시겠습니까?");

	if (conf){
		frm.mode.value="del"
		frm.win_idx.value=idx
		frm.submit();
	}


}
</script>





<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db2close.asp" -->