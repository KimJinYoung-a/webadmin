<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이벤트 당첨등록
' History : 2010.03.22 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<%
Dim evt_code 
	evt_code		= requestCheckVar(Request("evt_code"),10) 	
%>
<script language="javascript">
	
	//-- jsPopCal : 달력 팝업 --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
	
	
	function jsWinnerSubmit(frm,v){
			
		if(!frm.evt_ranking.value){
			alert("등수를 입력해주세요");
			frm.evt_ranking.focus();
			return;
		}
		
		if(!IsDigit(frm.evt_ranking.value)){
			alert("등수는 숫자만 입력가능합니다.");
			frm.evt_ranking.focus();
			return;
		}
		
		if(v == 'userseq'){
			if(!frm.evt_winner_seq.value){
				alert("당첨자를 입력해주세요");
				frm.evt_winner_seq.focus();
				return;
			}
		}else{
			if(!frm.evt_winner_user.value){
				alert("당첨자를 입력해주세요");
				frm.evt_winner_user.focus();
				return;			
			}		
		}
						
		if(confirm("등록하신 내용은 수정 또는 삭제가 불가능하며 고객에게 바로 적용됩니다.\n\n등록 하시겠습니까? ")){
			frm.smode.value = v;
			frm.submit();
		}
	}
	    
    //사은품 종류 등록
	function jsSetGiftKind(){
		var winkind;
		winkind = window.open('/admin/offshop/gift/popgiftKindReg.asp?giftkind_name='+document.frmWin.giftkind_name.value,'popkind','width=600, height=600;');
		winkind.focus();
	}

</script>

<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 당첨자 등록</div>
<table width="100%" border=0 align="center" class="a" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmWin" method="post" action="eventprize_process.asp">
<input type="hidden" name="evt_code" value="<%=evt_code%>">
<input type="hidden" name="mode" value="prize_add">
<input type="hidden" name="smode">
<tr>
	<td width="100" align="center" bgcolor="FFFFFF" colspan=3>기본정보</td>
</tr>
<tr>
	<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">구분</td>
	<td bgcolor="#FFFFFF">
		<%sbGetOptCommonCodeArr_off "evtprize_type", "", False,True,""%>
	</td>
</tr>		
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">등수</td>
	<td bgcolor="#FFFFFF"><input type="text" size="2" name="evt_ranking"> 없으면 0</td>
</tr>	
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">등수별칭<br>(티켓명등등)</td>
	<td bgcolor="#FFFFFF"><input type="text" name="evt_rankname" size="20" maxlength="32"></td>
</tr>	
<% if evt_code = "" then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">당첨확인기간</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="evtprize_startdate" value="<%= left(now(),10) %>"  size="10" maxlength="10" onClick="jsPopCal('evtprize_startdate');" style="cursor:hand;">
		~<input type="text" name="evtprize_enddate" size="10"  maxlength="10" value="<%=dateadd("d",14,date())%>" onClick="jsPopCal('evtprize_enddate');" style="cursor:hand;">
	</td>
</tr>	
<% end if %>
</table>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">당첨자<br>(오프라인고객번호)</td>
	<td bgcolor="#FFFFFF">
		콤머로 구분, 공백없이 ( EX: 30,111 )<br>
		<textarea name="evt_winner_seq" rows="2" cols="60"></textarea>
	</td>
</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right" height="40">
		<input type="button" onclick="jsWinnerSubmit(frmWin,'userseq');" value="오프라인고객번호로 저장" class="button">		
	</td>
</tr>
</table>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">당첨자<br>(이름으로 저장)</td>
	<td bgcolor="#FFFFFF">
		콤머로 구분, 공백없이 ( EX: 서동석,한용민 )<br>
		<textarea name="evt_winner_user" rows="2" cols="60"></textarea>
	</td>
</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right" height="40">
		<input type="button" onclick="jsWinnerSubmit(frmWin,'username');" value="이름으로 저장" class="button">	
	</td>
</tr>
</form>
</table>

<!-- #include virtual="/lib/db/dbclose.asp" -->


			

	
				