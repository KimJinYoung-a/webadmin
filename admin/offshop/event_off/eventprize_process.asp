<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이벤트 당첨등록
' History : 2010.03.24 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/email/mailLib.asp"-->
<!-- #include virtual="/lib/email/smsLib.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_Cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->

<%
dim mode , smode ,evt_code , evtprize_type , evt_ranking ,evt_rankname ,evtprize_startdate
dim evtprize_enddate , evt_winner_seq , evt_winner_user , evt_winner , j ,giftkind_code
dim give_evtprizecode , tmpCode , strSql
	mode 		= requestCheckVar(Request.Form("mode"),32)
	smode 		= requestCheckVar(Request.Form("smode"),32)
	evt_code  		= requestCheckVar(Request.Form("evt_code"),10)
	evtprize_type 		= requestCheckVar(Request.Form("evtprize_type"),10)
	evt_ranking 		= requestCheckVar(Request.Form("evt_ranking"),10)
	evt_rankname  		= requestCheckVar(Request.Form("evt_rankname"),32)
	evtprize_startdate 		= requestCheckVar(Request.Form("evtprize_startdate"),30)
	evtprize_enddate 		= requestCheckVar(Request.Form("evtprize_enddate"),30)
	evt_winner_seq  		= requestCheckVar(Request.Form("evt_winner_seq"),32)
	evt_winner_user  		= requestCheckVar(Request.Form("evt_winner_user"),32)

if smode = "username" then
	evt_winner = evt_winner_user
elseif smode = "userseq" then
	evt_winner = evt_winner_seq
end if	
evt_winner = split(evt_winner,",")

dim cEvtCont , evt_kind , evt_name
set cEvtCont = new cevent_list
	cEvtCont.frectevt_code = evt_code	'이벤트 코드	
		
	'이벤트 내용 가져오기	
	cEvtCont.fnGetEventCont_off
	evt_kind = cEvtCont.FOneItem.fevt_kind
	evt_name = cEvtCont.FOneItem.fevt_name
set cEvtCont = nothing

function event_gubun(v)
	if v = "username" then
		event_gubun = "evt_winner_name"
	else
		event_gubun = "evt_winner"
	end if		
end function

'//이벤트관리 등록
Function fnSetEventPrize_off(ByVal sType, ByVal eCode, ByVal evt_ranking,ByVal evt_rankname,ByVal iGiftKindCode,ByVal evt_winner,ByVal dAStartDate,ByVal dAEndDate, ByVal AdminID, ByVal iGiveEPCode,ByVal stitle , smode)			
	Dim iprizestatus : iprizestatus = 0
	IF 	(dAEndDate = "" OR sType="1") THEN iprizestatus = 3 '당첨확인기간 미 입력시 확인상태로
	IF iGiveEPCode = "" THEN iGiveEPCode = "NULL"
	
		strSql = "INSERT INTO [db_shop].[dbo].[tbl_event_prize_off] (evtprize_type, [evt_code],[evt_ranking] " &_
				" , [evt_rankname], giftkind_code, "&event_gubun(smode)&",  [evtprize_startdate], [evtprize_enddate], [evtprize_status]" &_
				" ,[AdminID],[give_evtprizecode],evtprize_name) values ("&_ 
				" "&sType&", "&eCode&", "&evt_ranking&",'"&html2db(evt_rankname)&"','"&iGiftKindCode&"', '"&evt_winner&"'"&_
				" , '"&dAStartDate &"','"&dAEndDate&"', "&iprizestatus&", '"& AdminID&"',"&iGiveEPCode&",'"&html2db(stitle)&"')"
		
	'response.write strSql &"<br>"		
	dbget.execute strSql	
	
	strSql = ""
	strSql = "update db_shop.dbo.tbl_event_off set"+vbcrlf
	strSql = strSql & " prizeyn = 'Y'"+vbcrlf
	strSql = strSql & " where evt_code = "&eCode&""+vbcrlf

	'response.write strSql&"<br>"			
	dbget.execute strSql				
	
	IF Err.Number <> 0 THEN
		dbget.RollBackTrans	  	
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "") 
	END IF
	
	strSql = "select @@IDENTITY "
	 
	rsget.Open strSql, dbget
	tmpCode = rsget(0)
	rsget.Close
End Function

'트랜잭션 
dbget.beginTrans

For j = 0 To UBound(evt_winner)	
    
	tmpCode = ""	
	'1. 이벤트관리 등록	
	fnSetEventPrize_off evtprize_type,evt_code,evt_ranking,evt_rankname,giftkind_code,Trim(evt_winner(j)),evtprize_startdate,evtprize_enddate,session("ssBctId"),give_evtprizecode,evt_name&"당첨" , smode
	
Next
	 	
dbget.CommitTrans
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

<script language="javascript">
	alert("OK");
	opener.location.reload();
	window.close();
</script>