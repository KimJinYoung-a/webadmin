<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/eventWinner/event_Sms_Process.asp
' Description :  이벤트 당첨자 SMS 발송 처리
' History : 2007.10.01 김정인
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/eventWinner_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventWinnerManageCls.asp"-->

<%

'// SMS 저장
Function SaveSmsLog(byval eCd ,byval sCont ,byval rpNo,byval rgUser,byval SendYn)

	dim oSms,arrSms
	set oSms = new ClsEventEntry
	oSms.FECode = evtCode
	arrSms = oSms.fnGetSms

	set oSms = nothing

	dim fnSQL

	If isArray(arrSms) Then
		fnSQL =" UPDATE db_event.dbo.tbl_event_sms_log " &_
				" set evt_code ='" & eCd & "' " &_
				" , SmsCont = '" & sCont & "' " &_
				" , replyNumber='" & rpNo & "' " &_
				" , regUser='" & rgUser & "' " &_
				" , regDate=getdate() " &_
				" , isSended ='" & SendYn & "'" &_
				" WHERE evt_code ='" & evtCode & "'"
	Else
		fnSQL =" INSERT INTO db_event.dbo.tbl_event_sms_log " &_
				" (evt_code,SmsCont,replyNumber,regUser,regDate,isSended) " &_
				" values " &_
				" ( '" & eCd & "' ,'" & sCont & "','" & rpNo & "','" & rgUser & "' ,getdate(),'" & SendYn & "' " &_
				" ) "

	End If

	dbget.execute(fnSQL)
End Function

'################-- 처리 Process 시작 --###################

dim mode

mode=request("mode")

dim evtCode,SmsCont,replyNumber,regUser,regDate,arridx

evtCode= request("eC")
SmsCont = request("msg")
replyNumber = request("reNo")
regUser = session("ssBctId")
arridx = chkarray(request("arridx"))

dim strSQL,msg,loopcnt


dbget.begintrans

if mode="save" then
	SaveSmsLog evtCode,SmsCont,replyNumber,regUser,"N"
	msg= "저장 되었습니다"
'// SMS 발송
elseIf mode="send" Then
	SaveSmsLog evtCode,SmsCont,replyNumber,regUser,"Y"

	fnWinnerSmsSended evtCode,arridx

	strSQL = strSQL &_
			" INSERT INTO [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) " &_
			" SELECT distinct usercell,'" & replyNumber & "','1',getdate(),'" & db2html(SmsCont) & "'" &_
			" FROM db_user.[dbo].tbl_user_n " &_
			" WHERE userid in (" & arridx & ")"
	msg= "문자가 발송 되었습니다"
	dbget.execute(strSQL)
End If


'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)

		response.write	"<script language='javascript'>"
		response.write	" alert('" & msg & "');"
		response.write	"</script>"
		dbget.close()	:	response.End
	Else
		dbget.RollBackTrans				'롤백(에러발생시)

		response.write	"<script language='javascript'>" &_
					"	alert('처리중 에러가 발생했습니다.');" &_
					"</script>"


	End If
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->