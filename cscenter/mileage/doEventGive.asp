<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<%

dim mode
dim userid, mileage, jukyo, jukyoCD
dim userid2, itemid, eventid, give_reason, itemoption, itemea
dim i, buf
'dim strSQL

mode = requestCheckvar(request("mode"),16)
userid = requestCheckvar(request("userid"),32)
mileage = requestCheckvar(request("mileage"),10)
jukyoCD = requestCheckvar(request("jukyoCD"),10)
jukyo = requestCheckvar(request("jukyo"),128)

userid2 = requestCheckvar(request("userid2"),32)
itemid = requestCheckvar(request("itemid"),10)
eventid = requestCheckvar(request("eventid"),10)
itemoption = requestCheckvar(request("itemoption"),4)
itemea = requestCheckvar(request("itemea"),1)
give_reason = requestCheckvar(request("give_reason"),128)


if (Not IsNumeric(mileage)) or (mileage="") then mileage = 0

if (userid="" and mode="mileagegive") then
    response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if

if (userid2="" and mode="itemgive") then
    response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if

'==============================================================================
dim strSQL
dim regUserID, checkItem
regUserID	= session("ssBctID")
if regUserID="" then
    response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if
if regUserID="corpse2" or regUserID="tozzinet" or regUserID="kobula" then
else
    response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if

if (mode = "mileagegive") then
	'마일리지 적립요청
	strSQL = " insert into [db_user].[dbo].tbl_mileagelog(userid,mileage,jukyocd,jukyo,regUserID)" & vbCrlf
	strSQL = strSQL + " values(" & vbCrlf
	strSQL = strSQL + " '" & userid & "'," & vbCrlf
	strSQL = strSQL + " " & mileage & "," & vbCrlf
	strSQL = strSQL + " " & jukyoCD & "," & vbCrlf
	strSQL = strSQL + " '" & jukyo & "'," & vbCrlf
	strSQL = strSQL + " '" & regUserID & "'" & vbCrlf
	strSQL = strSQL + " )"
	dbget.Execute strSQL
	'마일리지 재계산
	strSQL = "exec db_user.[dbo].[sp_Ten_ReCalcu_His_BonusMileage] '"& userid &"'"
	dbget.Execute strSQL

	response.write "<script>alert('적립 되었습니다.');</script>"
    response.write "<script>history.back();</script>"
elseif (mode = "itemgive") then

	'// CS 메모 저장
	strSQL = " select top 1 itemid "
	strSQL = strSQL + " from [db_my10x10].[dbo].[tbl_my_baguni] with(nolock)"
	strSQL = strSQL + " where userKey='" & userid2 & "'"
	strSQL = strSQL + " and itemid=" & itemid
	strSQL = strSQL + " and itemoption='" & itemoption & "'"
	rsget.CursorLocation = adUseClient
	rsget.Open strSQL,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		checkItem = rsget("itemid")
	end if
	rsget.close

	if (checkItem<>"") then
		response.write "<script>alert('이미 지급 되었습니다.'); history.back();</script>"
		dbget.close()	:	response.End
	end if

	'장바구니 입력
	strSQL = " insert into [db_my10x10].[dbo].[tbl_my_baguni](userKey,isLoginUser,itemid,itemoption,itemea,regdate,chkOrder)" & vbCrlf
	strSQL = strSQL + " values(" & vbCrlf
	strSQL = strSQL + " '" & userid2 & "'," & vbCrlf
    strSQL = strSQL + " 'Y'," & vbCrlf
	strSQL = strSQL + " " & itemid & "," & vbCrlf
	strSQL = strSQL + " '" & itemoption & "'," & vbCrlf
	strSQL = strSQL + " " & itemea & "," & vbCrlf
	strSQL = strSQL + " getdate(),'N'" & vbCrlf
	strSQL = strSQL + " )"
	dbget.Execute strSQL
	'로그 저장
	strSQL = " insert into [db_event].[dbo].[tbl_event_itemGiveLog](userid,itemid,itemoption,itemea,regdate,reguser)" & vbCrlf
	strSQL = strSQL + " values(" & vbCrlf
	strSQL = strSQL + " '" & userid2 & "'," & vbCrlf
	strSQL = strSQL + " " & itemid & "," & vbCrlf
	strSQL = strSQL + " '" & itemoption & "'," & vbCrlf
	strSQL = strSQL + " " & itemea & "," & vbCrlf
	strSQL = strSQL + " getdate()," & vbCrlf
	strSQL = strSQL + " '" & regUserID & "'" & vbCrlf
	strSQL = strSQL + " )"
	dbget.Execute strSQL

	response.write "<script>alert('지급 되었습니다.');</script>"
	response.write "<script>history.back();</script>"
else
	'
end if

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
