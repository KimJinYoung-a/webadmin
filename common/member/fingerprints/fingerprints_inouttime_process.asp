<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 지문인식 근태관리
' Hieditor : 2011.03.22 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/member/fingerprints/fingerprints_cls.asp" -->

<%
dim mode ,idx ,empno ,inoutType ,inoutTime ,inoutTimes ,isusing ,sqlStr
	mode = requestCheckvar(request("mode"),32)
	idx = requestCheckvar(request("idx"),10)
	empno = requestCheckvar(request("empno"),32)
	inoutType = requestCheckvar(request("inoutType"),1)		
	isusing = requestCheckvar(request("isusing"),1)
	inoutTime	= requestCheckvar(request("inoutTime"),10) & " " & replace(requestCheckvar(request("inoutTimes"),8),"24:00:00","23:59:59")	

'//입력
if mode = "fingerprintsedit" then
	
	'/수정
	if idx <> "" then
	
		dbget.begintrans
	
		sqlStr =	"Update db_partner.dbo.tbl_user_inouttime_log Set " &_
				"	empno = '" & empno & "' " &_
				"	,inoutType = '" & inoutType & "' " &_
				"	,isusing = '" & isusing & "' " &_
				"	,inoutTime = '" & inoutTime & "' " &_
				"	,lasteditupdate = getdate()" &_
				"	,lastedituserid = '"&session("ssBctID")&"'"&_
				"Where idx=" & idx
		
		'response.write SQL &"<br>"
		dbget.Execute sqlStr
		
		'/yyyymmdd 처리
		sqlStr = ""
		sqlStr = "update l set" + vbcrlf
		sqlStr = sqlStr & " yyyymmdd = convert(varchar(10),inouttime,121)" + vbcrlf
		sqlStr = sqlStr & " from db_partner.dbo.tbl_user_inouttime_log l" + vbcrlf
		sqlStr = sqlStr & " where idx = "&idx&""
	    
	    'response.write SQL &"<br>"
	    dbget.execute sqlStr

	    ''새벽 0시부터 새벽 6사이일 경우 yyyymmdd 이전날로 처리
		sqlStr = ""
		sqlStr = "update l set" + vbcrlf
		sqlStr = sqlStr & " yyyymmdd = convert(varchar(10),dateadd(dd,-1,inouttime),121)" + vbcrlf
		sqlStr = sqlStr & " from db_partner.dbo.tbl_user_inouttime_log l" + vbcrlf
		sqlStr = sqlStr & " where datediff(hh,convert(varchar(10),inouttime,121)+' 00:00:00',inouttime) <= 5  " + vbcrlf  
	    sqlStr = sqlStr & " and idx = "&idx&""
	    
	    'response.write SQL &"<br>"
	    dbget.execute sqlStr
	
		if err.number = 0 then
			dbget.committrans			
			response.write	"<script type='text/javascript'>" &_
							"	alert('OK');" &_
							"	location.href='/common/member/fingerprints/fingerprints_inouttime_edit.asp?idx="&idx&"';" &_
							"	opener.location.reload();" &_
							"</script>"
		else
			dbget.rollbacktrans
			response.write	"<script type='text/javascript'>" &_
							"	alert('예상하지 못한 에러가 발생되었습니다. 관리자 문의 하세요');" &_
							"	history.back();" &_							
							"</script>"		
		end if
	end if				

end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->