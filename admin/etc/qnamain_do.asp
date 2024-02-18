<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->


<style type="text/css">
<!--
.a {font-family:verdana-small;font-size:10pt;color:#000000}
-->
</style>
<%
dim mode
mode=request("mode")


'##############################
'				1:1 게시판
'##############################

if mode="qna" then '''1:1 게시판

		dim year,mon,day,wantdate
		year=request("year")
		mon=request("mon")
		day=request("day")

		wantdate=DateSerial(CStr(year),CStr(mon),CStr(day))


		dim sql
		sql="select userid,title,contents from [db_cs].[dbo].tbl_myqna" + vbcrlf
		sql = sql + " where regdate >= '" + left(wantdate,10)+ "'" + vbcrlf
		sql = sql + " and regdate < convert(varchar(10),getdate(),121)" + vbcrlf

		'response.write sql
		'dbget.close()	:	response.End
		rsget.open sql,dbget,1

		dim resultcount
		resultcount=rsget.recordcount

		dim userid(),contents(),title()

		redim preserve userid(resultcount)
		redim preserve contents(resultcount)
		redim preserve title(resultcount)
		dim i
		i=0
		do until rsget.eof
			userid(i) = rsget("userid")
			contents(i)=db2html(rsget("contents"))
			title(i)=rsget("title")
			i=i+1
			rsget.movenext
		loop

		if CStr(wantdate) = CStr(DateAdd("d",Date,"-1")) then

		Response.ContentType = "application/vnd.ms-excel"
		Response.AddHeader  "Content-Disposition" , "attachment; filename=상담" & Left(wantdate,10) & ".xls"
		else
		Response.ContentType = "application/vnd.ms-excel"
		Response.AddHeader  "Content-Disposition" , "attachment; filename=상담" & Left(wantdate,10) & "~" & DateAdd("d",Date,"-1") & ".xls"
		end if
%>

		<div align="center">
		<table border="1" cellpadding="4" cellspacing="2" width="1150">
			<tr align="center">
				<td align="center" width="50">번호</td>
				<td align="center" width="100">이름</td>
				<td width="1000">내용</td>
			</tr>
			<% for i=0 to resultcount-1 %>
			<tr align="center">
				<td class="a"><%= i+1 %></td>
				<td class="a"><% =userid(i) %></td>
				<td class="a" align="left"><% =nl2br(contents(i)) %></td>
			</tr>
			<% next %>
		</table>
		</div>
<%

'##############################
'				SMS보내기
'##############################


elseif mode="sms" then


dim inputmethod,inputArray,sendnumber
dim inputcnt,temp

inputmethod=request("inputmethod")
sendnumber=request("sendnumber")
sendmsg=html2db(request("sendmsg"))

temp=request("inputArray")

temp=split(temp,",")

inputcnt=Ubound(temp)

for i=0 to inputcnt

	if i<inputcnt then
		inputArray= inputArray & "'" & trim(temp(i)) & "',"
	else
		inputArray= inputArray & "'" & trim(temp(i)) & "'"
	end if

next


	if inputmethod="hp" then
		sql = "Insert into [db_sms].[ismsuser].em_tran (tran_phone, tran_callback, tran_status, tran_date, tran_msg)" + vbcrlf
		sql	= sql + " select distinct x.usercell,'" & sendnumber & "','1',getdate(),'" & db2html(sendmsg) & "'"+ vbcrlf
		sql = sql + " from (select distinct usercell" + vbcrlf
		sql = sql + " 			from db_user.[dbo].tbl_user_n where usercell in (" & CStr(inputArray) & ")) as x"+ vbcrlf


	elseif inputmethod="userid" then

		sql = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg )" + vbcrlf
		sql	= sql + " select distinct usercell,'" & sendnumber & "','1',getdate(),'" & db2html(sendmsg) & "'"
		sql = sql + " from db_user.[dbo].tbl_user_n " + vbcrlf
		sql = sql + " where userid in (" & CStr(inputArray) & ")"

	end if


	'response.write sql
	'dbget.close()	:	response.End

'// DB실행 및 페이지 이동 //

	'트랜젝션 시작
	dbget.beginTrans

	'실행
	dbget.execute(sql)

	'오류검사 및 반영
    If Err.Number = 0 Then
    	dbget.CommitTrans				'커밋(정상)
    	response.write "<script>alert('메시지를 발송하였습니다 ');</script>"
    	response.write "<script>history.go(-1);</script>"
    Else
			dbget.RollBackTrans				'롤백(에러발생시)
    End If

end if
 %>

<!-- #include virtual="/lib/db/dbclose.asp" -->