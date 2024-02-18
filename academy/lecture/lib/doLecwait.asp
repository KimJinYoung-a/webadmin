<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<%

function UpdateLecWaitCount(lec_idx)
	''대기자수 수정
	dim SQL

	'//상품테이블은 옵션의 총합
	SQL =	" update [db_academy].dbo.tbl_lec_item"
	SQL =	SQL +	" set wait_count=T.sumttl"
	SQL =	SQL +	" from ("
	SQL =	SQL +	" 	select lec_idx, sum(regcount) as sumttl"
	SQL =	SQL +	" 	from [db_academy].dbo.tbl_lec_waiting_user"
	SQL =	SQL +	" 	where isusing='Y'"
	SQL =	SQL +	" 	and ((currstate=0) or (currstate=3 and regEndDay>getdate()))"
	SQL =	SQL +	" 	and lec_idx=" + CStr(lec_idx)
	SQL =	SQL +	" 	group by lec_idx"
	SQL =	SQL +	" ) T"
	SQL =	SQL +	" where [db_academy].dbo.tbl_lec_item.idx=T.lec_idx"
	dbACADEMYget.execute(SQL)

	'//옵션테이블 적용
	SQL =	" update [db_academy].dbo.tbl_lec_item_option"
	SQL =	SQL +	" set wait_count=T.sumttl"
	SQL =	SQL +	" from ("
	SQL =	SQL +	" 	select lec_idx, lecOption, sum(regcount) as sumttl"
	SQL =	SQL +	" 	from [db_academy].dbo.tbl_lec_waiting_user"
	SQL =	SQL +	" 	where isusing='Y'"
	SQL =	SQL +	" 	and ((currstate=0) or (currstate=3 and regEndDay>getdate()))"
	SQL =	SQL +	" 	and lec_idx=" + CStr(lec_idx)
	SQL =	SQL +	" 	group by lec_idx, lecOption"
	SQL =	SQL +	" ) T"
	SQL =	SQL +	" where [db_academy].dbo.tbl_lec_item_option.lecIdx=T.lec_idx"
	SQL =	SQL +	"	and [db_academy].dbo.tbl_lec_item_option.lecOption=T.lecOption"
	dbACADEMYget.execute(SQL)
end function


function UpdateLecRankEdit(lec_idx)
	''대기자순위 수정
	dim SQL

	sql = " update [db_academy].dbo.tbl_lec_waiting_user"
	sql = sql + " set regrank=T.rank"
	sql = sql + " from ("
	sql = sql + " 		select idx,"
	sql = sql + " 		("
	sql = sql + " 			select count(*) from [db_academy].dbo.tbl_lec_waiting_user"
	sql = sql + " 			where lec_idx=" + CStr(lec_idx)
	sql = sql + " 			and regrank <>0"
	sql = sql +	" 			and isusing='Y'"
	sql = sql + " 			and idx<A.idx"
	sql = sql + " 			and lecOption=A.lecOption"
	sql = sql + " 		) + 1 as rank"
	sql = sql + " 		from  [db_academy].dbo.tbl_lec_waiting_user A"
	sql = sql + " 		where A.lec_idx=" + CStr(lec_idx)
	sql = sql +	" 		and A.isusing='Y'"
	sql = sql + " 		and A.regrank<>0"
	sql = sql + " 		) T"
	sql = sql + " where [db_academy].dbo.tbl_lec_waiting_user.lec_idx=" + CStr(lec_idx)
	sql = sql + " and [db_academy].dbo.tbl_lec_waiting_user.idx=T.idx"

	dbACADEMYget.execute(SQL)
end function


dim idx
dim lec_idx,userid,regcount,username,tel01,tel02,tel03,useremail,phone,isusing
dim mode, regrank, currstate, regendday, lecOption
dim SQL, msg, Previous_Rank
dim arridx

idx		= RequestCheckvar(request.form("idx"),10)
mode	= RequestCheckvar(request.form("mode"),16)
userid	= RequestCheckvar(request.form("userid"),32)
lec_idx = RequestCheckvar(request.Form("lec_idx"),10)
lecOption = RequestCheckvar(request.Form("lecOption"),4)
username = RequestCheckvar(request.Form("username"),16)
tel01 = RequestCheckvar(request.Form("tel01"),4)
tel02 = RequestCheckvar(request.Form("tel02"),4)
tel03 = RequestCheckvar(request.Form("tel03"),4)
useremail = Html2Db(request.Form("useremail"))
regrank = RequestCheckvar(request.Form("regrank"),10)
regcount = RequestCheckvar(request.Form("regcount"),10)
currstate = RequestCheckvar(request.Form("currstate"),6)
regendday = RequestCheckvar(request.Form("regendday"),10)

phone=CStr(Tel01) & "-" & CStr(Tel02) & "-" & CStr(Tel03)
isusing = RequestCheckvar(request.Form("isusing"),1)

arridx	=	trim(request("arridx"))
if arridx <> "" then
	if checkNotValidHTML(arridx) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if Right(arridx,1)="," then arridx= Left(arridx,Len(arridx)-1)

if ((mode="open") or (mode="del")) and (arridx="") then
	response.write "<script>alert('선택된 목록이 없습니다.');</script>"
	response.write "<script>location.replace('" + refer + "');</script>"
	dbACADEMYget.close()	:	response.End
end if

if (lec_idx="") then
	response.write "<script>alert('선택된 강좌가 없습니다.');</script>"
	response.write "<script>location.replace('" + refer + "');</script>"
	dbACADEMYget.close()	:	response.End
end if

Previous_Rank = 0


if mode="add" then
		dim mytotalcount
		Sql= 	" select Sum(regcount) as mycnt from [db_academy].[dbo].tbl_lec_waiting_user " &_
				" where lec_idx='" & lec_idx & "'" &_
				" and lecOption='" & lecOption & "'" &_
				" and userid='" & userid & "'" &_
				" and isusing='Y'" &_
				" and ((currstate=0) or (currstate=3 and regEndDay>getdate()))"
		rsACADEMYget.Open sql, dbACADEMYget, 1
		if not rsACADEMYget.eof then
			mytotalcount=rsACADEMYget("mycnt")
		end if
		rsACADEMYget.close

		''최대 2명까지 신청가능
		if (mytotalcount>=2) then
			response.write "<script>alert('한 강좌당 대기신청은 최대 2명까지 신청 가능하십니다.');</script>"
    		response.write "<script>history.back();</script>"
    		response.end
		end if

		Sql= 	" select Sum(regcount) as sumcnt from [db_academy].[dbo].tbl_lec_waiting_user " &_
				" where lec_idx='" & lec_idx & "'" &_
				" and lecOption='" & lecOption & "'" &_
				" and isusing='Y'"

		rsACADEMYget.Open sql, dbACADEMYget, 1

		if not rsACADEMYget.eof then
			Previous_Rank=rsACADEMYget("sumcnt")
		end if
		rsACADEMYget.close

		if IsNULL(Previous_Rank) or (Previous_Rank="") then Previous_Rank=0


		SQL =	" Insert into [db_academy].[dbo].tbl_lec_waiting_user " &_
				"	(lec_idx, lecOption, userid, user_name, user_phone, user_email, regrank, regcount) values " &_
				"	('" & lec_idx & "'" &_
				"	,'" & lecOption & "'" &_
				"	,'" & userid & "'" &_
				"	,'" & username & "'" &_
				"	,'" & phone & "'" &_
				"	,'" & useremail & "'" &_
				"	," & Previous_Rank+1 & "" &_
				"	,'" & regcount & "')"
		dbACADEMYget.execute(SQL)

	''대기자수 수정
	UpdateLecWaitCount(lec_idx)


	msg = "대기자 등록이 완료 되었습니다."

elseif mode="edit" then
	'' 수정.
	SQL =	" update [db_academy].[dbo].tbl_lec_waiting_user "
	SQL =	SQL +	" set lec_idx=" + CStr(lec_idx) + ","
	SQL =	SQL +	" userid='" + userid + "',"
	SQL =	SQL +	" user_name='" + username + "',"
	SQL =	SQL +	" user_phone='" + phone + "',"
	SQL =	SQL +	" user_email='" + useremail + "',"
	SQL =	SQL +	" regcount='" + regcount + "',"
	SQL =	SQL +	" isusing='" + isusing + "',"
	SQL =	SQL +	" currstate=" + currstate + ""

	if currstate="0" then
		SQL =	SQL +	" ,regendday=NULL"
	elseif regendday<>"" then
		SQL =	SQL +	" ,regendday='" + regendday + "'"

	end if
	SQL =	SQL +	" where idx=" + CStr(idx)

	dbACADEMYget.execute(SQL)

	''대기 순위 조정
	UpdateLecRankEdit(lec_idx)

	''대기자수 수정
	UpdateLecWaitCount(lec_idx)

	msg = "수정 되었습니다."


elseif mode="open" then
	''Open , 순위 0순위로 조정
	sql = "update [db_academy].[dbo].tbl_lec_waiting_user" ''강좌등록열경우 적용시킴
	sql = sql + " set currstate=3" + vbcrlf
	sql = sql + " ,regEndday=(convert(varchar(10),dateadd(d,1,getdate()),21) + ' 13:00:00') " + vbcrlf
	sql = sql + " ,regrank=0 " + vbcrlf
	sql = sql + " where idx in (" + arridx + ")" + vbcrlf

	rsACADEMYget.open sql,dbACADEMYget,1

	''SMS 전송
	'sql = "Insert into [110.93.128.72].[db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
	'sql = sql + " select distinct user_phone,"
	'sql = sql + " '02-741-9070',"
	'sql = sql + " '1',"
	'sql = sql + " getdate(),"
	'sql = sql + " '[핑거스]대기하신 강좌는 마이핑거스에서 확인후 결제가능합니다.(익일 오후 1시까지)'"
	'sql = sql + " from [db_academy].[dbo].tbl_lec_waiting_user where idx in (" + arridx + ")" + vbcrlf
	'sql = sql + " and isusing='Y'"
	'sql = sql + " and currstate=3"

    ''2015/10/16 수정
    sql = "insert into [SMSDB].[db_infoSMS].dbo.em_smt_tran (date_client_req, content, callback, service_type, broadcast_yn, msg_status,recipient_num) "
    sql = sql + " select distinct getdate(),'[더핑거스]대기하신 강좌는 마이핑거스에서 확인후 결제가능합니다.(익일 오후 1시까지)','027419070','0','N','1',user_phone" + vbcrlf
	sql = sql + " from [db_academy].[dbo].tbl_lec_waiting_user where idx in (" + arridx + ")" + vbcrlf
	sql = sql + " and isusing='Y'"
	sql = sql + " and currstate=3"
	    
	rsACADEMYget.Open sql,dbACADEMYget,1

	''대기 순위 조정
	UpdateLecRankEdit(lec_idx)

	''대기자수 수정
	UpdateLecWaitCount(lec_idx)

	msg = "결제 대기상태로 오픈 되었습니다. 유효기간은 (익일 오후 1시까지) 입니다."
elseif mode="del" then																	''삭제일 경우 리스트엔 보여짐 관리자도 삭제 가능
	sql = "update [db_academy].[dbo].tbl_lec_waiting_user"
	sql = sql + " set isusing='N'" + vbcrlf
	sql = sql + " where idx in (" + arridx + ")" + vbcrlf

	rsACADEMYget.open sql,dbACADEMYget,1

	''대기 순위 조정
	UpdateLecRankEdit(lec_idx)

	''대기자수 수정
	UpdateLecWaitCount(lec_idx)

	msg = "삭제 되었습니다."
end if

%>


<script>alert('<%= msg %>');</script>
<% if (mode="open") or (mode="del") then %>
	<script>document.location='/academy/lecture/wait_user_list2.asp?lec_idx=<%= lec_idx %>';</script>
<% else %>
	<script>document.location='/academy/lecture/lib/popwaitpersonreg.asp?idx=<%= idx %>&lec_idx=<%=lec_idx%>';</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->