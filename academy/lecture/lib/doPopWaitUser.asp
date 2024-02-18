<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'''doLecwait.asp 으로 수정.

dim mode,arridx,arrcnt
dim lec_idx
dim refer
refer = request.ServerVariables("HTTP_REFERER")

mode	=	RequestCheckvar(request("mode"),16)
lec_idx	=	RequestCheckvar(request("lec_idx"),10)
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

if (arridx="") then
	response.write "<script>alert('선택된 목록이 없습니다.');</script>"
	response.write "<script>location.replace('" + refer + "');</script>"
	dbget.close()	:	response.End
end if

if (lec_idx="") then
	response.write "<script>alert('선택된 강좌가 없습니다.');</script>"
	response.write "<script>location.replace('" + refer + "');</script>"
	dbget.close()	:	response.End
end if

dim sql

if mode="open" then
	''Open , 순위 1순위로 조정
	sql = "update [db_academy].[dbo].tbl_lec_waiting_user" ''강좌등록열경우 적용시킴
	sql = sql + " set currstate=3" + vbcrlf
	sql = sql + " ,regEndday=(convert(varchar(10),dateadd(d,1,getdate()),21) + ' 13:00:00') " + vbcrlf
	sql = sql + " ,regrank=0 " + vbcrlf
	sql = sql + " where idx in (" + arridx + ")" + vbcrlf

'response.write sql
	rsACADEMYget.open sql,dbACADEMYget,1

	''SMS 전송

	'sql = "Insert into [110.93.128.72].[db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
	'sql = sql + " select distinct '010-6324-9110'," ''user_phone
	'sql = sql + " '02-741-9070',"
	'sql = sql + " '1',"
	'sql = sql + " getdate(),"
	'sql = sql + " '[핑거스]대기신청 강좌가 결제 가능 상태로 변경되었습니다. 로그인후 결제 해주세요.'"
	'sql = sql + " from [db_academy].[dbo].tbl_lec_waiting_user where idx in (" + arridx + ")" + vbcrlf
	'sql = sql + " and isusing='Y'"
	'sql = sql + " and currstate=3"

	'rsACADEMYget.Open sql,dbACADEMYget,1

	''대기 순위 조정
	sql = " update [db_academy].dbo.tbl_lec_waiting_user"
	sql = sql + " set regrank=T.rank"
	sql = sql + " from ("
	sql = sql + " 		select idx,"
	sql = sql + " 		("
	sql = sql + " 			select count(*) from [db_academy].dbo.tbl_lec_waiting_user"
	sql = sql + " 			where lec_idx=" + CStr(lec_idx)
	sql = sql + " 			and idx not in (" + arridx + ")"
	sql = sql + " 			and idx<A.idx"
	sql = sql + " 		) + 1 as rank"
	sql = sql + " 		from  [db_academy].dbo.tbl_lec_waiting_user A"
	sql = sql + " 		where A.lec_idx=" + CStr(lec_idx)
	sql = sql + " 		and A.idx not in (" + arridx + ")"
	sql = sql + " 		) T"
	sql = sql + " where [db_academy].dbo.tbl_lec_waiting_user.lec_idx=" + CStr(lec_idx)
	sql = sql + " and [db_academy].dbo.tbl_lec_waiting_user.idx=T.idx"

	rsACADEMYget.open sql,dbACADEMYget,1

elseif mode="del" then																	''삭제일 경우 리스트엔 보여짐 관리자도 삭제 가능
	sql = "update [db_academy].[dbo].tbl_lec_waiting_user"
	sql = sql + " set isusing='N'" + vbcrlf
	sql = sql + " where idx in (" + arridx + ")" + vbcrlf

	rsACADEMYget.open sql,dbACADEMYget,1
'response.write sql & "<br>"
end if

%>

<script language='javascript'>
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>



<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->