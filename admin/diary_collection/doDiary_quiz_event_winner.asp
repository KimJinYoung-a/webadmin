<!-- #include virtual="/lib/db/db2open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim eventidx,mode,winnerList,win_idx

eventidx= request("eventidx")

mode=request("mode")

winnerList= trim(request("winnerList"))

win_idx= request("win_idx")


if mode="write" then
	dim arrayList,i

	arrayList= replace(winnerList,vbcrlf,",")
	arrayList= split(arrayList,",")

	for i = 0 to Ubound(arrayList)

		if arrayList(i)<>"" then

				sql = sql +	" insert into [db_cts].[dbo].[tbl_2007_diary_event_winner] (eventidx,userid) " &_
										" values(" & eventidx & ",'" & arrayList(i) & "')"
		end if

	next
	db2_rsget.open sql,db2_dbget,1
elseif mode="del" then
		sql	=	" delete [db_cts].[dbo].[tbl_2007_diary_event_winner] " &_
					" where win_idx = " & win_idx
		db2_rsget.open sql,db2_dbget,1

end if

response.redirect "/admin/sitemaster/diary_quiz_event_winnerList.asp?eventidx=" & eventidx
dbget.close()	:	response.End



%>

<!-- #include virtual="/lib/db/db2close.asp" -->