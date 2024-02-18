<%
'###########################################################
' Description : LMS발송관리
' Hieditor : 2020.03.16 한용민 생성
'###########################################################

Class clms_item
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

    public fridx
	public fsendmethod
    public ftitle
    public fcontents
    public fstate
    public ftestsend
    public fisusing
    public freservedate
    public fexception7dayyn
    public ftargetkey
    public ftargetstate
    public ftargetcnt
    public fregadminid
    public flastadminid
    public fregdate
    public flastupdate
    public frepeatlmsyn
    public fmember_smsok_checkyn
	public fmember_pushyn_checkyn
	public ftargetName
	public fuserlevel
	public fcnt
	public fordercnt
	public fsubtotalprice
	public fpushycnt
	public fsendafterpushycnt
	public fmakeridarr
	public fitemidarr
	public fkeywordarr
	public fbonuscouponidxarr
	public feventcodearr
	public fbutton_name
	public fbutton_url_mobile
	public fbutton_name2
	public fbutton_url_mobile2
	public ffailed_type
	public ffailed_subject
	public ffailed_msg
	public fRSLT
	public fsendmethodresult
	public forderitemidexceptionarr
	public ftemplate_code
	public fetc_template_code
	public freplacetagcode
	public fexceptionlogin
	public fexceptionuserlevelarr
	public fuserid
	public freguserid
	public flastuserid
	public fkakaoalrimyn
	public fmember_kakaoalrimyn_checkyn
	public fsendSuccessCount
	public fvalidmembercount

    public Function getTargetStateName()
        if (FtargetState=0) then
            getTargetStateName = "대기중"
        elseif (FtargetState=1) then
            getTargetStateName = "<font color='green'>타겟예약</font>"
        elseif (FtargetState=3) then
            getTargetStateName = "타겟중"
        elseif (FtargetState=7) then
            getTargetStateName = "<font color='red'>타겟완료</font>"
        else
            getTargetStateName = FtargetState
        end if
    end function

    public Function IsTargetActionValid()
        IsTargetActionValid = ((FtargetState=0) or (FtargetState=1)) and ((fstate=0) or (fstate=1))
    end function
end class

class clms_msg_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	Public FOneItem

	Public Frectridx
    Public Frectreservedate
    Public Frectstate
    Public Frectisusing
    Public Frecttitle
    Public Frecttargetkey
    Public frectrepeatlmsyn
	public frectsendmethod
	public frectuserid
	public TENDB
	public LOGISTICSDB

	' /admin/appmanage/lms/lms_reserve.asp
	public sub flmsmsglist()
		dim sqlStr, i, sqlsearch

		if Frectridx <> "" then
			sqlsearch = sqlsearch & " and r.ridx = '"& Frectridx &"'" & vbcrlf
		end If
		if Frectreservedate <> "" then
			sqlsearch = sqlsearch & " and convert(varchar(10),r.reservedate,120)='"& Frectreservedate &"'" & vbcrlf
		end If
		if Frectstate <> "" then
			sqlsearch = sqlsearch & " and r.state='"& Frectstate &"'" & vbcrlf
		end If
		if Frectisusing <> "" then
			sqlsearch = sqlsearch & " and r.isusing='"& Frectisusing &"'" & vbcrlf
		end if
		if Frecttitle <> "" then
			sqlsearch = sqlsearch & " and r.title like '%"& Frecttitle &"%'" & vbcrlf
		end if
		if Frecttargetkey <> "" then
			sqlsearch = sqlsearch & " and r.targetkey = "& Frecttargetkey &"" & vbcrlf
		end If
		if frectrepeatlmsyn <> "" then
			sqlsearch = sqlsearch & " and r.repeatlmsyn = '"& frectrepeatlmsyn &"'" & vbcrlf
		end If
		if frectsendmethod <> "" then
			sqlsearch = sqlsearch & " and r.sendmethod = '"& frectsendmethod &"'" & vbcrlf
		end If

		sqlStr = "select count(*) as cnt" & vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_lms_reserve r with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " left join db_contents.dbo.tbl_lms_targetQuery Q with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & "     on r.targetkey=Q.targetkey" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close
		
		if FTotalCount < 1 then exit sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)  & vbcrlf
		sqlStr = sqlStr & " r.ridx, r.sendmethod, r.title, r.contents, r.state, r.testsend, r.isusing, r.reservedate, r.exception7dayyn, r.targetkey" & vbcrlf
		sqlStr = sqlStr & " , r.targetstate, r.targetcnt, r.regadminid, r.lastadminid, r.regdate, r.lastupdate, r.repeatlmsyn" & vbcrlf
		sqlStr = sqlStr & " , r.member_smsok_checkyn, r.member_pushyn_checkyn, r.makeridarr, r.itemidarr, r.keywordarr, r.bonuscouponidxarr" & vbcrlf
		sqlStr = sqlStr & " , r.button_name, r.button_url_mobile, r.button_name2, r.button_url_mobile2, r.failed_type, r.failed_subject, r.failed_msg" & vbcrlf
		sqlStr = sqlStr & " , r.orderitemidexceptionarr, r.template_code, r.etc_template_code, r.exceptionlogin, r.exceptionuserlevelarr, r.eventcodearr" & vbcrlf
		sqlStr = sqlStr & " , r.member_kakaoalrimyn_checkyn" & vbcrlf
		sqlStr = sqlStr & " , Q.targetName, q.replacetagcode" & vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_lms_reserve r with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " left join db_contents.dbo.tbl_lms_targetQuery Q with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & "     on r.targetkey=Q.targetkey" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch & vbcrlf
		sqlStr = sqlStr & " order by r.reservedate Desc, r.ridx desc" & vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new clms_item

				FItemList(i).fridx			= rsget("ridx")
				FItemList(i).fsendmethod			= rsget("sendmethod")
				FItemList(i).ftitle			= db2html(rsget("title"))
				FItemList(i).fcontents			= db2html(rsget("contents"))
				FItemList(i).fstate			= rsget("state")
				FItemList(i).ftestsend			= rsget("testsend")
				FItemList(i).fisusing			= rsget("isusing")
				FItemList(i).freservedate			= rsget("reservedate")
				FItemList(i).fexception7dayyn			= rsget("exception7dayyn")
				FItemList(i).ftargetkey			= rsget("targetkey")
				FItemList(i).ftargetstate			= rsget("targetstate")
				FItemList(i).ftargetcnt			= rsget("targetcnt")
				FItemList(i).fregadminid			= rsget("regadminid")
				FItemList(i).flastadminid			= rsget("lastadminid")
				FItemList(i).fregdate			= rsget("regdate")
				FItemList(i).flastupdate			= rsget("lastupdate")
				FItemList(i).frepeatlmsyn			= rsget("repeatlmsyn")
				FItemList(i).fmember_smsok_checkyn			= rsget("member_smsok_checkyn")
				FItemList(i).fmember_pushyn_checkyn			= rsget("member_pushyn_checkyn")
				FItemList(i).ftargetName			= db2html(rsget("targetName"))
				FItemList(i).fmakeridarr			= rsget("makeridarr")
				FItemList(i).fitemidarr			= rsget("itemidarr")
				FItemList(i).fkeywordarr			= rsget("keywordarr")
				FItemList(i).fbonuscouponidxarr			= rsget("bonuscouponidxarr")
				FItemList(i).fbutton_name			= db2html(rsget("button_name"))
				FItemList(i).fbutton_url_mobile			= rsget("button_url_mobile")
				FItemList(i).fbutton_name2			= db2html(rsget("button_name2"))
				FItemList(i).fbutton_url_mobile2			= rsget("button_url_mobile2")
				FItemList(i).ffailed_type			= rsget("failed_type")
				FItemList(i).ffailed_subject			= db2html(rsget("failed_subject"))
				FItemList(i).ffailed_msg			= db2html(rsget("failed_msg"))
				FItemList(i).forderitemidexceptionarr			= rsget("orderitemidexceptionarr")
				FItemList(i).ftemplate_code			= db2html(rsget("template_code"))
				FItemList(i).fetc_template_code			= db2html(rsget("etc_template_code"))
				FItemList(i).freplacetagcode			= db2html(rsget("replacetagcode"))
				FItemList(i).fexceptionlogin			= rsget("exceptionlogin")
				FItemList(i).fexceptionuserlevelarr			= rsget("exceptionuserlevelarr")
				FItemList(i).feventcodearr			= rsget("eventcodearr")
				FItemList(i).fmember_kakaoalrimyn_checkyn			= rsget("member_kakaoalrimyn_checkyn")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub

	' /admin/appmanage/lms/poplmsmsg_report_realtime.asp
	public Function fLmsMsgListRealTime()
		dim sqlStr, i, sqlsearch, vreservedate
		vreservedate=""

		if Frectridx="" or isnull(Frectridx) then exit Function

		sqlStr = "select" & vbcrlf
		sqlStr = sqlStr & " r.reservedate" & vbcrlf
		sqlStr = sqlStr & " from "& TENDB &"db_contents.[dbo].[tbl_lms_reserve] r with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " where r.ridx = "& Frectridx &"" & vbcrlf

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			vreservedate = dateconvert(rsget("reservedate"))
		end if
		rsget.close

		sqlStr = "Create table #tmp_user (" & vbcrlf
		sqlStr = sqlStr & " 	ridx int not null" & vbcrlf
		sqlStr = sqlStr & " 	, userid nvarchar(32) not null primary key" & vbcrlf
		sqlStr = sqlStr & " 	, userlevel int not null default 99" & vbcrlf
		sqlStr = sqlStr & " 	, pushYn nvarchar(1) not null default 'N'" & vbcrlf
		sqlStr = sqlStr & " 	, prevOrder datetime null" & vbcrlf
		sqlStr = sqlStr & " 	, sendmethod nvarchar(16) null" & vbcrlf
		sqlStr = sqlStr & " 	, sendafterpushYn nvarchar(1) not null default 'N'" & vbcrlf
		sqlStr = sqlStr & " 	, RSLT nvarchar(10) null" & vbcrlf
		sqlStr = sqlStr & " 	, reservedate datetime null" & vbcrlf
		sqlStr = sqlStr & " )" & vbcrlf
		sqlStr = sqlStr & " CREATE NONCLUSTERED INDEX IX_userid ON #tmp_user(userid ASC)" & vbcrlf

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

		' 대상자 임시 테이블 입력
		sqlStr = "insert into #tmp_user" & vbcrlf
		sqlStr = sqlStr & " 	select distinct t.ridx, isnull(t.userid,'') as userid, 99 as userlevel, 'N' as pushYn, NULL as prevOrder" & vbcrlf
		sqlStr = sqlStr & " 	, NULL sendmethod, 'N' as sendafterpushYn, NULL as RSLT, NULL reservedate" & vbcrlf
		sqlStr = sqlStr & " 	from "& TENDB &"db_contents.[dbo].[tbl_lms_targettemp] t with (readuncommitted)" & vbcrlf
		'sqlStr = sqlStr & " 	left join "& TENDB &"db_user.dbo.tbl_user_n u with (nolock)" & vbcrlf
		'sqlStr = sqlStr & " 		on t.userid = u.userid" & vbcrlf
		sqlStr = sqlStr & " 	where t.ridx = "& Frectridx &"" & vbcrlf

		'response.write sqlStr & "<Br>"
		db3_dbget.CommandTimeout = 60*5   ' 5분
		db3_dbget.execute sqlStr

		' 회원정보
		sqlStr = "update u" & vbcrlf
		sqlStr = sqlStr & " set u.userlevel=l.userlevel" & vbcrlf
		sqlStr = sqlStr & " from #tmp_user as u" & vbcrlf
		sqlStr = sqlStr & " join "& TENDB &"db_user.dbo.tbl_logindata l with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on u.userid=l.userid" & vbcrlf

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

		' LMS정보
		sqlStr = "update u" & vbcrlf
		sqlStr = sqlStr & " set u.sendmethod=r.sendmethod, u.reservedate=r.reservedate" & vbcrlf
		sqlStr = sqlStr & " from #tmp_user as u" & vbcrlf
		sqlStr = sqlStr & " join "& TENDB &"db_contents.[dbo].[tbl_lms_reserve] r with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on u.ridx=r.ridx" & vbcrlf

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

		' 푸시허용자
		sqlStr = "update u" & vbcrlf
		sqlStr = sqlStr & " set u.pushYn='Y'" & vbcrlf
		sqlStr = sqlStr & " from #tmp_user as u" & vbcrlf
		sqlStr = sqlStr & " join "& TENDB &"db_contents.dbo.tbl_app_regInfo as B with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on u.userid=B.userid" & vbcrlf
		sqlStr = sqlStr & " 	and B.pushyn='Y'" & vbcrlf
		sqlStr = sqlStr & " 	and B.isusing='Y'" & vbcrlf
		sqlStr = sqlStr & " 	and ((B.appkey=6 and B.appVer>='36')" & vbcrlf
		sqlStr = sqlStr & " 	or (B.appkey=5 and B.appVer>='1'))" & vbcrlf

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

		' 친구톡정보. 발송디비
		'sqlStr = "update t set t.RSLT=l.RSLT" & vbcrlf
		'sqlStr = sqlStr & " from #tmp_user t" & vbcrlf
		'sqlStr = sqlStr & " join "& LOGISTICSDB &"db_kakaoMsg_v4_ft.dbo.KKF_MSG l with(noLock)" & vbcrlf
		'sqlStr = sqlStr & "		on t.userid=l.etc1" & vbcrlf
		'sqlStr = sqlStr & "		and t.ridx=l.etc2" & vbcrlf
		'sqlStr = sqlStr & " 	and l.reqdate>='"& vreservedate &"'" & vbcrlf		' 발송이후인거
		'sqlStr = sqlStr & " 	and l.reqdate <= dateadd(hh,+24,'"& vreservedate &"')" & vbcrlf
		'sqlStr = sqlStr & " where t.sendmethod='KAKAOFRIEND'" & vbcrlf
		'sqlStr = sqlStr & " and t.RSLT is null" & vbcrlf

		'response.write sqlStr & "<Br>"
		'db3_dbget.execute sqlStr

		' 친구톡정보. 로그디비
		'sqlStr = "update t set t.RSLT=l.RSLT" & vbcrlf
		'sqlStr = sqlStr & " from #tmp_user t" & vbcrlf
		'sqlStr = sqlStr & " join "& LOGISTICSDB &"db_kakaoMsg_v4_ft.dbo.KKF_MSG_LOG l with(noLock)" & vbcrlf
		'sqlStr = sqlStr & "		on t.userid=l.etc1" & vbcrlf
		'sqlStr = sqlStr & "		and t.ridx=l.etc2" & vbcrlf
		'sqlStr = sqlStr & " 	and l.reqdate>='"& vreservedate &"'" & vbcrlf		' 발송이후인거
		'sqlStr = sqlStr & " 	and l.reqdate <= dateadd(hh,+24,'"& vreservedate &"')" & vbcrlf
		'sqlStr = sqlStr & " where t.sendmethod='KAKAOFRIEND'" & vbcrlf
		'sqlStr = sqlStr & " and t.RSLT is null" & vbcrlf

		'response.write sqlStr & "<Br>"
		'db3_dbget.execute sqlStr

		' 발송이후 푸시허용자
		sqlStr = "update t" & vbcrlf
		sqlStr = sqlStr & " set t.sendafterpushYn='Y'" & vbcrlf
		sqlStr = sqlStr & " from #tmp_user as t" & vbcrlf
		sqlStr = sqlStr & " join "& TENDB &"db_contents.dbo.tbl_app_wish_userinfo as f with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on t.userid=f.userid" & vbcrlf
		sqlStr = sqlStr & " 	and isnull(f.lastpushyn,'N')='Y'" & vbcrlf
		sqlStr = sqlStr & " 	and isnull(f.lastpushyndate,'')>='"& vreservedate &"'" & vbcrlf		' 발송이후인거
		sqlStr = sqlStr & " 	and isnull(f.lastpushyndate,'') <= dateadd(hh,+24,'"& vreservedate &"')" & vbcrlf

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

		' 이전구매일수
		sqlStr = "update t set t.prevOrder=m.regdate" & vbcrlf
		sqlStr = sqlStr & " from #tmp_user t" & vbcrlf
		sqlStr = sqlStr & " join "& TENDB &"db_order.dbo.tbl_order_master as m with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on t.userid=m.userid" & vbcrlf
		sqlStr = sqlStr & " 	and m.ipkumdiv>3" & vbcrlf
		sqlStr = sqlStr & " 	and m.cancelyn='N'" & vbcrlf
		sqlStr = sqlStr & " 	and m.jumundiv not in ('6','9')" & vbcrlf
		sqlStr = sqlStr & " 	and m.regdate<='"& vreservedate &"'" & vbcrlf		' 발송일 이전인거

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

		' 주문건저장
		sqlStr = " select" & vbcrlf
		sqlStr = sqlStr & " t.userid, count(m.idx) as ordercnt, sum(isnull(m.subtotalprice,0)) as subtotalprice" & vbcrlf
		sqlStr = sqlStr & " into #tmporder" & vbcrlf
		sqlStr = sqlStr & " from #tmp_user as T" & vbcrlf
		sqlStr = sqlStr & " join "& TENDB &"db_order.dbo.tbl_order_master as m with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on T.userid=m.userid" & vbcrlf
		sqlStr = sqlStr & " 	and m.ipkumdiv>3" & vbcrlf
		sqlStr = sqlStr & " 	and m.jumundiv not in ('6','9')" & vbcrlf
		sqlStr = sqlStr & " 	and m.cancelyn='N'" & vbcrlf
		sqlStr = sqlStr & " 	and m.regdate >= '"& vreservedate &"'" & vbcrlf
		sqlStr = sqlStr & " 	and m.regdate <= dateadd(hh,+24,'"& vreservedate &"')" & vbcrlf
		sqlStr = sqlStr & " 	and m.userid<>''" & vbcrlf
		sqlStr = sqlStr & " where t.userid<>''" & vbcrlf
		sqlStr = sqlStr & " group by t.userid" & vbcrlf

		'response.write sqlStr & "<Br>"
		'response.end
		db3_dbget.execute sqlStr

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)  & vbcrlf
		sqlStr = sqlStr & "  t.userlevel" & vbcrlf
		sqlStr = sqlStr & " , count(t.userid) as cnt" & vbcrlf
		sqlStr = sqlStr & " , sum(isnull(m.ordercnt,0)) as ordercnt" & vbcrlf
		sqlStr = sqlStr & " , sum(isnull(m.subtotalprice,0)) as subtotalprice" & vbcrlf
		sqlStr = sqlStr & " , sum((case when t.pushYn='Y' then 1 else 0 end)) as pushycnt" & vbcrlf
		sqlStr = sqlStr & " , sum((case when t.sendafterpushYn='Y' then 1 else 0 end)) as sendafterpushycnt" & vbcrlf
		sqlStr = sqlStr & " , (case when sendmethod='KAKAOFRIEND' then (case when RSLT='0' then t.sendmethod else 'LMS' end)" & vbcrlf
		sqlStr = sqlStr & "  	else t.sendmethod end) as sendmethodresult" & vbcrlf
		sqlStr = sqlStr & " from #tmp_user as T" & vbcrlf
		sqlStr = sqlStr & " left join #tmporder as m" & vbcrlf
		sqlStr = sqlStr & " 	on t.userid=m.userid" & vbcrlf
		sqlStr = sqlStr & " group by t.userlevel" & vbcrlf
		sqlStr = sqlStr & " , (case when sendmethod='KAKAOFRIEND' then (case when RSLT='0' then t.sendmethod else 'LMS' end)" & vbcrlf
		sqlStr = sqlStr & "  	else t.sendmethod end)" & vbcrlf
		sqlStr = sqlStr & " order by sendmethodresult asc, t.userlevel asc" & vbcrlf

		'response.write sqlStr &"<br>"
		db3_rsget.pagesize = FPageSize
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.RecordCount
		ftotalcount = db3_rsget.RecordCount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.EOF
				set FItemList(i) = new clms_item

				FItemList(i).fuserlevel	= db3_rsget("userlevel")
				FItemList(i).fcnt		= db3_rsget("cnt")
				FItemList(i).fordercnt		= db3_rsget("ordercnt")
				FItemList(i).fsubtotalprice		= db3_rsget("subtotalprice")
				FItemList(i).fpushycnt		= db3_rsget("pushycnt")
				FItemList(i).fsendafterpushycnt		= db3_rsget("sendafterpushycnt")
				FItemList(i).fsendmethodresult		= db3_rsget("sendmethodresult")

				db3_rsget.movenext
				i=i+1
			loop
		end if
		db3_rsget.Close
	end Function

	' /admin/appmanage/lms/poplmsmsgReportStatisticsDB.asp
	public Function fLmsMsgListStatisticsDB()
		dim sqlStr, i, sqlsearch

		if Frectridx="" or isnull(Frectridx) then exit Function

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)  & vbcrlf
		sqlStr = sqlStr & " t.ridx, t.sendmethod, t.userlevel, t.cnt, t.orderCnt, t.subTotalPrice, t.pushyCnt" & vbcrlf
		sqlStr = sqlStr & " , t.sendafterpushycnt, t.sendSuccessCount, t.validmembercount" & vbcrlf
		sqlStr = sqlStr & " from db_datamart.[dbo].[tbl_lmsReserveSummary] as T with (nolock)" & vbcrlf
		sqlStr = sqlStr & " where ridx="& Frectridx &"" & vbcrlf
		sqlStr = sqlStr & " order by t.ridx asc, t.sendmethod asc, t.userlevel asc"

		'response.write sqlStr &"<br>"
		db3_rsget.pagesize = FPageSize
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.RecordCount
		ftotalcount = db3_rsget.RecordCount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.EOF
				set FItemList(i) = new clms_item

				FItemList(i).fuserlevel	= db3_rsget("userlevel")
				FItemList(i).fcnt		= db3_rsget("cnt")
				FItemList(i).fordercnt		= db3_rsget("ordercnt")
				FItemList(i).fsubtotalprice		= db3_rsget("subtotalprice")
				FItemList(i).fpushycnt		= db3_rsget("pushycnt")
				FItemList(i).fsendafterpushycnt		= db3_rsget("sendafterpushycnt")
				FItemList(i).fsendmethod		= db3_rsget("sendmethod")
				FItemList(i).fsendSuccessCount		= db3_rsget("sendSuccessCount")
				FItemList(i).fvalidmembercount		= db3_rsget("validmembercount")

				db3_rsget.movenext
				i=i+1
			loop
		end if
		db3_rsget.Close
	end Function

	'//admin/appmanage/lms/msg/poplmsmsg_edit.asp
	public sub lmsmsg_getrow()
		dim sqlStr, i, sqlsearch

		if Frectridx <> "" then
			sqlsearch = sqlsearch & " and r.ridx="&Frectridx&"" & vbcrlf
		end if
		if Frectisusing <> "" then
			sqlsearch = sqlsearch & " and r.isusing='"&Frectisusing&"'" & vbcrlf
		end if

		sqlStr = "select top 1"
		sqlStr = sqlStr & " r.ridx, r.sendmethod, r.title, r.contents, r.state, r.testsend, r.isusing, r.reservedate, r.exception7dayyn, r.targetkey" & vbcrlf
		sqlStr = sqlStr & " , r.targetstate, r.targetcnt, r.regadminid, r.lastadminid, r.regdate, r.lastupdate, r.repeatlmsyn" & vbcrlf
		sqlStr = sqlStr & " , r.member_smsok_checkyn, r.member_pushyn_checkyn, r.makeridarr, r.itemidarr, r.keywordarr, r.bonuscouponidxarr" & vbcrlf
		sqlStr = sqlStr & " , r.button_name, r.button_url_mobile, r.button_name2, r.button_url_mobile2, r.failed_type, r.failed_subject, r.failed_msg" & vbcrlf
		sqlStr = sqlStr & " , r.orderitemidexceptionarr, r.template_code, r.etc_template_code, r.exceptionlogin, r.exceptionuserlevelarr, r.eventcodearr" & vbcrlf
		sqlStr = sqlStr & " , r.member_kakaoalrimyn_checkyn" & vbcrlf
		sqlStr = sqlStr & " , Q.targetName, q.replacetagcode" & vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_lms_reserve r with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " left join db_contents.dbo.tbl_lms_targetQuery Q" & vbcrlf
		sqlStr = sqlStr & "     on r.targetkey=Q.targetkey" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch		
		sqlStr = sqlStr & " order by r.ridx Desc" & vbcrlf

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		ftotalcount = rsget.RecordCount
		set FOneItem = new clms_item

		if Not rsget.Eof then
			FOneItem.fridx			= rsget("ridx")
			FOneItem.fsendmethod			= rsget("sendmethod")
			FOneItem.ftitle			= db2html(rsget("title"))
			FOneItem.fcontents			= db2html(rsget("contents"))
			FOneItem.fstate			= rsget("state")
			FOneItem.ftestsend			= rsget("testsend")
			FOneItem.fisusing			= rsget("isusing")
			FOneItem.freservedate			= rsget("reservedate")
			FOneItem.fexception7dayyn			= rsget("exception7dayyn")
			FOneItem.ftargetkey			= rsget("targetkey")
			FOneItem.ftargetstate			= rsget("targetstate")
			FOneItem.ftargetcnt			= rsget("targetcnt")
			FOneItem.fregadminid			= rsget("regadminid")
			FOneItem.flastadminid			= rsget("lastadminid")
			FOneItem.fregdate			= rsget("regdate")
			FOneItem.flastupdate			= rsget("lastupdate")
			FOneItem.frepeatlmsyn			= rsget("repeatlmsyn")
			FOneItem.fmember_smsok_checkyn			= rsget("member_smsok_checkyn")
			FOneItem.fmember_pushyn_checkyn			= rsget("member_pushyn_checkyn")
			FOneItem.ftargetName			= db2html(rsget("targetName"))
			FOneItem.fmakeridarr			= rsget("makeridarr")
			FOneItem.fitemidarr			= rsget("itemidarr")
			FOneItem.fkeywordarr			= rsget("keywordarr")
			FOneItem.fbonuscouponidxarr			= rsget("bonuscouponidxarr")
			FOneItem.fbutton_name			= db2html(rsget("button_name"))
			FOneItem.fbutton_url_mobile			= rsget("button_url_mobile")
			FOneItem.fbutton_name2			= db2html(rsget("button_name2"))
			FOneItem.fbutton_url_mobile2			= rsget("button_url_mobile2")
			FOneItem.ffailed_type			= rsget("failed_type")
			FOneItem.ffailed_subject			= db2html(rsget("failed_subject"))
			FOneItem.ffailed_msg			= db2html(rsget("failed_msg"))
			FOneItem.forderitemidexceptionarr			= rsget("orderitemidexceptionarr")
			FOneItem.ftemplate_code			= db2html(rsget("template_code"))
			FOneItem.fetc_template_code			= db2html(rsget("etc_template_code"))
			FOneItem.freplacetagcode			= db2html(rsget("replacetagcode"))
			FOneItem.fexceptionlogin			= rsget("exceptionlogin")
			FOneItem.fexceptionuserlevelarr			= rsget("exceptionuserlevelarr")
			FOneItem.feventcodearr			= rsget("eventcodearr")
			FOneItem.fmember_kakaoalrimyn_checkyn			= rsget("member_kakaoalrimyn_checkyn")
		end if

		rsget.Close
	end Sub

	' //admin/appmanage/lms/lms_agree.asp
	public sub flms_agree_list()
		dim sqlStr, i, sqlsearch

		if frectuserid <> "" then
			sqlsearch = sqlsearch & " and a.userid = '"& frectuserid &"'" & vbcrlf
		end If

		sqlStr = "select count(*) as cnt" & vbcrlf
		sqlStr = sqlStr & " from db_contents.[dbo].[tbl_lms_agree] a with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close
		
		if FTotalCount < 1 then exit sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)  & vbcrlf
		sqlStr = sqlStr & " userid, regdate, lastupdate, reguserid, lastuserid, kakaoalrimyn" & vbcrlf
		sqlStr = sqlStr & " from db_contents.[dbo].[tbl_lms_agree] a with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by isnull(a.lastupdate,a.regdate) desc" & vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new clms_item

				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).flastupdate = rsget("lastupdate")
				FItemList(i).freguserid = rsget("reguserid")
				FItemList(i).flastuserid = rsget("lastuserid")
				FItemList(i).fkakaoalrimyn = rsget("kakaoalrimyn")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub

	' /admin/appmanage/lms/lms_agree_edit.asp
	public sub flms_agree_one()
		dim sqlStr, i, sqlsearch

		if frectuserid <> "" then
			sqlsearch = sqlsearch & " and a.userid = '"& frectuserid &"'" & vbcrlf
		end If

		sqlStr = "select top 1"
		sqlStr = sqlStr & " userid, regdate, lastupdate, reguserid, lastuserid, kakaoalrimyn" & vbcrlf
		sqlStr = sqlStr & " from db_contents.[dbo].[tbl_lms_agree] a with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		ftotalcount = rsget.RecordCount
		set FOneItem = new clms_item

		if Not rsget.Eof then
			FOneItem.fuserid = rsget("userid")
			FOneItem.fregdate = rsget("regdate")
			FOneItem.flastupdate = rsget("lastupdate")
			FOneItem.freguserid = rsget("reguserid")
			FOneItem.flastuserid = rsget("lastuserid")
			FOneItem.fkakaoalrimyn = rsget("kakaoalrimyn")
		end if

		rsget.Close
	end Sub

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

		IF application("Svr_Info")="Dev" THEN
			TENDB="tendb."
		else
			LOGISTICSDB="LOGISTICSDB."
		end if
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function
	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

Class ClmstargetCommonCode
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	Public FOneItem

	public Ftargetkey
	public FtargetName
	public FtargetQuery
	public Fisusing
	public Frepeatlmsyn
	public ftarget_procedureyn
	public ftidx
	public fsendmethod
	public ftemplate_code
	public ftemplate_name
	public fcontents
	public fbutton_name
	public fbutton_url_mobile
	public fbutton_name2
	public fbutton_url_mobile2
	public ffailed_type
	public ffailed_subject
	public ffailed_msg
	public fregadminid
	public flastadminid
	public fregdate
	public flastupdate
	public fsortno
	public freplacetagcode

	public frectrepeatlmsyn
	public frecttargetkey
	public Frectisusing
	public frecttidx
	public frectsendmethod

	public Function GetlmstargetList
		Dim strSql, sqlsearch

		if frectrepeatlmsyn<>"" then
			sqlsearch = sqlsearch & " and q.repeatlmsyn='"&frectrepeatlmsyn&"'" & vbcrlf
		end if

		strSql = "SELECT q.targetkey,q.targetName,q.targetQuery,q.isusing,q.repeatlmsyn,q.target_procedureyn,q.replacetagcode" & vbcrlf
		strSql = strSql & " From db_contents.dbo.tbl_lms_targetQuery q with (readuncommitted)" & vbcrlf
		strSql = strSql & " WHERE 1=1 " & sqlsearch
		strSql = strSql & " order by q.repeatlmsyn asc, q.targetkey desc" & vbcrlf

		'response.write strSql &"<br>"		
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.recordcount
		Fresultcount = rsget.recordcount
		IF not rsget.EOF THEN
			GetlmstargetList = rsget.getRows()
		End IF
		rsget.Close		
	End Function
	
	public Function GetlmstargetCont
		Dim strSql, sqlsearch

		if frecttargetkey<>"" then
			sqlsearch = sqlsearch & " and q.targetkey="&frecttargetkey&"" & vbcrlf
		end if
		if Frectisusing<>"" then
			sqlsearch = sqlsearch & " and q.isusing='"&Frectisusing&"'" & vbcrlf
		end if

		strSql = "SELECT top 1 q.targetkey,q.targetName,q.targetQuery,q.isusing,q.repeatlmsyn,q.target_procedureyn,q.replacetagcode" & vbcrlf
		strSql = strSql & " From db_contents.dbo.tbl_lms_targetQuery q with (readuncommitted)" & vbcrlf
		strSql = strSql & " WHERE 1=1 " & sqlsearch

		'response.write strSql &"<br>"		
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.recordcount
		Fresultcount = rsget.recordcount
		IF not rsget.EOF THEN
			ftargetkey 	= rsget("targetkey")
			ftargetName 	= db2html(rsget("targetName"))
			ftargetQuery 	= db2html(rsget("targetQuery"))
			fisusing 	= rsget("isusing")
			frepeatlmsyn 	= rsget("repeatlmsyn")
			ftarget_procedureyn 	= rsget("target_procedureyn")
			freplacetagcode 	= db2html(rsget("replacetagcode"))
		End IF			
		rsget.Close		
	End Function

	public Function GetlmstemplateList
		Dim strSql, sqlsearch

		if frectsendmethod<>"" then
			sqlsearch = sqlsearch & " and t.sendmethod='"&frectsendmethod&"'" & vbcrlf
		end if

		strSql = "SELECT" & vbcrlf
		strSql = strSql & " t.tidx, t.sendmethod, t.template_code, t.template_name, t.contents, t.button_name, t.button_url_mobile" & vbcrlf
		strSql = strSql & " , t.button_name2, t.button_url_mobile2, t.failed_type, t.failed_subject, t.failed_msg, t.isusing, t.regadminid" & vbcrlf
		strSql = strSql & " , t.lastadminid, t.regdate, t.lastupdate, t.sortno" & vbcrlf
		strSql = strSql & " From db_contents.dbo.tbl_lms_template t with (readuncommitted)" & vbcrlf
		strSql = strSql & " WHERE 1=1 " & sqlsearch
		strSql = strSql & " order by t.sendmethod asc, t.sortno desc, t.tidx desc" & vbcrlf

		'response.write strSql &"<br>"		
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.recordcount
		Fresultcount = rsget.recordcount
		IF not rsget.EOF THEN
			GetlmstemplateList = rsget.getRows()
		End IF
		rsget.Close		
	End Function

	public Function GetlmstemplateCont
		Dim strSql, sqlsearch

		if frecttidx<>"" then
			sqlsearch = sqlsearch & " and t.tidx="&frecttidx&"" & vbcrlf
		end if
		if Frectisusing<>"" then
			sqlsearch = sqlsearch & " and t.isusing='"&Frectisusing&"'" & vbcrlf
		end if

		strSql = "SELECT" & vbcrlf
		strSql = strSql & " t.tidx, t.sendmethod, t.template_code, t.template_name, t.contents, t.button_name, t.button_url_mobile" & vbcrlf
		strSql = strSql & " , t.button_name2, t.button_url_mobile2, t.failed_type, t.failed_subject" & vbcrlf
		strSql = strSql & " , t.failed_msg, t.isusing, t.regadminid, t.lastadminid, t.regdate, t.lastupdate, t.sortno" & vbcrlf
		strSql = strSql & " From db_contents.dbo.tbl_lms_template t with (readuncommitted)" & vbcrlf
		strSql = strSql & " WHERE 1=1 " & sqlsearch

		'response.write strSql &"<br>"		
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.recordcount
		Fresultcount = rsget.recordcount
		IF not rsget.EOF THEN
			ftidx 	= rsget("tidx")
			fsendmethod 	= rsget("sendmethod")
			ftemplate_code 	= db2html(rsget("template_code"))
			ftemplate_name 	= db2html(rsget("template_name"))
			fcontents 	= db2html(rsget("contents"))
			fbutton_name 	= db2html(rsget("button_name"))
			fbutton_url_mobile 	= db2html(rsget("button_url_mobile"))
			fbutton_name2 	= db2html(rsget("button_name2"))
			fbutton_url_mobile2 	= db2html(rsget("button_url_mobile2"))
			ffailed_type 	= rsget("failed_type")
			ffailed_subject 	= db2html(rsget("failed_subject"))
			ffailed_msg 	= db2html(rsget("failed_msg"))
			fisusing 	= rsget("isusing")
			fregadminid 	= rsget("regadminid")
			flastadminid 	= rsget("lastadminid")
			fregdate 	= rsget("regdate")
			flastupdate 	= rsget("lastupdate")
			fsortno 	= rsget("sortno")
		End IF			
		rsget.Close		
	End Function

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

' 발송상태 공통함수
function Drawlmsstatename(selectBoxName,selectedId,changeFlag, statetype)
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %> id="<%=selectBoxName%>">
		<option value="" <% if selectedId="" then response.write "selected" %>>전체</option>
		<option value="0" <% if selectedId="0" then response.write "selected" %>>작성중</option>
		<option value="1" <% if selectedId="1" then response.write "selected" %>>발송예약</option>

		<% if statetype="" then %>
			<option value="7" <% if selectedId="7" then response.write "selected" %>>타겟중</option>
			<option value="9" <% if selectedId="9" then response.write "selected" %>>발송완료</option>
		<% end if %>
	</select>
<%
end Function

' LMS구분
function Drawrepeatgubun(selectBoxName,selectedId,changeFlag,allyn)
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %> id="<%=selectBoxName%>" class="select" >
		<% if allyn="Y" then %>
			<option value="" <% if selectedId="" then response.write "selected" %>>전체</option>
		<% end if %>

		<option value="N" <% if selectedId="N" then response.write "selected" %>>일반발송</option>
		<option value="Y" <% if selectedId="Y" then response.write "selected" %>>반복발송</option>
	</select>
<%
end Function

' 알림톡,친구톡 실패시
function Drawfailed_type(selectBoxName,selectedId,changeFlag)
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %> id="<%=selectBoxName%>" class="select" >
		<option value="" <% if selectedId="" then response.write "selected" %>>실패시문자발송안함</option>
		<option value="LMS" <% if selectedId="LMS" then response.write "selected" %>>LMS</option>
	</select>
<%
end Function

' 발송방법
function Drawsendmethod(selectBoxName,selectedId,changeFlag,allyn)
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %> id="<%=selectBoxName%>" class="select" >
		<% if allyn="Y" then %>
			<option value="" <% if selectedId="" then response.write "selected" %>>전체</option>
		<% end if %>

		<option value="LMS" <% if selectedId="LMS" then response.write "selected" %>>LMS</option>
		<option value="KAKAOFRIEND" <% if selectedId="KAKAOFRIEND" then response.write "selected" %>>친구톡</option>
		<option value="KAKAOALRIM" <% if selectedId="KAKAOALRIM" then response.write "selected" %>>알림톡</option>
	</select>
<%
end Function

' 발송방법
Function Selectsendmethodname(v)
	dim tmpval

	if v = "LMS" then
		tmpval = "LMS"
	elseif v = "KAKAOFRIEND" then
		tmpval = "친구톡"
	elseif v = "KAKAOALRIM" then
		tmpval = "알림톡"
	else
		tmpval = ""
	end If
	Selectsendmethodname=tmpval
End Function

' LMS구분
Function Selectlmsgubunname(v)
	dim tmpval

	if v = "N" then
		tmpval = "일반발송"
	elseif v = "Y" then
		tmpval = "반복발송"
	else
		tmpval = "일반발송"
	end If
	Selectlmsgubunname=tmpval
End Function

Function lmsmsgstate(v)
	dim tmpval

	if v = "0" then
		tmpval = "작성중"
	elseif v = "1" then
		tmpval = "<font color='red'>발송예약</font>"
	elseif v = "7" then
		tmpval = "<font color='green'>발송중</font>"
	elseif v = "9" then
		tmpval = "<font color='blue'>발송완료</font>"
	Else
		tmpval = "전체"
	end If
	lmsmsgstate=tmpval
End Function

' 발송타켓
Sub drawSelectBoxlmsTarget(selectBoxName, selectedId, addStr, repeatlmsyn)
   dim tmp_str,query1
   %>
	<select name="<%=selectBoxName%>" <%=addStr%>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>

   <%
    query1 = " select targetkey,targetName from db_contents.dbo.tbl_lms_targetQuery with (readuncommitted)"
    query1 = query1 & " where (isusing='Y' "
 
    if selectedId<>"" then
        query1 = query1 & " or targetkey='"&selectedId&"'"
    end if
 
    query1 = query1 & " )"

    if repeatlmsyn<>"" then
        query1 = query1 & " and repeatlmsyn='"& repeatlmsyn &"'"
    end if
	query1 = query1 & " order by targetkey"

	'response.write query1 &"<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open query1, dbget, adOpenForwardOnly, adLockReadOnly

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("targetkey")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("targetkey")&"' "&tmp_str&">"&rsget("targetName")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   
   response.write("</select>")
End Sub

' 템플릿
Sub drawSelectBoxtemplate(selectBoxName, selectedId, chplg, sendmethod)
   dim tmp_str,query1

	if sendmethod="" or isnull(sendmethod) then exit Sub

   %>
	<select name="<%=selectBoxName%>" <%= chplg %>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>

   <%
    query1 = " select sendmethod, template_code, template_name from db_contents.dbo.tbl_lms_template with (readuncommitted)" & vbcrlf
    query1 = query1 & " where isusing='Y'" & vbcrlf

	if sendmethod<>"" then
		query1 = query1 & " and sendmethod=N'"& sendmethod &"'" & vbcrlf
	end if

	query1 = query1 & " order by sendmethod asc, sortno desc, tidx desc" & vbcrlf

	'response.write query1 &"<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open query1, dbget, adOpenForwardOnly, adLockReadOnly

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("template_code")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"& db2html(rsget("template_code")) &"' "&tmp_str&">["& db2html(rsget("template_code")) &"]"& db2html(rsget("template_name")) &"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   
   response.write("</select>")
End Sub

%>