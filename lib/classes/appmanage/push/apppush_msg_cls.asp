<%
'###########################################################
' Description : 푸시관리
' Hieditor : 서동석 생성
'			 2017.03.27 한용민 수정
'###########################################################

Class cpush_item
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx
	Public freservedate
	Public fpushtitle
	Public fpushurl
	Public fpushimg
	Public fpushimg2
	Public fpushimg3
	Public fpushimg4
	Public fpushimg5
	public fimgtype
	Public fstate
	Public ftestpush
	Public fisusing
	'// regInfo_log
	Public Fregidx
	Public fappkey
	Public fdeviceid
	Public fappVer
	Public flastact
	Public fregdate
	Public fttlCnt
	Public fwaitCnt
	Public fsentCnt
	Public fsuccCnt
	Public ffailCnt
	Public ffirstSentDate
	Public flastSentDate
	Public fdiffnomuts
	public fmakeridarr
	public fitemidarr
	public fkeywordarr
	public fbonuscouponidxarr
	public fnotclickyn

    Public fistargetMsg
    Public fnoduppDate  '' 금일 발송건에 대해 중복발송 안함.
	public fnoduppDate2
	public fnoduppDate3
    Public ftargetKey   '' 
    Public fadmcomment  '' 코멘트
    Public fbaseIdx     '' 타케팅 마스터 키 - 안드로이드에서 중복 수신 방지하기위한 키값 pkey - 이값이 동일하면 메세지가 여러개 와도 받지 않는다.

    Public ftargetName
    Public FtargetState
    Public FmayTargetCnt
	Public fregadminid
	Public flastadminid
	Public flastupdate
	Public freviewCount
	public fpushcontents
	public fscheduleidx
	public frepeatidx
	public fdaterepeatgubun
	public fcountrepeatgubun
	public frepeatdate
	public fpushschedule
	public frepeatpushyn
	public fmultiPskey
	public fdiffseconds
	public fclickCnt
	public fsendranking
	public fprivateYN

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
        IsTargetActionValid = (fistargetMsg=1) and ((FtargetState=0) or (FtargetState=1)) and ((fstate=0) or (fstate=1))
    end function
end class

class cpush_msg_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	Public FOneItem

	public farrList

	public frectcomplete
	public Frectdate
	Public Frectidx
	Public FRectuserid
	Public Fstate
	Public Fisusing
	public Frectpushtitle
	public Frectpushurl
	public FrecttargetKey
	public Frectrepeatidx
	public frectrepeatpushyn
	public Frectappkey
	public Frectdeviceid


	' /admin/appmanage/push/msg/index.asp
	public sub fpushmsglist()
		dim sqlStr, i, sqlsearch

		if Frectidx <> "" then
			sqlsearch = sqlsearch & " and P.idx = '"& Frectidx &"'"
		end If
		if Frectdate <> "" then
			sqlsearch = sqlsearch & " and convert(varchar(10),P.reservedate,120)='"& Frectdate &"'"
		end If
		if Fstate <> "" then
			sqlsearch = sqlsearch & " and P.state='"& Fstate &"'"
		end If
		if Fisusing <> "" then
			sqlsearch = sqlsearch & " and P.isusing='"& Fisusing &"'"
		end if
		if Frectpushtitle <> "" then
			sqlsearch = sqlsearch & " and P.pushtitle like '%"& Frectpushtitle &"%'"
		end if
		if Frectpushurl <> "" then
			sqlsearch = sqlsearch & " and P.pushurl like '%"& Frectpushurl &"%'"
		end if
		if FrecttargetKey <> "" then
			sqlsearch = sqlsearch & " and p.targetKey = "& FrecttargetKey &""
		end If
		if frectrepeatpushyn <> "" then
			sqlsearch = sqlsearch & " and P.repeatpushyn = '"& frectrepeatpushyn &"'"
		end If
		if frectrepeatidx <> "" then
			sqlsearch = sqlsearch & " and P.repeatidx = '"& frectrepeatidx &"'"
		end If

		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_app_push_reserve P"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch		

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close
		
		if FTotalCount < 1 then exit sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " P.idx , P.pushtitle , P.pushurl , P.pushimg , P.state , P.testpush , P.isusing , P.reservedate, P.istargetMsg, p.pushcontents"
		sqlStr = sqlStr & " , P.noduppDate, P.targetKey, P.baseIdx, P.targetState, isNULL(P.mayTargetCnt,0) as mayTargetCnt, p.regadminid, p.lastadminid"
		sqlStr = sqlStr & " , p.regdate, p.lastupdate, p.repeatpushyn, p.sendranking, p.privateYN"
		sqlStr = sqlStr & " ,Q.targetName, isNULL(E.reviewCount,0) as reviewCount"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_app_push_reserve P"
		sqlStr = sqlStr & " left join db_contents.dbo.tbl_app_targetQuery Q"
		sqlStr = sqlStr & "     on P.targetKey=Q.targetKey"
		sqlStr = sqlStr & " left join db_temp.dbo.tbl_evaluated_count_push E on P.idx=E.pushidx"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch		
		sqlStr = sqlStr & " order by P.reservedate Desc,(CASE WHEN P.targetKey=9999 THEN 0 WHEN P.targetKey=1 THEN 99999 ELSE 1 END) desc, P.idx " ''idx=>reservedate
		
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
				set FItemList(i) = new cpush_item

				FItemList(i).fidx			= rsget("idx")
				FItemList(i).fpushtitle		= rsget("pushtitle")
				FItemList(i).fpushurl		= rsget("pushurl")
				FItemList(i).fpushimg		= rsget("pushimg")
				FItemList(i).fstate			= rsget("state")
				FItemList(i).ftestpush		= rsget("testpush")
				FItemList(i).fisusing		= rsget("isusing")
				FItemList(i).freservedate	= rsget("reservedate")
				FItemList(i).fistargetMsg   = rsget("istargetMsg")
				FItemList(i).fnoduppDate    = rsget("noduppDate")
				FItemList(i).ftargetKey     = rsget("targetKey")
				FItemList(i).fbaseIdx       = rsget("baseIdx")
				FItemList(i).ftargetState   = rsget("targetState")
				FItemList(i).fmayTargetCnt   = rsget("mayTargetCnt")
				FItemList(i).ftargetName    = rsget("targetName")
				FItemList(i).fregadminid    = rsget("regadminid")
				FItemList(i).flastadminid    = rsget("lastadminid")
				FItemList(i).fregdate    = rsget("regdate")
				FItemList(i).flastupdate    = rsget("lastupdate")
				FItemList(i).freviewCount    = rsget("reviewCount")
				FItemList(i).fpushcontents    = rsget("pushcontents")
				FItemList(i).frepeatpushyn    = rsget("repeatpushyn")
				FItemList(i).fsendranking    = rsget("sendranking")
				FItemList(i).fprivateYN			= rsget("privateYN")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub

	' /admin/appmanage/push/msg/push_repeat.asp
	public sub fPush_RepeatList()
		dim sqlStr, i, sqlsearch

		if Frectrepeatidx <> "" then
			sqlsearch = sqlsearch & " and P.repeatidx = '"& Frectrepeatidx &"'"
		end If
		if Fstate <> "" then
			sqlsearch = sqlsearch & " and P.state='"& Fstate &"'"
		end If
		if Fisusing <> "" then
			sqlsearch = sqlsearch & " and P.isusing='"& Fisusing &"'"
		end if
		if Frectpushtitle <> "" then
			sqlsearch = sqlsearch & " and P.pushtitle like '%"& Frectpushtitle &"%'"
		end if
		if Frectpushurl <> "" then
			sqlsearch = sqlsearch & " and P.pushurl like '%"& Frectpushurl &"%'"
		end if
		if FrecttargetKey <> "" then
			sqlsearch = sqlsearch & " and P.targetKey = "& FrecttargetKey &""
		end If

		sqlStr = "select count(*) as cnt" & vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_app_push_repeat P" & vbcrlf
		sqlStr = sqlStr & " left join db_contents.dbo.tbl_app_targetQuery Q" & vbcrlf
		sqlStr = sqlStr & "     on P.targetKey=Q.targetKey" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch		

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close
		
		if FTotalCount < 1 then exit sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " p.repeatidx, p.pushtitle, p.pushcontents, p.pushurl, p.pushimg, p.imgtype, p.state, p.testpush" & vbcrlf
		sqlStr = sqlStr & " , p.isusing, p.noduppDate, p.targetKey, p.admcomment, p.targetstate, isNULL(p.mayTargetCnt,0) as mayTargetCnt" & vbcrlf
		sqlStr = sqlStr & " , p.makeridarr, p.itemidarr, p.keywordarr, p.bonuscouponidxarr, p.notclickyn, p.regadminid" & vbcrlf
		sqlStr = sqlStr & " , p.lastadminid, p.regdate, p.lastupdate, p.sendranking, p.privateYN" & vbcrlf
		sqlStr = sqlStr & " , Q.targetName" & vbcrlf
		sqlStr = sqlStr & " , STUFF((" & vbcrlf
		sqlStr = sqlStr & "  	SELECT Top 100 '|^|' + cast(daterepeatgubun as nvarchar(10)) + '|*|' + cast(countrepeatgubun as nvarchar(10)) + '|*|' + convert(nvarchar(19), repeatdate, 121)" & vbcrlf
		sqlStr = sqlStr & "  	FROM [db_contents].[dbo].[tbl_app_push_repeat_schedule] as s" & vbcrlf
		sqlStr = sqlStr & "  	WHERE p.repeatidx = s.repeatidx" & vbcrlf
		sqlStr = sqlStr & "  	ORDER BY s.scheduleidx asc" & vbcrlf
		sqlStr = sqlStr & "  	FOR XML PATH('')), 1, 3, '') as pushschedule" & vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_app_push_repeat P" & vbcrlf
		sqlStr = sqlStr & " left join db_contents.dbo.tbl_app_targetQuery Q" & vbcrlf
		sqlStr = sqlStr & "     on P.targetKey=Q.targetKey" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch		
		sqlStr = sqlStr & " order by P.repeatidx desc" & vbcrlf
		
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
				set FItemList(i) = new cpush_item

				FItemList(i).frepeatidx = rsget("repeatidx")
				FItemList(i).fpushtitle = rsget("pushtitle")
				FItemList(i).fpushcontents = rsget("pushcontents")
				FItemList(i).fpushurl = rsget("pushurl")
				FItemList(i).fpushimg = rsget("pushimg")
				FItemList(i).fimgtype = rsget("imgtype")
				FItemList(i).fstate = rsget("state")
				FItemList(i).ftestpush = rsget("testpush")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fnoduppDate = rsget("noduppDate")
				FItemList(i).ftargetKey = rsget("targetKey")
				FItemList(i).fadmcomment = rsget("admcomment")
				FItemList(i).ftargetstate = rsget("targetstate")
				FItemList(i).fmayTargetCnt = rsget("mayTargetCnt")
				FItemList(i).fmakeridarr = rsget("makeridarr")
				FItemList(i).fitemidarr = rsget("itemidarr")
				FItemList(i).fkeywordarr = rsget("keywordarr")
				FItemList(i).fbonuscouponidxarr = rsget("bonuscouponidxarr")
				FItemList(i).fnotclickyn = rsget("notclickyn")
				FItemList(i).fregadminid = rsget("regadminid")
				FItemList(i).flastadminid = rsget("lastadminid")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).flastupdate = rsget("lastupdate")
				FItemList(i).ftargetName    = rsget("targetName")
				FItemList(i).fpushschedule = rsget("pushschedule")
				FItemList(i).fsendranking = rsget("sendranking")
				FItemList(i).fprivateYN = rsget("privateYN")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub

	'//admin/appmanage/push/msg/poppushrepeat_edit.asp
	public sub fpush_RepeatOne_Getrow()
		dim sqlStr, i, sqlsearch

		if FRectidx <> "" then
			sqlsearch = sqlsearch & " and p.repeatidx="& FRectidx &"" & vbcrlf
		end if

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " p.repeatidx, p.pushtitle, p.pushcontents, p.pushurl, p.pushimg, p.pushimg2, p.pushimg3, p.pushimg4, p.pushimg5, p.imgtype, p.state" & vbcrlf
		sqlStr = sqlStr & " , p.testpush, p.isusing, p.noduppDate, p.targetKey, p.admcomment, p.targetstate, isNULL(p.mayTargetCnt,0) as mayTargetCnt" & vbcrlf
		sqlStr = sqlStr & " , p.makeridarr, p.itemidarr, p.keywordarr, p.bonuscouponidxarr, p.notclickyn, p.regadminid, p.lastadminid, p.regdate, p.lastupdate" & vbcrlf
		sqlStr = sqlStr & " , p.sendranking, p.privateYN" & vbcrlf
		sqlStr = sqlStr & " , (select top 1 targetName from db_contents.[dbo].[tbl_app_targetQuery] where isusing='Y' and targetKey=p.targetKey) as targetName" & vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_app_push_repeat p" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch		
		sqlStr = sqlStr & " order by p.repeatidx Desc" & vbcrlf
		
		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		set FOneItem = new cpush_item
		if Not rsget.Eof then
			FOneItem.frepeatidx = rsget("repeatidx")
			FOneItem.fpushtitle = rsget("pushtitle")
			FOneItem.fpushcontents = rsget("pushcontents")
			FOneItem.fpushurl = rsget("pushurl")
			FOneItem.fpushimg = rsget("pushimg")
			FOneItem.fpushimg2 = rsget("pushimg2")
			FOneItem.fpushimg3 = rsget("pushimg3")
			FOneItem.fpushimg4 = rsget("pushimg4")
			FOneItem.fpushimg5 = rsget("pushimg5")
			FOneItem.fimgtype = rsget("imgtype")
			FOneItem.fstate = rsget("state")
			FOneItem.ftestpush = rsget("testpush")
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fnoduppDate = rsget("noduppDate")
			FOneItem.ftargetKey = rsget("targetKey")
			FOneItem.fadmcomment = rsget("admcomment")
			FOneItem.ftargetstate = rsget("targetstate")
			FOneItem.fmayTargetCnt = rsget("mayTargetCnt")
			FOneItem.fmakeridarr = rsget("makeridarr")
			FOneItem.fitemidarr = rsget("itemidarr")
			FOneItem.fkeywordarr = rsget("keywordarr")
			FOneItem.fbonuscouponidxarr = rsget("bonuscouponidxarr")
			FOneItem.fnotclickyn = rsget("notclickyn")
			FOneItem.fregadminid = rsget("regadminid")
			FOneItem.flastadminid = rsget("lastadminid")
			FOneItem.fregdate = rsget("regdate")
			FOneItem.flastupdate = rsget("lastupdate")
			FOneItem.ftargetName = rsget("targetName")
			FOneItem.fsendranking = rsget("sendranking")
			FOneItem.fprivateYN = rsget("privateYN")

		end if

		rsget.Close
	end Sub

	' /admin/appmanage/push/msg/popPushRepeat_edit.asp
	public sub fpush_Repeat_Schedule_List()
		dim sqlStr, i, sqlsearch

		if Frectrepeatidx <> "" then
			sqlsearch = sqlsearch & " and s.repeatidx="& Frectrepeatidx &""
		end if

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " s.scheduleidx, s.repeatidx, s.daterepeatgubun, s.countrepeatgubun, s.repeatdate"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_app_push_repeat_schedule as s"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by s.scheduleidx asc"
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cpush_item

				FItemList(i).Fscheduleidx = rsget("scheduleidx")
				FItemList(i).Frepeatidx = rsget("repeatidx")
				FItemList(i).Fdaterepeatgubun = rsget("daterepeatgubun")
				FItemList(i).fcountrepeatgubun = rsget("countrepeatgubun")
				FItemList(i).frepeatdate = rsget("repeatdate")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub

	'//admin/appmanage/push/msg/poppushmsg_edit.asp
	public sub pushmsgtest_getrow()
		dim sqlStr, i, sqlsearch

		if Frectidx <> "" then
			sqlsearch = sqlsearch & " and idx="&Frectidx&"" & vbcrlf
		end if

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " idx , pushtitle , pushurl , pushimg, pushimg2, pushimg3, pushimg4, pushimg5 , state , testpush , isusing , reservedate" & vbcrlf
		sqlStr = sqlStr & " , istargetMsg, noduppDate, targetKey, admcomment, baseIdx, targetState, isNULL(mayTargetCnt,0) as mayTargetCnt, makeridarr, itemidarr" & vbcrlf
		sqlStr = sqlStr & " , keywordarr, bonuscouponidxarr, notclickyn, noduppDate2, noduppDate3, regadminid, lastadminid, regdate, lastupdate, pushcontents" & vbcrlf
		sqlStr = sqlStr & " , sendranking, privateYN" & vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_app_push_reserve with (nolock)" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch		
		sqlStr = sqlStr & " order by idx Desc" & vbcrlf
		
		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		set FOneItem = new cpush_item
		if Not rsget.Eof then
			FOneItem.Fidx			= rsget("idx")
			FOneItem.fpushtitle		= rsget("pushtitle")
			FOneItem.fpushurl		= rsget("pushurl")
			FOneItem.fpushimg		= rsget("pushimg")
			FOneItem.fpushimg2		= rsget("pushimg2")
			FOneItem.fpushimg3		= rsget("pushimg3")
			FOneItem.fpushimg4		= rsget("pushimg4")
			FOneItem.fpushimg5		= rsget("pushimg5")
			FOneItem.fstate			= rsget("state")
			FOneItem.ftestpush		= rsget("testpush")
			FOneItem.fisusing		= rsget("isusing")
			FOneItem.freservedate	= rsget("reservedate")
	        FOneItem.fistargetMsg   = rsget("istargetMsg")
	        FOneItem.fnoduppDate    = rsget("noduppDate")
			FOneItem.fnoduppDate2   = rsget("noduppDate2")
			FOneItem.fnoduppDate3   = rsget("noduppDate3")
	        FOneItem.ftargetKey     = rsget("targetKey") 
	        FOneItem.fadmcomment    = rsget("admcomment")
	        FOneItem.fbaseIdx       = rsget("baseIdx")
	        FOneItem.ftargetState   = rsget("targetState")
	        FOneItem.fmayTargetCnt   = rsget("mayTargetCnt")
	        FOneItem.fmakeridarr   = rsget("makeridarr")
	        FOneItem.fitemidarr   = rsget("itemidarr")
	        FOneItem.fkeywordarr   = rsget("keywordarr")
	        FOneItem.fbonuscouponidxarr   = rsget("bonuscouponidxarr")
	        FOneItem.fnotclickyn   = rsget("notclickyn")
			FOneItem.fregadminid    = rsget("regadminid")
			FOneItem.flastadminid    = rsget("lastadminid")
			FOneItem.fregdate    = rsget("regdate")
			FOneItem.flastupdate    = rsget("lastupdate")
			FOneItem.fpushcontents    = rsget("pushcontents")
			FOneItem.fsendranking    = rsget("sendranking")
			FOneItem.fprivateYN    = rsget("privateYN")
		end if

		rsget.Close
	end Sub

	public sub pushmsg_userinfo()
		dim sqlStr, i, sqlsearch

		if Frectuserid <> "" then
			sqlsearch = sqlsearch & " and l.userid='"&FRectuserid&"'"
		end if
		
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_app_regInfo as l "
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten as p "
		sqlStr = sqlStr & " on l.userid = p.userid "
		sqlStr = sqlStr & " where ((l.appkey=6 and l.appVer>='36') or (l.appkey=5 and l.appVer>='1')) "
		sqlStr = sqlStr & " and l.isusing = 'Y' " & sqlsearch		
        sqlStr = sqlStr & " and ((p.userid is Not NULL)"
        sqlStr = sqlStr & "    or (l.userid in ('10x10green','10x10yellow'))"
        sqlStr = sqlStr & " )" 
        
		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")
		rsget.Close
		
		if FTotalCount < 1 then exit sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " l.regidx , l.appkey , l.deviceid , l.appVer , l.lastact , l.regdate "
		sqlStr = sqlStr & " from db_contents.dbo.tbl_app_regInfo as l "
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten as p "
		sqlStr = sqlStr & " on l.userid = p.userid "
		sqlStr = sqlStr & " where ((l.appkey=6 and l.appVer>='36') or (l.appkey=5 and l.appVer>='1')) "
		sqlStr = sqlStr & " and l.isusing = 'Y' " & sqlsearch	
		sqlStr = sqlStr & " and ((p.userid is Not NULL)"
        sqlStr = sqlStr & "    or (l.userid in ('10x10green','10x10yellow'))"
        sqlStr = sqlStr & " )" 
		sqlStr = sqlStr & " order by l.lastupdate desc"
		
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
				set FItemList(i) = new cpush_item

				FItemList(i).Fregidx		= rsget("regidx")
				FItemList(i).fappkey		= rsget("appkey")
				FItemList(i).fdeviceid		= rsget("deviceid")
				FItemList(i).fappVer		= rsget("appVer")
				FItemList(i).flastact		= rsget("lastact")
				
				FItemList(i).fappkey        = rsget("appkey")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub

	' /admin/appmanage/push/msg/poppushmsg_report.asp
	public sub fpushsummary_report()
		dim sqlStr, i, sqlsearch

		if Frectidx <> "" then
			sqlsearch = sqlsearch & " and r.idx = "& Frectidx &""
		end If
		if Frectdate <> "" then
			sqlsearch = sqlsearch & " and convert(varchar(10),r.reservedate,120)='"& Frectdate &"'"
		end If
		if Frectpushtitle <> "" then
			sqlsearch = sqlsearch & " and r.pushtitle like '%"& Frectpushtitle &"%'"
		end if
		if Frectpushurl <> "" then
			sqlsearch = sqlsearch & " and r.pushurl like '%"& Frectpushurl &"%'"
		end if
		if FrecttargetKey <> "" then
			if FrecttargetKey="99999" then
				sqlsearch = sqlsearch & " and r.istargetMsg = 0"
			else
				sqlsearch = sqlsearch & " and r.targetKey = "& FrecttargetKey &""
			end if
		end If
		if frectrepeatpushyn <> "" then
			sqlsearch = sqlsearch & " and r.repeatpushyn = '"& frectrepeatpushyn &"'"
		end If
		if Frectappkey <> "" then
			sqlsearch = sqlsearch & " and s.appkey = "& Frectappkey &""
		end If

		sqlStr = "select count(*) as cnt" & vbcrlf
		sqlStr = sqlStr & " from db_AppNoti.dbo.tbl_AppPushMsgSummary as s with (nolock) " & vbcrlf
		sqlStr = sqlStr & " join tendb.db_contents.dbo.tbl_app_push_reserve as r with (nolock) " & vbcrlf
		sqlStr = sqlStr & " 	on s.multiPskey=r.idx " & vbcrlf
		sqlStr = sqlStr & " left join tendb.db_contents.dbo.tbl_app_targetQuery Q with (nolock) "
		sqlStr = sqlStr & "     on r.targetKey=Q.targetKey "
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsAppNotiget.CursorLocation = adUseClient
		rsAppNotiget.Open sqlStr, dbAppNotiget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsAppNotiget("cnt")
		rsAppNotiget.Close

		if FTotalCount < 1 then exit sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " s.multiPskey, s.appkey, s.ttlCnt, s.waitCnt, s.sentCnt, s.succCnt, s.failCnt, s.firstSentDate, s.lastSentDate " & vbcrlf
		sqlStr = sqlStr & " , datediff(n,s.firstSentDate,s.lastSentDate) as diffnomuts, s.regdate, s.lastupdate, datediff(s,firstSentDate,s.lastSentDate) as diffseconds " & vbcrlf
		sqlStr = sqlStr & " , s.clickCnt " & vbcrlf
		sqlStr = sqlStr & " , r.pushtitle , r.pushurl , r.pushimg , r.state , r.reservedate, r.targetKey, isNULL(r.mayTargetCnt,0) as mayTargetCnt " & vbcrlf
		sqlStr = sqlStr & " , r.pushcontents, r.istargetMsg, r.targetState, r.sendranking" & vbcrlf
		sqlStr = sqlStr & " , Q.targetName " & vbcrlf
		sqlStr = sqlStr & " from db_AppNoti.dbo.tbl_AppPushMsgSummary as s with (nolock) " & vbcrlf
		sqlStr = sqlStr & " join tendb.db_contents.dbo.tbl_app_push_reserve as r with (nolock) " & vbcrlf
		sqlStr = sqlStr & " 	on s.multiPskey=r.idx " & vbcrlf
		sqlStr = sqlStr & " left join tendb.db_contents.dbo.tbl_app_targetQuery Q with (nolock) "
		sqlStr = sqlStr & "     on r.targetKey=Q.targetKey "
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by s.multiPskey desc, s.appkey asc" & vbcrlf

		'response.write sqlStr &"<br>"
		rsAppNotiget.pagesize = FPageSize
		rsAppNotiget.CursorLocation = adUseClient
		rsAppNotiget.Open sqlStr, dbAppNotiget, adOpenForwardOnly, adLockReadOnly

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
		if  not rsAppNotiget.EOF  then
			rsAppNotiget.absolutepage = FCurrPage
			do until rsAppNotiget.EOF
				set FItemList(i) = new cpush_item

				FItemList(i).fmultiPskey		= rsAppNotiget("multiPskey")
				FItemList(i).fappkey		= rsAppNotiget("appkey")
				FItemList(i).fttlCnt		= rsAppNotiget("ttlCnt")
				FItemList(i).fwaitCnt		= rsAppNotiget("waitCnt")
				FItemList(i).fsentCnt		= rsAppNotiget("sentCnt")
				FItemList(i).fsuccCnt		= rsAppNotiget("succCnt")
				FItemList(i).ffailCnt		= rsAppNotiget("failCnt")
				FItemList(i).ffirstSentDate		= rsAppNotiget("firstSentDate")
				FItemList(i).flastSentDate		= rsAppNotiget("lastSentDate")
				FItemList(i).fdiffnomuts		= rsAppNotiget("diffnomuts")
				FItemList(i).fregdate		= rsAppNotiget("regdate")
				FItemList(i).flastupdate		= rsAppNotiget("lastupdate")
				FItemList(i).fdiffseconds		= rsAppNotiget("diffseconds")
				FItemList(i).fclickCnt		= rsAppNotiget("clickCnt")

				FItemList(i).freservedate		= rsAppNotiget("reservedate")
				FItemList(i).fpushtitle		= rsAppNotiget("pushtitle")
				FItemList(i).fpushcontents		= rsAppNotiget("pushcontents")
				FItemList(i).fpushurl		= rsAppNotiget("pushurl")
				FItemList(i).fpushimg		= rsAppNotiget("pushimg")
				FItemList(i).fistargetMsg		= rsAppNotiget("istargetMsg")
				FItemList(i).ftargetState		= rsAppNotiget("targetState")
				FItemList(i).fstate		= rsAppNotiget("state")

				FItemList(i).ftargetName		= rsAppNotiget("targetName")
				FItemList(i).fsendranking		= rsAppNotiget("sendranking")

				rsAppNotiget.movenext
				i=i+1
			loop
		end if
		rsAppNotiget.Close
	end Sub

	' /admin/appmanage/push/msg/poppushmsg_report.asp
	public sub fpushsummary_Repeat_report()
		dim sqlStr, i, sqlsearch

		if Frectdate <> "" then
			sqlsearch = sqlsearch & " and s.senddate='"& Frectdate &"'"
		end If
		if Frectpushtitle <> "" then
			sqlsearch = sqlsearch & " and r.pushtitle like '%"& Frectpushtitle &"%'"
		end if
		if Frectpushurl <> "" then
			sqlsearch = sqlsearch & " and r.pushurl like '%"& Frectpushurl &"%'"
		end if
		if FrecttargetKey <> "" then
			sqlsearch = sqlsearch & " and r.targetKey = "& FrecttargetKey &""
		end If
		if Frectappkey <> "" then
			sqlsearch = sqlsearch & " and s.appkey = "& Frectappkey &""
		end If

		sqlStr = "select count(*) as cnt" & vbcrlf
		sqlStr = sqlStr & " from db_AppNoti.dbo.tbl_AppPushMsgSummary_repeat as s with (nolock) " & vbcrlf
		sqlStr = sqlStr & " join tendb.db_contents.dbo.tbl_app_push_repeat as r with (nolock) " & vbcrlf
		sqlStr = sqlStr & " 	on s.targetkey=r.targetkey " & vbcrlf
		sqlStr = sqlStr & " left join tendb.db_contents.dbo.tbl_app_targetQuery Q with (nolock) "
		sqlStr = sqlStr & "     on r.targetKey=Q.targetKey "
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsAppNotiget.CursorLocation = adUseClient
		rsAppNotiget.Open sqlStr, dbAppNotiget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsAppNotiget("cnt")
		rsAppNotiget.Close

		if FTotalCount < 1 then exit sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " s.senddate as reservedate, s.appkey, s.targetkey, s.ttlCnt, s.waitCnt, s.sentCnt, s.succCnt, s.failCnt, s.firstSentDate, s.lastSentDate " & vbcrlf
		sqlStr = sqlStr & " , datediff(n,s.firstSentDate,s.lastSentDate) as diffnomuts, s.regdate, s.lastupdate, datediff(s,firstSentDate,s.lastSentDate) as diffseconds " & vbcrlf
		sqlStr = sqlStr & " , s.clickCnt, s.repeatidx " & vbcrlf
		sqlStr = sqlStr & " , r.pushtitle , r.pushurl , r.pushimg , r.state, r.pushcontents, r.sendranking" & vbcrlf
		sqlStr = sqlStr & " , Q.targetName " & vbcrlf
		sqlStr = sqlStr & " from db_AppNoti.dbo.tbl_AppPushMsgSummary_repeat as s with (nolock) " & vbcrlf
		sqlStr = sqlStr & " join tendb.db_contents.dbo.tbl_app_push_repeat as r with (nolock) " & vbcrlf
		sqlStr = sqlStr & " 	on s.targetkey=r.targetkey " & vbcrlf
		sqlStr = sqlStr & " left join tendb.db_contents.dbo.tbl_app_targetQuery Q with (nolock) "
		sqlStr = sqlStr & "     on r.targetKey=Q.targetKey "
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by s.senddate desc, s.targetkey asc, s.appkey asc" & vbcrlf

		'response.write sqlStr &"<br>"
		rsAppNotiget.pagesize = FPageSize
		rsAppNotiget.CursorLocation = adUseClient
		rsAppNotiget.Open sqlStr, dbAppNotiget, adOpenForwardOnly, adLockReadOnly

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
		if  not rsAppNotiget.EOF  then
			rsAppNotiget.absolutepage = FCurrPage
			do until rsAppNotiget.EOF
				set FItemList(i) = new cpush_item
				
				FItemList(i).ftargetkey		= rsAppNotiget("targetkey")
				FItemList(i).fappkey		= rsAppNotiget("appkey")
				FItemList(i).fttlCnt		= rsAppNotiget("ttlCnt")
				FItemList(i).fwaitCnt		= rsAppNotiget("waitCnt")
				FItemList(i).fsentCnt		= rsAppNotiget("sentCnt")
				FItemList(i).fsuccCnt		= rsAppNotiget("succCnt")
				FItemList(i).ffailCnt		= rsAppNotiget("failCnt")
				FItemList(i).ffirstSentDate		= rsAppNotiget("firstSentDate")
				FItemList(i).flastSentDate		= rsAppNotiget("lastSentDate")
				FItemList(i).fdiffnomuts		= rsAppNotiget("diffnomuts")
				FItemList(i).fregdate		= rsAppNotiget("regdate")
				FItemList(i).flastupdate		= rsAppNotiget("lastupdate")
				FItemList(i).fdiffseconds		= rsAppNotiget("diffseconds")
				FItemList(i).fclickCnt		= rsAppNotiget("clickCnt")
				FItemList(i).frepeatidx		= rsAppNotiget("repeatidx")

				FItemList(i).freservedate		= rsAppNotiget("reservedate")
				FItemList(i).fpushtitle		= rsAppNotiget("pushtitle")
				FItemList(i).fpushcontents		= rsAppNotiget("pushcontents")
				FItemList(i).fpushurl		= rsAppNotiget("pushurl")
				FItemList(i).fpushimg		= rsAppNotiget("pushimg")
				FItemList(i).fstate		= rsAppNotiget("state")

				FItemList(i).ftargetName		= rsAppNotiget("targetName")
				FItemList(i).fsendranking		= rsAppNotiget("sendranking")

				rsAppNotiget.movenext
				i=i+1
			loop
		end if
		rsAppNotiget.Close
	end Sub

	' /admin/appmanage/push/msg/poppushmsg_report.asp
	public Function fpushmessage_report()
		dim sqlStr, i, sqlsearch

		if Frectidx="" or isnull(Frectidx) then exit Function

		sqlStr = "exec [db_AppNoti].[dbo].[sp_Ten_Report_MultiMsg] '" & Frectidx & "'"

		'response.write sqlStr &"<br>"
		rsAppNotiget.pagesize = FPageSize
		rsAppNotiget.CursorLocation = adUseClient
		dbAppNotiget.CommandTimeout = 60*5   ' 5분
		rsAppNotiget.Open sqlStr, dbAppNotiget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsAppNotiget.RecordCount
		ftotalcount = rsAppNotiget.RecordCount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsAppNotiget.EOF  then
			rsAppNotiget.absolutepage = FCurrPage
			do until rsAppNotiget.EOF
				set FItemList(i) = new cpush_item

				FItemList(i).fmultiPskey		= rsAppNotiget("multiPskey")
				FItemList(i).fappkey		= rsAppNotiget("appkey")
				FItemList(i).fttlCnt		= rsAppNotiget("ttlCnt")
				FItemList(i).fwaitCnt		= rsAppNotiget("waitCnt")
				FItemList(i).fsentCnt		= rsAppNotiget("sentCnt")
				FItemList(i).fsuccCnt		= rsAppNotiget("succCnt")
				FItemList(i).ffailCnt		= rsAppNotiget("failCnt")
				FItemList(i).ffirstSentDate		= rsAppNotiget("firstSentDate")
				FItemList(i).flastSentDate		= rsAppNotiget("lastSentDate")
				FItemList(i).fdiffnomuts		= rsAppNotiget("diffnomuts")
				FItemList(i).fregdate		= rsAppNotiget("regdate")
				FItemList(i).flastupdate		= rsAppNotiget("lastupdate")
				FItemList(i).fdiffseconds		= rsAppNotiget("diffseconds")
				FItemList(i).fclickCnt		= rsAppNotiget("clickCnt")

				FItemList(i).freservedate		= rsAppNotiget("reservedate")
				FItemList(i).fpushtitle		= rsAppNotiget("pushtitle")
				FItemList(i).fpushcontents		= rsAppNotiget("pushcontents")
				FItemList(i).fpushurl		= rsAppNotiget("pushurl")
				FItemList(i).fpushimg		= rsAppNotiget("pushimg")
				FItemList(i).fistargetMsg		= rsAppNotiget("istargetMsg")
				FItemList(i).ftargetState		= rsAppNotiget("targetState")
				FItemList(i).fstate		= rsAppNotiget("state")

				FItemList(i).ftargetName		= rsAppNotiget("targetName")

				rsAppNotiget.movenext
				i=i+1
			loop
		end if
		rsAppNotiget.Close
	end Function

	' /admin/appmanage/push/msg/poppushmsg_report_clickdetail.asp
	public sub fpushreport_clicklist()
		dim sqlStr, i

		if Frectidx="" and FrecttargetKey="" then exit sub

        sqlStr = "exec db_AppNoti.dbo.usp_Ten_PushReport_click_Count '"& Frectidx &"','"& Frectdeviceid &"','"& Frectappkey &"','"& FrecttargetKey &"','"& Frectrepeatpushyn &"', '"& Frectuserid &"'" & vbcrlf

		'response.write sqlStr &"<br>"
		'response.end
		rsAppNotiget.CursorType = adOpenStatic
		rsAppNotiget.LockType = adLockOptimistic
		rsAppNotiget.pagesize = FPageSize
		rsAppNotiget.CursorLocation = adUseClient
		rsAppNotiget.Open sqlStr, dbAppNotiget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsAppNotiget("cnt")
		rsAppNotiget.Close

		if FTotalCount < 1 then exit sub

		sqlStr = "exec db_AppNoti.dbo.usp_Ten_PushReport_click_List '"&CStr((FPageSize*(FCurrPage-1)) + 1)&"','"&CStr(FPageSize*FCurrPage)&"','"& Frectidx &"','"& Frectdeviceid &"','"& Frectappkey &"','"& FrecttargetKey &"','"& Frectrepeatpushyn &"', '"& Frectuserid &"'" & vbcrlf

		'response.write sqlStr &"<br>"
		rsAppNotiget.CursorType = adOpenStatic
		rsAppNotiget.LockType = adLockOptimistic
		rsAppNotiget.pagesize = FPageSize
		rsAppNotiget.CursorLocation = adUseClient
		rsAppNotiget.Open sqlStr, dbAppNotiget, adOpenForwardOnly, adLockReadOnly

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
		if  not rsAppNotiget.EOF  then
			farrList = rsAppNotiget.getrows()
		end if
		rsAppNotiget.Close
	end Sub

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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

Class CpushtargetCommonCode
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	Public FOneItem

	public FtargetKey
	public FtargetName
	public FtargetQuery
	public Fisusing
	public Frepeatpushyn
	public ftarget_procedureyn
	public freplacetagcode

	public frectrepeatpushyn
	public frecttargetKey

	public Function GetpushtargetList
		Dim strSql, sqlsearch

		if frectrepeatpushyn<>"" then
			sqlsearch = sqlsearch & " and repeatpushyn='"&frectrepeatpushyn&"'" & vbcrlf
		end if

		strSql = "SELECT top 1000 targetKey,targetName,targetQuery,isusing,repeatpushyn, target_procedureyn, replacetagcode" & vbcrlf
		strSql = strSql & " From db_contents.[dbo].[tbl_app_targetQuery] with (nolock)" & vbcrlf
		strSql = strSql & " WHERE 1=1 " & sqlsearch
		strSql = strSql & " order by targetKey desc" & vbcrlf

		'response.write strSql &"<br>"		
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.recordcount
		Fresultcount = rsget.recordcount
		IF not rsget.EOF THEN
			GetpushtargetList = rsget.getRows()
		End IF
		rsget.Close		
	End Function
	
	public Function GetpushtargetCont
		Dim strSql, sqlsearch

		if frecttargetKey<>"" then
			sqlsearch = sqlsearch & " and targetKey="&frecttargetKey&"" & vbcrlf
		end if

		strSql = "SELECT top 1 targetKey,targetName,targetQuery,isusing,repeatpushyn, target_procedureyn, replacetagcode" & vbcrlf
		strSql = strSql & " From db_contents.[dbo].[tbl_app_targetQuery] with (nolock)" & vbcrlf
		strSql = strSql & " WHERE 1=1 " & sqlsearch

		'response.write strSql &"<br>"		
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.recordcount
		Fresultcount = rsget.recordcount
		IF not rsget.EOF THEN
			ftargetKey 	= rsget("targetKey")
			ftargetName 	= rsget("targetName")
			ftargetQuery 	= rsget("targetQuery")
			fisusing 	= rsget("isusing")
			frepeatpushyn 	= rsget("repeatpushyn")
			ftarget_procedureyn 	= rsget("target_procedureyn")
			freplacetagcode 	= rsget("replacetagcode")
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
function Drawpushstatename(selectBoxName,selectedId,changeFlag, statetype)
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

' 푸시 os 종류		' 2019.06.20 한용민 생성
function Drawpushappkeyname(selectBoxName,selectedId,changeFlag)
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %> id="<%=selectBoxName%>">
		<option value="" <% if selectedId="" then response.write "selected" %>>전체</option>
		<option value="5" <% if selectedId="5" then response.write "selected" %>>IOS</option>
		<option value="6" <% if selectedId="6" then response.write "selected" %>>ANDROID</option>
	</select>
<%
end Function

' 푸시구분		' 2019.06.20 한용민 생성
function Drawpushgubun(selectBoxName,selectedId,changeFlag,allyn)
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %> id="<%=selectBoxName%>" class="select" >
		<% if allyn="Y" then %>
			<option value="" <% if selectedId="" then response.write "selected" %>>전체</option>
		<% end if %>

		<option value="N" <% if selectedId="N" then response.write "selected" %>>일반푸시</option>
		<option value="Y" <% if selectedId="Y" then response.write "selected" %>>반복푸시</option>
	</select>
<%
end Function

' 타푸시중복우선순위		' 2019.10.24 한용민 생성
function Drawsendranking(selectBoxName,selectedId,changeFlag)
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %> id="<%=selectBoxName%>" class="select" >
		<option value="" <% if selectedId="" then response.write "selected" %>>전체</option>
		<option value="3" <% if selectedId="3" then response.write "selected" %>>높음</option>
		<option value="6" <% if selectedId="6" then response.write "selected" %>>보통</option>
		<option value="9" <% if selectedId="9" then response.write "selected" %>>낮음</option>
	</select>
<%
end Function

' 발송우선순위		' 2019.10.24 한용민 생성
Function getsendrankingname(v)
	dim tmpval

	if v = "3" then
		tmpval = "높음"
	elseif v = "6" then
		tmpval = "보통"
	elseif v = "9" then
		tmpval = "낮음"
	else
		tmpval = ""
	end If
	getsendrankingname=tmpval
End Function

' 푸시구분		' 2019.06.20 한용민 생성
Function Selectpushgubunname(v)
	dim tmpval

	if v = "N" then
		tmpval = "일반푸시"
	elseif v = "Y" then
		tmpval = "반복푸시"
	else
		tmpval = "일반푸시"
	end If
	Selectpushgubunname=tmpval
End Function

Function Selectappname(v)
	dim tmpval

	if v = "3" then
		tmpval = "colorApp(ios)"
	elseif v = "4" then
		tmpval = "colorApp(android)"
	elseif v = "5" then
		tmpval = "wishApp(ios)"
	elseif v = "6" then
		tmpval = "wishApp(android)"
	end If
	Selectappname=tmpval
End Function

Function pushmsgstate(v)
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
	pushmsgstate=tmpval
End Function

'//admin/appmanage/push/msg/poppushmsg_edit.asp		'//admin/appmanage/push/msg/index.asp
Sub drawSelectBoxTarget(selectBoxName, selectedId, addStr, repeatpushyn, targetallyn)
   dim tmp_str,query1
   %>
	<select name="<%=selectBoxName%>" <%=addStr%>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>

		<% if targetallyn="Y" then %>
			<% if repeatpushyn="N" then %>
				<option value='99999' <%if selectedId="99999" then response.write " selected"%>>회원전체</option>
			<% end if %>
		<% end if %>
   <%
    query1 = " select targetKey,targetName from db_contents.dbo.tbl_app_targetQuery with (nolock)"
    query1 = query1 & " where (isusing='Y' "
 
    if selectedId<>"" then
        query1 = query1 & " or targetKey='"&selectedId&"'"
    end if
 
    query1 = query1 & " )"

    if repeatpushyn<>"" then
        query1 = query1 & " and repeatpushyn='"& repeatpushyn &"'"
    end if
	query1 = query1 & " order by targetKey"

	'response.write query1 &"<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open query1, dbget, adOpenForwardOnly, adLockReadOnly

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("targetKey")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("targetKey")&"' "&tmp_str&">"&rsget("targetName")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   
   response.write("</select>")
End Sub

'//admin/appmanage/push/msg/poppushmsg_report.asp
Sub drawSelectBoxrepeatpush(selectBoxName, selectedId, chplg)
   dim tmp_str,sqlStr
   %>
	<select name="<%=selectBoxName%>" <%=chplg%> >
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>

	<%
	sqlStr = "select "
	sqlStr = sqlStr & " p.repeatidx, p.pushtitle,p.isusing " & vbcrlf
	sqlStr = sqlStr & " from db_contents.dbo.tbl_app_push_repeat p" & vbcrlf
	sqlStr = sqlStr & " where 1=1 "		' and p.isusing='Y'
	sqlStr = sqlStr & " group by p.repeatidx, p.pushtitle,p.isusing " & vbcrlf
	sqlStr = sqlStr & " order by p.isusing Desc, p.repeatidx Desc" & vbcrlf

	'response.write sqlStr &"<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("repeatidx")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("repeatidx")&"' "&tmp_str&">"&rsget("pushtitle")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   
   response.write("</select>")
End Sub
%>