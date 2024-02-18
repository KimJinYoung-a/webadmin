<%
'###########################################################
' Description : 게시판
' Hieditor : 2009.04.17 이상구 생성
'			 2010.01.03 한용민 수정
'###########################################################

Class CMyQNAExtItem
	Public Forderserial
	Public Fbuyname

	Private Sub Class_Initialize()
		''
	End Sub
	Private Sub Class_Terminate()
		''
	End Sub
End Class

class cmyqna_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fid
	public fuserid
	public fusername
	public forderserial
	public fqadiv
	public ftitle
	public fusermail
	public Fuserphone
	public femailok
	public fcontents
	public fregdate
	public freplyuser
	public Fchargeid
	public freplytitle
	public freplycontents
	public freplydate
	public fitemid
	public fisusing
	public fFextsitename
	public fFuserlevel
	public fFExpectReplyDate
	public fFrealnamecheck
	public fshopid
	public Fuserlevel
	public fssBctId
	public fcompany_name
	public flevel_sn
	public fuserdiv
	public fbigo

	'/권한체크 '/매장소속
	public function getmemberdisp()
	    '/level_sn 6:매장직원 ,7:매장캐셔권한 , 8:매출조회권한  '/직원권한
	    if flevel_sn <> "6" and flevel_sn <> "7" and flevel_sn <> "8" and fuserdiv <= 9 then
			getmemberdisp = true
		else
			getmemberdisp = false
		end if
	end function

	'/권한체크 '직원 아이디로 접속 했는지 아닌지 // 매장아이디가 항상 streetshop으로 시작하지 않음.. :: 수정해야함.
	public function getmemberofficedisp()
	    if (left(fssBctId,10) <> "streetshop") and (left(fssBctId,9) <> "wholesale") then
			getmemberofficedisp = true
		else
			getmemberofficedisp = false
		end if
	end function

	'/사용중지 공통함수에 공용펑션으로 쓸것		'/2016.06.30 한용민
    public function GetUserLevelStr()
        if IsNULL(Fuserlevel) then
    		GetUserLevelStr = "&nbsp;"
    	    Exit function
        end if

        Select Case CStr(Fuserlevel)
    		Case "5"
    			GetUserLevelStr = "<span class='member_orange'>ORANGE</span>"
    		Case "0"
    			GetUserLevelStr = "<span class='member_yellow'>YELLOW</span>"
    		Case "1"
    			GetUserLevelStr = "<span class='member_green'>GREEN</span>"
    		Case "2"
    			GetUserLevelStr = "<span class='member_blue'>BLUE</span>"
    		Case "3"
    			GetUserLevelStr = "<span class='member_vipsilver'>VIP SILVER</span>"
    			''GetUserLevelStr = "<span class='member_vip'>VIP</span>"
    		Case "4"
    			GetUserLevelStr = "<span class='member_vipgold'>VIP GOLD</span>"
    		Case "6"
    			GetUserLevelStr = "<span class='member_vvipgold'>VVIP</span>"
    		Case "7"
    			GetUserLevelStr = "<span class='member_staff'>STAFF</span>"
    		Case "8"
    			GetUserLevelStr = "<span class='member_red'>FAMILY</span>"
    		Case "9"
    			GetUserLevelStr = "<span class='member_red'>BIZ</span>"
    		Case Else
    			GetUserLevelStr = "<span class='member_orange'>ORANGE</span>"
    	end Select
	end function
end class

class CMyQNA_list
	public FItemList()
	public FOneItem
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public frectisusing
	public frectshopid
	public frectidx
	public FRectStartDay
	public FRectEndDay
	public FRectmakerid
	public frectid
	public frectssBctId

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

	'//common/offshop/board/online_cscenter_qna_reply.asp
    public Sub fqnaread()
        dim sqlStr , sqlsearch

		if frectid <> "" then
			sqlsearch = sqlsearch & " and id = "&frectid&""
		end if

		sqlStr = "SELECT top 1" & vbcrlf
		sqlStr = sqlStr & " a.id, a.userID, a.userName, a.userLevel, a.replyUser, a.chargeid, a.qaDiv, a.title, a.regDate" + vbcrlf
		sqlStr = sqlStr & " ,a.orderSerial, a.itemID, a.extSiteName, a.isUsing, a.dispYN, a.usermail, a.userphone" + vbcrlf
		sqlStr = sqlStr & " ,a.emailok, a.contents, a.replytitle, a.replycontents, a.replydate, a.itemid" + vbcrlf
		sqlStr = sqlStr & " ,a.ExpectReplyDate,a.shopid, u.realnamecheck" + vbcrlf
		sqlStr = sqlStr & " FROM db_cs.dbo.tbl_myQna a" + vbcrlf
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_n as u" + vbcrlf
		sqlStr = sqlStr & " on a.userid=u.userid" + vbcrlf
		sqlStr = sqlStr & " WHERE 1=1 " & sqlsearch

        'response.write sqlStr&"<br>"
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        ftotalcount = rsget.RecordCount

        set FOneItem = new cmyqna_oneitem

        if Not rsget.Eof then
			FOneItem.fshopid = rsget("shopid")
			FOneItem.fid = rsget("id")
			FOneItem.fuserid = rsget("userid")
			FOneItem.fusername = db2html(rsget("username"))
			FOneItem.forderserial = rsget("orderserial")
			FOneItem.fqadiv = rsget("qadiv")
			FOneItem.ftitle = db2html(rsget("title"))
			FOneItem.fusermail = rsget("usermail")
			FOneItem.fuserphone = rsget("userphone")
			FOneItem.femailok = rsget("emailok")
			FOneItem.fcontents = db2html(rsget("contents"))
			FOneItem.fregdate = rsget("regdate")
			FOneItem.freplyuser = rsget("replyuser")
			FOneItem.Fchargeid = rsget("chargeid")
			FOneItem.freplytitle = db2html(rsget("replytitle"))
			FOneItem.freplycontents = db2html(rsget("replycontents"))
			FOneItem.freplydate = rsget("replydate")
			FOneItem.fitemid = rsget("itemid")
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fFextsitename = rsget("extsitename")
			FOneItem.fuserlevel = rsget("userlevel")
			FOneItem.fFExpectReplyDate = rsget("ExpectReplyDate")
			FOneItem.fFrealnamecheck = rsget("realnamecheck")
        end if
        rsget.Close
    end Sub

	'/common/offshop/board/online_cscenter_qna_list.asp
    public Sub fmembercheck()
        dim sqlStr , sqlsearch

		if frectssBctId <> "" then
			sqlsearch = sqlsearch & " and id = '"&frectssBctId&"'"
		end if

        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " id, company_name, level_sn,userdiv,bigo" + vbcrlf
		sqlStr = sqlStr & " from db_partner.dbo.tbl_partner" + vbcrlf
		sqlStr = sqlStr & " where isusing='Y' " & sqlsearch

        'response.write sqlStr&"<br>"
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        ftotalcount = rsget.RecordCount

        set FOneItem = new cmyqna_oneitem

        if Not rsget.Eof then
			FOneItem.fssBctId = rsget("id")
			FOneItem.fcompany_name = rsget("company_name")
			FOneItem.flevel_sn = rsget("level_sn")
			FOneItem.fuserdiv = rsget("userdiv")
			FOneItem.fbigo = rsget("bigo")
        end if
        rsget.Close
    end Sub

end class

'id, userid, username, orderserial, qadiv, title, usermail, userphone, emailok, contents, regdate, replyuser, chargeid, replytitle, replycontents, replydate, isusing
Class CMyQNAItem
	private Fid
	private Fuserid
	private Fusername
	private Forderserial
	private Fqadiv
	private Ftitle
	private Fusermail
	private Fuserphone
	private Femailok
	private Fcontents
	private Fregdate
	private Freplyuser
	private Fchargeid
	private Freplytitle
	private Freplycontents
	private Freplydate
	private Fdispyn
	private Fisusing
	public Fitemid
	public Fisupchebeasong
	public Fitemname
	public Fitemoption
	public Fitemoptionname
	public Fmakerid
	public Fdeliverytype
	public Frealnamecheck
	public FExtSiteName
	public FExpectReplyDate
	public Fuserlevel
	public fcompany_name
	public flevel_sn
	public fuserdiv
	public fssBctId
	public fshopid
	public FEvalPoint
	public Freplyqadiv
	public FattachFile
	public FattachFile2
	public fqadivname
	Public Fdevice
	Public FOS
	Public FOSetc
	Public Fsitename
	Public FuserGubun
	Public Fupchereplyuser
	Public Fupchereplycontents
	Public Fupchereplydate
	Public FtargetMakerID
	Public FsendYN
	Public Fupcheviewdate

	'/사용중지 공통함수에 공용펑션으로 쓸것		'/2016.06.30 한용민
    public function GetUserLevelStr()
        if IsNULL(Fuserlevel) then
    		GetUserLevelStr = "&nbsp;"
    	    Exit function
        end if

        Select Case CStr(Fuserlevel)
    		Case "5"
    			GetUserLevelStr = "<span class='member_orange'>ORANGE</span>"
    		Case "0"
    			GetUserLevelStr = "<span class='member_yellow'>YELLOW</span>"
    		Case "1"
    			GetUserLevelStr = "<span class='member_green'>GREEN</span>"
    		Case "2"
    			GetUserLevelStr = "<span class='member_blue'>BLUE</span>"
    		Case "3"
    			GetUserLevelStr = "<span class='member_vipsilver'>VIP SILVER</span>"
    			''GetUserLevelStr = "<span class='member_vip'>VIP</span>"
    		Case "4"
    			GetUserLevelStr = "<span class='member_vipgold'>VIP GOLD</span>"
    		Case "6"
    			GetUserLevelStr = "<span class='member_vvipgold'>VVIP</span>"
    		Case "7"
    			GetUserLevelStr = "<span class='member_staff'>STAFF</span>"
    		Case "8"
    			GetUserLevelStr = "<span class='member_red'>FAMILY</span>"
    		Case "9"
    			GetUserLevelStr = "<span class='member_red'>BIZ</span>"
    		Case Else
    			GetUserLevelStr = "<span class='member_orange'>ORANGE</span>"
    	end Select
	end function

    public function IsUpchebeasong()
    	if (Fdeliverytype="2") or (Fdeliverytype="5") or (Fdeliverytype="9") then
    		IsUpchebeasong = true
    	else
    		IsUpchebeasong = false
    	end if
	end function

	Property Get id()
		id = Fid
	end Property

	Property Get userid()
		userid = Fuserid
	end Property

	Property Get username()
		username = Fusername
	end Property

	Property Get orderserial()
		orderserial = Forderserial
	end Property

	Property Get qadiv()
		qadiv = Fqadiv
	end Property

	Property Get title()
		title = Ftitle
	end Property

	Property Get usermail()
		usermail = Fusermail
	end Property

	Property Get userphone()
		userphone = Fuserphone
	end Property

	Property Get emailok()
		emailok = Femailok
	end Property

	Property Get contents()
		contents = Fcontents
	end Property

	Property Get regdate()
		regdate = Fregdate
	end Property

	Property Get replyuser()
		replyuser = Freplyuser
	end Property

	Property Get chargeid()
		chargeid = Fchargeid
	end Property

	Property Get replytitle()
		replytitle = Freplytitle
	end Property

	Property Get replycontents()
		replycontents = Freplycontents
	end Property

	Property Get replydate()
		replydate = Freplydate
	end Property

	Property Get isusing()
		isusing = Fisusing
	end Property

	Property Get dispyn()
		dispyn = Fdispyn
	end Property

	Property Get itemid()
		itemid = Fitemid
	end Property

	Property Let id(byVal v)
		Fid = v
	end Property

	Property Let userid(byVal v)
		Fuserid = v
	end Property

	Property Let username(byVal v)
		Fusername = v
	end Property

	Property Let orderserial(byVal v)
		Forderserial = v
	end Property

	Property Let qadiv(byVal v)
		Fqadiv = v
	end Property

	Property Let title(byVal v)
		Ftitle = v
	end Property

	Property Let usermail(byVal v)
		Fusermail = v
	end Property

	Property Let userphone(byVal v)
		Fuserphone = v
	end Property

	Property Let emailok(byVal v)
		Femailok = v
	end Property

	Property Let contents(byVal v)
		Fcontents = v
	end Property

	Property Let regdate(byVal v)
		Fregdate = v
	end Property

	Property Let replyuser(byVal v)
		Freplyuser = v
	end Property

	Property Let chargeid(byVal v)
		Fchargeid = v
	end Property

	Property Let replytitle(byVal v)
		Freplytitle = v
	end Property

	Property Let replycontents(byVal v)
		Freplycontents = v
	end Property

	Property Let replydate(byVal v)
		Freplydate = v
	end Property

	Property Let isusing(byVal v)
		Fisusing = v
	end Property

	Property Let dispyn(byVal v)
		Fdispyn = v
	end Property

	Property Let itemid(byVal v)
		Fitemid = v
	end Property

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CMyQNA
    public results()
	public FItemList()
	public FOneItem
	public FCurrPage
	public FTotalPage
	public FTotalCount
	public FPageSize
	public FResultCount
	public FScrollCount
	private FIDBefore
	private FIDAfter
	public FSearchUserID
	public FSearchOrderSerial
	public FSearchItemID
	public FSearchNew
	public FRectItemNotInclude
	public FRectOnlyItemInclude
	public FRectDesigner
	public FSearchWriteId
	public FSearchChargeId
	public FSearchStartDate
	public FSearchEndDate
	Public FreplyDate1
	Public FreplyDate2
    public frectshopid
    public frectshopflg
	public FRectReplyQADiv
	public FSearchUserLevel
	public FRectEvalPoint
	public FSearchDiv
	public FSearchText
	public FRectIsUsing
	public FRectSiteName
	public FRectUpcheOnly
	public FRectTargetMakerID
	public FRectCurrState
	public FRectUserGubun
	public FRectItemID

	Private Sub Class_Initialize()
		redim results(0)
		FSearchUserID = ""
		FSearchOrderSerial = ""
		FSearchNew = "N"
	End Sub
	Private Sub Class_Terminate()
	End Sub

	Property Get CurrPage()
		CurrPage = FCurrPage
	end Property

	Property Get TotalPage()
		TotalPage = FTotalPage
	end Property

	Property Get TotalCount()
		TotalCount = FTotalCount
	end Property

	Property Get PageSize()
		PageSize = FPageSize
	end Property

	Property Get ResultCount()
		ResultCount = FResultCount
	end Property

	Property Get ScrollCount()
		ScrollCount = FScrollCount
	end Property

	Property Get IDBefore()
		IDBefore = FIDBefore
	end Property

	Property Get IDAfter()
		IDAfter = FIDAfter
	end Property

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = TotalPage > StartScrollPage + ScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((Currpage-1)\ScrollCount)*ScrollCount +1
	end Function

	Property Let CurrPage(byVal v)
		FCurrPage = v
	end Property

	Property Let PageSize(byVal v)
		FPageSize = v
	end Property

	Property Let ScrollCount(byVal v)
		FScrollCount = v
	end Property

	Property Let SearchUserID(byVal v)
		FSearchUserID = v
	end Property

	Property Let SearchOrderSerial(byVal v)
		FSearchOrderSerial = v
	end Property

	Property Let SearchNew(byVal v)
		FSearchNew = v
	end Property

	private FRectQadiv

	Property Get RectQadiv()
		RectQadiv = FRectQadiv
	end Property

	Property Let RectQadiv(byVal v)
		FRectQadiv = v
	end Property

	Public Sub GetExtOrder()
		dim sqlStr, i

		sqlStr = " select distinct top " & FPageSize & " m.orderserial, m.buyname "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_order].[dbo].[tbl_order_master] m "
		sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_detail] d on m.orderserial = d.orderserial "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and m.sitename = '" & FRectSiteName & "' "
		sqlStr = sqlStr + " 	and d.itemid = " & FRectItemID
		sqlStr = sqlStr + " 	and m.regdate >= DateAdd(day, -21, getdate()) "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	m.orderserial desc "
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if not rsget.EOF then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CMyQNAExtItem

				FItemList(i).Forderserial = rsget("orderserial")
				FItemList(i).Fbuyname = rsget("buyname")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	End Sub

    Public Function getLecQnalist()
    	dim sqlStr, i
    	sqlStr = " select count(q.id) as cnt from [db_cs].[dbo].tbl_myqna q, [db_contents].[dbo].tbl_lecture_item i "
    	sqlStr = sqlStr + " where (q.itemid=i.linkitemid) "
    	sqlStr = sqlStr + " and q.itemid<>0"
    	sqlStr = sqlStr + " and q.isusing = 'Y' "
    	if (FSearchNew = "Y") then
                sqlStr = sqlStr + " and (replyuser = '') "
        end if
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " m.id, m.userid, m.username, m.orderserial, m.qadiv, m.title, m.regdate, m.replyuser, m.isusing, m.itemid, m.extsitename,i.lecturer "
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_myqna m "
        sqlStr = sqlStr + " , [db_contents].[dbo].tbl_lecture_item i "
        sqlStr = sqlStr + " where m.itemid=i.linkitemid "
		sqlStr = sqlStr + " and m.itemid<>0"
		sqlStr = sqlStr + " and m.isusing = 'Y' "
        if (FSearchNew = "Y") then
                sqlStr = sqlStr + " and (replyuser = '') "
        end if

		sqlStr = sqlStr + " order by m.regdate desc "

		if FPageSize<>0 then
			rsget.pagesize = PageSize
		end if

        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        FTotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount - (CurrPage-1)*PageSize

        redim preserve results(FResultCount)

        if not rsget.EOF then
        i = 0
        rsget.absolutepage = FCurrPage
        do until ( rsget.eof or (i > FResultCount))

		set results(i) = new CMyQNAItem

            results(i).id = rsget("id")
            results(i).userid = rsget("userid")
            results(i).username = db2html(rsget("username"))
            results(i).orderserial = rsget("orderserial")
            results(i).qadiv = rsget("qadiv")
            results(i).title = db2html(rsget("title"))
            'results(i).usermail = rsget("usermail")
            'results(i).emailok = rsget("emailok")
            'results(i).contents = rsget("contents")
            results(i).regdate = rsget("regdate")
            results(i).replyuser = rsget("replyuser")
            'results(i).replytitle = rsget("replytitle")
            'results(i).replycontents = rsget("replycontents")
			results(i).isusing = rsget("isusing")
            results(i).Fitemid = rsget("itemid")
			results(i).Fmakerid = rsget("lecturer")
			'results(i).Fdeliverytype = rsget("deliverytype")
			results(i).FExtSiteName = rsget("extsitename")

		rsget.MoveNext
		i = i + 1
                loop
        end if
        rsget.close
    end function

	'//cscenter/board/cscenter_qna_board_list.asp
    Public Function fqnalist()
        dim sqlStr, i , sqlsearch

		if (FSearchDiv <> "") and (FSearchText <> "") then
			Select Case FSearchDiv
				Case "title"
					sqlsearch = sqlsearch + " and m.title like '%" + CStr(FSearchText) + "%' "
				Case "contents"
					sqlsearch = sqlsearch + " and convert(varchar(200), m.contents) like '%" + CStr(FSearchText) + "%' "
				Case "makerid"
					sqlsearch = sqlsearch + " and i.makerid = '" & FSearchText & "' "
				Case "username"
					sqlsearch = sqlsearch + " and m.username = '" & FSearchText & "' "
				Case Else
					''
			End Select
		end if

		if (FRectUpcheOnly = "Y") then
			sqlsearch = sqlsearch + " and m.qadiv in ('00', '02', '03', '04', '06', '07', '14', '20') "
		end if

		if frectshopflg = "Y" then
			sqlsearch = sqlsearch + " and m.shopid is not null"
		elseif frectshopflg = "N" then
			sqlsearch = sqlsearch + " and m.shopid is null"
		end if

        if FSearchUserID <> "" then
			sqlsearch = sqlsearch + " and m.userid = '" + FSearchUserID + "' "
        end if

        if FRectTargetMakerID <> "" then
			sqlsearch = sqlsearch + " and m.makerid = '" + FRectTargetMakerID + "' "
        end if

        if FSearchOrderSerial <> "" then
			sqlsearch = sqlsearch + " and m.orderserial = '" + FSearchOrderSerial + "'"
        end if

        if FSearchItemID <> "" then
			sqlsearch = sqlsearch + " and i.itemid = '" + FSearchItemID + "'"
        end if

        IF FSearchWriteId<>"" Then
			sqlsearch = sqlsearch + " and replyuser = '" + FSearchWriteId + "' "
		End IF

		if (FRectEvalPoint <> "") then
			sqlsearch = sqlsearch + " and IsNull(m.EvalPoint, 0) > 0 "
			if (FRectEvalPoint = "3DN") then
				sqlsearch = sqlsearch + " and IsNull(m.EvalPoint, 0) <= 3 "
			else
				sqlsearch = sqlsearch + " and IsNull(m.EvalPoint, 0) = " + CStr(FRectEvalPoint) + " "
			end if
		end if

        if FRectReplyQADiv<>"" then
			if (FRectReplyQADiv = "all") then
				sqlsearch = sqlsearch + " and IsNull(replyqadiv, '01') in ('02', '03', '10', '99') "
			else
				sqlsearch = sqlsearch + " and IsNull(replyqadiv, '01') = '" + replyqadiv + "' "
			end if

		end if

		'//당담자 검색일 경우
        IF FSearchChargeId<>"" Then
			sqlsearch = sqlsearch + " and IsNull(m.chargeid, '') = '" + FSearchChargeId + "' "

			'/문의자가 직원일경우와 고객건의사항 문의의 경우 담당자 분배를 안함
			sqlsearch = sqlsearch + " and not(m.userlevel=7 or m.qadiv=26)"
		End IF

        if (FSearchNew = "N") then
			sqlsearch = sqlsearch + " and replyDate is null"
		elseif (FSearchNew = "VV") then
			sqlsearch = sqlsearch + " and replyDate is null and m.userlevel in (6) "
        elseif (FSearchNew = "VE") then
			sqlsearch = sqlsearch + " and replyDate is null and m.userlevel in (6) and qadiv <> '00' "
        elseif (FSearchNew = "VD") then
			sqlsearch = sqlsearch + " and replyDate is null and m.userlevel in (6) and qadiv  = '00' "
		elseif (FSearchNew = "V") then
			sqlsearch = sqlsearch + " and replyDate is null and m.userlevel in (3,4) "
        elseif (FSearchNew = "E") then
			sqlsearch = sqlsearch + " and replyDate is null and m.userlevel in (3,4) and qadiv <> '00' "
        elseif (FSearchNew = "D") then
			sqlsearch = sqlsearch + " and replyDate is null and m.userlevel in (3,4) and qadiv  = '00' "
        elseif (FSearchNew = "UN") then
			sqlsearch = sqlsearch + " and replyDate is null and upcheReplyDate is null "
        end if

		if FRectCurrState <> "" then
			select case FRectCurrState
				case "B001":
					sqlsearch = sqlsearch + " and replyDate is null "
				case "B006":
					sqlsearch = sqlsearch + " and replyDate is null and upcheReplyDate is not null "
				case "B007":
					sqlsearch = sqlsearch + " and replyDate is not null "
				case "B008":
					sqlsearch = sqlsearch + " and replyDate is not null and s.cs_status <= '002' "
				case "B009":
					sqlsearch = sqlsearch + " and replyDate is not null and s.cs_status > '002' "
				case else
					'//
			end select
		end if

		if FSearchUserLevel <> "" then
			sqlsearch = sqlsearch + " and m.userlevel = '" & FSearchUserLevel & "'"
		end if

		if FRectQadiv <> "" then
			sqlsearch = sqlsearch + " and m.qadiv = '" & FRectQadiv & "'"

			'/고객건의사항 문의의 경우 직원은 제외함
			sqlsearch = sqlsearch + " and not(m.userlevel=7)"
		end if

		' 등록일기준
		if FSearchStartDate<>"" then
			sqlsearch = sqlsearch + " and m.regdate >='"& FSearchStartDate & "'"
		end if
		if FSearchEndDate<>"" then
			sqlsearch = sqlsearch + " and m.regdate < '"& dateAdd("d",1,FSearchEndDate)  & "'"
		end if

		' 답변일기준
		if FreplyDate1<>"" then
			sqlsearch = sqlsearch + " and replyDate >='"& FreplyDate1 & "'"
		end if
		if FreplyDate2<>"" then
			sqlsearch = sqlsearch + " and replyDate < '"& dateAdd("d",1,FreplyDate2)  & "'"
		end if

		if frectshopid <> "" then
			sqlsearch = sqlsearch + " and shopid = '"&frectshopid&"'"
		end if

		if (FRectIsUsing <> "") then
			sqlsearch = sqlsearch + " and m.isusing = '" & FRectIsUsing & "'"
			if (FRectIsUsing = "Y") then
				sqlsearch = sqlsearch + " and m.dispyn = 'Y'"
			end if
		end if

		if (FRectSiteName <> "") then
			if (FRectSiteName = "extall") then
				sqlsearch = sqlsearch + " and m.sitename <> '10x10' "
			else
				sqlsearch = sqlsearch + " and m.sitename = '" & FRectSiteName & "'"
			end if
		end if

		if (FRectUserGubun <> "") then
			if (FRectUserGubun = "C") then
				sqlsearch = sqlsearch + " and m.userGubun = 'C' "
			else
				sqlsearch = sqlsearch + " and m.userGubun = 'M' "
			end if
		end if

		sqlStr = " select count(id) as cnt" & vbcrlf
		sqlStr = sqlStr + " from [db_cs].[dbo].[tbl_myqna] as m with (noLock) " & vbcrlf
		if (FSearchDiv = "makerid") then
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i" & vbcrlf
			sqlStr = sqlStr + "	on m.itemid=i.itemid" & vbcrlf
		end if
		if (FRectCurrState = "B008") or (FRectCurrState = "B009") then
			sqlStr = sqlStr + " left join [db_temp].[dbo].[tbl_Sabannet_Detail] s with (noLock) on m.extQnaIdx = s.idx " & vbcrlf
		end if
		''sqlStr = sqlStr + " where m.isusing = 'Y' " & sqlsearch
		sqlStr = sqlStr + " where 1=1 " & sqlsearch

		''response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

        sqlStr = " select top " + CStr(FPageSize*FCurrPage)
        sqlStr = sqlStr + " m.id, m.userid, m.username,m.orderserial, m.qadiv, m.title, m.regdate, m.replyuser, m.chargeid, m.replyDate" & vbcrlf
        sqlStr = sqlStr + " ,m.isusing, m.dispyn,m.itemid, m.extsitename, m.userlevel ,m.shopid, i.makerid, i.deliverytype " & vbcrlf
        sqlStr = sqlStr + " ,d.itemname, d.itemoption, d.itemoptionname, IsNull(m.EvalPoint, 0) as EvalPoint, IsNull(m.replyqadiv, '01') as replyqadiv" & vbcrlf
        sqlStr = sqlStr + " , IsNull(m.attach01, '') as attachFile, c.comm_name as qadivname, m.sitename, m.userGubun " & vbcrlf
		sqlStr = sqlStr + " , m.upchereplyuser, m.upchereplyDate, m.makerid as targetMakerID " & vbcrlf	'// 001 : 접수, 002 : 답변완료, 003 : 답변전송, 004 : 강제저장(&전송)
		sqlStr = sqlStr + " , m.upcheviewdate " & vbcrlf
		if (FRectCurrState = "B008") or (FRectCurrState = "B009") then
			sqlStr = sqlStr + " , (case when replyDate is not null and s.cs_status > '002' then 'Y' else '' end) as sendYN " & vbcrlf
		elseif FRectSiteName = "10x10" then
			sqlStr = sqlStr + " , '' as sendYN " & vbcrlf
		else
			sqlStr = sqlStr + " , 'N/A' as sendYN " & vbcrlf
		end if
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_myqna as m with (noLock) " & vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i" & vbcrlf
		sqlStr = sqlStr + "	on m.itemid=i.itemid" & vbcrlf
		sqlStr = sqlStr + " left join db_order.dbo.tbl_order_detail d " & vbcrlf
		sqlStr = sqlStr + " 	on m.orderdetailidx = d.idx " & vbcrlf
		sqlStr = sqlStr + " left join [db_cs].[dbo].[tbl_cs_comm_code] c" & vbcrlf
		sqlStr = sqlStr + " 	on m.qadiv = right(c.comm_cd,2)" & vbcrlf
		sqlStr = sqlStr + " 	and c.comm_isdel='N'" & vbcrlf
		sqlStr = sqlStr + " 	and left(c.comm_group,3)='D00'" & vbcrlf
		if (FRectCurrState = "B008") or (FRectCurrState = "B009") then
			sqlStr = sqlStr + " left join [db_temp].[dbo].[tbl_Sabannet_Detail] s with (noLock) on m.extQnaIdx = s.idx " & vbcrlf
		end if
		''sqlStr = sqlStr + " where m.isusing = 'Y' " & sqlsearch
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		If (FRectSiteName = "nvstorefarm") Then
			sqlStr = sqlStr + " order by m.id desc "
		Else
			sqlStr = sqlStr + " order by m.regdate desc "
		End If

		if FPageSize<>0 then
			rsget.pagesize = PageSize
		end if

		''response.write sqlStr & "<br>"
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount - (CurrPage-1)*PageSize

        if (FResultCount<1) then FResultCount=0

        redim preserve results(FResultCount)

        if not rsget.EOF then
            i = 0
            rsget.absolutepage = FCurrPage
			do until ( rsget.eof or (i > FResultCount))
			set results(i) = new CMyQNAItem

				results(i).fshopid = rsget("shopid")
                results(i).id = rsget("id")
                results(i).userid = rsget("userid")
                results(i).username = db2html(rsget("username"))
                results(i).orderserial = rsget("orderserial")
                results(i).qadiv = rsget("qadiv")
                results(i).title = db2html(rsget("title"))
                results(i).regdate = rsget("regdate")
                results(i).replyuser = rsget("replyuser")
                results(i).chargeid = rsget("chargeid")
                results(i).replyDate = rsget("replyDate")
				results(i).dispyn 	= rsget("dispyn")
                results(i).isusing = rsget("isusing")
                results(i).Fitemid = rsget("itemid")

                if (results(i).Fitemid <> 0) then
					results(i).Fmakerid = rsget("makerid")
					results(i).Fdeliverytype = rsget("deliverytype")

					results(i).Fitemname = db2html(rsget("itemname"))
					results(i).Fitemoption = rsget("itemoption")
					results(i).Fitemoptionname = db2html(rsget("itemoptionname"))
                end if

				results(i).FExtSiteName = rsget("extsitename")
				results(i).Fuserlevel = rsget("userlevel")
				results(i).FEvalPoint = rsget("EvalPoint")
				results(i).Freplyqadiv = rsget("replyqadiv")
				results(i).FattachFile = rsget("attachFile")
				results(i).fqadivname = db2html(rsget("qadivname"))

				'/글쓴이가 직원이거나 고객건의사항일경우 담당자 배정 안함으로 뿌림
				if results(i).Fuserlevel="7" or results(i).qadiv="26" then
					results(i).chargeid=""
				end if

				results(i).Fsitename = rsget("sitename")
				results(i).FuserGubun = rsget("userGubun")

				results(i).Fupchereplyuser = rsget("upchereplyuser")
				results(i).Fupchereplydate = rsget("upchereplydate")
				results(i).FtargetMakerID = rsget("targetMakerID")
				results(i).FsendYN = rsget("sendYN")
				results(i).Fupcheviewdate = rsget("upcheviewdate")
		    rsget.MoveNext
		    i = i + 1
            loop
        end if
        rsget.close
	end Function

    Public Function list()
        dim sqlStr, i , sqlsearch

		if (FSearchDiv <> "") and (FSearchText <> "") then
			Select Case FSearchDiv
				Case "title"
					sqlsearch = sqlsearch + " and m.title like '%" + CStr(FSearchText) + "%' "
				Case "contents"
					sqlsearch = sqlsearch + " and convert(varchar(200), m.contents) like '%" + CStr(FSearchText) + "%' "
				Case Else
					''
			End Select
		end if

		if frectshopflg = "Y" then
			sqlsearch = sqlsearch + " and m.shopid is not null"
		elseif frectshopflg = "N" then
			sqlsearch = sqlsearch + " and m.shopid is null"
		end if

        if FSearchUserID <> "" then
			sqlsearch = sqlsearch + " and m.userid = '" + FSearchUserID + "' "
        end if

        if FSearchOrderSerial <> "" then
			sqlsearch = sqlsearch + " and m.orderserial = '" + FSearchOrderSerial + "'"
        end if

        IF FSearchWriteId<>"" Then
			sqlsearch = sqlsearch + " and replyuser = '" + FSearchWriteId + "' "
		End IF

		if (FRectEvalPoint <> "") then
			sqlsearch = sqlsearch + " and IsNull(m.EvalPoint, 0) > 0 "
			if (FRectEvalPoint = "3DN") then
				sqlsearch = sqlsearch + " and IsNull(m.EvalPoint, 0) <= 3 "
			else
				sqlsearch = sqlsearch + " and IsNull(m.EvalPoint, 0) = " + CStr(FRectEvalPoint) + " "
			end if
		end if

        if FRectReplyQADiv<>"" then
			if (FRectReplyQADiv = "all") then
				sqlsearch = sqlsearch + " and IsNull(replyqadiv, '01') in ('02', '03', '10', '99') "
			else
				sqlsearch = sqlsearch + " and IsNull(replyqadiv, '01') = '" + replyqadiv + "' "
			end if

		end if

        IF FSearchChargeId<>"" Then
			sqlsearch = sqlsearch + " and IsNull(chargeid, '') = '" + FSearchChargeId + "' "
		End IF

        if (FSearchNew = "N") then
			sqlsearch = sqlsearch + " and replyDate is null"
		elseif (FSearchNew = "VV") then
			sqlsearch = sqlsearch + " and replyDate is null and m.userlevel in (6) "
        elseif (FSearchNew = "VE") then
			sqlsearch = sqlsearch + " and replyDate is null and m.userlevel in (6) and qadiv <> '00' "
        elseif (FSearchNew = "VD") then
			sqlsearch = sqlsearch + " and replyDate is null and m.userlevel in (6) and qadiv  = '00' "
		elseif (FSearchNew = "V") then
			sqlsearch = sqlsearch + " and replyDate is null and m.userlevel in (3,4) "
        elseif (FSearchNew = "E") then
			sqlsearch = sqlsearch + " and replyDate is null and m.userlevel in (3,4) and qadiv <> '00' "
        elseif (FSearchNew = "D") then
			sqlsearch = sqlsearch + " and replyDate is null and m.userlevel in (3,4) and qadiv  = '00' "
        end if

		if FSearchUserLevel <> "" then
			sqlsearch = sqlsearch + " and m.userlevel = '" & FSearchUserLevel & "'"
		end if

		if FRectQadiv <> "" then
			sqlsearch = sqlsearch + " and m.qadiv = '" & FRectQadiv & "'"
		end if

		'if (FRectItemNotInclude<>"") then
			'sqlsearch = sqlsearch + " and ((itemid=0) or (itemid is NULL) or (qadiv not in ('02','03')))"
		'end if

		'if (FRectOnlyItemInclude<>"") then
			'sqlsearch = sqlsearch + " and ((itemid>0))"
		'end if

		' 등록일기준
		if FSearchStartDate<>"" then
			sqlsearch = sqlsearch + " and m.regdate >='"& FSearchStartDate & "'"
		end if
		if FSearchEndDate<>"" then
			sqlsearch = sqlsearch + " and m.regdate < '"& dateAdd("d",1,FSearchEndDate)  & "'"
		end if

		' 답변일기준
		if FreplyDate1<>"" then
			sqlsearch = sqlsearch + " and replyDate >='"& FreplyDate1 & "'"
		end if
		if FreplyDate2<>"" then
			sqlsearch = sqlsearch + " and replyDate < '"& dateAdd("d",1,FreplyDate2)  & "'"
		end if

		if frectshopid <> "" then
			sqlsearch = sqlsearch + " and shopid = '"&frectshopid&"'"
		end if

		sqlStr = " select count(id) as cnt" & vbcrlf
		sqlStr = sqlStr + " from [db_cs].[dbo].[tbl_myqna] as m" & vbcrlf
		''sqlStr = sqlStr + " left join [db_user].[dbo].[tbl_logindata] as l on m.userid = l.userid " & vbcrlf
		sqlStr = sqlStr + " where isusing = 'Y' " & sqlsearch
		''rw sqlStr

		rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

        sqlStr = " select top " + CStr(FPageSize*FCurrPage)
        sqlStr = sqlStr + " m.id, m.userid, m.username,m.orderserial, m.qadiv, m.title, m.regdate, m.replyuser, m.chargeid, m.replyDate" & vbcrlf
        sqlStr = sqlStr + " ,m.isusing, m.dispyn,m.itemid, m.extsitename, m.userlevel ,m.shopid, i.makerid, i.deliverytype " & vbcrlf
        sqlStr = sqlStr + " ,d.itemname, d.itemoption, d.itemoptionname, IsNull(m.EvalPoint, 0) as EvalPoint, IsNull(m.replyqadiv, '01') as replyqadiv" & vbcrlf
        sqlStr = sqlStr + " , IsNull(m.attach01, '') as attachFile, c.comm_name as qadivname, m.sitename, m.userGubun" & vbcrlf
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_myqna as m" & vbcrlf
		''sqlStr = sqlStr + " left join [db_user].[dbo].[tbl_logindata] as l on m.userid = l.userid " & vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i" & vbcrlf
		sqlStr = sqlStr + "	on m.itemid=i.itemid" & vbcrlf
		sqlStr = sqlStr + " left join db_order.dbo.tbl_order_detail d " & vbcrlf
		sqlStr = sqlStr + " 	on m.orderdetailidx = d.idx " & vbcrlf
		sqlStr = sqlStr + " left join [db_cs].[dbo].[tbl_cs_comm_code] c" & vbcrlf
		sqlStr = sqlStr + " 	on m.qadiv = right(c.comm_cd,2)" & vbcrlf
		sqlStr = sqlStr + " 	and c.comm_isdel='N'" & vbcrlf
		sqlStr = sqlStr + " 	and left(c.comm_group,3)='D00'" & vbcrlf
        sqlStr = sqlStr + " where m.isusing = 'Y' " & sqlsearch
        sqlStr = sqlStr + " order by m.regdate desc "
		''response.write sqlStr

		if FPageSize<>0 then
			rsget.pagesize = PageSize
		end if

		'response.write sqlStr &"<br>"
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount - (CurrPage-1)*PageSize

		'if (FResultCount > PageSize) then
		'	FResultCount = PageSize
		'end if
        if (FResultCount<1) then FResultCount=0

        redim preserve results(FResultCount)

        if not rsget.EOF then
            i = 0
            rsget.absolutepage = FCurrPage
			do until ( rsget.eof or (i > FResultCount))
			set results(i) = new CMyQNAItem
				results(i).fshopid = rsget("shopid")
                results(i).id = rsget("id")
                results(i).userid = rsget("userid")
                results(i).username = db2html(rsget("username"))
                results(i).orderserial = rsget("orderserial")
                results(i).qadiv = rsget("qadiv")
                results(i).title = db2html(rsget("title"))
                'results(i).usermail = rsget("usermail")
                'results(i).emailok = rsget("emailok")
                'results(i).contents = rsget("contents")
                results(i).regdate = rsget("regdate")
                results(i).replyuser = rsget("replyuser")
                results(i).chargeid = rsget("chargeid")
                results(i).replyDate = rsget("replyDate")
                'results(i).replytitle = rsget("replytitle")
                'results(i).replycontents = rsget("replycontents")
				results(i).dispyn 	= rsget("dispyn")
                results(i).isusing = rsget("isusing")
                results(i).Fitemid = rsget("itemid")

                if (results(i).Fitemid <> 0) then
					results(i).Fmakerid = rsget("makerid")
					results(i).Fdeliverytype = rsget("deliverytype")

					results(i).Fitemname = db2html(rsget("itemname"))
					results(i).Fitemoption = rsget("itemoption")
					results(i).Fitemoptionname = db2html(rsget("itemoptionname"))
                end if

				results(i).FExtSiteName = rsget("extsitename")
				results(i).Fuserlevel = rsget("userlevel")
				results(i).FEvalPoint = rsget("EvalPoint")
				results(i).Freplyqadiv = rsget("replyqadiv")
				results(i).FattachFile = rsget("attachFile")
				results(i).fqadivname = db2html(rsget("qadivname"))
		    rsget.MoveNext
		    i = i + 1
            loop
        end if
        rsget.close
	end Function

	' QNA목록 Prepared Statement 버전
	' TODO : 검색조건중 FSearchChargeId 무시됨!!!
    public Sub list2()
		'response.write "에러 : 시스템팀 문의"
		'response.end

        Dim i
		Dim paramInfo
		Dim sqlStr, sqlColumn, sqlTable, sqlWhere, sqlOrder, sqlGroup	' 쿼리문 변수 선언

		sqlWhere = " and m.isusing = 'Y' "
        if (FSearchNew = "Y") then
            sqlWhere = sqlWhere + " and replyDate is null "
        end if

        if (FSearchUserID <> "") then
            sqlWhere = sqlWhere + " and m.userid = ? "
			Call redimParam(paramInfo, "@userid"		, adVarchar	, adParamInput	, 32	, FSearchUserID)
        end if

        if (FSearchOrderSerial <> "") then
            sqlWhere = sqlWhere + " and m.orderserial = ? "
			Call redimParam(paramInfo, "@orderserial"		, adVarchar	, adParamInput	, 11	, FSearchOrderSerial)
        end if

		IF FSearchWriteId<>"" Then
            sqlWhere = sqlWhere + " and m.replyuser = ? "
			Call redimParam(paramInfo, "@replyuser"		, adVarchar	, adParamInput	, 32	, FSearchWriteId)
		End IF

        IF FSearchChargeId<>"" Then
            sqlWhere = sqlWhere + " and m.chargeid = ? "
			Call redimParam(paramInfo, "@chargeid"		, adVarchar	, adParamInput	, 32	, FSearchChargeId)
		End IF

		if (FRectQadiv <> "") then
            sqlWhere = sqlWhere + " and m.qadiv = ? "
			Call redimParam(paramInfo, "@qadiv"		, adVarchar	, adParamInput	, 2	, FRectQadiv)
		end if

		' 등록일기준
		if (FSearchStartDate<>"") then
            sqlWhere = sqlWhere + " and m.regdate >= ? "
			Call redimParam(paramInfo, "@regdate1"		, adVarchar	, adParamInput	, 10		, FSearchStartDate)
		end if
		if (FSearchEndDate<>"") then
            sqlWhere = sqlWhere + " and m.regdate < ? "
			Call redimParam(paramInfo, "@regdate2"		, adVarchar	, adParamInput	, 10		, dateAdd("d",1,FSearchEndDate))
		end if

		' 답변일기준
		if (FreplyDate1<>"") then
            sqlWhere = sqlWhere + " and m.replyDate >= ? "
			Call redimParam(paramInfo, "@replyDate1"		, adVarchar	, adParamInput	, 10		, FreplyDate1)
		end if
		if (FreplyDate2<>"") then
            sqlWhere = sqlWhere + " and m.replyDate < ? "
			Call redimParam(paramInfo, "@replyDate2"		, adVarchar	, adParamInput	, 10		, dateAdd("d",1,FreplyDate2))
		end if

		' 쿼리문 조합용 변수 설정
		sqlColumn = " m.id, m.userid, m.username, m.orderserial, m.qadiv, m.title, m.regdate, m.replyuser" & vbcrlf
        sqlColumn = sqlColumn & " , m.replyDate, m.itemid, m.extsitename, m.userlevel,  m.isusing, m.dispyn, m.chargeid" & vbcrlf
        sqlColumn = sqlColumn & " , IsNull(m.attach01, '') as attachFile" & vbcrlf
        sqlTable = " from [db_cs].[dbo].tbl_myqna m" & vbcrlf
        sqlOrder = " order by m.id desc" & vbcrlf

		sqlStr = makeQuery("", sqlTable, sqlWhere, sqlOrder, "", "", "")	' 카운트 쿼리
		''response.write sqlColumn & "<Br>"
		Call RecordSQL(sqlStr, paramInfo)

		If Not rsget.EOF Then
			FTotalCount = rsget(0)
		End If
		rsget.Close

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		sqlStr = makeQuery(sqlColumn, sqlTable, sqlWhere, sqlOrder, FCurrPage, FPageSize, "")	' 페이징 쿼리
		''response.write sqlStr
		Call RecordSQL(sqlStr, paramInfo)

		i=0
		if  not rsget.EOF  then
		do until rsget.eof
			redim preserve results(i)
			set results(i) = new CMyQNAItem

			results(i).id = rsget("id")
			results(i).userid = rsget("userid")
			results(i).username = db2html(rsget("username"))
			results(i).orderserial = rsget("orderserial")
			results(i).qadiv = rsget("qadiv")
			results(i).title = db2html(rsget("title"))
			'results(i).usermail = rsget("usermail")
			'results(i).emailok = rsget("emailok")
			'results(i).contents = rsget("contents")
			results(i).regdate = rsget("regdate")
			results(i).replyuser = rsget("replyuser")
			results(i).replyDate = rsget("replyDate")
			'results(i).replytitle = rsget("replytitle")
			'results(i).replycontents = rsget("replycontents")
			results(i).dispyn 	= rsget("dispyn")
			results(i).isusing = rsget("isusing")
			results(i).Fitemid = rsget("itemid")
			'results(i).Fmakerid = rsget("makerid")
			'results(i).Fdeliverytype = rsget("deliverytype")
			results(i).FExtSiteName = rsget("extsitename")
			results(i).Fuserlevel = rsget("userlevel")
			results(i).chargeid = rsget("chargeid")
			results(i).FattachFile = rsget("attachFile")

			i=i+1
			rsget.moveNext
		loop
		end If
		FREsultCount = i
		rsget.Close
    end Sub

	' Q&A 목록 SP버전
	Public Function list_NewVer(ByVal replyEnd)
		response.write "에러 : 시스템팀 문의"
		response.end

		Dim i, sqlStr, objRs ,paramInfo

		paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
			,Array("@PageSize"		, adInteger	, adParamInput	,		, FPageSize)	_
			,Array("@CurrPage"		, adInteger	, adParamInput	,		, FCurrPage) _
			,Array("@searchSDate"	, adVarchar	, adParamInput	, 10	, FSearchStartDate) _
			,Array("@searchEDate"	, adVarchar	, adParamInput	, 10	, FSearchEndDate) _
			,Array("@qaDiv"			, adVarchar	, adParamInput	, 2		, FRectQadiv) _
			,Array("@askUserID"		, adVarchar	, adParamInput	, 32	, FSearchUserID) _
			,Array("@replyUserID"	, adVarchar	, adParamInput	, 32	, FSearchWriteId) _
			,Array("@chargeid"		, adVarchar	, adParamInput	, 32	, FSearchChargeId) _
			,Array("@orderSerial"	, adVarchar	, adParamInput	, 11	, FSearchOrderSerial) _
			,Array("@replyEnd"		, adVarchar	, adParamInput	, 1		, replyEnd) _
		)

		sqlStr = "db_cs.dbo.sp_Ten_QnaList"
		Call fnExecSPReturnRSOutput(sqlStr, paramInfo)

		FTotalCount = CDbl(GetValue(paramInfo, "@RETURN_VALUE"))	' 토탈카운트
		FtotalPage  = Int ( (FTotalCount - 1) / FPageSize ) + 1
		If FTotalCount = 0 Then	FtotalPage = 1

		i=0
		if  not rsget.EOF  then
			do until rsget.eof

			redim preserve results(i)
			set results(i) = new CMyQNAItem

			results(i).id = rsget("id")
			results(i).userid = rsget("userid")
			results(i).username = db2html(rsget("username"))
			results(i).orderserial = rsget("orderserial")
			results(i).qadiv = rsget("qadiv")
			results(i).title = db2html(rsget("title"))
			'results(i).usermail = rsget("usermail")
			'results(i).emailok = rsget("emailok")
			'results(i).contents = rsget("contents")
			results(i).regdate = rsget("regdate")
			results(i).replyuser = rsget("replyuser")
			'results(i).replytitle = rsget("replytitle")
			'results(i).replycontents = rsget("replycontents")
			results(i).dispyn 	= rsget("dispyn")
			results(i).isusing = rsget("isusing")
			results(i).Fitemid = rsget("itemid")
			'results(i).Fmakerid = rsget("makerid")
			'results(i).Fdeliverytype = rsget("deliverytype")
			results(i).FExtSiteName = rsget("extsitename")
			results(i).Fuserlevel = rsget("userlevel")

			i=i+1
			rsget.moveNext
			loop
		end if
		FResultCount = i
		rsget.Close
	end Function

	Public Function read(ByVal PKID)
		Set results(0) = new CMyQNAItem

		If PKID <> "" Then
			Dim i, sqlStr
			Dim paramInfo
			paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
				,Array("@PKID"		, adInteger	, adParamInput	,		, PKID)	_
			)
			sqlStr = "db_cs.dbo.sp_Ten_QnaOne_New"
			Call fnExecSPReturnRSOutput(sqlStr, paramInfo)

			Dim regEx
			Set regEx = New RegExp

			With regEx
				.Pattern = "[ ]{2,}"
				.IgnoreCase = True
				.Global = True
				.MultiLine = True
			End With

			If Not rsget.EOF Then

				results(0).id = rsget("id")
				results(0).userid = rsget("userid")
				results(0).username = db2html(rsget("username"))
				results(0).orderserial = rsget("orderserial")
				results(0).qadiv = rsget("qadiv")
				results(0).title = db2html(rsget("title"))
				results(0).usermail = rsget("usermail")
				results(0).userphone = rsget("userphone")
				results(0).emailok = rsget("emailok")
				results(0).contents = db2html(rsget("contents"))
				results(0).regdate = rsget("regdate")
				results(0).replyuser = rsget("replyuser")
				results(0).replytitle = db2html(rsget("replytitle"))
				results(0).replycontents = db2html(rsget("replycontents"))
				results(0).replydate = rsget("replydate")
				results(0).Fmakerid = rsget("makerid")
				results(0).itemid = rsget("itemid")
				results(0).Fisupchebeasong = rsget("isupchebeasong")
				results(0).isusing = rsget("isusing")

				results(0).Fextsitename = rsget("extsitename")

				results(0).Fuserlevel = rsget("userlevel")
				results(0).FExpectReplyDate = rsget("ExpectReplyDate")
				results(0).Frealnamecheck = rsget("realnamecheck")

				results(0).Fitemname = db2html(rsget("itemname"))
				results(0).Fitemoption = rsget("itemoption")
				results(0).Fitemoptionname = db2html(rsget("itemoptionname"))

				results(0).FEvalPoint = rsget("EvalPoint")
				results(0).Freplyqadiv = rsget("replyqadiv")

				results(0).FattachFile = rsget("attachFile")
				results(0).FattachFile2 = rsget("attachFile2")
				results(0).Fdevice = rsget("device")
				results(0).FOS = rsget("OS")
				results(0).FOSetc = rsget("OSetc")

				results(i).Fsitename = rsget("sitename")
				results(i).FuserGubun = rsget("userGubun")

				if IsNull(results(i).Fsitename) then
					results(i).Fsitename = "10x10"
				end if

				if (results(i).Fsitename <> "") then
					results(i).contents = regEx.Replace(results(i).contents, vbCrLf)
				end if

				if IsNull(results(0).orderserial) then
					results(0).orderserial = ""
				end if

				results(0).Fupchereplyuser = rsget("upchereplyuser")
				results(0).Fupchereplycontents = rsget("upchereplycontents")
				results(0).Fupchereplydate = rsget("upchereplydate")
				results(0).FtargetMakerID = rsget("targetMakerID")
				results(0).Fupcheviewdate = rsget("upcheviewdate")
                results(0).chargeid = rsget("chargeid")
			End If
			rsget.close
			set regEx = nothing
		End If
	end Function


	' 답변, 유형수정
    Public Function BackProcData(ByVal mode)
		Dim ErrCode, ErrMsg

        rw results(0).qaDiv
		Dim sqlStr
		Dim paramInfo
		paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
			,Array("@mode"			, adVarchar	, adParamInput	, 10	, mode)	_

			,Array("@id"			, adInteger	, adParamInput	, 4	, results(0).id) _
			,Array("@qadiv"			, adChar	, adParamInput	, 2	, results(0).qaDiv) _
			,Array("@replyUser"		, adVarchar	, adParamInput	, 32	, results(0).replyUser) _

			,Array("@replyTitle"	, adVarchar	, adParamInput	, 128	, results(0).replyTitle) _
			,Array("@replyContents"	, adVarchar	, adParamInput	, 8000	, results(0).replyContents) _
			,Array("@MD5KEY"		, adVarchar	, adParamInput	, 32	, MD5(results(0).id) ) _
			,Array("@replyqadiv"	, adVarchar	, adParamInput	, 2	, results(0).Freplyqadiv ) _
		)

		sqlStr = "db_cs.dbo.sp_Ten_QnaProc"
		Call fnExecSP(sqlStr, paramInfo)
	End Function

	'/사용안함(DB화 시킴) 용만
	Public Function code2name(byval v)
        if (v = "00") then
                code2name = "배송문의"
        elseif (v = "01") then
                code2name = "주문문의"
        elseif (v = "02") then
                code2name = "상품문의"
        elseif (v = "03") then
                code2name = "재고문의"
        elseif (v = "04") then
                code2name = "취소문의"
        elseif (v = "05") then
                code2name = "환불문의"
        elseif (v = "06") then
                code2name = "교환문의"
        elseif (v = "07") then
                code2name = "AS문의"
        elseif (v = "08") then
                code2name = "이벤트문의"
        elseif (v = "09") then
                code2name = "증빙서류문의"
        elseif (v = "10") then
                code2name = "시스템문의"
        elseif (v = "11") then
                code2name = "회원제도문의"
        elseif (v = "12") then
                code2name = "회원정보문의"
        elseif (v = "13") then
                code2name = "당첨문의"
        elseif (v = "14") then
                code2name = "반품문의"
        elseif (v = "15") then
                code2name = "입금문의"
        elseif (v = "16") then
                code2name = "오프라인문의"
        elseif (v = "17") then
                code2name = "쿠폰/마일리지"
        elseif (v = "18") then
                code2name = "결제방법문의"
        elseif (v = "20") then
                code2name = "기타문의"
        elseif (v = "21") then
                code2name = "아이띵소문의"
        elseif (v = "22") then
                code2name = "이벤트문의()"
        elseif (v = "23") then
                code2name = "사은품문의"
        elseif (v = "24") then
                code2name = "POINT1010문의"
        elseif (v = "25") then
                code2name = "선물포장문의"
        else
                code2name = ""
        end if
	end Function

end Class

Class CMyQNAOrderInfo
	private Fcount
	private Ftotalprice
	private FMcount
	private FMtotalprice

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	Property Get OrderCount()
		OrderCount = Fcount
	end Property

	Property Get TotalPrice()
		TotalPrice = Ftotalprice
	end Property

	Property Get MOrderCount()
		MOrderCount = FMcount
	end Property

	Property Get MTotalPrice()
		MTotalPrice = FMtotalprice
	end Property

	Property Let OrderCount(byVal v)
		Fcount = v
	end Property

	Property Let TotalPrice(byVal v)
		Ftotalprice = v
	end Property

	Property Let MOrderCount(byVal v)
		FMcount = v
	end Property

	Property Let MTotalPrice(byVal v)
		FMtotalprice = v
	end Property

	Public Function UserOrderInfo(byval userid)
		dim sqlStr

		sqlStr = "select count(orderserial) as cnt, isnull(sum(subtotalprice),0) as totalprice"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master"
		sqlStr = sqlStr + " where userid = '" + userid + "'"
		sqlStr = sqlStr + " and ipkumdiv not in ('0','1')"
		sqlStr = sqlStr + " and cancelyn = 'N'"
		rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if  not rsget.EOF  then
			Fcount = rsget("cnt")
			Ftotalprice = rsget("totalprice")
		end if
		rsget.close
	end Function

	Public Function UserMinusOrderInfo(byval userid)
		dim sqlStr

		sqlStr = "select count(orderserial) as cnt, isnull(sum(subtotalprice),0) as totalprice"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master"
		sqlStr = sqlStr + " where userid = '" + userid + "'"
		sqlStr = sqlStr + " and ipkumdiv >=5"
		sqlStr = sqlStr + " and cancelyn = 'Y'"
		rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if  not rsget.EOF  then
			FMcount = rsget("cnt")
			FMtotalprice = rsget("totalprice")
		end if
		rsget.close
	end Function

end class

function drawSelectBoxqadiv(selectBoxName, selectedId, chplg, allyn, comm_isdel, dispyn)
	dim tmp_str, sqlStr, tmpqadivname

	sqlStr = "select" & vbcrlf
	sqlStr = sqlStr & " c.comm_cd, c.comm_name, c.comm_group, c.comm_isDel, c.comm_color, c.sortno, c.dispyn" & vbcrlf
	sqlStr = sqlStr & " from [db_cs].[dbo].[tbl_cs_comm_code] c" & vbcrlf
	sqlStr = sqlStr & " where left(comm_group,3)='D00'" & vbcrlf

	if comm_isdel <> "" then
		sqlStr = sqlStr & " and comm_isdel='"& comm_isdel &"'" & vbcrlf
	end if
	if dispyn <> "" then
		sqlStr = sqlStr & " and dispyn='"& dispyn &"'" & vbcrlf
	end if

	sqlStr = sqlStr & " order by comm_group asc, sortno asc"

	'response.write sqlStr & "<Br>"
	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
%>
	<select name="<%=selectBoxName%>" <%=chplg%>>
		<% if allyn="Y" then %>
			<option value='' <%if selectedId="" then response.write " selected"%>>전체</option>
		<% end if %>
<%
	if  not rsget.EOF  then
	rsget.Movefirst

	do until rsget.EOF
		if isarray(split(rsget("comm_name"),"!@#")) then
			if ubound(split(rsget("comm_name"),"!@#")) > 0 then
				tmpqadivname =  split(rsget("comm_name"),"!@#")(1)
			end if
		end if

		if selectedId = right(rsget("comm_cd"),2) then
			tmp_str = " selected"
		end if
		response.write "<option value='"& right(rsget("comm_cd"),2) &"' "&tmp_str&">"& db2html(tmpqadivname) &"</option>" & vbcrlf
		tmp_str = ""
		rsget.MoveNext
	loop
	end if
	rsget.close
	response.write("</select>")
end function

function GetExtItemURL(sitename, itemid)
	dim extItemURL, extItemId
	dim sqlStr

	extItemURL = ""
	sqlStr = " select top 1 findURL "
	sqlStr = sqlStr + " from [db_cs].[dbo].[tbl_ext_item_div] where sitename = '" & sitename & "' and useyn = 'Y' "
	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF then
		extItemURL = rsget("findURL")
	end if
	rsget.close

	if extItemURL = "" then
		GetExtItemURL = ""
		exit function
	end if

	'extItemId = 0
	extItemId = "X"		'0에서 X로 변경..by 2019-11-07 진영
	sqlStr = " exec [db_cs].[dbo].[usp_Ten_ExtItemCode] '" & sitename & "', " & itemid
	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF then
		extItemId = rsget("extitemid")
	end if
	rsget.close

	if extItemId = "X" then
		GetExtItemURL = ""
		exit function
	end if

	GetExtItemURL = extItemURL & extItemId
end function

%>
