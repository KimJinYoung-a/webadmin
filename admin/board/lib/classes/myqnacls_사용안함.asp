<%

'id, userid, username, orderserial, qadiv, title, usermail, emailok, contents, regdate, replyuser, replytitle, replycontents, replydate, isusing
Class CMyQNAItem
	private Fid
	private Fuserid
	private Fusername
	private Forderserial
	private Fqadiv
	private Ftitle
	private Fusermail
	private Femailok
	private Fcontents
	private Fregdate
	private Freplyuser
	private Freplytitle
	private Freplycontents
	private Freplydate

	private Fdispyn
	private Fisusing
	public Fitemid
	public Fmakerid
	public Fdeliverytype

	public FExtSiteName

	public Fuserlevel

	'/사용금지		'/공통펑션에 공용함수 쓸것.		'/2016.07.21 한용민
    public function GetUserLevelStr()
    	if IsNULL(Fuserlevel) then
    		GetUserLevelStr = "&nbsp;"
    	elseif CStr(Fuserlevel)="0" then
    		GetUserLevelStr = "<font color=#FFCC00>YELLOW</font>"
    	elseif CStr(Fuserlevel)="1" then
    		GetUserLevelStr = "<font color=#33ff66>GREEN</font>"
    	elseif CStr(Fuserlevel)="2" then
    		GetUserLevelStr = "<font color=#3366ff>BLUE</font>"
    	elseif CStr(Fuserlevel)="3" then
    		GetUserLevelStr = "<font color=#ff3366>VIP</font>"
    	elseif CStr(Fuserlevel)="9" then
    		GetUserLevelStr = "<font color=#ff33ff>MANIA</font>"
    	else
    		GetUserLevelStr = CStr(Fuserlevel)
    	end if
	end function

    public function IsUpchebeasong()
    	if (Fdeliverytype="2") or (Fdeliverytype="5") then
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

        '==========================================================================
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

        '==========================================================================
	Private Sub Class_Initialize()
                '
	End Sub

	Private Sub Class_Terminate()
        '
	End Sub
end Class

Class CMyQNA
    public results()

	private FCurrPage
	private FTotalPage
	private FTotalCount
	private FPageSize
	private FResultCount
	private FScrollCount

	private FIDBefore
	private FIDAfter

	public FSearchUserID
	public FSearchOrderSerial
	private FSearchNew

	public FRectItemNotInclude
	public FRectOnlyItemInclude
	public FRectDesigner


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

	Private Sub Class_Initialize()
	redim results(0)
		FSearchUserID = ""
		FSearchOrderSerial = ""
		FSearchNew = "N"
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub

        '======================================================================
    Public Function getLecQnalist()
    	dim sql, i
    	sql = " select count(q.id) as cnt from [db_cs].[10x10].tbl_myqna q, [db_contents].[dbo].tbl_lecture_item i "
    	sql = sql + " where (q.itemid=i.linkitemid) "
    	sql = sql + " and q.itemid<>0"
    	sql = sql + " and q.isusing = 'Y' "
    	if (FSearchNew = "Y") then
                sql = sql + " and (replyuser = '') "
        end if
        rsget.Open sql, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.Close

		sql = " select top " + CStr(FPageSize*FCurrPage) + " m.id, m.userid, m.username, m.orderserial, m.qadiv, m.title, m.regdate, m.replyuser, m.isusing, m.itemid, m.extsitename,i.lecturer "
        sql = sql + " from [db_cs].[10x10].tbl_myqna m "
        sql = sql + " , [db_contents].[dbo].tbl_lecture_item i "
        sql = sql + " where m.itemid=i.linkitemid "
		sql = sql + " and m.itemid<>0"
		sql = sql + " and m.isusing = 'Y' "
        if (FSearchNew = "Y") then
                sql = sql + " and (replyuser = '') "
        end if

		sql = sql + " order by m.regdate desc "

		if FPageSize<>0 then
			rsget.pagesize = PageSize
		end if

        rsget.Open sql, dbget, 1

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

	Public Function list()
        dim sql, i

		sql = " select count(id) as cnt from [db_cs].[10x10].tbl_myqna "
		sql = sql + " where (isusing = 'Y') "
		if (FSearchUserID <> "") then
                sql = sql + " and (userid = '" + FSearchUserID + "') "
        end if

        if (FSearchOrderSerial <> "") then
                sql = sql + " and (orderserial = '" + FSearchOrderSerial + "') "
        end if

        if (FSearchNew = "Y") then
                sql = sql + " and (replyuser = '') and (dispyn <> 'N') "
        end if

        if (FRectQadiv <> "") then
                sql = sql + " and (qadiv = '" & FRectQadiv & "') "
        end if

		if (FRectItemNotInclude<>"") then
				sql = sql + " and ((itemid=0) or (itemid is NULL) or (qadiv not in ('02','03')))"
		end if

		if (FRectOnlyItemInclude<>"") then
			sql = sql + " and ((itemid>0))"
		end if

		rsget.Open sql, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.Close

        sql = " select top " + CStr(FPageSize*FCurrPage) + " m.id, m.userid, m.username, "
        sql = sql + " m.orderserial, m.qadiv, m.title, m.regdate, m.replyuser, m.isusing, m.dispyn, "
        sql = sql + " m.itemid, i.makerid, i.deliverytype, m.extsitename, m.userlevel "
        sql = sql + " from [db_cs].[10x10].tbl_myqna m "
        sql = sql + " left join [db_item].[10x10].tbl_item i on m.itemid=i.itemid"
        sql = sql + " where (m.isusing = 'Y') "

        if (FSearchUserID <> "") then
                sql = sql + " and (m.userid = '" + FSearchUserID + "') "
        end if

        if (FSearchOrderSerial <> "") then
                sql = sql + " and (m.orderserial = '" + FSearchOrderSerial + "') "
        end if

        if (FSearchNew = "Y") then
                sql = sql + " and (m.replyuser = '') and (m.dispyn <> 'N')"
        end if

		if (FRectQadiv <> "") then
				sql = sql + " and (m.qadiv = '" & FRectQadiv & "') "
		end if

		if (FRectItemNotInclude<>"") then
				sql = sql + " and ((m.itemid=0) or (m.itemid is NULL) or (qadiv not in ('02','03')) )"
		end if

		if (FRectOnlyItemInclude<>"") then
			sql = sql + " and (m.itemid>0)"
		end if

		if (FRectDesigner<>"") then
			sql = sql + " and (i.makerid ='" + FRectDesigner + "') "
		end if

        sql = sql + " order by m.regdate desc "

		if FPageSize<>0 then
			rsget.pagesize = PageSize
		end if
                rsget.Open sql, dbget, 1
                'response.write sql


		FTotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount - (CurrPage-1)*PageSize

		'if (FResultCount > PageSize) then
		'	FResultCount = PageSize
		'end if

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
																results(i).dispyn 	= rsget("dispyn")
                                results(i).isusing = rsget("isusing")

                                results(i).Fitemid = rsget("itemid")
								results(i).Fmakerid = rsget("makerid")
								results(i).Fdeliverytype = rsget("deliverytype")
								results(i).FExtSiteName = rsget("extsitename")

								results(i).Fuserlevel = rsget("userlevel")
				rsget.MoveNext
				i = i + 1
                        loop
                end if
                rsget.close
	end Function

	Public Function read(byval id)
                dim sql, i

                sql = " select id, userid, username, orderserial, qadiv, title,"
                sql = sql + " usermail, emailok, contents, regdate, replyuser, replytitle,"
                sql = sql + " replycontents, replydate,itemid, isusing, extsitename, userlevel"
                sql = sql + " from [db_cs].[10x10].tbl_myqna "
                sql = sql + " where (id = " + id + ") "

                if (FSearchUserID <> "") then
                        sql = sql + " and (userid = '" + FSearchUserID + "') "
                end if

                if (FSearchOrderSerial <> "") then
                        sql = sql + " and (orderserial = '" + FSearchOrderSerial + "') "
                end if

                rsget.Open sql, dbget, 1
                'response.write sql

                redim preserve results(rsget.RecordCount)

                if not rsget.EOF then
                        set results(0) = new CMyQNAItem

                                results(i).id = rsget("id")
                                results(i).userid = rsget("userid")
                                results(i).username = db2html(rsget("username"))
                                results(i).orderserial = rsget("orderserial")
                                results(i).qadiv = rsget("qadiv")
                                results(i).title = db2html(rsget("title"))
                                results(i).usermail = rsget("usermail")
                                results(i).emailok = rsget("emailok")
                                results(i).contents = db2html(rsget("contents"))
                                results(i).regdate = rsget("regdate")
                                results(i).replyuser = rsget("replyuser")
                                results(i).replytitle = db2html(rsget("replytitle"))
                                results(i).replycontents = db2html(rsget("replycontents"))
                                results(i).replydate = rsget("replydate")
                                results(i).itemid = rsget("itemid")
                                results(i).isusing = rsget("isusing")

                                results(i).Fextsitename = rsget("extsitename")

                                results(i).Fuserlevel = rsget("userlevel")
                end if
                rsget.close
	end Function

	Public Function write(byval boarditem)
                dim sql, i

                sql = " insert into [db_cs].[10x10].tbl_myqna(userid, username, orderserial, qadiv, title, usermail, emailok, contents, regdate, replyuser, replytitle, replycontents, isusing) "
                sql = sql + " values('" + boarditem.userid + "', '" + boarditem.username + "', '" + boarditem.orderserial + "', '" + boarditem.qadiv + "', '" + boarditem.title + "', '" + boarditem.usermail + "', '" + boarditem.emailok + "', '" + boarditem.contents + "', getdate(), '', '', '', 'Y') "
                rsget.Open sql, dbget, 1
	end Function

	Public Function modify(byval boarditem)
                dim sql, i

                sql = "update [db_cs].[10x10].tbl_myqna " + VbCRlf
                sql = sql + " set qadiv = '" + boarditem.qadiv + "'," + VbCRlf
                sql = sql + " title = '" + boarditem.title + "'," + VbCRlf
                sql = sql + " usermail = '" + boarditem.usermail + "'," + VbCRlf
                sql = sql + " emailok = '" + boarditem.emailok + "'," + VbCRlf
                sql = sql + " contents = '" + boarditem.contents + "' " + VbCRlf
                sql = sql + " where (id = " + boarditem.id + ") " + VbCRlf
                sql = sql + " and (userid = '" + boarditem.userid + "') " + VbCRlf
                sql = sql + " and (orderserial = '" + boarditem.orderserial + "') "
                'response.write sql
                'dbget.close()	:	response.End
                rsget.Open sql, dbget, 1
	end Function

	Public Function reply(byval boarditem)
                dim sql, i

                sql = "update [db_cs].[10x10].tbl_myqna " + VbCRlf
                sql = sql + " set replyuser = '" + boarditem.replyuser + "'," + VbCRlf
                sql = sql + " replytitle = '" + boarditem.replytitle + "'," + VbCRlf
                sql = sql + " replycontents = '" + boarditem.replycontents + "', " + VbCRlf
                sql = sql + " replydate = getdate()" + VbCRlf
                sql = sql + " where (id = " + boarditem.id + ") "
                'response.write sql
                'dbget.close()	:	response.End
                rsget.Open sql, dbget, 1
	end Function

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

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()
                '
	End Sub

	Public Function UserOrderInfo(byval userid)
		dim sql
			sql = "select count(orderserial) as cnt, isnull(sum(subtotalprice),0) as totalprice"
			sql = sql + " from [db_order].[10x10].tbl_order_master"
			sql = sql + " where userid = '" + userid + "'"
			sql = sql + " and ipkumdiv not in ('0','1')"
			sql = sql + " and cancelyn = 'N'"
			rsget.Open sql,dbget,1
			if  not rsget.EOF  then
				Fcount = rsget("cnt")
				Ftotalprice = rsget("totalprice")
			end if
			rsget.close
	end Function

	Public Function UserMinusOrderInfo(byval userid)
		dim sql
			sql = "select count(orderserial) as cnt, isnull(sum(subtotalprice),0) as totalprice"
			sql = sql + " from [db_order].[10x10].tbl_order_master"
			sql = sql + " where userid = '" + userid + "'"
			sql = sql + " and ipkumdiv >=5"
			sql = sql + " and cancelyn = 'Y'"
			rsget.Open sql,dbget,1
			if  not rsget.EOF  then
				FMcount = rsget("cnt")
				FMtotalprice = rsget("totalprice")
			end if
			rsget.close
	end Function

end class
%>