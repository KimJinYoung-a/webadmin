<%

'id, userid, username, orderserial, qadiv, title, usermail, emailok, contents, regdate, replyuser, replytitle, replycontents, replydate, isusing
Class CMyQNAItem
	private Fidx
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
	private Fgubun

	private Fisusing
	public Fitemid
	private Fbrandname
	private Fitemname
	private Fsellcash
	private Flistimage
	
	public Fmakerid
	public Fdeliverytype

	public FExtSiteName
	public Fshopid
	public Fshopname
        '==========================================================================

   public function IsUpchebeasong()
    	if (Fdeliverytype="2") or (Fdeliverytype="5") then
    		IsUpchebeasong = true
    	else
    		IsUpchebeasong = false
    	end if
	end function

	Property Get idx()
		idx = Fidx
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
	
	Property Get gubun()
		gubun = Fgubun
	end Property
	
	Property Get isusing()
		isusing = Fisusing
	end Property

	Property Get itemid()
		itemid = Fitemid
	end Property

	Property Get itemname()
		itemname = Fitemname
	end Property
	
	Property Get brandname()
		brandname = Fbrandname
	end Property
	
	Property Get sellcash()
		sellcash = Fsellcash
	end Property
	
	Property Get listimage()
		listimage = Flistimage
	end Property
	
        '==========================================================================
	Property Let idx(byVal v)
		Fidx = v
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
	
	Property Let gubun(byVal v)
		Fgubun = v
	end Property

	Property Let isusing(byVal v)
		Fisusing = v
	end Property

	Property Let itemid(byVal v)
		Fitemid = v
	end Property
	
	Property Let itemname(byVal v)
		Fitemname = v
	end Property
	
	Property Let brandname(byVal v)
		Fbrandname = v
	end Property
	
	Property Let sellcash(byVal v)
		Fsellcash = v
	end Property
	
	Property Let listimage(byVal v)
		Flistimage = v
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

	public FCurrPage
	private FTotalPage
	private FTotalCount
	public FPageSize
	private FResultCount
	public FScrollCount

	private FIDBefore
	private FIDAfter

	private FSearchUserID
	private FSearchOrderSerial
	private FSearchNew

	public FRectItemNotInclude
	public FRectOnlyItemInclude
	public FRectDesigner
	public FRectSearchKey
	public FRectSearchString

	Property Get TotalPage()
		TotalPage = FTotalPage
	end Property

	Property Get TotalCount()
		TotalCount = FTotalCount
	end Property

	Property Get ResultCount()
		ResultCount = FResultCount
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
		HasNextScroll = TotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrPage-1)\FScrollCount)*FScrollCount +1
	end Function


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
	Public Function list()
        dim sql, i, sqladd
	
		'매장 아이디라면 해당 내용만 쿼리
		if fnChkAuth(session("ssBctDiv"),session("ssBctID"),session("ssBctBigo"))<>"" then
			FRectDesigner = fnChkAuth(session("ssBctDiv"),session("ssBctID"),session("ssBctBigo"))
		end if
		
		'// 추가 쿼리
		IF ( FRectDesigner <> "" )THEN
			sqladd = " and m.shopid = '" & FRectDesigner & "' "
		END IF

		if (FRectSearchString <> "") then
			sqladd = sqladd + " and m." & FRectSearchKey & " like '%" + FRectSearchString + "%' "
        end if

        if (FSearchNew = "Y") then
			sqladd = sqladd + " and (m.replyuser = '' OR m.replyuser is null) "
		elseif (FSearchNew = "N") then
			sqladd = sqladd + " and m.replyuser<>'' and m.replyuser is Not null "
        end if

		'// 결과 카운트
		sql = " select count(m.idx) as cnt "
		sql = sql + " from [db_shop].[dbo].tbl_offshop_qna as m  INNER JOIN  [db_shop].[dbo].tbl_shop_user as b on m.shopid =b.userid  "
		sql = sql + " where (m.isusing = 'Y') and b.vieworder <> 0 " + sqladd
		
		rsget.Open sql, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.Close

        '// 목록 접수
        sql = " select top " + CStr(FPageSize*FCurrPage) + " m.idx, m.userid,  m.title, m.regdate, m.replyuser, m.isusing, m.shopid, b.shopname "
        sql = sql + " from [db_shop].[dbo].tbl_offshop_qna m INNER JOIN  [db_shop].[dbo].tbl_shop_user as b on m.shopid =b.userid  "
        sql = sql + " where (m.isusing = 'Y') and b.vieworder <> 0 " + sqladd
        sql = sql + " order by m.regdate desc "

		if FPageSize<>0 then
			rsget.pagesize = FPageSize
		end if
                rsget.Open sql, dbget, 1
                'response.write sql


		FTotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount - (FCurrPage-1)*FPageSize

		'if (FResultCount > FPageSize) then
		'	FResultCount = FPageSize
		'end if

                redim preserve results(FResultCount)

                if not rsget.EOF then
                        i = 0
                        rsget.absolutepage = FCurrPage
                        do until ( rsget.eof or (i > FResultCount))
                                set results(i) = new CMyQNAItem

                                results(i).idx = rsget("idx")
                                results(i).userid = rsget("userid")
                               ' results(i).username = db2html(rsget("username"))
'                                results(i).orderserial = rsget("orderserial")
'                                results(i).qadiv = rsget("qadiv")
                                results(i).title = db2html(rsget("title"))
                                'results(i).usermail = rsget("usermail")
                                'results(i).emailok = rsget("emailok")
                                'results(i).contents = rsget("contents")
                                results(i).regdate = rsget("regdate")
                                results(i).replyuser = rsget("replyuser")
                                'results(i).replytitle = rsget("replytitle")
                                'results(i).replycontents = rsget("replycontents")

                                results(i).isusing = rsget("isusing")
								results(i).Fshopid = rsget("shopid")
								results(i).Fshopname = rsget("shopname")
'                                results(i).Fitemid = rsget("itemid")
'								results(i).Fmakerid = rsget("makerid")
'								results(i).Fdeliverytype = rsget("deliverytype")
'								results(i).FExtSiteName = rsget("extsitename")
				rsget.MoveNext
				i = i + 1
                        loop
                end if
                rsget.close
	end Function

	Public Function read(byval idx)
                dim sql, i

                sql = " select idx, userid, title, usermail, emailok, itemid, contents, regdate, replyuser,  replycontents, replydate, isusing from [db_shop].[dbo].tbl_offshop_qna "
                sql = sql + " where (idx = " + idx + ") "

                if (FSearchUserID <> "") then
                        sql = sql + " and (userid = '" + FSearchUserID + "') "
                end if

                rsget.Open sql, dbget, 1
                'response.write sql

                redim preserve results(rsget.RecordCount)
				
				
                if not rsget.EOF then
                		i=0
                        set results(i) = new CMyQNAItem

								results(i).idx = rsget("idx")
								results(i).userid = rsget("userid")
								'results(i).username = db2html(rsget("username"))
								'results(i).orderserial = rsget("orderserial")
								'results(i).qadiv = rsget("qadiv")
								results(i).title = db2html(rsget("title"))
								results(i).usermail = rsget("usermail")
								results(i).emailok = rsget("emailok")
								results(i).contents = db2html(rsget("contents"))
								results(i).regdate = rsget("regdate")
								results(i).replyuser = rsget("replyuser")
								'results(i).replytitle = db2html(rsget("replytitle"))
								 results(i).replycontents = db2html(rsget("replycontents"))
								 results(i).replydate = rsget("replydate")
								 results(i).itemid = rsget("itemid")
								 results(i).isusing = rsget("isusing")
								
								'results(i).Fextsitename = rsget("extsitename")
				end if
                
                rsget.close
				
				IF results(i).itemid <> 0 OR not isnull(results(i).itemid) THEN
					call fnGetItemInfo(results(i).itemid)
				END IF	
	end Function
	
	private Function fnGetItemInfo(ByVal itemID)
		Dim strSql,i
		i=0
		IF itemID = "" THEN EXIT FUNCTION
		strSql = " SELECT brandname, itemname, sellcash,listimage FROM [db_item].[10x10].tbl_item WHERE itemid ="&itemID
		rsget.Open strSql, dbget, 1
		IF not rsget.Eof THEN
			results(i).brandname 	= rsget("brandname")
			results(i).itemname 	= rsget("itemname")
			results(i).sellcash 	= rsget("sellcash")
			results(i).listimage 	= "http://webimage.10x10.co.kr/image/list/"&GetImageSubFolderByItemid(itemID)&"/"&rsget("listimage")
		END IF	
		rsget.Close		
	End Function	
	
	Public Function write(byval boarditem)
                dim sql, i

                sql = " insert into [db_shop].[dbo].tbl_offshop_qna(userid, username, orderserial, qadiv, title, usermail, emailok, contents, regdate, replyuser, replytitle, replycontents, isusing) "
                sql = sql + " values('" + boarditem.userid + "', '" + boarditem.username + "', '" + boarditem.orderserial + "', '" + boarditem.qadiv + "', '" + boarditem.title + "', '" + boarditem.usermail + "', '" + boarditem.emailok + "', '" + boarditem.contents + "', getdate(), '', '', '', 'Y') "
                rsget.Open sql, dbget, 1
	end Function

	Public Function modify(byval boarditem)
                dim sql, i

                sql = "update [db_shop].[dbo].tbl_offshop_qna " + VbCRlf
                sql = sql + " set title = '" + boarditem.title + "'," + VbCRlf
                sql = sql + " usermail = '" + boarditem.usermail + "'," + VbCRlf
                sql = sql + " emailok = '" + boarditem.emailok + "'," + VbCRlf
                sql = sql + " contents = '" + boarditem.contents + "' " + VbCRlf
                sql = sql + " where (idx = " + boarditem.idx + ") " + VbCRlf
                sql = sql + " and (userid = '" + boarditem.userid + "') " + VbCRlf
                'response.write sql
                'dbget.close()	:	response.End
                rsget.Open sql, dbget, 1
	end Function

	Public Function reply(byval boarditem)
                dim sql, i

                sql = "update [db_shop].[dbo].tbl_offshop_qna " + VbCRlf
                sql = sql + " set replyuser = '" + boarditem.replyuser + "'," + VbCRlf                  
             '   sql = sql + " replytitle = '" + boarditem.replytitle + "'," + VbCRlf             	
                sql = sql + " replycontents = '" + boarditem.replycontents + "', " + VbCRlf
                sql = sql + "  replydate = getdate()" + VbCRlf
                sql = sql + " where (idx = " + boarditem.idx + ") "
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
                        code2name = "취소,환불문의"
                elseif (v = "06") then
                        code2name = "교환문의"
                elseif (v = "08") then
                        code2name = "사은품문의"
                elseif (v = "10") then
                        code2name = "시스템문의"
                elseif (v = "12") then
                        code2name = "개인정보관련"
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

'// 오프샾 Select Box 생성
Sub printOffShopSelectBox(isNew, nowShop)
	Dim strPrint, strSQL
	strPrint = "<select name='shopid'>" & vbCrLf &_
				"<option value=''>전체</option>" & vbCrLf

	strSQL = "select t1.shopid, t2.shopname, count(*) as cnt " &_
				"from [db_shop].[dbo].tbl_offshop_qna as t1 " &_
				"	Join [db_shop].[dbo].tbl_shop_user as t2 " &_
				"		on t1.shopid=t2.userid " &_
				"where t1.isusing='Y' " &_
				"	and t2.vieworder<>0 "
		if isNew="Y" then
			strSQL = strSQL & "	and (t1.replyuser = '' OR t1.replyuser is null) "
		elseif isNew="N" then
			strSQL = strSQL & "	and t1.replyuser <>'' and t1.replyuser is Not null "
		end if
	strSQL = strSQL & "group by t1.shopid, t2.shopname"

	rsget.Open strSQL,dbget,1

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			if nowShop=rsget("shopid") then
				strPrint = strPrint & "<option value='" & rsget("shopid") & "' selected>" & rsget("shopname") & "[" & rsget("cnt") & "]</option>" & vbCrLf
			else
				strPrint = strPrint & "<option value='" & rsget("shopid") & "'>" & rsget("shopname") & "[" & rsget("cnt") & "]</option>" & vbCrLf
			end if
			rsget.MoveNext
		Loop
	end if
	rsget.Close
	
	strPrint = strPrint & "</select>" & vbCrLf
	
	'화면에 출력
	Response.Write strPrint
End Sub
%>