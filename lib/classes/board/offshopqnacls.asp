<%
'###########################################################
' Description : 오프샾이용문의 클래스
' Hieditor : 2009.04.07 서동석 생성
'			 2011.05.03 한용민 수정
'###########################################################

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
	
	public FUserCell
	public FCellOK
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
	

	Property Get usercell()
		usercell = FUserCell
	end Property		
	
	Property Get cellok()
		cellok = FCellOK
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
	
	Property Let usercell(byVal v)
		FUserCell = v
	end Property
	
	Property Let cellok(byVal v)
		FCellOK = v
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
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FIDBefore
	public FIDAfter
	public FSearchUserID
	public FSearchOrderSerial
	public FSearchNew
	public FRectItemNotInclude
	public FRectOnlyItemInclude
	public FRectDesigner
	public FRectSearchKey
	public FRectSearchString
	public frectidx
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

		FSearchUserID = ""
		FSearchOrderSerial = ""
		FSearchNew = "N"		
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

	'//admin/board/offshop_qna_board_list.asp
	Public sub list()
        dim sql, i, sqladd
	
		'매장 아이디라면 해당 내용만 쿼리
		if fnChkAuth(session("ssBctDiv"),session("ssBctID"),session("ssBctBigo"))<>"" then
			FRectDesigner = fnChkAuth(session("ssBctDiv"),session("ssBctID"),session("ssBctBigo"))
		end if

		'// 추가 쿼리
		IF ( FRectDesigner <> "" )THEN
			
			'/대학로 매장이거나 대학로 리빙매장 일경우 둘다 글보임
			if FRectDesigner = "streetshop011" or FRectDesigner = "streetshop017" then
				sqladd = sqladd & " and m.shopid in ('streetshop011','streetshop017')"
			
			'/아닐경우 지정 매장만	
			else
				sqladd = sqladd & " and m.shopid = '" & FRectDesigner & "' "
			end if
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
		
		'response.write sql &"<br>"					
		rsget.Open sql,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		if FTotalCount < 1 then exit sub
		
		'데이터 리스트 
		sql = "select top " & Cstr(FPageSize * FCurrPage)		
        sql = sql + " m.idx, m.userid,  m.title, m.regdate, m.replyuser, m.isusing, m.shopid, b.shopname "
        sql = sql + " from [db_shop].[dbo].tbl_offshop_qna m"
        sql = sql + " INNER JOIN [db_shop].[dbo].tbl_shop_user as b on m.shopid =b.userid  "
        sql = sql + " where (m.isusing = 'Y') and b.vieworder <> 0 " + sqladd
        sql = sql + " order by m.regdate desc "		

		'response.write sql &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sql,dbget,1

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
				set FItemList(i) = new CMyQNAItem
				
                FItemList(i).idx = rsget("idx")
                FItemList(i).userid = rsget("userid")
                FItemList(i).title = db2html(rsget("title"))
                FItemList(i).regdate = rsget("regdate")
                FItemList(i).replyuser = rsget("replyuser")
                FItemList(i).isusing = rsget("isusing")
				FItemList(i).Fshopid = rsget("shopid")
				FItemList(i).Fshopname = rsget("shopname")	
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'//admin/board/offshop_qna_board_reply.asp
	Public sub read()
        dim sql, i

        sql = " select idx, shopid, userid, title, usermail, emailok, itemid, contents, regdate, replyuser,  replycontents, replydate, isusing, "
        sql = sql + " isNull(usercell,'') AS usercell, isNull(cellok,'') AS cellok from [db_shop].[dbo].tbl_offshop_qna "
        sql = sql + " where (idx = " + frectidx + ") "

        if (FSearchUserID <> "") then
			sql = sql + " and (userid = '" + FSearchUserID + "') "
        end if
		
		'response.write sql &"<Br>"
        rsget.Open sql, dbget, 1

        redim preserve FItemList(rsget.RecordCount)		
		
        if not rsget.EOF then
		i=0
        set FItemList(i) = new CMyQNAItem

			FItemList(i).idx = rsget("idx")
			FItemList(i).userid = rsget("userid")
			FItemList(i).Fshopid = rsget("shopid")
			FItemList(i).title = db2html(rsget("title"))
			FItemList(i).usermail = rsget("usermail")
			FItemList(i).emailok = rsget("emailok")
			FItemList(i).contents = db2html(rsget("contents"))
			FItemList(i).regdate = rsget("regdate")
			FItemList(i).replyuser = rsget("replyuser")
			FItemList(i).replycontents = db2html(rsget("replycontents"))
			FItemList(i).replydate = rsget("replydate")
			FItemList(i).itemid = rsget("itemid")
			FItemList(i).isusing = rsget("isusing")
			FItemList(i).usercell = rsget("usercell")
			FItemList(i).cellok = rsget("cellok")
			
		end if
        
        rsget.close
		
		IF FItemList(i).itemid <> 0 OR not isnull(FItemList(i).itemid) THEN
			call fnGetItemInfo(FItemList(i).itemid)
		END IF	
	end sub

	'//admin/board/offshop_qna_board_reply.asp
	private Function fnGetItemInfo(ByVal itemID)
		Dim strSql,i
		
		i=0
		IF itemID = "" THEN EXIT FUNCTION
			
		strSql = " SELECT brandname, itemname, sellcash,listimage FROM [db_item].[dbo].tbl_item WHERE itemid ="&itemID
		
		'response.write sql &"<Br>"
		rsget.Open strSql, dbget, 1
		
		IF not rsget.Eof THEN
			FItemList(i).brandname 	= rsget("brandname")
			FItemList(i).itemname 	= rsget("itemname")
			FItemList(i).sellcash 	= rsget("sellcash")
			FItemList(i).listimage 	= "http://webimage.10x10.co.kr/image/list/"&GetImageSubFolderByItemid(itemID)&"/"&rsget("listimage")
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
			sql = sql + " from [db_order].[dbo].tbl_order_master"
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
			sql = sql + " from [db_order].[dbo].tbl_order_master"
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