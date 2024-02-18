<%
'###########################################################
' Description : 온라인 1:1 게시판 문의 보기
' Hieditor : 2009.04.17 이상구 생성
'			 2023.10.25 한용민 수정(쿼리튜닝. 이메일발송 cdo->메일러로 변경)
'###########################################################

Class CBoardItem
	private Fid
	private Fuserid
	private Fuseremail
	private Ftitle
	private Flinkurl
	private Fcomment
	private Fimage1
	private Fimage2
	private Fimage3
	private Fregdate
	private Fhitcount
	private Fdeleteyn
	private FIIcon
	private FImgExplain1
	private FImgExplain2
	private FImgExplain3
	private FCntCount
	private FPoints
	
	public Ffixnotics
	public Ffixsdate
	public Ffixedate
	public FisPopup
	public FpopSdate
	public FpopEdate

	Property Get id()
		id = Fid
	end Property

	Property Get userid()
		userid = Fuserid
	end Property

	Property Get useremail()
		useremail = Fuseremail
	end Property

	Property Get title()
		title = Ftitle
	end Property

	Property Get linkurl()
		linkurl = Flinkurl
	end Property

	Property Get comment()
		comment = Fcomment
	end Property

	Property Get image1()
		image1 = Trim(Fimage1)
	end Property

	Property Get image2()
		image2 = Trim(Fimage2)
	end Property

	Property Get image3()
		image3 = Trim(Fimage3)
	end Property

	Property Get regdate()
		regdate = Fregdate
	end Property

	Property Get hitcount()
		hitcount = Fhitcount
	end Property

	Property Get deleteyn()
		deleteyn = Fdeleteyn
	end Property

	Property Get IIcon()
		IIcon = FIIcon
	end Property
	
	Property Get ImgExplain1()
		if isnull(FImgExplain1) then
			ImgExplain1 = ""
		else
			ImgExplain1 = FImgExplain1
		end if
	end Property
	
	Property Get ImgExplain2()
		if isnull(FImgExplain2) then
			ImgExplain2 = ""
		else
			ImgExplain2 = FImgExplain2
		end if
	end Property
	
	Property Get ImgExplain3()
		if isnull(FImgExplain3) then
			ImgExplain3 = ""
		else
			ImgExplain3 = FImgExplain3
		end if
	end Property
	
	Property Get CommentCount()
		if IsNumeric(FCntCount) then
			CommentCount = FCntCount
		else
			CommentCount = 0
		end if
	end Property
	
	Property Get Points()
		if IsNull(FPoints) then
			Points = 0
		else
			Points = FPoints
		end if
	end Property

	Property Let id(byVal v)
		Fid = v
	end Property

	Property Let userid(byVal v)
		Fuserid = v
	end Property

	Property Let useremail(byVal v)
		Fuseremail = v
	end Property

	Property Let title(byVal v)
		Ftitle = v
	end Property

	Property Let linkurl(byVal v)
		Flinkurl = v
	end Property

	Property Let comment(byVal v)
		Fcomment = v
	end Property

	Property Let image1(byVal v)
		Fimage1 = v
	end Property

	Property Let image2(byVal v)
		Fimage2 = v
	end Property

	Property Let image3(byVal v)
		Fimage3 = v
	end Property

	Property Let regdate(byVal v)
		Fregdate = v
	end Property

	Property Let hitcount(byVal v)
		Fhitcount = v
	end Property

	Property Let deleteyn(byVal v)
		Fdeleteyn = v
	end Property
	
	Property Let IIcon(byVal v)
		FIIcon = v
	end Property
	
	Property Let ImgExplain1(byVal v)
		FImgExplain1 = v
	end Property
	
	Property Let ImgExplain2(byVal v)
		FImgExplain2 = v
	end Property
	
	Property Let ImgExplain3(byVal v)
		FImgExplain3 = v
	end Property
	
	Property Let CommentCount(byVal v)
		FCntCount = v
	end Property
	
	Property Let Points(byval v)
		FPoints = v
	end Property

    public function IsImageExists()
    	IsImageExists = (image1<>"") or (image2<>"") or (image3<>"")
	end function

	public function GetImageCount()
		dim cnt 
		cnt=0
		if image1<>"" then cnt = cnt +1
		if image2<>"" then cnt = cnt +1
		if image3<>"" then cnt = cnt +1
		GetImageCount = cnt
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CCommentItem
	public Fid
	public Fname
	public Ftitle
	public Fregdate
    public FRectID
	public FRectName
	public FRectEmail
    public FRectTitle
	public FRectContents
	public FRectIdx
    public FRectWriteday
	private FScrollCount

end class

Class CBoard
    public FTableName
    public BoardItem()
	public CommentItem()
	public Fmode

	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
    public Fint_total

	public FCurrPage
	public FPageCount
    public FTotalCount
	public FRectsearch
	public FRectsearch2

    public FRectID
	public FRectName
	public FRectEmail
    public FRectTitle
	public FRectContents
	public FRectIdx
    public FRectWriteday
	public Fregdate
	public FRectDesignerID
	public FRectFixonly
	public FRectDispCate
	public Fdispcate1
	public Fdispcatename
	public FRectRef
	public FRectLevel
	public FRectSerial
    public FRectRefuserid
	public FRectNum
	public FRectDeleteyn
	public FFixNotics
	public FRectMDid
	public FRectCatCD
	public FRectTarget
	public MailSendedCount
	public WillMailSendCount
	public FRectfileName

	public FcomCnt

 	public FFixSdate
 	public FFixEdate
 	public FIsPopup
 	public FpopSDate
 	public FpopEdate
 		
	Private Sub Class_Initialize()
		redim BoardItem(0)
		redim CommentItem(0)
		FScrollCount = 10
		FCurrPage =1
		FPageSize = 10
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	Public Function write(byval userid, byval title, byval linkurl, byval image1, byval image2, byval image3, byval useremail, byval comment,byval iicon,byval imgexplain1,byval imgexplain2,byval imgexplain3)
        dim sql

        sql = "insert into " + FTableName + "(userid, title, linkurl, image1, image2, image3, regdate, hitcount, useremail, comment, deleteyn, iicon, imgexplain1, imgexplain2, imgexplain3) "
        sql = sql + " values('" + userid + "', '" + title + "', '" + linkurl + "', '" + image1 + "', '" + image2 + "', '" + image3 + "', getdate(), 0, '" + useremail + "', '" + comment + "', 'N','" + iicon + "','" + imgexplain1 + "','" + imgexplain2 + "','" + imgexplain3 +"')"
 
        on error resume next

        dbget.execute sql
        if err then
            write = err.description
            on error goto 0
        else
            write = ""
        end if
	end Function

	Public Function list(byval pgno, byval pgsize)
        dim sql, tbl2

		tbl2 = FTableName + "_com"
		tbl2 = "(select masterid, count(id) as cnt, sum(points) as pt from " + tbl2 + " group by masterid)"
        sql = "select top " + CStr(FPageSize * CInt(pgsize)) + "t1.id, t1.userid, t1.title, t1.linkurl, t1.image1, t1.image2, t1.image3, t1.regdate, t1.hitcount, t1.useremail, t1.comment,t1.iicon, t1.imgexplain1, t1.imgexplain2, t1.imgexplain3"
        sql = sql + " , t2.cnt, t2.pt"
        sql = sql + " from " + FTableName + " as t1 with (nolock) left join " + tbl2 + " as t2 with (nolock) on t1.id=t2.masterid"
        sql = sql + " where (t1.deleteyn = 'N') "
        sql = sql + " order by t1.id desc "

        CurrPage = pgsize
        rsget.pagesize = PageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalPage = rsget.PageCount
		FResultCount = rsget.RecordCount - ((CurrPage-1) * PageSize)
        if (FResultCount>PageSize) then
			FResultCount = PageSize
		end if

        redim preserve BoardItem(FResultCount)
        dim i, tmp
        i = 0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until (i = FResultCount)
                set BoardItem(i) = new CBoardItem
                BoardItem(i).id = rsget("id")
                BoardItem(i).userid = rsget("userid")
                BoardItem(i).title = db2html(rsget("title"))
                BoardItem(i).linkurl = rsget("linkurl")
                BoardItem(i).image1 = rsget("image1")
                BoardItem(i).image2 = rsget("image2")
                BoardItem(i).image3 = rsget("image3")
                BoardItem(i).regdate = rsget("regdate")
                BoardItem(i).hitcount = rsget("hitcount")
                BoardItem(i).useremail = rsget("useremail")
                BoardItem(i).comment = db2html(rsget("comment"))
				BoardItem(i).IIcon = rsget("iicon")
				BoardItem(i).ImgExplain1 = db2html(rsget("imgexplain1"))
				BoardItem(i).ImgExplain2 = db2html(rsget("imgexplain2"))
				BoardItem(i).ImgExplain3 = db2html(rsget("imgexplain3"))
				BoardItem(i).CommentCount = rsget("cnt")
				BoardItem(i).Points = rsget("pt")
				rsget.MoveNext
				i = i + 1
			loop
		end if
        rsget.close

        'if err then
        '    list = err.description
        '    on error goto 0
        'else
        '    list = ""
        'end if
	end Function

	Public Function read(byval pid)
        dim sql

		sql = "update " + FTableName
        sql = sql + " set hitcount=hitcount+1"
        sql = sql + " where id = " + CStr(pid)

        dbget.execute sql

        sql = "select id, userid, title, linkurl, image1, image2, image3, regdate, hitcount, useremail, comment, iicon, imgexplain1, imgexplain2, imgexplain3 "
        sql = sql + " from " + FTableName + " with (nolock)"
        sql = sql + " where (deleteyn = 'N') and (id = " + CStr(pid) + ")"

		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

        redim preserve BoardItem(1)
        dim i, tmp
        i = 0
		if not rsget.EOF then
			do until (i = 1)
                set BoardItem(i) = new CBoardItem
                BoardItem(i).id = rsget("id")
                BoardItem(i).userid = rsget("userid")
                BoardItem(i).title = db2html(rsget("title"))
                BoardItem(i).linkurl = rsget("linkurl")
                BoardItem(i).image1 = rsget("image1")
                BoardItem(i).image2 = rsget("image2")
                BoardItem(i).image3 = rsget("image3")
                BoardItem(i).regdate = rsget("regdate")
                BoardItem(i).hitcount = rsget("hitcount")
                BoardItem(i).useremail = rsget("useremail")
                BoardItem(i).comment = db2html(rsget("comment"))
				BoardItem(i).IIcon = rsget("iicon")
				BoardItem(i).ImgExplain1 = db2html(rsget("imgexplain1"))
				BoardItem(i).ImgExplain2 = db2html(rsget("imgexplain2"))
				BoardItem(i).ImgExplain3 = db2html(rsget("imgexplain3"))
				rsget.MoveNext
				i = i + 1
			loop
		end if
        rsget.close


	end Function

	Public Sub ReadComments(byval pid)
		dim sql

		sql = "select id,userid,comment,regdate,iicon,points "
		sql = sql + " from " + FTableName + "_com with (nolock)"
        sql = sql + " where masterid = " + CStr(pid)
        sql = sql + " and deleteyn='N'"

		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
        redim preserve CommentItem(rsget.PageCount)
        dim i
        i=0
		if not rsget.EOF then
        do until rsget.eof
        	set CommentItem(i) = new CCommentItem
        	CommentItem(i).id       = rsget("id")
			CommentItem(i).userid   = rsget("userid")
			CommentItem(i).comment  = db2html(rsget("comment"))
			CommentItem(i).regdate  = rsget("regdate")
			CommentItem(i).iicon    = rsget("iicon")
			CommentItem(i).points   = rsget("points")
        	rsget.movenext
        	i=i+1
        loop
		end if
        rsget.close
	end Sub
 
	
	Public Function design_notice()
        dim sql, wheredetail,i

		wheredetail = ""
		if FRectFixonly="on" then
			wheredetail = " and fixnotics='Y' and( (fixsdate <=getdate() and fixedate >=getdate() ) or (fixsdate is Null and fixedate is null)) "
		elseif FRectFixonly="off" then
			wheredetail = " and (fixnotics<>'Y' or (fixnotics='Y' and (fixsdate>getdate() or fixedate <getdate()) )) "
		end if

		if (FRectsearch <> "" and FRectsearch2 <> "" ) then
			wheredetail = " and " + FRectsearch + " like '%" + FRectsearch2 + "%'"
		end if
  
		sql = "select  board_idx, name, email, title, writeday, isNull(dispcate1,0) as dispcate1, d.catename"
		sql = sql + " ,(select count(comidx) from db_board.dbo.tbl_partnerA_notice_comment with (nolock) where isusing =1 and board_idx = a.board_idx) as comCnt "
		sql = sql + " ,fixnotics, fixsdate, fixedate, ispopup, popsdate, popedate "
		sql = sql + " from " + FTableName + " as A with (nolock)"
		sql = sql + " left outer join db_item.dbo.tbl_display_cate as d with (nolock) on d.catecode = A.dispcate1 "
		sql = sql + " where board_idx<>0"
		sql = sql + " and deleteyn='N'"
		sql = sql + wheredetail
		sql = sql + " order by board_idx DESC"

        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

        'redim preserve BoardItem(FResultCount)
        if (FResultCount<1) then FResultCount=0
		redim BoardItem(FResultCount)

		i = 0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until (i >= FResultCount)
                set BoardItem(i) = new CBoard
                BoardItem(i).FRectIdx = rsget("board_idx")
                BoardItem(i).FRectName = db2html(rsget("name"))
				BoardItem(i).FRectTitle = db2html(rsget("title"))
				BoardItem(i).Fdispcatename = rsget("catename")
                BoardItem(i).Fregdate = rsget("writeday")
                BoardItem(i).FcomCnt = rsget("comCnt")
                BoardItem(i).Ffixnotics = rsget("fixnotics")
                BoardItem(i).Ffixsdate = rsget("fixsdate")
                BoardItem(i).Ffixedate = rsget("fixedate")
                BoardItem(i).Fispopup = rsget("ispopup")
                BoardItem(i).Fpopsdate = rsget("popsdate")
                BoardItem(i).Fpopedate = rsget("popedate")
				rsget.MoveNext
				i = i + 1
			loop
		end if
        rsget.close
	end Function

	'// 해당 카테고리에 속한 상품이 있는 업체만 공지 볼 수 있도록 기능 추가 2014.12.30 정윤정
	Public Function design_notice_dispcate()
        dim sql, wheredetail,i

		wheredetail = ""
		if FRectFixonly="on" then
			wheredetail = " and fixnotics='Y' and( (fixsdate <=getdate() and fixedate >=getdate() ) or (fixsdate is Null and fixedate is null)) "
		elseif FRectFixonly="off" then
			wheredetail = " and (fixnotics<>'Y' or (fixnotics='Y' and (fixsdate>getdate() or fixedate <getdate()) )) "
		end if

		if (FRectsearch <> "" and FRectsearch2 <> "" ) then
			wheredetail = " and " + FRectsearch + " like '%" + FRectsearch2 + "%'"
		end if

		sql = "select  board_idx, name, email, title, writeday"
		sql = sql + " ,( select count(comidx) from db_board.dbo.tbl_partnerA_notice_comment with (nolock) where isusing = 1 and board_idx = n.board_idx)  as comCnt "			
        sql = sql + " from [db_board].[dbo].tbl_designer_notice as n with (nolock)"  
        sql = sql + " left outer join db_partner.dbo.tbl_partner_dispcate as p with (nolock)"
        sql = sql + "	on n.dispcate1 = p.catecode  and p.makerid ='"+FRectDesignerID+"'"
        sql = sql + " where board_idx<>0"
		sql = sql + " and deleteyn='N'"
        sql = sql + wheredetail
        sql = sql + " and (n.dispcate1 is null or n.dispcate1 ='' or (n.dispcate1 is not null and p.catecode is not null) )"
        sql = sql + " order by board_idx DESC"

        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

        'redim preserve BoardItem(FResultCount)
        if (FResultCount<1) then FResultCount=0
		redim BoardItem(FResultCount)

		i = 0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until (i >= FResultCount)
                set BoardItem(i) = new CBoard
                BoardItem(i).FRectIdx = rsget("board_idx")
                BoardItem(i).FRectName = db2html(rsget("name"))
				BoardItem(i).FRectTitle = db2html(rsget("title"))
                BoardItem(i).Fregdate = rsget("writeday")
                   BoardItem(i).FcomCnt  = rsget("comCnt")
				rsget.MoveNext
				i = i + 1
			loop
		end if
        rsget.close
	end Function

	Public Function design_notice_write()
        dim sql,tmpeCode,fileName,strSql

		if FFixNotics ="" then FFixNotics ="N"
	   if FIsPopup ="" then FIsPopup ="N"
	   	
        sql = "insert into " &FTableName & "(userid, name, email, title, content, writeday, readnum, ref, ref_level, ref_serial, num, fixnotics, dispcate1, fixSdate, fixEdate, isPopup, popSdate, popEdate) "&vbcrlf
        sql = sql & " values('"& FRectID & "', '" & FRectName & "', '" & FRectEmail & "', '" &  FRectTitle & "', '"& FRectContents & "', getdate(), 0,0,0,0,0,'" & FFixNotics & "','"&FRectDispCate&"' "&vbcrlf
        if FFixNotics="Y" then
        sql = sql & " ,'"&FFixSdate&"','"&FFixEdate&"'"&vbcrlf
    	  else
    	  	 sql = sql & " ,NULL,NULL"&vbcrlf
    	end if
    
        sql = sql & ",'"&FIsPopup&"'"&vbcrlf
        if FisPopup ="Y" then
        sql = sql & ",'"&FpopSdate&"','"&FpopEdate&"')"
       else
       	 sql = sql & ",NULL, NULL)"
    	end if
 
         on error resume next

		dbget.execute sql
        if err then
            design_notice_write = err.description
            on error goto 0
        else 
        	if FRectfileName <> "" then
        		strSql = "select SCOPE_IDENTITY()" 
						rsget.CursorLocation = adUseClient
						rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
						tmpeCode = rsget(0)
						rsget.Close
						
        		'첨부파일 등록
							fileName = split(FRectfileName,",")
							For i = 0 To UBound(fileName)
							if (trim(fileName(i)) <> "") then
								strSql = " INSERT INTO db_board.dbo.tbl_partnerA_notice_attachfile(board_idx, fileLink) "
								strSql = strSql & " VALUES ("&tmpeCode&",'"&trim(fileName(i))&"' ) " 
								dbget.execute strSql
							end if
							Next
						end if	
            design_notice_write = ""
        end if
	end Function

	'// 업체 공지메일 목록 보기
	Public Function design_notice_mail_preview()
		dim sql, addsql, selMail, i

		Select Case FRectTarget
			Case "basic"
				selMail = "email"
			Case "deliver"
				selMail = "deliver_email"
			Case "account"
				selMail = "jungsan_email"
		End Select

		addsql = ""
		if FRectMDid<>"" then
			addsql = addsql & " and c.mduserid='" & FRectMDid & "' "
		end if
		if FRectCatCD<>"" then
			addsql = addsql & " and c.catecode='" & FRectCatCD & "' "
		end if

		'# 목록 카운트
		sql = " SELECT count(t.email) as cnt ,CEILING(CAST(Count(t.email) AS FLOAT)/" + CStr(FPageSize) + ") as totalpage " &_
				" ,(select count(t2.email) from ( " &_
				" 	SELECT distinct p2." & selMail & " as email" &_
				" 	FROM [db_partner].[dbo].tbl_partner p2 with (nolock)" &_
				" 	JOIN [db_user].[dbo].tbl_user_c c2 with (nolock)" &_
				" 		on p2.id=c2.userid " &_
				" 	WHERE p2.userdiv = 9999 " &_
				" 	and c2.userdiv<'10' " &_
				" 	and c2.isusing='Y' " &_
				" 	and p2.isusing='Y' " &_
				" 	and p2." & selMail & " <>'' " &_
				" 	and p2." & selMail & " like '%@%' " &_
				" 	and p2." & selMail & " <> 'partner@1300k.com' " & replace(addsql,"c.","c2.") &_
				" 	) as t2) as totMailCnt " &_
					" FROM ( " &_
					" 	SELECT distinct c.userid, p." & selMail & " as email " &_
					" 	FROM [db_partner].[dbo].tbl_partner p with (nolock)" &_
					" 	JOIN [db_user].[dbo].tbl_user_c c with (nolock)" &_
					" 		on p.id=c.userid " &_
					" 	WHERE p.userdiv = 9999 " &_
					" 	and c.userdiv<'10' " &_
					" 	and c.isusing='Y' " &_
					" 	and p.isusing='Y' " &_
					" 	and p." & selMail & " <>'' " &_
					" 	and p." & selMail & " like '%@%' " &_
					" 	and p." & selMail & " <> 'partner@1300k.com' " & addsql &_
					" 	) as t "
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		if not rsget.eof then
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totalpage")
			Fint_total = rsget("totMailCnt")
		end if

		rsget.close

		'# 목록 접수
		sql = " SELECT DISTINCT TOP " & CStr(FPageSize * CInt(FCurrPage)) & "c.userid, c.socname_kor, p." & selMail & " as email " &_
					" FROM [db_partner].[dbo].tbl_partner p with (nolock)" &_
					" JOIN [db_user].[dbo].tbl_user_c c with (nolock)" &_
					" 	on p.id=c.userid " &_
					" WHERE p.userdiv = 9999 " &_
					" and c.userdiv<'10' " &_
					" and c.isusing='Y' " &_
					" and p.isusing='Y' " &_
					" and p." & selMail & " <>'' " &_
					" and p." & selMail & " like '%@%' " &_
					" and p." & selMail & " <> 'partner@1300k.com' " & addsql &_
					" ORDER BY p." & selMail & " asc "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		if FResultCount<1 then FResultCount=0

		redim preserve BoardItem(FResultCount)
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			
			i=0
			do until rsget.EOF
				set BoardItem(i) = new CBoard
				BoardItem(i).FRectName			= rsget("socname_kor")
				BoardItem(i).FRectEmail			= rsget("email")
				BoardItem(i).FRectDesignerID	= rsget("userid")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		
    	rsget.close
	end Function

	' /admin/board/dodesignernoticemail.asp
	Public Function design_notice_mail_send()
        dim sql, addsql, selMail, mailcontent, reqemail, i
        dim dfPath, fso, ffso
        Const CPagingCount = 50

		' '(구)이메일 템플릿 접수
		' '//실섭,테섭구분
		' IF application("Svr_Info")="Dev" THEN
		' 	dfPath = "C:\testweb\admin2009scm\lib\email\mailtemplate" 		'// 테섭(scm)
		' ELSE
		'     dfPath = Server.MapPath("\lib\email\mailtemplate")				'// 실섭(scm)
		' END IF

		' '/* 파일 불러오기 */
		' Set fso = server.CreateObject("Scripting.FileSystemObject")
		' 	IF fso.FileExists(dfPath & "\mail_u01.htm") then
		' 		set ffso = fso.OpenTextFile(dfPath & "\mail_u01.htm",1)
		' 		mailcontent = ffso.ReadAll
		' 		ffso.close
		' 		set ffso = nothing
		' 	ELSE
		' 		mailcontent = ""
		' 	End IF
		' Set fso = nothing

		'//=======  메일 발송 =========/
		dim oMail

		Select Case FRectTarget
			Case "basic"
				selMail = "email"
			Case "deliver"
				selMail = "deliver_email"
			Case "account"
				selMail = "jungsan_email"
		End Select

		addsql = ""
		if FRectMDid<>"" then
			addsql = addsql & " and c.mduserid='" & FRectMDid & "' "
		end if
		if FRectCatCD<>"" then
			addsql = addsql & " and c.catecode='" & FRectCatCD & "' "
		end if
		'if session("ssBctId")="tozzinet" then
		'	addsql = addsql & " and p.id='temp'"
		'end if
		sql = " SELECT count(t.email) as cnt ,CEILING(CAST(Count(t.email) AS FLOAT)/" + CStr(CPagingCount) + ") as totalpage " &_
					" FROM ( " &_
					" 	SELECT distinct p." & selMail & " as email " &_
					" 	FROM [db_partner].[dbo].tbl_partner p with (nolock)" &_
					" 	LEFT JOIN [db_user].[dbo].tbl_user_c c with (nolock)" &_
					" 		on p.id=c.userid " &_
					" 	WHERE p.userdiv = 9999 " &_
					" 	and c.userdiv<'10' " &_
					" 	and c.isusing='Y' " &_
					" 	and p.isusing='Y' " &_
					" 	and p." & selMail & " <>'' " &_
					" 	and p." & selMail & " like '%@%' " &_
					" 	and p." & selMail & " <> 'partner@1300k.com' " & addsql &_
					" 	GROUP BY p." & selMail &_
					" 	) as t "

		'response.write sql & "<br>"
		'response.end
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		if not rsget.eof then
			WillMailSendCount = rsget("totalpage") - 1
		end if

		rsget.close

		sql = " SELECT DISTINCT TOP " + CStr(CPagingCount) + " p." & selMail & " as email " &_
					" FROM [db_partner].[dbo].tbl_partner p with (nolock)" &_
					" LEFT JOIN [db_user].[dbo].tbl_user_c c with (nolock)" &_
					" 	on p.id=c.userid " &_
					" WHERE p.userdiv = 9999 " &_
					" and c.userdiv<'10' " &_
					" and c.isusing='Y' " &_
					" and p.isusing='Y' " &_
					" and p." & selMail & " <>'' " &_
					" and p." & selMail & " like '%@%' " &_
					" and p." & selMail & " <> 'partner@1300k.com' " & addsql

					if MailSendedCount>0 then '메일 나눠서 보내기용
						sql = sql + "" &_
						" and p." & selMail & " >( " &_
						" 	SELECT top 1 t.email " &_
						" 	FROM ( " &_
						" 			SELECT DISTINCT top " & (CPagingCount * MailSendedCount) & " tp." & selMail & " as email " &_
						" 			FROM [db_partner].[dbo].tbl_partner tp with (nolock)" &_
						" 			LEFT JOIN [db_user].[dbo].tbl_user_c tc with (nolock)" &_
						" 				on tp.id=tc.userid " &_
						" 			WHERE tp.userdiv = 9999 " &_
						" 			and tc.userdiv<'10' " &_
						" 			and tc.isusing='Y' " &_
						" 			and tp.isusing='Y' " &_
						" 			and tp." & selMail & " <>'' " &_
						" 			and tp." & selMail & " like '%@%' " &_
						" 			and tp." & selMail & " <> 'partner@1300k.com' " & Replace(addsql,"c.","tc.") &_
						" 			ORDER BY tp." & selMail & " asc " &_
						" 			) as t " &_
						" 	ORDER BY t.email desc " &_
						" 	) "

					end if

					sql = sql + "" &_
					" ORDER BY p." & selMail & " asc "

		'response.write sql & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		i = 0
		if not rsget.EOF then
			do until rsget.eof
			    reqemail = db2html(trim(rsget("email")))
			    On Error resume Next
					IF reqemail<>"" and not(isnull(reqemail)) THEN
						set oMail = New MailCls         '' mailLib2
							oMail.MailTitles	= FRectTitle
							oMail.SenderNm		= "텐바이텐"
							'oMail.SenderMail	= "mailzine@10x10.co.kr"
							oMail.SenderMail	= "customer@10x10.co.kr"
							oMail.AddrType		= "string"
							oMail.ReceiverNm	= reqemail
							oMail.ReceiverMail	= reqemail
							oMail.MailType = "17"
							mailcontent = oMail.getMailTemplate
							''parsing
							mailcontent = Replace(mailcontent,":mailtitle:","텐바이텐 업체 공지 메일")	'메일 타이틀
							mailcontent = Replace(mailcontent,":mailcontents:",FRectContents)			'메일 본문

							oMail.MailConts 	= mailcontent
							'response.write mailcontent
							'response.end
							oMail.MailerMailGubun = 11		' 메일러 자동메일 번호
							oMail.Send_TMSMailer()
						set oMail = nothing
					end if

					' (구)이메일 발송
					'mailcontent = Replace(mailcontent,":mailtitle:","텐바이텐 업체 공지 메일")	'메일 타이틀
					'mailcontent = Replace(mailcontent,":mailcontents:",FRectContents)			'메일 본문
    				'call sendmail(FrecteMail, reqemail, FRectTitle, mailcontent)

    				if Err then
    				    response.write " [ " & reqemail & " ] " & Err.Description & "<br>"
    				end if
				On Error Goto 0
				rsget.MoveNext
				
				response.flush
			loop
		end if
    	rsget.close
    
    	response.write CStr(MailSendedCount) & "/" & CStr(WillMailSendCount)
    	response.flush
	end Function

	Public Function design_notice_read(byval pid)
        dim sql

		sql = "update " + FTableName
        sql = sql + " set readnum = readnum + 1"
        sql = sql + " where board_idx = " + CStr(pid)

		dbget.execute sql

        sql = "select userid, name, email, title, content, writeday, fixnotics, isNull(dispcate1,0) as dispcate1, d.catename "
        sql = sql +" , convert(varchar(16),fixSdate,121) as fixsdate, convert(varchar(16),fixEdate,121) as fixedate, isPopup, convert(varchar(16),popSdate,121) as popSdate, convert(varchar(16),popEdate,121) as popEdate  "
        sql = sql + " from " + FTableName + " with (nolock)"
        sql = sql + " left outer join db_item.dbo.tbl_display_cate as d with (nolock) on d.catecode = "+FTableName+".dispcate1 "
        sql = sql + " where board_idx = " + CStr(pid)

		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
		if not rsget.EOF then
			FRectID = rsget("userid")
            FRectName = rsget("name")
			FRectEmail = rsget("email")
            FRectTitle = db2html(rsget("title"))
            FRectContents = db2html(rsget("content"))
			'FRectContents = Replace(FRectContents, vbCrLf, "<BR>")
			Fregdate = rsget("writeday")
			Ffixnotics = rsget("fixnotics")
			Fdispcate1 = rsget("dispcate1")
			Fdispcatename = rsget("catename")
			FfixSdate = rsget("fixSdate")
			FfixEdate = rsget("fixEdate")
			FisPopup = rsget("isPopup")
			FpopSdate = rsget("popSdate")
			FpopEdate = rsget("popEdate")
        end if
		rsget.close
	end Function

	Public Function design_notice_modify(byval pid)
        dim sql,strSql

		sql = "update " + FTableName
        sql = sql + " set name = '" + FRectName + "'," + VBCRLF
        sql = sql + " email = '" + FRectEmail + "'," + VBCRLF
        sql = sql + " title = '" + FRectTitle + "'," + VBCRLF
        sql = sql + " content = '" + FRectContents + "'," + VBCRLF
        sql = sql + " fixnotics = '" + Ffixnotics + "'," + VBCRLF
        sql = sql + " dispcate1 = '"+FRectDispCate+"'" +VBCRLF
        if Ffixnotics="Y" then
        sql = sql + ", fixSdate ='"&FfixSdate&"'"&VBCRLF
        sql = sql + ", fixEdate ='"&FfixEdate&"'"&VBCRLF
    	 ELSE
    	 sql = sql + ", fixSdate =NULL"&VBCRLF
        sql = sql + ", fixEdate =NULL"&VBCRLF
    	 end if
    
        sql = sql + ", isPopup ='"&FisPopup&"'"&VBCRLF
        
           if FisPopup="Y" then
      sql = sql + ", popSdate ='"&FpopSdate&"'"&VBCRLF
        sql = sql + ", popEdate ='"&FpopEdate&"'"&VBCRLF
    	 ELSE
    	 sql = sql + ", popSdate =NULL"&VBCRLF
        sql = sql + ", popEdate =NULL"&VBCRLF
    	 end if    	 
        
        sql = sql + " where board_idx = " + CStr(pid)

		dbget.execute sql
        
        IF fileName <> "" then
				'첨부파일 등록
				strSql = "DELETE FROM db_board.dbo.tbl_partnerA_notice_attachfile	where board_idx = "&CStr(pid)
				dbget.execute strSql
				
					fileName = split(fileName,",")
					For i = 0 To UBound(fileName)
					if (trim(fileName(i)) <> "") then
						strSql = " INSERT INTO db_board.dbo.tbl_partnerA_Notice_attachfile(board_idx, fileLink) "
						strSql = strSql & " VALUES ("&CStr(pid)&",'"&trim(fileName(i))&"' ) "
						dbget.execute strSql
					end if
					Next
				END IF	
	end Function

	Public Function design_notice_del()
        dim sql,deletecount,deletelistcount,i

		FRectIdx = Request("deletelist")
		deletecount =  split(FRectIdx,",")
		deletelistcount = ubound(deletecount)

         for i=0 to deletelistcount-1
         	sql = "update " + FTableName + " set deleteyn='Y' where board_idx = " & deletecount(i)

         	'response.write sql
           'sql = "delete from " + FTableName + " where board_idx = " & deletecount(i)
           dbget.execute sql
		 next
	end Function

	Public Function design_board()
        dim sql, wheredetail,i

		wheredetail = ""

		if (FRectsearch <> "" and FRectsearch2 <> "" ) then
		wheredetail = " and " + FRectsearch + " like '%" + FRectsearch2 + "%'"
        wheredetail = wheredetail + " order by  board_idx desc"
        else
		wheredetail = wheredetail + " order by  ref desc , ref_serial"
		end if

		sql = "select  board_idx, name, email, title, writeday,ref_level,deleteyn"
        sql = sql + " from" & FTableName & " with (nolock)"
        sql = sql + " where ref_userid='" & FRectDesignerID & "'"
		sql = sql + wheredetail

        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

        redim preserve BoardItem(FResultCount)

		i = 0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until (i >= FResultCount)
                set BoardItem(i) = new CBoard
                BoardItem(i).FRectIdx = rsget("board_idx")
                BoardItem(i).FRectName = rsget("name")
				BoardItem(i).FRectTitle = rsget("title")
                BoardItem(i).Fregdate = rsget("writeday")
 '               BoardItem(i).FRectRef = rsget("ref")
                BoardItem(i).FRectLevel = rsget("ref_level")
   '             BoardItem(i).FRectSerial = rsget("ref_serial")
  '              BoardItem(i).FRectNum = rsget("num")
                BoardItem(i).FRectDeleteyn = rsget("deleteyn")
				rsget.MoveNext
				i = i + 1
			loop
		end if
        rsget.close
	end Function

	Public Function design_board_write()
        dim sql,number

		'현재 등록된 최대 인덱스 번호를 구한다
		sql = "Select max(board_idx) as maxcount from " + FTableName + " with (nolock)"
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		if isnull(rsget("maxcount")) then
			number = 1
		else
			number = rsget("maxcount") +1
		end if

        rsget.close

		if FRectIdx <> "" then
			FRectRef = FRectRef
			sql = "update " + FTableName + " set ref_serial=ref_serial+1" & _
					" where  ref = " & FRectRef & " and ref_serial > " & FRectSerial
            dbget.execute sql

			FRectLevel = FRectLevel + 1
			FRectSerial = FRectSerial + 1
'			      rsget.close
		else
			FRectRef = number
			FRectLevel = 0
			FRectSerial = 0
			''FRectRefuserid = FRectDesignerID
		end if

        sql = "insert into " &  FTableName  & "(userid, name, email, title, content, writeday, readnum, ref, ref_level, ref_serial, num,ref_userid) "
        sql = sql + " values('" & FRectDesignerID & "','" & FRectName & "','" & FRectEmail & "','" &  FRectTitle & "','" & FRectContents & "',getdate(),0," & FRectRef & "," & FRectLevel & "," & FRectSerial & "," & number & ",'" & FRectRefuserid & "')"

		dbget.execute sql
	end Function

	Public Function design_board_read(byval pid)
        dim sql

		sql = "update " + FTableName
        sql = sql + " set readnum = readnum + 1"
        sql = sql + " where board_idx = " + CStr(pid)

        dbget.execute sql

        sql = "select userid, name, email, title, content, writeday"
        sql = sql + " from " + FTableName + " with (nolock)"
        sql = sql + " where board_idx = " + CStr(pid)

		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
		if not rsget.EOF then
			FRectID = rsget("userid")
            FRectName = rsget("name")
			FRectEmail = rsget("email")
            FRectTitle = rsget("title")
            FRectContents = rsget("content")
			FRectContents = Replace(FRectContents, vbCrLf, "<BR>")
			Fregdate = rsget("writeday")
        end if
		rsget.close
	end Function

	Public Function design_board_reply()
        dim sql,number,ref,ref_level,ref_serial

        sql = "select userid,content, ref, ref_level, ref_serial"
        sql = sql + " from " & FTableName + " with (nolock)"
        sql = sql + " where board_idx=" & FRectIdx

		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
		if not rsget.EOF then
			FRectIdx = FRectIdx
			FRectContents = rsget("content")
			FRectContents = ">>" & FRectContents
			FRectRef = rsget("ref")
			FRectLevel = rsget("ref_level")
			FRectSerial = rsget("ref_serial")
			FRectRefuserid = rsget("userid")
        end if
		rsget.close
	end Function

	public Function fnGetAttachFile 
		dim strSql
		strSql = "SELECT attachFileidx,board_idx,fileLink FROM db_board.dbo.tbl_partnerA_notice_attachfile with (nolock) WHERE board_idx ="&FRectIdx
		 
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			IF not rsget.EOF THEN
				 fnGetAttachFile = rsget.getRows()
			END IF
		rsget.close
	End Function
	
	Public Function design_board_modify(byval pid)
        dim sql

        if Fmode = "" then
	        sql = "select userid, name, email, title, content, writeday"
	        sql = sql + " from " + FTableName + " with (nolock)"
	        sql = sql + " where board_idx = " + CStr(pid)
	
			rsget.CursorLocation = adUseClient
			rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
			if not rsget.EOF then
					FRectID = rsget("userid")
	                FRectName = rsget("name")
					FRectEmail = rsget("email")
	                FRectTitle = rsget("title")
	                FRectContents = rsget("content")
					Fregdate = rsget("writeday")
			end if
	        rsget.close
		else
			sql = "update " + FTableName
	        sql = sql + " set name = '" + FRectName + "',"
	        sql = sql + " email = '" + FRectEmail + "',"
	        sql = sql + " title = '" + FRectTitle + "',"
	        sql = sql + " content = '" + FRectContents + "'"
	        sql = sql + " where board_idx = " + CStr(pid)
	
	        dbget.execute sql
       end if
	end Function

	Public Function design_board_del()
        dim sql

		sql = "update " + FTableName
		sql = sql + " set deleteyn = 'Y'"
		sql = sql + " where board_idx=" + FRectIdx

		dbget.execute sql
	end Function

	Public Function admin_design_board()
        dim sql, wheredetail,i

		wheredetail = ""

		if (FRectsearch <> "" and FRectsearch2 <> "" ) then
			wheredetail = " and " + FRectsearch + " like '%" + FRectsearch2 + "%'"
	        wheredetail = wheredetail + " order by  board_idx desc"
        else
			wheredetail = wheredetail + " order by  ref desc , ref_serial"
		end if

		sql = "select  board_idx, name, email, title, writeday,ref_level,deleteyn"
        sql = sql + " from " + FTableName + " with (nolock)"
        sql = sql + " where board_idx<>0"
		sql = sql + wheredetail

		'response.write sql
        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.RecordCount

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

        redim preserve BoardItem(FResultCount)

		i = 0
		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until (i >= FResultCount)
                set BoardItem(i) = new CBoard
                BoardItem(i).FRectIdx = rsget("board_idx")
                BoardItem(i).FRectName = rsget("name")
				BoardItem(i).FRectTitle = rsget("title")
                BoardItem(i).Fregdate = rsget("writeday")
 '               BoardItem(i).FRectRef = rsget("ref")
                BoardItem(i).FRectLevel = rsget("ref_level")
   '             BoardItem(i).FRectSerial = rsget("ref_serial")
  '              BoardItem(i).FRectNum = rsget("num")
                BoardItem(i).FRectDeleteyn = rsget("deleteyn")
				rsget.MoveNext
				i = i + 1
			loop
		end if
        rsget.close
	end Function

	public function GetBestItemID()
		dim i,maxpoints,re
		re =0
		maxpoints=-1
		for i=0 to Ubound(BoardItem)-1
			if BoardItem(i).Points>maxpoints then
				re=i
				maxpoints = BoardItem(i).Points
			end if
		next
		GetBestItemID = re
	end Function

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
%>