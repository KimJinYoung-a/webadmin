<%
'##########################################################
'	history : 서동석 생성
'			2022.09.13 한용민 수정(엑셀다운로드,검색조건 추가)
'##########################################################

'id, reqcd, companyname, chargename, chargeposition, zipcode, address, phone, hp, email, companyurl, companycomments, reqcomment, attachfile, attachfile2, regdate, finishdate
Class CCompanyRequestItem
	private Fid
	private Freqcd
	private Fcompanyname
	private Fzipcode
	private Faddress
	private Fchargename
	private Fchargeposition
	private Fphone
	private Fhp
	private Femail
	private Fcompanyurl
	private Fcompanycomments
	private Freqcomment
	private Fattachfile
	private Fattachfile2
	private Fregdate
	private Ffinishdate
	private Fcategubun
	private Freplyuser
	private Freplycomment
	private Fsellgubun
	private FipjumYN
	private fService
	private futong
	private fcur_target
	private fmanufacturing_name
	private fmanufacturing
	private fphysical_name
	private fphysical
	private flicense
	private findustrial
    private ftax
    private flicense_no
	private fcd1
    private fcd2
    private fcd3

    public cd1name
    public dispcatename1
    public dispcatename2
    public dispcate
	public Fworkid
	public fisusing

	Property Get cd1()
		cd1 = fcd1
	end Property

	Property Get cd2()
		cd2 = fcd2
	end Property

	Property Get cd3()
		cd3 = fcd3
	end Property

	Property Get license_no()
		license_no = flicense_no
	end Property

	Property Get tax()
		tax = ftax
	end Property

	Property Get Service()
		Service = fService
	end Property

	Property Get utong()
		utong = futong
	end Property

	Property Get cur_target()
		cur_target = fcur_target
	end Property

	Property Get manufacturing_name()
		manufacturing_name = fmanufacturing_name
	end Property

	Property Get manufacturing()
		manufacturing = fmanufacturing
	end Property

	Property Get physical_name()
		physical_name = fphysical_name
	end Property

	Property Get physical()
		physical = fphysical
	end Property

	Property Get license()
		license = flicense
	end Property

	Property Get industrial()
		industrial = findustrial
	end Property

     Property Get ipjumYN()
    	ipjumYN    =FipjumYN
    end Property

    Property Get sellgubun()
    	sellgubun    =Fsellgubun
    end Property

    Property Get replyuser()
    	replyuser=Freplyuser
    end Property

    Property Get replycomment()
    	replycomment=Freplycomment
    end Property

    Property Get categubun()
    	categubun=Fcategubun
    end Property

	Property Get id()
		id = Fid
	end Property

	Property Get reqcd()
		reqcd = Freqcd
	end Property

	Property Get companyname()
		companyname = Fcompanyname
	end Property

	Property Get zipcode()
		zipcode = Fzipcode
	end Property

	Property Get address()
		address = Faddress
	end Property

	Property Get chargename()
		chargename = Fchargename
	end Property

	Property Get chargeposition()
		chargeposition = Fchargeposition
	end Property

	Property Get phone()
		phone = Fphone
	end Property

	Property Get hp()
		hp = Fhp
	end Property

	Property Get email()
		email = Femail
	end Property

	Property Get companyurl()
		companyurl = Fcompanyurl
	end Property

	Property Get companycomments()
		companycomments = Fcompanycomments
	end Property

	Property Get reqcomment()
		reqcomment = Freqcomment
	end Property

	Property Get attachfile()
		attachfile = Fattachfile
	end Property

	Property Get attachfile2()
		attachfile2 = Fattachfile2
	end Property

	Property Get regdate()
		regdate = Fregdate
	end Property

	Property Get finishdate()
		finishdate = Ffinishdate
	end Property

	Property Let cd1(byVal v)
		fcd1 = v
	end Property

	Property Let cd2(byVal v)
		fcd2 = v
	end Property

	Property Let cd3(byVal v)
		fcd3 = v
	end Property

	Property Let license_no(byVal v)
		flicense_no = v
	end Property

	Property Let tax(byVal v)
		ftax = v
	end Property

	Property Let Service(byVal v)
		fService = v
	end Property

	Property Let utong(byVal v)
		futong = v
	end Property

	Property Let cur_target(byVal v)
		fcur_target = v
	end Property

	Property Let manufacturing_name(byVal v)
		fmanufacturing_name = v
	end Property

	Property Let manufacturing(byVal v)
		fmanufacturing = v
	end Property

	Property Let physical_name(byVal v)
		fphysical_name = v
	end Property

	Property Let physical(byVal v)
		fphysical = v
	end Property

	Property Let license(byVal v)
		flicense = v
	end Property

	Property Let industrial(byVal v)
		findustrial = v
	end Property

	Property Let ipjumYN(byVal v)
    	FipjumYN = v
     end Property

	 Property Let sellgubun(byVal v)
    	Fsellgubun = v
     end Property

	 Property Let replycomment(byVal v)
    	Freplycomment = v
    end Property

     Property Let replyuser(byVal v)
    	Freplyuser = v
    end Property

	Property Let categubun(byVal v)
		Fcategubun = v
	end Property

	Property Let id(byVal v)
		Fid = v
	end Property

	Property Let reqcd(byVal v)
		Freqcd = v
	end Property

	Property Let companyname(byVal v)
		Fcompanyname = v
	end Property

	Property Let zipcode(byVal v)
		Fzipcode = v
	end Property

	Property Let address(byVal v)
		Faddress = v
	end Property

	Property Let chargename(byVal v)
		Fchargename = v
	end Property

	Property Let chargeposition(byVal v)
		Fchargeposition = v
	end Property

	Property Let phone(byVal v)
		Fphone = v
	end Property

	Property Let hp(byVal v)
		Fhp = v
	end Property

	Property Let email(byVal v)
		Femail = v
	end Property

	Property Let companyurl(byVal v)
		Fcompanyurl = v
	end Property

	Property Let companycomments(byVal v)
		Fcompanycomments = v
	end Property

	Property Let reqcomment(byVal v)
		Freqcomment = v
	end Property

	Property Let attachfile(byVal v)
		Fattachfile = v
	end Property

	Property Let attachfile2(byVal v)
		Fattachfile2 = v
	end Property

	Property Let regdate(byVal v)
		Fregdate = v
	end Property

	Property Let finishdate(byVal v)
		Ffinishdate = v
	end Property

	public function getAllianceGubun()
		Select Case Fsellgubun
			Case "1"
				getAllianceGubun = "공급 제휴"
			Case "2"
				getAllianceGubun = "컨텐츠 제휴"
			Case "3"
				getAllianceGubun = "공동마케팅 및 프로모션 제휴"
			Case "4"
				getAllianceGubun = "문화이벤트 제휴"
			Case "5"
				getAllianceGubun = "기술 및 솔루션 관련 제휴"
			Case "6"
				getAllianceGubun = "광고관련"
			Case Else
				getAllianceGubun = "기타제휴"
		End Select
	end function

	Private Sub Class_Initialize()
		'
	End Sub
	Private Sub Class_Terminate()
        '
	End Sub
end Class

Class CCompanyRequest
    public results()

	private FCurrPage
	private FTotalPage
	private FTotalCount
	private FPageSize
	private FResultCount
	private FScrollCount

	private FIDBefore
	private FIDAfter

	public FReqcd
	public FOnlyNotFinish
	public FRectSearchKey
	public FRectCatevalue
	public FipjumYN
	public FRectDispCate
	public FRectSellgubun
	public FRectWorkid
	public FRectlicense_no
	public FRectReqcomment
	public FRectID
	public FRectstartdate
	public FRectenddate
	public fArrList
	
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

	Private Sub Class_Initialize()
		redim results(0)
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub

        'id, reqcd, companyname, chargename, chargeposition, zipcode, address, phone, hp, email, companyurl, companycomments, reqcomment, attachfile, attachfile2, regdate, finishdate
        '======================================================================

	' 밑에 함수를 고칠경우 getReqListNotpaging() 함수 도 같이 고쳐야 한다.
	Public Function list()
        dim sql, i, sqlsearch

        if (FRectID<>"") then
            sqlsearch = sqlsearch & " and id="&FRectID&VbCRLF
        end if
        
		if FReqcd<>"" then
			sqlsearch = sqlsearch & " and reqcd='" & FReqcd & "'"
		end if

		if FOnlyNotFinish<>"" then
			sqlsearch = sqlsearch & " and finishdate is NULL"
			sqlsearch = sqlsearch & " and replycomment is NULL"
		end if

		if FRectSearchKey<>"" then
			sqlsearch = sqlsearch & " and companyname like '%" & FRectSearchKey & "%'"
		end if

		if FRectCatevalue <>"" then
			if FRectCatevalue="etc" then
        		sqlsearch = sqlsearch & " and categubun='etc'" &vbcrlf
			else
        		sqlsearch = sqlsearch & " and categubun='" & FRectCatevalue & "'"
			end if
		end if

		if FipjumYN <>"" then
			sqlsearch = sqlsearch & " and ipjumYN='" & FipjumYN & "'" &vbcrlf
		end if

		if FRectDispCate <> "" then
			sqlsearch = sqlsearch & " and dispcate like '" & FRectDispCate & "%'" &vbcrlf
		end if

		if FRectSellgubun <> "" then
			sqlsearch = sqlsearch & " and sellgubun = '" & FRectSellgubun & "'" &vbcrlf
		end if

		if FRectWorkid <> "" then
			sqlsearch = sqlsearch & " and workid = '" & FRectWorkid & "'" &vbcrlf
		end if
		if FRectlicense_no <> "" then
			sqlsearch = sqlsearch & " and replace(license_no,'-','') = '" & replace(FRectlicense_no,"-","") & "'" &vbcrlf
		end if
		if FRectReqcomment <> "" then
			sqlsearch = sqlsearch & " and reqcomment like '%" & FRectReqcomment & "%'" &vbcrlf
		end if
		if FRectstartdate<>"" and FRectenddate<>"" then
			if FRectstartdate<>"" then
				sqlsearch = sqlsearch & " and a.regdate >= '"& FRectstartdate &"'" &vbcrlf
			end if
			if FRectenddate<>"" then
				sqlsearch = sqlsearch & " and a.regdate < '"& FRectenddate &"'" &vbcrlf
			end if
		end if

        sql = " select count(id) as cnt, CEILING(CAST(Count(*) AS FLOAT)/'"&FPageSize&"' ) as totPg"
		sql = sql & " from [db_cs].[dbo].tbl_company_request as a with (nolock)"
        sql = sql & " where id<>0"
        sql = sql & " and isusing='Y' " & sqlsearch

		'response.write sql & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		
		if FTotalCount < 1 then exit function
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit Function
		end if

        sql = " select top " & CStr(FPageSize*FCurrPage)
        sql = sql & " id, reqcd, companyname, chargename, chargeposition, zipcode, address, phone, hp, email"
        sql = sql & " , companyurl, companycomments, reqcomment, attachfile, attachfile2, regdate, finishdate,categubun,categubun2,categubun3"
        sql = sql & " ,sellgubun,ipjumYN,replycomment ,utong,cur_target,tax,license_no"
        sql = sql & " ,industrial,license,Service,physical,physical_name,manufacturing,manufacturing_name, b.code_nm, a.dispcate, d.catename as catename1, e.catename as catename2 "
        sql = sql & " from [db_cs].[dbo].tbl_company_request as a with (nolock)"
        sql = sql & " left outer join db_item.dbo.tbl_Cate_large as b with (nolock)" 
        sql = sql & " 	on a.categubun = b.code_large "
        sql = sql & " left outer join [db_item].[dbo].[tbl_display_cate] as d with (nolock)"
        sql = sql & " 	on d.catecode = left(a.dispcate,3) and d.useyn ='Y' "
        sql = sql & " left outer join [db_item].[dbo].[tbl_display_cate] as e with (nolock)"
        sql = sql & " 	on e.catecode = a.dispcate  and e.useyn ='Y' "
        sql = sql & " where id<>0"
       	sql = sql & " and a.isusing='Y' " & sqlsearch
        sql = sql & " order by a.regdate desc "

		if FPageSize<>0 then
			rsget.pagesize = PageSize
		end if

		'response.write sql & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalPage =  CInt(FTotalCount\FPageSize)

		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

	        redim preserve results(FResultCount)

	        if not rsget.EOF then
                i = 0
                rsget.absolutepage = FCurrPage
                do until ( rsget.eof or (i > FResultCount))

                    set results(i) = new CCompanyRequestItem
		                results(i).id = rsget("id")
		                results(i).reqcd = rsget("reqcd")
		                results(i).companyname = db2html(rsget("companyname"))
		                results(i).chargename = db2html(rsget("chargename"))
		                results(i).chargeposition = rsget("chargeposition")
		                results(i).zipcode = rsget("zipcode")
		                results(i).address = db2html(rsget("address"))
		                results(i).phone = rsget("phone")
		                results(i).hp = rsget("hp")
		                results(i).email = rsget("email")
		                results(i).companyurl = rsget("companyurl")
		                results(i).companycomments = db2html(rsget("companycomments"))
		                results(i).reqcomment = db2html(rsget("reqcomment"))
		                results(i).attachfile = rsget("attachfile")
		                results(i).attachfile2 = rsget("attachfile2")
		                results(i).regdate = rsget("regdate")
		                results(i).finishdate = rsget("finishdate")
		                results(i).replycomment=db2html(rsget("replycomment"))
		                results(i).cd1 = rsget("categubun")
		                results(i).cd2 = rsget("categubun2")
		                results(i).cd3 = rsget("categubun3")
		                results(i).sellgubun=rsget("sellgubun")
		                results(i).ipjumYN=rsget("ipjumYN")
		                results(i).industrial = db2html(rsget("industrial"))
		                results(i).license = db2html(rsget("license"))
		                results(i).Service = db2html(rsget("Service"))
		                results(i).physical = rsget("physical")
		                results(i).physical_name = rsget("physical_name")
		                results(i).manufacturing = rsget("manufacturing")
		                results(i).manufacturing_name=db2html(rsget("manufacturing_name"))
		                results(i).utong = rsget("utong")
		                results(i).cur_target = db2html(rsget("cur_target"))
		                results(i).tax = rsget("tax")
		                results(i).license_no = rsget("license_no")
		                results(i).dispcate = rsget("dispcate")
		                results(i).dispcatename1 = rsget("catename1")
		                results(i).dispcatename2 = rsget("catename2")
		                results(i).cd1name = rsget("code_nm")
				rsget.MoveNext
				i = i + 1
				loop
            end if
            rsget.close
	end Function

	' 밑에 함수를 고칠경우 list() 함수 도 같이 고쳐야 한다.
	Public Function getReqListNotpaging()
        dim sql, i, sqlsearch

        if (FRectID<>"") then
            sqlsearch = sqlsearch & " and id="&FRectID&VbCRLF
        end if
        
		if FReqcd<>"" then
			sqlsearch = sqlsearch & " and reqcd='" & FReqcd & "'"
		end if

		if FOnlyNotFinish<>"" then
			sqlsearch = sqlsearch & " and finishdate is NULL"
			sqlsearch = sqlsearch & " and replycomment is NULL"
		end if

		if FRectSearchKey<>"" then
			sqlsearch = sqlsearch & " and companyname like '%" & FRectSearchKey & "%'"
		end if

		if FRectCatevalue <>"" then
			if FRectCatevalue="etc" then
        		sqlsearch = sqlsearch & " and categubun='etc'" &vbcrlf
			else
        		sqlsearch = sqlsearch & " and categubun='" & FRectCatevalue & "'"
			end if
		end if

		if FipjumYN <>"" then
			sqlsearch = sqlsearch & " and ipjumYN='" & FipjumYN & "'" &vbcrlf
		end if

		if FRectDispCate <> "" then
			sqlsearch = sqlsearch & " and dispcate like '" & FRectDispCate & "%'" &vbcrlf
		end if

		if FRectSellgubun <> "" then
			sqlsearch = sqlsearch & " and sellgubun = '" & FRectSellgubun & "'" &vbcrlf
		end if

		if FRectWorkid <> "" then
			sqlsearch = sqlsearch & " and workid = '" & FRectWorkid & "'" &vbcrlf
		end if
		if FRectlicense_no <> "" then
			sqlsearch = sqlsearch & " and replace(license_no,'-','') = '" & replace(FRectlicense_no,"-","") & "'" &vbcrlf
		end if
		if FRectReqcomment <> "" then
			sqlsearch = sqlsearch & " and reqcomment like '%" & FRectReqcomment & "%'" &vbcrlf
		end if
		if FRectstartdate<>"" and FRectenddate<>"" then
			if FRectstartdate<>"" then
				sqlsearch = sqlsearch & " and a.regdate >= '"& FRectstartdate &"'" &vbcrlf
			end if
			if FRectenddate<>"" then
				sqlsearch = sqlsearch & " and a.regdate < '"& FRectenddate &"'" &vbcrlf
			end if
		end if

        sql = " select top " & CStr(FPageSize*FCurrPage)
        sql = sql & " id, reqcd, companyname, chargename, chargeposition, zipcode, address, phone, hp, email"
        sql = sql & " , companyurl, companycomments, reqcomment, attachfile, attachfile2, regdate, finishdate,categubun,categubun2,categubun3"
        sql = sql & " ,sellgubun,ipjumYN,replycomment ,utong,cur_target,tax,license_no"
        sql = sql & " ,industrial,license,Service,physical,physical_name,manufacturing,manufacturing_name, b.code_nm, a.dispcate, d.catename as catename1, e.catename as catename2 "
        sql = sql & " from [db_cs].[dbo].tbl_company_request as a with (nolock)"
        sql = sql & " left outer join db_item.dbo.tbl_Cate_large as b with (nolock)" 
        sql = sql & " 	on a.categubun = b.code_large "
        sql = sql & " left outer join [db_item].[dbo].[tbl_display_cate] as d with (nolock)"
        sql = sql & " 	on d.catecode = left(a.dispcate,3) and d.useyn ='Y' "
        sql = sql & " left outer join [db_item].[dbo].[tbl_display_cate] as e with (nolock)"
        sql = sql & " 	on e.catecode = a.dispcate  and e.useyn ='Y' "
        sql = sql & " where id<>0"
       	sql = sql & " and a.isusing='Y' " & sqlsearch
        sql = sql & " order by a.regdate desc "

		if FPageSize<>0 then
			rsget.pagesize = PageSize
		end if

		'response.write sql & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount=rsget.RecordCount
		FTotalCount=rsget.RecordCount

	        if not rsget.EOF then
                fArrList = rsget.getrows()
            end if
            rsget.close
	end Function

	Public Function read(byval id)
                dim sql, i

                sql = " select top 1"
		        sql = sql & " id, reqcd, companyname, chargename, chargeposition, zipcode, address, phone, hp, email"
		        sql = sql & " , companyurl, companycomments, reqcomment, attachfile, attachfile2, regdate, finishdate,categubun,categubun2,categubun3,replyuser"
		        sql = sql & " ,sellgubun,ipjumYN,replycomment ,utong,cur_target,tax,license_no"
		        sql = sql & " ,industrial,license,Service,physical,physical_name,manufacturing,manufacturing_name, a.dispcate, d.catename as catename1, e.catename as catename2, workid, a.isusing"
		        sql = sql & " from [db_cs].[dbo].tbl_company_request as a with (nolock)"
		        sql = sql & "  left outer join  [db_item].[dbo].[tbl_display_cate] as d with (nolock) on d.catecode = left(a.dispcate,3) and d.useyn ='Y' "
        		sql = sql & "  left outer join  [db_item].[dbo].[tbl_display_cate] as e with (nolock) on e.catecode = a.dispcate  and e.useyn ='Y' "
                sql = sql & " where a.id = " & id & ""

				'response.write sql & "<br>"
				rsget.CursorLocation = adUseClient
				rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

                redim preserve results(rsget.RecordCount)

                if not rsget.EOF then
                        set results(0) = new CCompanyRequestItem
						results(i).fisusing = rsget("isusing")
                        results(i).id = rsget("id")
                        results(i).reqcd = rsget("reqcd")
                        results(i).companyname = db2html(rsget("companyname"))
                        results(i).chargename = db2html(rsget("chargename"))
                        results(i).chargeposition = rsget("chargeposition")
                        results(i).zipcode = rsget("zipcode")
                        results(i).address = db2html(rsget("address"))
                        results(i).phone = rsget("phone")
                        results(i).hp = rsget("hp")
                        results(i).email = rsget("email")
                        results(i).companyurl = db2html(rsget("companyurl"))
                        results(i).companycomments = db2html(rsget("companycomments"))
                        results(i).reqcomment = db2html(rsget("reqcomment"))
                        results(i).attachfile = rsget("attachfile")
                        results(i).attachfile2 = rsget("attachfile2")
                        results(i).regdate = rsget("regdate")
                        results(i).finishdate = rsget("finishdate")
                        results(i).replyuser =db2html(rsget("replyuser"))
                        results(i).replycomment=db2html(rsget("replycomment"))
                        results(i).cd1 = rsget("categubun")
                        results(i).cd2 = rsget("categubun2")
                        results(i).cd3 = rsget("categubun3")
                        results(i).sellgubun=rsget("sellgubun")
                        results(i).ipjumYN=rsget("ipjumYN")
		                results(i).industrial = db2html(rsget("industrial"))
		                results(i).license = db2html(rsget("license"))
		                results(i).Service = db2html(rsget("Service"))
		                results(i).physical = rsget("physical")
		                results(i).physical_name = rsget("physical_name")
		                results(i).manufacturing = rsget("manufacturing")
		                results(i).manufacturing_name=db2html(rsget("manufacturing_name"))
		                results(i).utong = rsget("utong")
		                results(i).cur_target = db2html(rsget("cur_target"))
		                results(i).tax = rsget("tax")
		                results(i).license_no = rsget("license_no")
		                results(i).dispcate = rsget("dispcate")
		                results(i).dispcatename1 = rsget("catename1")
		                results(i).dispcatename2 = rsget("catename2")
		                results(i).Fworkid = rsget("workid")
                end if
                rsget.close
	end Function

	Public Function write(byval boarditem)
                dim sql, i

                sql = " insert into [db_cs].[dbo].tbl_company_request(reqcd, companyname, chargename, chargeposition, zipcode, address, phone, hp, email, companyurl, companycomments, reqcomment, attachfile, attachfile2, regdate) "
                sql = sql + " values('" + boarditem.reqcd + "', '" + boarditem.companyname + "', '" + boarditem.chargename + "', '" + boarditem.chargeposition + "', '" + boarditem.zipcode + "', '" + boarditem.address + "', '" + boarditem.phone + "', '" + boarditem.hp + "', '" + boarditem.email + "', '" + boarditem.companyurl + "', '" + boarditem.companycomments + "', '" + boarditem.reqcomment + "', '" + boarditem.attachfile + "', '" + boarditem.attachfile2 + "', getdate()) "
                rsget.Open sql, dbget, 1
	end Function

	Public Function finish(byval id)
                dim sql, i

                sql = "update [db_cs].[dbo].tbl_company_request set finishdate = getdate() " + VbCrlf
                sql = sql + " where (id = '" + id + "') "
                rsget.Open sql, dbget, 1
	end Function


	Public Function delitem(byval id)
		        dim sql, i

		        sql = "update [db_cs].[dbo].tbl_company_request set isusing = 'N'" + VbCrlf
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1

	end Function

	Public Function sellchange(byval id,sellgubun)
                dim sql, i

                sql = "update [db_cs].[dbo].tbl_company_request set sellgubun = '" + sellgubun + "'" + VbCrlf
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1
	end Function

	Public Function ipjumchange(byval id,ipjumYN)
                dim sql, i

                sql = "update [db_cs].[dbo].tbl_company_request set ipjumYN = '" + ipjumYN + "'" + VbCrlf
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1
	end Function

	Public Function catechange(byval id,dispcate)
                dim sql, i

                sql = "update [db_cs].[dbo].tbl_company_request set dispcate = '" & dispcate& "'" + VbCrlf
                sql = sql & " where (id = " & id & ") "
                rsget.Open sql, dbget, 1
	end Function

	Public Function writecomm(byval id,user,comment)
                dim sql, i

                sql = "update [db_cs].[dbo].tbl_company_request " + vbcrlf
                sql = sql + " set replyuser='" + user + "'" + vbcrlf
                sql = sql + ", replycomment='" + comment + "'" + vbcrlf
                sql = sql + ", finishdate=getdate()" + vbcrlf
                sql = sql + " where (id = " + id + ") "+vbcrlf
                rsget.Open sql, dbget, 1

	end Function

	Public Function modify(byval boarditem)
                dim sql, i

                sql = "update [db_cs].[dbo].tbl_notice " + VbCrlf
                sql = sql + " set title = '" + boarditem.title + "'," + VbCrlf
                sql = sql + " contents = '" + boarditem.contents + "'," + VbCrlf
                sql = sql + " yuhyostart = '" + boarditem.yuhyostart + "'," + VbCrlf
                sql = sql + " yuhyoend = '" + boarditem.yuhyoend + "' " + VbCrlf
                sql = sql + " where (id = " + boarditem.id + ") "
                rsget.Open sql, dbget, 1
	end Function

	Public Function delete(byval id)
                dim sql, i

                sql = "update [db_cs].[dbo].tbl_notice set isusing = 'N' "
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1
	end Function

	Public Function code2name(byval v)
                if (v = "01") then
                        code2name = "입점의뢰서"
                elseif (v = "02") then
                        code2name = "사업제휴의뢰서"
                elseif (v = "03") then
                        code2name = "특정상품의뢰"
                elseif (v = "04") then
                        code2name = "추천상품의뢰"
                else
                        code2name = ""
                end if
	end Function

	public function commentcheck(byval value)
		if len(value)<>0 then
				commentcheck="Y"
		else
				commentcheck="N"
		end if
	end function

end Class

'카데고리 대분류
Sub Drawcatelarge(selectBoxName,selectedId)
   dim query1, qyery2

	 '옵션 내용 DB에서 가져오기
   query1 = " select code_large,code_nm from db_item.dbo.tbl_Cate_large"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then

       '도돌이 시작
       do until rsget.EOF
			if Lcase(selectedId) = Lcase(rsget("code_large")) then
				response.write db2html(rsget("code_nm"))
			end if
           rsget.MoveNext
       loop
   end if
   rsget.close

End Sub

'카데고리 중분류
Sub Drawcatemid(selectcd1,selectBoxName,selectedId)
   dim query1, qyery2

	 '옵션 내용 DB에서 가져오기
   query1 = " select code_large,code_mid,code_nm from db_item.dbo.tbl_Cate_mid where code_large='" & selectcd1 & "'"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then

       '도돌이 시작
       do until rsget.EOF
			if Lcase(selectedId) = Lcase(rsget("code_mid")) then
				response.write db2html(rsget("code_nm"))
			end if
           rsget.MoveNext
       loop
   end if
   rsget.close

End Sub

Sub sbGetwork(ByVal selName, ByVal sIDValue, ByVal sScript)
	Dim strSql, arrList, intLoop
	strSql = " SELECT userid, username "
	strSql = strSql & " FROM db_partner.dbo.tbl_user_tenbyten  "
	strSql = strSql & " WHERE userid = '"&sIDValue&"' and userid <> '' "
	rsget.Open strSql,dbget
	IF not rsget.eof THEN
		arrList = rsget.getRows()
	End IF
	rsget.close

	IF isArray(arrList) THEN
%>
		<input type="text" class="text" name="doc_workername" value="<%=arrList(1,0)%>" size="10" readonly>
		<input type="button" class="button" value="지정" onClick="upcheworkerlist('<%=companyrequest.results(0).id%>')">
		<input type="button" class="button" value="삭제" onClick="upcheworkerDel('<%=companyrequest.results(0).id%>')">
<%
	Else
%>
		<input type="text" class="text" name="doc_workername" value="" size="10" readonly>
		<input type="button" class="button" value="지정" onClick="upcheworkerlist('<%=companyrequest.results(0).id%>')">
<%
	End IF
End Sub

Function getUpcheoneWorkname(eC)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT T.username "
	strSql = strSql & " FROM db_partner.dbo.tbl_user_tenbyten as T "
	strSql = strSql & " JOIN [db_cs].[dbo].tbl_company_request as D on T.userid = D.workid "
	strSql = strSql & " WHERE D.id = '"&eC&"' and T.userid <> '' "
	rsget.Open strSql,dbget
	IF not rsget.eof THEN
		getUpcheoneWorkname = rsget("username")
	End IF
	rsget.close
End Function

Function DrawWorkIdCombo(selectBoxName,selectedId)
	Dim tmp_str, strSql
%>
	<select name="<%=selectBoxName%>" class="select">
<%
	response.write("<option value='' selected>선택</option>")
	strSql = ""
	strSql = strSql & " select D.userid, D.username from "
	strSql = strSql & " [db_partner].[dbo].tbl_partner as A "
	strSql = strSql & " INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS D ON A.id = D.userid "
	strSql = strSql & " where A.isusing = 'Y' AND A.userdiv < 999 AND A.id <> '' AND Left(A.id,10) <> 'streetshop'  "
	strSql = strSql & " AND A.part_sn IN (14) "
	strSql = strSql & " ORDER BY A.part_sn ASC, A.posit_sn ASC, A.regdate ASC "
	rsget.Open strSql,dbget,1

	If not rsget.EOF Then
		Do Until rsget.EOF
			If rsget("userid") = selectedId Then
				tmp_str = " selected"
			End If
			response.write("<option value='"&rsget("userid")&"' "&tmp_str&">" + db2html(rsget("username")) + "</option>")
			tmp_str = ""
			rsget.MoveNext
		Loop
	End If
	rsget.close

	response.write("</select>")
End Function
%>

