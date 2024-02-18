<%

'id, reqcd, companyname, chargename, chargeposition, zipcode, address, phone, hp, email, companyurl, companycomments, reqcomment, attachfile, regdate, finishdate
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
	private Fregdate
	private Ffinishdate
	private Fcategubun
	private Freplyuser
	private Freplycomment
	private Fsellgubun
	private FipjumYN

        '==========================================================================

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

	Property Get regdate()
		regdate = Fregdate
	end Property

	Property Get finishdate()
		finishdate = Ffinishdate
	end Property

        '==========================================================================

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

	Property Let regdate(byVal v)
		Fregdate = v
	end Property

	Property Let finishdate(byVal v)
		Ffinishdate = v
	end Property

        '==========================================================================
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

        'id, reqcd, companyname, chargename, chargeposition, zipcode, address, phone, hp, email, companyurl, companycomments, reqcomment, attachfile, regdate, finishdate
        '======================================================================
	Public Function list()
                dim sql, i

                sql = " select count(id) as cnt from [db_cs].[10x10].tbl_company_request "
				sql = sql + " where id<>0"


				if FReqcd<>"" then
					sql = sql + " and reqcd='" + FReqcd + "'"
				end if

				if FOnlyNotFinish<>"" then
					sql = sql + " and finishdate is NULL"
					sql = sql + " and replycomment is NULL"
				end if

				if FRectSearchKey<>"" then
					sql = sql + " and companyname like '%" + FRectSearchKey + "%'"
				end if

				if FRectCatevalue <>"" then
					if FRectCatevalue="etc" then
						sql = sql + " and categubun='etc'" +vbcrlf
					else
						sql = sql + " and categubun='" + FRectCatevalue + "'"
					end if
				end if

				if FipjumYN <>"" then
					sql = sql + " and ipjumYN='" + FipjumYN + "'" +vbcrlf
				end if

				sql=sql + " and isusing='Y'"

		rsget.Open sql, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.Close

                sql = " select top " + CStr(FPageSize*FCurrPage) + " id, reqcd, companyname, chargename, chargeposition, zipcode, address, phone, hp, email, companyurl, companycomments, reqcomment, attachfile, regdate, finishdate,categubun,sellgubun,ipjumYN,replycomment from [db_cs].[10x10].tbl_company_request "
                sql = sql + " where id<>0"

                if FReqcd<>"" then
					sql = sql + " and reqcd='" + FReqcd + "'"
				end if

				if FOnlyNotFinish<>"" then
					sql = sql + " and finishdate is NULL"
				end if

				if FRectSearchKey<>"" then
					sql = sql + " and companyname like '%" + FRectSearchKey + "%'"
				end if

				if FRectCatevalue <>"" then
					if FRectCatevalue="etc" then
						sql = sql + " and categubun='etc'" +vbcrlf
					else
						sql = sql + " and categubun='" + FRectCatevalue + "'"
					end if
				end if

				if FipjumYN <>"" then
					sql = sql + " and ipjumYN='" + FipjumYN + "'" +vbcrlf
				end if
               	sql=sql + " and isusing='Y'"
                sql = sql + " order by regdate desc "

		if FPageSize<>0 then
			rsget.pagesize = PageSize
		end if
                rsget.Open sql, dbget, 1
                'response.write sql

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
                                results(i).regdate = rsget("regdate")
                                results(i).finishdate = rsget("finishdate")
                                results(i).replycomment=db2html(rsget("replycomment"))
                                results(i).categubun = rsget("categubun")
                                results(i).sellgubun=rsget("sellgubun")
                                results(i).ipjumYN=rsget("ipjumYN")


				rsget.MoveNext
				i = i + 1
                        loop
                end if
                rsget.close
	end Function

	Public Function read(byval id)
                dim sql, i

                sql = " select top 1 id, reqcd, companyname, chargename, chargeposition, zipcode, address, phone, hp, email, companyurl, companycomments, reqcomment, attachfile, regdate, finishdate,replyuser,replycomment,categubun,sellgubun,ipjumYN from [db_cs].[10x10].tbl_company_request "
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1
                'response.write sql

                redim preserve results(rsget.RecordCount)

                if not rsget.EOF then
                        set results(0) = new CCompanyRequestItem

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
                        results(i).regdate = rsget("regdate")
                        results(i).finishdate = rsget("finishdate")
                        results(i).replyuser =db2html(rsget("replyuser"))
                        results(i).replycomment=db2html(rsget("replycomment"))
                        results(i).categubun = rsget("categubun")
                        results(i).sellgubun=rsget("sellgubun")
                        results(i).ipjumYN=rsget("ipjumYN")
                end if
                rsget.close
	end Function

	Public Function write(byval boarditem)
                dim sql, i

                sql = " insert into [db_cs].[10x10].tbl_company_request(reqcd, companyname, chargename, chargeposition, zipcode, address, phone, hp, email, companyurl, companycomments, reqcomment, attachfile, regdate) "
                sql = sql + " values('" + boarditem.reqcd + "', '" + boarditem.companyname + "', '" + boarditem.chargename + "', '" + boarditem.chargeposition + "', '" + boarditem.zipcode + "', '" + boarditem.address + "', '" + boarditem.phone + "', '" + boarditem.hp + "', '" + boarditem.email + "', '" + boarditem.companyurl + "', '" + boarditem.companycomments + "', '" + boarditem.reqcomment + "', '" + boarditem.attachfile + "', getdate()) "
                rsget.Open sql, dbget, 1
	end Function

	Public Function finish(byval id)
                dim sql, i

                sql = "update [db_cs].[10x10].tbl_company_request set finishdate = getdate() " + VbCrlf
                sql = sql + " where (id = '" + id + "') "
                rsget.Open sql, dbget, 1
	end Function


	Public Function delitem(byval id)
		        dim sql, i

		        sql = "update [db_cs].[10x10].tbl_company_request set isusing = 'N'" + VbCrlf
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1

	end Function

	Public Function sellchange(byval id,sellgubun)
                dim sql, i

                sql = "update [db_cs].[10x10].tbl_company_request set sellgubun = '" + sellgubun + "'" + VbCrlf
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1
	end Function

	Public Function ipjumchange(byval id,ipjumYN)
                dim sql, i

                sql = "update [db_cs].[10x10].tbl_company_request set ipjumYN = '" + ipjumYN + "'" + VbCrlf
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1
	end Function

	Public Function catechange(byval id,categubun)
                dim sql, i

                sql = "update [db_cs].[10x10].tbl_company_request set categubun = '" + categubun + "'" + VbCrlf
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1
	end Function

	Public Function writecomm(byval id,user,comment)
                dim sql, i

                sql = "update [db_cs].[10x10].tbl_company_request " + vbcrlf
                sql = sql + " set replyuser='" + user + "'" + vbcrlf
                sql = sql + ", replycomment='" + comment + "'" + vbcrlf
                sql = sql + ", finishdate=getdate()" + vbcrlf
                sql = sql + " where (id = " + id + ") "+vbcrlf
                rsget.Open sql, dbget, 1

	end Function

	Public Function modify(byval boarditem)
                dim sql, i

                sql = "update [db_cs].[10x10].tbl_notice " + VbCrlf
                sql = sql + " set title = '" + boarditem.title + "'," + VbCrlf
                sql = sql + " contents = '" + boarditem.contents + "'," + VbCrlf
                sql = sql + " yuhyostart = '" + boarditem.yuhyostart + "'," + VbCrlf
                sql = sql + " yuhyoend = '" + boarditem.yuhyoend + "' " + VbCrlf
                sql = sql + " where (id = " + boarditem.id + ") "
                rsget.Open sql, dbget, 1
	end Function

	Public Function delete(byval id)
                dim sql, i

                sql = "update [db_cs].[10x10].tbl_notice set isusing = 'N' "
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1
	end Function

	Public Function code2name(byval v)
                if (v = "01") then
                        code2name = "ÀÔÁ¡ÀÇ·Ú¼­"
                elseif (v = "02") then
                        code2name = "»ç¾÷Á¦ÈÞÀÇ·Ú¼­"
                elseif (v = "03") then
                        code2name = "Æ¯Á¤»óÇ°ÀÇ·Ú"
                elseif (v = "04") then
                        code2name = "ÃßÃµ»óÇ°ÀÇ·Ú"
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
%>

    