<%
Class ipkumListitem
	public Fidx
	public FBankdate
	public Fgubun
	public Fipkumuser
	public Fipkumsum
	public Fchulkumsum
	public Fremainsum
	public Fbankname
	public Forderserial
	public Ffinishstr
	public Fipkumstate
	public Fregdate
	public Ffinishuser
	public Ftenbank

	public Fpaperexist

	public Fipkumgubun

	public FipkumCause
	Private sub Class_Intialize()

	end sub

	Private sub Class_Terminate()

	end Sub

end Class

Class IpkumListNetteller

	public Fbkno
	public Fbkacctno
	public Fbkname
	public Fbkdate
	public Fbkjukyo
	public Fbkcontent
	public Fbketc
	public Fbkinput
	public Fbkoutput
	public Fbkbalance
	public Fbkjango
	public Fbkxferdatetime
	public Ftlno
	public Fbkmemo1
	public Fbkmemo2
	public Fbkmemo3
	public Fbkmemo4
	public Fbktag1
	public Fbktag2
	public Fbktag3

	Private sub Class_Intialize()

	end sub

	Private sub Class_Terminate()

	end Sub

end Class

Class IpkumChecklist
	public Fipkumitem()
	public FipkumListNetteller()
	public Fipkumoneitem

	public Fckdate
	public FRectRegStart
	public FRectRegEnd
	public yyyy1,yyyy2 ''검색 날짜
	public mm1,mm2
	public dd1,dd2

	public FSearchtype01
	public FOrderby	'검색순서(최근일)
	public FTotalCount
	public CTenbank ''검색하는 은행
	public ipkumname '입금자명
	public ipkumstate  '입금 미확인. 미처리

	public FRectIpkumGubun
	public FRectIpkumIdx

	public FScrollCount
	public FCurrpage
	public FPagesize
	public FTotalPage
	public FResultCount
	public cksheep
	public Frectidx
	Private Sub Class_Initialize()
		dim Fipkumitem()
		FScrollCount=5
		FPagesize=200


	end Sub

	Private Sub Class_Terminate()

	end Sub

	Public Sub Getipkumlist()
	dim sqlStr,othersql,i
	if Ctenbank<> "" then
		otherSql = otherSql + " and i.tenbank='" & Ctenbank & "'" + vbcrlf
	end if

	if FRectIpkumIdx<> "" then
		otherSql = otherSql + " and i.idx='" & FRectIpkumIdx & "'" + vbcrlf
	end if

	if ipkumname<>"" then
		otherSql= otherSql + " and i.jukyo like '%" & ipkumname & "%'" + vbcrlf
	end if

	if FRectIpkumGubun<>"" then
		otherSql= otherSql + " and i.ipkumgubun = '" + CStr(FRectIpkumGubun) + "'" + vbcrlf
	end if

	if ipkumstate="1" then '매칭실패
		otherSql = otherSql + " and i.ipkumstate='1'" + vbcrlf
	elseif ipkumstate="0" then '미처리
		otherSql = otherSql + " and i.ipkumstate='0'" + vbcrlf
	else
	end if


	sqlStr = "select Count(idx) as cnt" + vbcrlf
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_ipkum_list i " + vbcrlf
	sqlStr = sqlStr + " where i.bankdate between '" & dateserial(CStr(yyyy1),CStr(mm1),CStr(dd1)) & "' and '" & dateserial(CStr(yyyy2),CStr(mm2),CStr(dd2)) & "'" + vbcrlf
	sqlStr = sqlStr + othersql
	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	FtotalCount=rsget("cnt")

	rsget.close

	if cksheep=1 then
		sqlStr = "select " + vbcrlf
	else
		sqlStr = "select top " & CStr(FCurrpage*FPagesize) + vbcrlf
	end if
	sqlStr = sqlStr + " i.idx,i.bankdate,i.gubun,i.jukyo,i.ipkumsum,i.chulkumsum,i.remainsum" + vbcrlf
	sqlStr = sqlStr + ",i.bankname,i.orderserial,i.finishstr,i.ipkumstate,i.regdate,i.finishuser,i.tenbank" + vbcrlf

	sqlStr = sqlStr + ", (case when IsNull(o.orderserial, '') = '' then 'N' else 'Y' end) as paperexist, i.ipkumgubun, i.ipkumCause " + vbcrlf
	sqlStr = sqlStr + " from " + vbcrlf
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_ipkum_list i " + vbcrlf
	sqlStr = sqlStr + " 	left join [db_order].[dbo].tbl_order_master o " + vbcrlf
	sqlStr = sqlStr + " 	on " + vbcrlf
	sqlStr = sqlStr + " 		1 = 1 " + vbcrlf
	sqlStr = sqlStr + " 		and i.orderserial = o.orderserial " + vbcrlf
	sqlStr = sqlStr + " 		and o.accountdiv = '7' " + vbcrlf
	sqlStr = sqlStr + " 		and ( " + vbcrlf
	sqlStr = sqlStr + " 			IsNull(o.authcode, '') <> '' " + vbcrlf
	sqlStr = sqlStr + " 			or " + vbcrlf
	sqlStr = sqlStr + " 			IsNull(o.cashreceiptreq, '') in ('R', 'T') " + vbcrlf
	sqlStr = sqlStr + " 	) " + vbcrlf

	sqlStr = sqlStr + " where i.bankdate between '" & dateserial(CStr(yyyy1),CStr(mm1),CStr(dd1)) & "' and '" & dateserial(CStr(yyyy2),CStr(mm2),CStr(dd2)) & "'" + vbcrlf
	sqlStr = sqlStr + othersql
	sqlStr = sqlStr + " order by i.regdate desc "

	rsget.pagesize=FPagesize
	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

	if FResultCount<1 then FResultCount=0

	FTotalPage = FTotalCount\FPagesize
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FTotalPage = FtotalPage +1
		end if


	i=0
	redim preserve Fipkumitem(FResultCount)


	if not rsget.eof then
	rsget.absolutepage=FCurrPage

	Do until rsget.eof
		set Fipkumitem(i) = new ipkumlistitem

			Fipkumitem(i).Fidx  =rsget("idx")
			Fipkumitem(i).FBankdate=rsget("bankdate")
			Fipkumitem(i).Fgubun=rsget("gubun")
			Fipkumitem(i).Fipkumuser=rsget("jukyo")
			Fipkumitem(i).Fipkumsum=rsget("ipkumsum")
			Fipkumitem(i).Fchulkumsum=rsget("chulkumsum")
			Fipkumitem(i).Fremainsum=rsget("remainsum")
			Fipkumitem(i).Fbankname=rsget("bankname")
			Fipkumitem(i).Forderserial=rsget("orderserial")
			Fipkumitem(i).Ffinishstr=rsget("finishstr")
			Fipkumitem(i).Fipkumstate=rsget("ipkumstate")
			Fipkumitem(i).Fregdate=rsget("regdate")
			Fipkumitem(i).Ffinishuser=rsget("finishuser")
			Fipkumitem(i).Ftenbank=rsget("tenbank")

			Fipkumitem(i).Fpaperexist=rsget("paperexist")
			Fipkumitem(i).Fipkumgubun=rsget("ipkumgubun")

			Fipkumitem(i).FipkumCause=rsget("ipkumCause")
		rsget.movenext
		i=i+1
		loop
		rsget.close
	end if
	end Sub


	Public Sub GetipkumlistByIdx()
		dim sqlStr
		dim i

		sqlStr = " select idx,bankdate,gubun,jukyo,ipkumsum,chulkumsum,remainsum" + vbcrlf
		sqlStr = sqlStr + ",bankname,orderserial,finishstr,ipkumstate,regdate,finishuser,tenbank" + vbcrlf
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_ipkum_list" + vbcrlf
		sqlStr = sqlStr + " where idx = '" + CStr(idx) + "'" + vbcrlf

		rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		set Fipkumoneitem = new ipkumlistitem

		if Not rsget.Eof then
			Fipkumoneitem.Fidx  =rsget("idx")
			Fipkumoneitem.FBankdate=rsget("bankdate")
			Fipkumoneitem.Fgubun=rsget("gubun")
			Fipkumoneitem.Fipkumuser=rsget("jukyo")
			Fipkumoneitem.Fipkumsum=rsget("ipkumsum")
			Fipkumoneitem.Fchulkumsum=rsget("chulkumsum")
			Fipkumoneitem.Fremainsum=rsget("remainsum")
			Fipkumoneitem.Fbankname=rsget("bankname")
			Fipkumoneitem.Forderserial=rsget("orderserial")
			Fipkumoneitem.Ffinishstr=rsget("finishstr")
			Fipkumoneitem.Fipkumstate=rsget("ipkumstate")
			Fipkumoneitem.Fregdate=rsget("regdate")
			Fipkumoneitem.Ffinishuser=rsget("finishuser")
			Fipkumoneitem.Ftenbank=rsget("tenbank")
		end if

		rsget.close

	end Sub

	Public Sub GetipkumlistAccounts()
		dim sqlStr,i

		sqlStr = "select Count(bkno) as cnt "
		sqlStr = sqlStr + " from [db_log].[dbo].tblbank L, "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_bank_div A "
		sqlStr = sqlStr + " where  L.bkacctno=A.accountno "

		if Ctenbank<> "" then
		sqlStr = sqlStr + " and L.bkacctno='" & Ctenbank & "'"
		end if

		if (FSearchtype01<>"") then
			sqlStr= sqlStr + " and L.bkjukyo like '%" & ipkumname & "%'"
		end if

		if (Fckdate<>"") then
			sqlStr = sqlStr + " and bkdate >='" + CStr(Replace(FRectRegStart,"-","")) + "'"
			sqlStr = sqlStr + " and bkdate <'" + CStr(Replace(FRectRegEnd,"-","")) + "'"
		end if



		rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalCount=rsget("cnt")

		rsget.close


		sqlStr = "select top " & CStr(FCurrpage*FPagesize)
		sqlStr = sqlStr + " L.* "
		sqlStr = sqlStr + " from [db_log].[dbo].tblbank L, "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_bank_div A "
		sqlStr = sqlStr + " where  L.bkacctno=A.accountno "

		if Ctenbank<> "" then
		sqlStr = sqlStr + " and L.bkacctno='" & Ctenbank & "'"
		end if

		if (FSearchtype01<>"") then
			sqlStr= sqlStr + " and L.bkjukyo like '%" & ipkumname & "%'"
		end if

		if (Fckdate<>"") then
			sqlStr = sqlStr + " and bkdate >='" + CStr(Replace(FRectRegStart,"-","")) + "'"
			sqlStr = sqlStr + " and bkdate <'" + CStr(Replace(FRectRegEnd,"-","")) + "'"
		end if

		if (FOrderby<>"") then
			sqlStr = sqlStr + " order by L.bkno desc"  '최근일순
		end if

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve Fipkumitem(FResultCount)
		i=0

		if not rsget.Eof then

		rsget.absolutepage = FCurrPage

		Do until rsget.eof
			set Fipkumitem(i) = new IpkumListNetteller

				Fipkumitem(i).Fbkno				= rsget("bkno")
				Fipkumitem(i).Fbkacctno         = rsget("bkacctno")
				Fipkumitem(i).Fbkname           = rsget("bkname")
				Fipkumitem(i).Fbkdate           = rsget("bkdate")
				Fipkumitem(i).Fbkjukyo          = rsget("bkjukyo")
				Fipkumitem(i).Fbkcontent        = rsget("bkcontent")
				Fipkumitem(i).Fbketc            = rsget("bketc")
				Fipkumitem(i).Fbkinput          = rsget("bkinput")
				Fipkumitem(i).Fbkoutput         = rsget("bkoutput")
				Fipkumitem(i).Fbkbalance        = rsget("bkbalance")
				Fipkumitem(i).Fbkjango          = rsget("bkjango")
				Fipkumitem(i).Fbkxferdatetime   = rsget("bkxferdatetime")
				Fipkumitem(i).Ftlno             = rsget("tlno")
				Fipkumitem(i).Fbkmemo1          = rsget("bkmemo1")
				Fipkumitem(i).Fbkmemo2          = rsget("bkmemo2")
				Fipkumitem(i).Fbkmemo3          = rsget("bkmemo3")
				Fipkumitem(i).Fbkmemo4          = rsget("bkmemo4")
				Fipkumitem(i).Fbktag1           = rsget("bktag1")
				Fipkumitem(i).Fbktag2           = rsget("bktag2")
				Fipkumitem(i).Fbktag3           = rsget("bktag3")

			rsget.movenext
			i=i+1
		loop
		rsget.close
		end if

	end Sub


	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrPage-1)\FScrollCount)*FScrollCount +1
	end Function
end class
%>
