<%
Class CUpcheDeliverPayJungsanItem
    public Fid
    public Fgubun01_name
    public Fgubun02_name
    public Forderserial
    public Fcustomername
    public Fuserid
    public Fwriteuser
    public Ffinishuser
    public Fcontents_jupsu
    public Fcontents_finish
    public Fregdate
    public Ffinishdate
    public Fmakerid
    public FreturnMethod
    public Frefundrequire
    public Fcanceltotal
    public Frefunditemcostsum
    public Frefundbeasongpay
    public Frefunddeliverypay

    public FjungsanDetailId
    public FjungsanDetailName
    public FjungsanSuplycash

    public Fadd_upchejungsandeliverypay
    public Fsitename

    public function IsJungsanDataExists()
        IsJungsanDataExists = Not IsNULL(FjungsanDetailId)
    end function

    Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CUpcheDeliverPayJungsan
    public FItemList()

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public FRectMakerid
	public FRectYYYYMM
	public FRectUserID
	public FRectOrderserial
	public FRectOnlyNotReged

    public FRectcksiteMM

	public Sub GetMonthlyDeliverPayJungsanList()
	    dim sqlStr, i
	    dim regStartDate, finishStartDate, finishEndDate

	    ''A001,A100,A002

	    finishStartDate = FRectYYYYMM + "-01"
	    finishEndDate   = CStr(DateSerial(Left(FRectYYYYMM,4),Right(FRectYYYYMM,2)+1,1))
	    ''regStartDate    = CStr(dateAdd("m",-5,finishStartDate))                                 ''2=>5

	    ''if (regStartDate<"2007-04-23") then regStartDate = "2007-04-23"

	    sqlStr = " select count(A.id) as cnt from "
	    sqlStr = sqlStr + " [db_cs].[dbo].tbl_new_as_list A WITH(NOLOCK)"
	    sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_as_refund_info R WITH(NOLOCK)" + VbCrlf
        sqlStr = sqlStr + " 	on  A.id=R.asid" + VbCrlf
        sqlStr = sqlStr + " 	left join db_cs.dbo.tbl_as_upcheAddjungsan U WITH(NOLOCK)" + VbCrlf
        sqlStr = sqlStr + " 	on  A.id=U.asid" + VbCrlf
        sqlStr = sqlStr + " 	left join [db_jungsan].[dbo].tbl_designer_jungsan_detail J WITH(NOLOCK)"
        sqlStr = sqlStr + " 	on J.gubuncd in ('witakchulgo','DT')" '','DL' lotteComM '수수료로
        sqlStr = sqlStr + " 	and J.itemid=0"
        sqlStr = sqlStr + " 	and J.detailidx=A.id"
        ''sqlStr = sqlStr + " 	and J.mastercode=A.orderserial"
        sqlStr = sqlStr + " 	left join [db_order].[dbo].tbl_order_master m WITH(NOLOCK)" + VbCrlf
        sqlStr = sqlStr + " 	on A.orderserial=m.orderserial " + VbCrlf

	    sqlStr = sqlStr + " where A.divcd in ('A004','A700','A000','A001','A100','A002','A200','A999')" + VbCrlf

	    if (FRectOrderserial<>"") or (FRectuserid<>"") then

	    else
            ''sqlStr = sqlStr + " and A.regdate>='" + regStartDate + "'" + VbCrlf
            sqlStr = sqlStr + " and A.finishdate>='" + finishStartDate + "'" + VbCrlf
            sqlStr = sqlStr + " and A.finishdate<'" + finishEndDate + "'" + VbCrlf
        end if
        sqlStr = sqlStr + " and A.currstate='B007'" + VbCrlf
        sqlStr = sqlStr + " and A.deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + " and ((A.divcd in ('A004','A000','A001','A100','A002','A200','A999') and A.requireupche='Y') or (A.divcd='A700'))" + VbCrlf
        sqlStr = sqlStr + " and (U.add_upchejungsandeliverypay<>0)" + VbCrlf
        ''sqlStr = sqlStr + " and ((R.refunddeliverypay <>0) or (U.add_upchejungsandeliverypay<>0))" + VbCrlf

        if FRectMakerid<>"" then
            sqlStr = sqlStr + " and A.makerid='" + FRectMakerid + "'" + VbCrlf
        end if

        if FRectOrderserial<>"" then
            sqlStr = sqlStr + " and A.orderserial='" + FRectOrderserial + "'" + VbCrlf
        end if

        if FRectUserID<>"" then
            sqlStr = sqlStr + " and A.userid='" + FRectUserID + "'" + VbCrlf
        end if

        if FRectOnlyNotReged<>"" then
            sqlStr = sqlStr + " and J.id Is NULL"
        end if

        sqlStr = sqlStr + " and a.makerid<>'10x10logistics'"

        if (FRectcksiteMM<>"") then
            sqlStr = sqlStr + " and m.sitename in ('lotteComM','Gs25')"
        end if
'rw  sqlStr
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		    FTotalCount = rsget("cnt")
		rsget.Close

        sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " A.id, C1.comm_name as gubun01_name, C2.comm_name as gubun02_name, " + VbCrlf
        sqlStr = sqlStr + " A.orderserial, A.customername, A.userid, A.writeuser,A.finishuser, " + VbCrlf
        sqlStr = sqlStr + " A.contents_jupsu, A.contents_finish, A.regdate, A.finishdate, " + VbCrlf
        sqlStr = sqlStr + " A.makerid, R.returnMethod, R.refundrequire, R.canceltotal, R.refunditemcostsum, IsNULL(R.refundbeasongpay,0) as refundbeasongpay,IsNULL(R.refunddeliverypay,0) as refunddeliverypay" + VbCrlf '', R.refundrequire-(R.refunditemcostsum+R.refunddeliverypay)" + VbCrlf
        sqlStr = sqlStr + " , J.id as jungsanDetailId, J.itemname as jungsanDetailName , J.suplycash as jungsanSuplycash"
        sqlStr = sqlStr + " , IsNULL(U.add_upchejungsandeliverypay,0) as add_upchejungsandeliverypay "
        sqlStr = sqlStr + " , m.sitename"
        ''sqlStr = sqlStr + " ,R.*" + VbCrlf
        sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list A WITH(NOLOCK)" + VbCrlf
        sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_as_refund_info R WITH(NOLOCK)" + VbCrlf
        sqlStr = sqlStr + " 	on  A.id=R.asid" + VbCrlf
        sqlStr = sqlStr + " 	left join db_cs.dbo.tbl_as_upcheAddjungsan U WITH(NOLOCK)" + VbCrlf
        sqlStr = sqlStr + " 	on  A.id=U.asid" + VbCrlf
        sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_cs_comm_code C1 WITH(NOLOCK)" + VbCrlf
        sqlStr = sqlStr + " 	on A.gubun01=C1.comm_cd " + VbCrlf
        sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_cs_comm_code C2 WITH(NOLOCK)" + VbCrlf
        sqlStr = sqlStr + " 	on A.gubun02=C2.comm_cd " + VbCrlf
        sqlStr = sqlStr + " 	left join [db_order].[dbo].tbl_order_master m WITH(NOLOCK)" + VbCrlf
        sqlStr = sqlStr + " 	on A.orderserial=m.orderserial " + VbCrlf

        sqlStr = sqlStr + " 	left join [db_jungsan].[dbo].tbl_designer_jungsan_detail J WITH(NOLOCK)"
        sqlStr = sqlStr + " 	on J.gubuncd in ('witakchulgo','DT')"  ''DL추가 2015/05/01 =>원복 '','DL' lotteComM '수수료로
        sqlStr = sqlStr + " 	and J.itemid=0"
        sqlStr = sqlStr + " 	and J.detailidx=A.id"
        ''sqlStr = sqlStr + " 	and J.mastercode=A.orderserial"

        sqlStr = sqlStr + " where A.divcd in ('A004','A700','A000','A001','A100','A002','A200','A999')" + VbCrlf

        if (FRectOrderserial<>"") or (FRectuserid<>"") then

	    else
            ''sqlStr = sqlStr + " and A.regdate>='" + regStartDate + "'" + VbCrlf
            sqlStr = sqlStr + " and A.finishdate>='" + finishStartDate + "'" + VbCrlf
            sqlStr = sqlStr + " and A.finishdate<'" + finishEndDate + "'" + VbCrlf
        end if
        sqlStr = sqlStr + " and A.currstate='B007'" + VbCrlf
        sqlStr = sqlStr + " and A.deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + " and ((A.divcd in ('A004','A000','A001','A100','A002','A200','A999') and A.requireupche='Y') or (A.divcd='A700'))" + VbCrlf
        sqlStr = sqlStr + " and (U.add_upchejungsandeliverypay<>0)" + VbCrlf
        ''sqlStr = sqlStr + " and ((R.refunddeliverypay <>0) or (U.add_upchejungsandeliverypay<>0))" + VbCrlf

        if FRectMakerid<>"" then
            sqlStr = sqlStr + " and A.makerid='" + FRectMakerid + "'" + VbCrlf
        end if

        if FRectOrderserial<>"" then
            sqlStr = sqlStr + " and A.orderserial='" + FRectOrderserial + "'" + VbCrlf
        end if

        if FRectUserID<>"" then
            sqlStr = sqlStr + " and A.userid='" + FRectUserID + "'" + VbCrlf
        end if

        if FRectOnlyNotReged<>"" then
            sqlStr = sqlStr + " and J.id Is NULL"
        end if

        if (FRectcksiteMM<>"") then
            sqlStr = sqlStr + " and m.sitename in ('lotteComM','Gs25')"
        end if

        sqlStr = sqlStr + " and a.makerid<>'10x10logistics'"
        sqlStr = sqlStr + " order by A.makerid, A.id desc" + VbCrlf
'rw   sqlStr
        rsget.PageSize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CUpcheDeliverPayJungsanItem

				FItemList(i).Fid                 = rsget("id")
                FItemList(i).Fgubun01_name       = db2html(rsget("gubun01_name"))
                FItemList(i).Fgubun02_name       = db2html(rsget("gubun02_name"))
                FItemList(i).Forderserial        = rsget("orderserial")
                FItemList(i).Fcustomername       = db2html(rsget("customername"))
                FItemList(i).Fuserid             = rsget("userid")
                FItemList(i).Fwriteuser          = rsget("writeuser")
                FItemList(i).Ffinishuser         = rsget("finishuser")
                FItemList(i).Fcontents_jupsu     = db2html(rsget("contents_jupsu"))
                FItemList(i).Fcontents_finish    = db2html(rsget("contents_finish"))
                FItemList(i).Fregdate            = rsget("regdate")
                FItemList(i).Ffinishdate         = rsget("finishdate")
                FItemList(i).Fmakerid            = rsget("makerid")
                FItemList(i).FreturnMethod       = rsget("returnMethod")
                FItemList(i).Frefundrequire      = rsget("refundrequire")
                FItemList(i).Fcanceltotal        = rsget("canceltotal")
                FItemList(i).Frefunditemcostsum  = rsget("refunditemcostsum")
                FItemList(i).Frefundbeasongpay   = rsget("refundbeasongpay")
                FItemList(i).Frefunddeliverypay  = rsget("refunddeliverypay")

                FItemList(i).FjungsanDetailId    = rsget("jungsanDetailId")
                FItemList(i).FjungsanDetailName  = rsget("jungsanDetailName")
                FItemList(i).FjungsanSuplycash   = rsget("jungsanSuplycash")

                FItemList(i).Fadd_upchejungsandeliverypay = rsget("add_upchejungsandeliverypay")
                FItemList(i).Fsitename = rsget("sitename")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end Sub

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage = 1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
End Class
%>
