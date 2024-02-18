<%

'// ============================================================================
'// SCM SCM SCM !!!!
'// ============================================================================

'####################################################
' Description :  출고지시 클래스
' History : 2009.03.28 서동석 생성
'			2011.05.18 한용민 수정
'####################################################

class CBaljuIpgoItem
	public FBaljuCode
	''public FMakerid
	public FSiteSeq
	public FBrandName
	public FItemGubun
	public FItemID
	public FItemName
	public FItemOption
	public FItemOptionName
	public FOrgSellcash
	public FBaljuNo
	public fitemno
	public FIpgoNo
	public FPrintNo
	public FMiIpgoNo
	public FTmpIpgoNo
	public FImageSmall
	public FImageList
	public FPageNo
	public Fpublicbarcode
	'''public FRecCode			''브랜드 대표랙코드
	public FItemRackCode	''상품   랙코드
    public FdasPackNo
    public FacctItemNo
    public Fdasindex
    public ForderItemNo
    public Fsellyn
    public Flimityn
    public Flimitno
    public Flimitsold
	public fdeliverytype
	public fmakerid

    public FwarehouseCd

    public Function IsUpcheBeasong()
		if Fdeliverytype="2" or Fdeliverytype="5" or Fdeliverytype="9" or Fdeliverytype="7" then
			IsUpcheBeasong = true
		else
			IsUpcheBeasong = false
		end if
	end function

	public function IsSoldOut
		IsSoldOut = ((FSellYn="N") or ((FLimitYn="Y") and (GetLimitEa=0)))
	end function

	public function GetIsSlodOutText
		if IsSoldOut then
			GetIsSlodOutText = "SoldOut"
		else
			GetIsSlodOutText = ""
		end if
	end function

	public function GetIsLimitText
		if (FLimitYn="Y") then
			GetIsLimitText = "한정(" + CStr(GetLimitEa) + ")"
		else
			GetIsLimitText = ""
		end if
	end function

	public function GetLimitEa
		GetLimitEa = FLimitNo-FLimitSold
		if GetLimitEa<1 then GetLimitEa=0
	end function

	public function GetBarCode()
		if ((Not IsNull(FItemGubun)) and (CStr(FItemGubun) <> "")) then
			GetBarCode = FItemGubun + Format00(6,FItemID) + FItemOption
		else
			GetBarCode = Format00(2,FSiteSeq) + Format00(6,FItemID) + FItemOption
		end if

		if (FItemID >= 1000000) then
			'// TODO : 차후에 /lib/BarcodeFunction.asp 를 이용하도록 변경해야 한다.
			if ((Not IsNull(FItemGubun)) and (CStr(FItemGubun) <> "")) then
				GetBarCode = FItemGubun + Format00(8,FItemID) + FItemOption
			else
				GetBarCode = Format00(2,FSiteSeq) + Format00(8,FItemID) + FItemOption
			end if
		end if

		'GetBarCode = Format00(2,FSiteSeq) + Format00(6,FItemID) + FItemOption
	end function

	public function GetItemRackCode()
		GetItemRackCode = FItemRackCode

		if IsNULL(FItemRackCode) then GetItemRackCode  = "9999"
	end function

	public function GetIpgoNoColor()
		if (FBaljuNo-FIpgoNo)>0 then
			GetIpgoNoColor = "#3333CC"
		elseif (FBaljuNo-FIpgoNo)<0 then
			GetIpgoNoColor = "#CC3333"
		else
			GetIpgoNoColor = "#000000"
		end if
	end function

	public function GetMiIpgoColor()
		dim miipgocolor
		miipgocolor = GetMiIpgoNo
		if miipgocolor>0 then
			GetMiIpgoColor = "#CC3333"
		elseif miipgocolor<0 then
			GetMiIpgoColor = "#3333CC"
		else
			GetMiIpgoColor = "#000000"
		end if
	end function

	public function GetMiIpgoNo()
		GetMiIpgoNo = FBaljuNo - FIpgoNo
	end function

	Private Sub Class_Initialize()
		FOrgSellcash	= 0
		FBaljuNo	= 0
		FIpgoNo		= 0
		FMiIpgoNo	= 0
		FTmpIpgoNo  = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CBaljuChulgoItem
    public FsiteSeq

	public Forderserial

	public Fjumundiv
	public Fuserid

	public Ftotalcost
	public Ftotalmileage
	public Ftotalsum
	public FOrderStatus
	public Fipkumdate
	public Fregdate

	public Fcancelyn
	public Fbuyname
	public Fbuyphone
	public Fbuyhp
	public Fbuyemail
	public Freqname
	public Freqzipcode
	public Freqaddress
	public Freqphone
	public Freqhp
	public Fdeliverymemo
	public Fdeliverno
	''public Fsitename

	public Fsubtotalprice
	public Freqzipaddr

	public Fsongjangdiv
	public Frdsite
	public Ftencardspend
	public Fbeasongmemo
	public Freqdate
	public Freqtime
	public Fcardribbon
	public Fmessage
	public Ffromname
	public Fcashreceiptreq
	public Finireceipttid
	public Freferip
	public Fuserlevel
	public Flinkorderserial
	public Fspendmembership
	public Fsentenceidx

	public FBaljuKey

	public Fbaljuflag
	public Ffixdate
	public Fprintdate

	public FlocalDlvInclude
    public Fdasindex
    public FtenbeaCnt

    public Fitem1cnt
    public Fitem2cnt
    public Fitem3cnt
    public Fitem1Packcnt
    public Fitem2Packcnt
    public Fitem3Packcnt

    'public Fbeadaldiv
	'public Fbeadaldate
    'public Faccountname
	'public Faccountdiv
	'public Faccountno
	'public Ftotalvat
    'public Fidx
    'public Fpaygatetid
	'public Fdiscountrate
	''public FBaljudetailid
	'public Fresultmsg
	'public Frduserid
	'public Fmilelogid
	'public Fmiletotalprice
	'public Fjungsanflag
	'public Fauthcode

	public function GetBaljuSiteName()

		GetBaljuSiteName = fnGetSiteNameBySiteSeq(CStr(FsiteSeq))

    end function

    public function getDASBaljuFlagColor()
        if IsNULL(Fbaljuflag) then
			getDASBaljuFlagColor = "#FFFFFF"
		elseif CStr(Fbaljuflag)="0" then
		    getDASBaljuFlagColor = "#FFFFFF"
		elseif CStr(Fbaljuflag)="1" then
		    getDASBaljuFlagColor = "#777777"
		elseif CStr(Fbaljuflag)="2" then
		    getDASBaljuFlagColor = "#999999"
		elseif CStr(Fbaljuflag)="3" then
		    getDASBaljuFlagColor = "#99CC99"
		elseif CStr(Fbaljuflag)="5" then
		    getDASBaljuFlagColor = "#CC99CC"
		elseif CStr(Fbaljuflag)="7" then
		    getDASBaljuFlagColor = "#FF6666"
		else
		    getDASBaljuFlagColor = "#FFFFFF"
		end if
    end function

	public function getBaljuStateName()
		''if IsNULL(FlocalDlvInclude) then
		''	getBaljuStateName = "-"
		if CStr(Fbaljuflag)="0" then
			getBaljuStateName = "입고대기"
		elseif CStr(Fbaljuflag)="1" then
			getBaljuStateName = "취소"
		elseif CStr(Fbaljuflag)="2" then
			getBaljuStateName = "미배"
		elseif CStr(Fbaljuflag)="3" then
			getBaljuStateName = "출력대기"
		elseif CStr(Fbaljuflag)="5" then
			getBaljuStateName = "출력완료"
		elseif CStr(Fbaljuflag)="6" then
			getBaljuStateName = "-"
		elseif CStr(Fbaljuflag)="7" then
			getBaljuStateName = "출고완료"
		elseif CStr(Fbaljuflag)="8" then
			getBaljuStateName = "전송완료"
		elseif CStr(Fbaljuflag)="9" then
			getBaljuStateName = "기타완료"
		else
			getBaljuStateName = ""
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CBaljuIpgo
	public FItemList()
    public FOneItem
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FRectSiteSeq
	public FRectBaljuKey
	public FRectSiteBaljuKey
	public FRectSearchType
	public FRectItemGubun
	public FRectItemID
	public FRectItemOption
	public FRectIpgoNo
	public FRectBarcode
	public FRectMakePageSize
	public FRectPreMakeItemNo
	public FRectPageNo
	public FRectSectionSize
	public FPageNoStart
	public FPageNoEnd

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	'온라인
	'// 텐텐 사이트 출고지시키로 로직스 출고지시키 찾기
	public function GetBaljuKeyWithSiteBaljuKey()
		dim sqlstr

		sqlstr = " select top 1 baljukey "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_aLogistics.dbo.tbl_Logistics_baljumaster "
		sqlStr = sqlStr + " where siteBaljuid = " + CStr(FRectSiteBaljuKey) + " and siteSeq = 10 "
		''response.write sqlstr
		rsget_Logistics.Open sqlStr,dbget_Logistics,1

		if rsget_Logistics.Eof then
			GetBaljuKeyWithSiteBaljuKey = 0
		else
			GetBaljuKeyWithSiteBaljuKey = rsget_Logistics("baljukey")
		end if
		rsget_Logistics.close

	end function

	'온라인
	public function GetMaxPage()
		dim sqlstr

		sqlstr = " select IsNull(max(pageno), 0) as maxpage from db_aLogistics.dbo.tbl_logistics_baljuipgo" + VbCrlf
		sqlStr = sqlStr + " where BaljuKey=" + CStr(FRectBaljuKey) + VbCrlf
		''response.write sqlstr
		rsget_Logistics.Open sqlStr,dbget_Logistics,1

		if rsget_Logistics.Eof then
			GetMaxPage = 0
		else
			GetMaxPage = rsget_Logistics("maxpage")
		end if
		rsget_Logistics.close

	end function

	'온라인
	public sub MakeBaljuPage()
		dim sqlstr,i
		dim maxpage
		dim remaincount
		dim maxloop
		dim notpageingExists
		dim currrackno

		remaincount = 0

		sqlstr = "select count(*) as cnt "
		sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljuipgo" + VbCrlf
		sqlStr = sqlStr + " where BaljuKey=" + CStr(FRectBaljuKey) + VbCrlf
		sqlStr = sqlStr + " and pageno=0"

		rsget_Logistics.Open sqlStr,dbget_Logistics,1
			remaincount = rsget_Logistics("cnt")
		rsget_Logistics.close
		'response.write sqlStr


		if (remaincount>0) then
			''최종페이지
			sqlstr = " select max(pageno) as maxpage "
			sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljuipgo" + VbCrlf
			sqlStr = sqlStr + " where BaljuKey=" + CStr(FRectBaljuKey) + VbCrlf
			rsget_Logistics.Open sqlStr,dbget_Logistics,1
			if rsget_Logistics.Eof then
				maxpage = 1
			else
				maxpage = rsget_Logistics("maxpage") + 1
			end if
			rsget_Logistics.close
		end if

		if (remaincount>0) then
			'' N개 이상인 상품 먼저.
			maxloop = 200 '(remaincount \ FRectMakePageSize) + 1

			sqlstr = "select count(*) as cnt "
			sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljuipgo" + VbCrlf
			sqlStr = sqlStr + " where BaljuKey=" + CStr(FRectBaljuKey) + VbCrlf
			sqlStr = sqlStr + " and pageno=0" + VbCrlf
			sqlStr = sqlStr + " and baljuno>=" + CStr(FRectPreMakeItemNo) + VbCrlf

			rsget_Logistics.Open sqlStr,dbget_Logistics,1
				notpageingExists = rsget_Logistics("cnt")>0
			rsget_Logistics.close


			for i=0 to maxloop-1
				''랙코드 앞 두자리 별로 페이징 - 변경
				currrackno = "NONE"

				''sqlstr = " select top 1 convert(int,IsNULL(itemrackcode,9999)/100) as itemrackcode" + VbCrlf
				sqlstr = " select top 1 IsNULL(itemrackcode,'9999') as itemrackcode" + VbCrlf
				sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljuipgo b" + VbCrlf
				sqlStr = sqlStr + "     left join db_aLogistics.dbo.tbl_logistics_item i"
				sqlStr = sqlStr + "     on b.siteSeq=i.siteSeq"
				sqlStr = sqlStr + "     and b.itemid=i.siteitemid"
				sqlStr = sqlStr + "     and b.itemoption=i.siteitemoption"
				sqlStr = sqlStr + " where BaljuKey=" + CStr(FRectBaljuKey) + VbCrlf
				sqlStr = sqlStr + " and pageno=0" + VbCrlf
				sqlStr = sqlStr + " order by IsNULL(i.itemrackcode,'9999'), b.itemid, b.itemoption" + VbCrlf

                rsget_Logistics.CursorLocation = adUseClient
		        rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly
				if not rsget_Logistics.Eof then
					currrackno = rsget_Logistics("itemrackcode")
				end if
				rsget_Logistics.close

				if currrackno="NONE" then exit for

				sqlstr = " update db_aLogistics.dbo.tbl_logistics_baljuipgo" + VbCrlf
				sqlstr = sqlStr + " set pageno=" + CStr(maxpage) + VbCrlf
				sqlstr = sqlStr + " from " + VbCrlf
				sqlstr = sqlStr + " ( " + VbCrlf
				sqlstr = sqlStr + " 	select top " + CStr(FRectMakePageSize) + " b.BaljuKey,b.siteSeq,b.itemid,b.itemoption "
				sqlStr = sqlStr + " 	from db_aLogistics.dbo.tbl_logistics_baljuipgo b" + VbCrlf
				sqlStr = sqlStr + "         left join db_aLogistics.dbo.tbl_logistics_item i"
				sqlStr = sqlStr + "         on b.siteSeq=i.siteSeq"
				sqlStr = sqlStr + "         and b.itemid=i.siteitemid"
				sqlStr = sqlStr + "         and b.itemoption=i.siteitemoption"
				sqlStr = sqlStr + " 	where BaljuKey=" + CStr(FRectBaljuKey) + VbCrlf
				sqlStr = sqlStr + " 	and pageno=0"
				sqlStr = sqlStr + " 	order by IsNULL(i.itemrackcode,'9999'), b.itemid, b.itemoption"
				sqlStr = sqlStr + " ) T " + VbCrlf
				sqlStr = sqlStr + " where db_aLogistics.dbo.tbl_logistics_baljuipgo.BaljuKey=T.BaljuKey" + VbCrlf
				sqlStr = sqlStr + " and db_aLogistics.dbo.tbl_logistics_baljuipgo.siteSeq=T.siteSeq" + VbCrlf
				sqlStr = sqlStr + " and db_aLogistics.dbo.tbl_logistics_baljuipgo.itemid=T.itemid" + VbCrlf
				sqlStr = sqlStr + " and db_aLogistics.dbo.tbl_logistics_baljuipgo.itemoption=T.itemoption" + VbCrlf

				dbget_Logistics.Execute sqlStr


'				''랙코드 앞 두자리 별로 페이징 - 변경
'				sqlstr = " update db_aLogistics.dbo.tbl_logistics_baljuipgo" + VbCrlf
'				sqlstr = sqlStr + " set pageno=0" + VbCrlf
'				sqlStr = sqlStr + " where BaljuKey=" + CStr(FRectBaljuKey) + VbCrlf
'				sqlStr = sqlStr + " and pageno=" + CStr(maxpage) + VbCrlf
'				sqlStr = sqlStr + " and convert(int,IsNULL(itemrackcode,9999)/100)<>" + CStr(currrackno)
'
'				dbget_Logistics.Execute sqlStr

				maxpage = maxpage + 1

				sqlstr = "select count(*) as cnt from db_aLogistics.dbo.tbl_logistics_baljuipgo" + VbCrlf
				sqlStr = sqlStr + " where BaljuKey=" + CStr(FRectBaljuKey) + VbCrlf
				sqlStr = sqlStr + " and pageno=0" + VbCrlf

				rsget_Logistics.Open sqlStr,dbget_Logistics,1
					notpageingExists = rsget_Logistics("cnt")>0
				rsget_Logistics.close

				if Not notpageingExists then exit for
			next

		end if
	end sub

	'온라인	' /admin/ordermaster/pop_logistics_baljuitemlist.asp
	public Sub GetBaljuIpgoByPageRect()
		dim sqlStr,i

		sqlStr = " select top 1000"
		sqlStr = sqlStr & " b.baljukey, b.SiteSeq, b.itemid, b.itemoption,b.baljuno, b.ipgono, b.pageno"
		sqlStr = sqlStr & " , i.siteitemname, i.siteoptionname, IsNULL(i.orgsellprice,0) as orgsellprice"
		sqlStr = sqlStr & " , i.itemrackcode, i.imageSmall, i.brandName"
		sqlStr = sqlStr & " from db_aLogistics.dbo.tbl_logistics_baljuipgo b with (nolock)"
		sqlStr = sqlStr & " left join db_aLogistics.dbo.tbl_logistics_item i with (nolock)"
		sqlStr = sqlStr & "     on b.siteSeq=i.siteSeq"
		sqlStr = sqlStr & "     and b.itemid=i.siteitemid"
		sqlStr = sqlStr & "     and b.itemoption=i.siteitemoption"
		sqlStr = sqlStr & " where BaljuKey=" + CStr(FRectBaljuKey) + VbCrlf
		sqlStr = sqlStr & " and pageno>=" + CStr(FPageNoStart)
		sqlStr = sqlStr & " and pageno<=" + CStr(FPageNoEnd)
		sqlStr = sqlStr & " order by b.pageno, i.itemrackcode, b.itemid, b.itemoption"

		'response.write sqlStr & "<br>"
        rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget_Logistics.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget_Logistics.EOF  then
			do until rsget_Logistics.eof
				set FItemList(i) = new CBaljuIpgoItem

				FItemList(i).FBaljuCode      = FRectBaljuKey
				FItemList(i).FBrandName	     = rsget_Logistics("brandName")
				FItemList(i).FSiteSeq	     = rsget_Logistics("SiteSeq")
				FItemList(i).FItemID         = rsget_Logistics("itemid")
				FItemList(i).FItemOption     = rsget_Logistics("itemoption")
				FItemList(i).FItemName       = db2html(rsget_Logistics("siteitemname"))
				FItemList(i).FItemOptionName = db2html(rsget_Logistics("siteoptionname"))
				FItemList(i).FOrgSellcash    = rsget_Logistics("orgsellprice")
				FItemList(i).FBaljuNo        = rsget_Logistics("baljuno")
				FItemList(i).FIpgoNo         = rsget_Logistics("ipgono")
				FItemList(i).FPageNo		 = rsget_Logistics("pageno")
				FItemList(i).FItemRackCode	 = rsget_Logistics("itemrackcode")

				IF (FItemList(i).FSiteSeq="10") THEN
				    FItemList(i).FimageSmall    = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget_Logistics("imageSmall")
				ELSE
				    FItemList(i).FimageSmall     = rsget_Logistics("imageSmall")
			    END IF

				i=i+1
				rsget_Logistics.moveNext
			loop
		end if
		rsget_Logistics.Close

	end Sub

	' /admin/ordermaster/pop_logistics_baljuitem.asp
	public Sub GetBaljuIpgoitem()
		dim sqlStr,i

		if FRectBaljuKey="" or isnull(FRectBaljuKey) then exit Sub

		sqlStr = " select" & vbcrlf
		sqlStr = sqlStr & " '10' as itemgubun,d.itemid, isnull(d.itemoption,'0000') as itemoption, IsNull(i.warehouseCd, 'BLK') as warehouseCd " & vbcrlf
		sqlStr = sqlStr & " , sum(isnull(d.itemno,0)) as itemno, isnull(d.makerid,'') as makerid" & vbcrlf
		sqlStr = sqlStr & " from [db_aLogistics].[dbo].[tbl_Logistics_baljumaster] bm with (nolock)" & vbcrlf
		sqlStr = sqlStr & " join [db_aLogistics].[dbo].[tbl_Logistics_baljudetail] bd with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on bm.baljuKey=bd.baljuKey" & vbcrlf
		sqlStr = sqlStr & " join [db_aLogistics].[dbo].[tbl_Logistics_order_detail] d with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on bd.orderserial=d.orderserial" & vbcrlf
		sqlStr = sqlStr & " 	and d.cancelyn<>'Y'" & vbcrlf
		sqlStr = sqlStr & " 	and d.itemid not in (0,100)" & vbcrlf
		sqlStr = sqlStr & " 	and d.isupchebeasong='N'" & vbcrlf
        sqlStr = sqlStr & " 	and bd.siteseq = d.siteseq " & vbcrlf
		sqlStr = sqlStr & " 	left join [db_aLogistics].[dbo].[tbl_Logistics_item] i "
		sqlStr = sqlStr & " 	on "
		sqlStr = sqlStr & " 		1 = 1 "
		sqlStr = sqlStr & " 		and d.siteSeq = i.siteSeq "
        sqlStr = sqlStr & " 		and d.itemgubun = i.siteItemGubun "
		sqlStr = sqlStr & " 		and d.itemid = i.siteItemid "
		sqlStr = sqlStr & " 		and d.itemoption = i.siteItemOption "
		sqlStr = sqlStr & " where bm.sitebaljuid = "& FRectBaljuKey &"" & vbcrlf
		sqlStr = sqlStr & " group by d.itemid, isnull(d.itemoption,'0000'), isnull(d.makerid,''), IsNull(i.warehouseCd, 'BLK') " & vbcrlf
		sqlStr = sqlStr & " order by makerid asc, d.itemid asc, itemoption asc" & vbcrlf

		'sqlStr = " select isnull(i.brandid,'') as makerid, '10' as itemgubun, b.itemid, isnull(b.itemoption,'0000') as itemoption" & vbcrlf
		'sqlStr = sqlStr & " , sum(isnull(b.baljuno,0)) as baljuno" & vbcrlf
		'sqlStr = sqlStr & " from db_aLogistics.dbo.tbl_logistics_baljuipgo b with (nolock)" & vbcrlf
		'sqlStr = sqlStr & " left join db_aLogistics.dbo.tbl_logistics_item i with (nolock)" & vbcrlf
		'sqlStr = sqlStr & "     on b.siteSeq=i.siteSeq" & vbcrlf
		'sqlStr = sqlStr & "     and b.itemid=i.siteitemid" & vbcrlf
		'sqlStr = sqlStr & "     and b.itemoption=i.siteitemoption" & vbcrlf
		'sqlStr = sqlStr & " where BaljuKey="& FRectBaljuKey &"" & vbcrlf
		'sqlStr = sqlStr & " group by i.brandid , b.itemid , isnull(b.itemoption,'0000')" & vbcrlf
		'sqlStr = sqlStr & " order by i.brandid asc, b.itemid asc, itemoption asc" & vbcrlf

		'response.write sqlStr & "<br>"
        rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget_Logistics.RecordCount
		ftotalcount = rsget_Logistics.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget_Logistics.EOF  then
			do until rsget_Logistics.eof
				set FItemList(i) = new CBaljuIpgoItem

				FItemList(i).fmakerid	     = rsget_Logistics("makerid")
				FItemList(i).fitemgubun	     = rsget_Logistics("itemgubun")
				FItemList(i).fitemid         = rsget_Logistics("itemid")
				FItemList(i).fitemoption     = rsget_Logistics("itemoption")
				FItemList(i).fitemno	     = rsget_Logistics("itemno")

                FItemList(i).FwarehouseCd    = rsget_Logistics("warehouseCd")

				i=i+1
				rsget_Logistics.moveNext
			loop
		end if
		rsget_Logistics.Close
	end Sub

	Private Sub Class_Terminate()

	End Sub

end class

%>
