<%
'###########################################################
' Description : 회원등급별 매출
' Hieditor : 2009.04.17 허진원 생성
'			 2016.07.20 한용민 수정
'###########################################################

class CDesignerJumunList
	public Fmaeipdiv
	public Fdefaultmargine
	public FM_margin
	public FW_margin
	public FU_margin
	public FMasterItemList()
	public Fselltotal
	public Fseldate
	public Fsellcnt
	public FResultCount
	public FCancelyn
	public FItemCount
	public FItemID
	public FItemName
	public Fsitename
	public Fmakerid
	public FsmallImage
	public Fpurchasetype
	public Fsex
	public Fselltotal2
	public Fsellcnt2
	public Fcash
	public Fonlinecnt
	public FSocname
	public Fdpart
	public Fdpartcount
	public Fitemgubun
	public FItemgubunNm
	public FItemNo
	public FItemCost
	public FItemOptionStr
	public FBuycash
	public Fipkumdiv
	public FItemSellprice
	public Faccountdiv
	public Fcode_nm
	public Fsubtotalprice
	public FDate
	public FDayselltotal
	public FDaysellcnt
	public Fminustotal
	public Fminusbuytotal
	public Fminuscount
	public FYYYYMMDDHHNNSS
	public FCLarge
	Public Flecturer
	Public Fcate_large
	Public FCateName
	Public FTcnt
	Public Fmiletotalprice
	Public Fspendmembership
	Public Ftencardspend
	Public Fallatdiscountprice
	public Fjumuntotalsum
	public Fbuytotal
	public Ftenbeasongcount
    public Fordercount

	public function GetChannelName()
		GetChannelName = "CH" + Fitemgubun
	end function

	public function GetChannelName_Kor()
		Select Case Fitemgubun
			Case "010"
				GetChannelName_Kor = "디자인문구"
			Case "020"
				GetChannelName_Kor = "오피스/개인소품"
			Case "030"
				GetChannelName_Kor = "키덜트/취미"
			Case "040"
				GetChannelName_Kor = "가구/패브릭"
			Case "050"
				GetChannelName_Kor = "조명/데코"
			Case "060"
				GetChannelName_Kor = "주방/욕실"
			Case "070"
				GetChannelName_Kor = "가방/슈즈/쥬얼리"
			Case "080"
				GetChannelName_Kor = "Women"
			Case "090"
				GetChannelName_Kor = "Men"
			Case "100"
				GetChannelName_Kor = "베이비"
			Case "110"
				GetChannelName_Kor = "감성채널"
			Case Else
				GetChannelName_Kor = "기타"
		End Select
	end Function

	public function GetCAName()
		Select Case Fcate_large
			Case "010"
				GetCAName = "디자인문구"
			Case "020"
				GetCAName = "오피스/개인소품"
			Case "030"
				GetCAName = "키덜트/취미"
			Case "040"
				GetCAName = "가구/패브릭"
			Case "050"
				GetCAName = "조명/데코"
			Case "060"
				GetCAName = "주방/욕실"
			Case "070"
				GetCAName = "가방/슈즈/쥬얼리"
			Case "080"
				GetCAName = "Women"
			Case "090"
				GetCAName = "Men"
			Case "100"
				GetCAName = "베이비"
			Case "110"
				GetCAName = "감성채널"
			Case Else
				GetCAName = "기타"
		End Select
	end function

	public function IsAvailJumun()
		IsAvailJumun = Not ((CStr(Fipkumdiv)="0") or (CStr(Fipkumdiv)="1") or (CStr(FCancelyn)="D") or (CStr(FCancelyn)="Y"))
	end function

	public function GetDpartName()
		if Fdpart=1 then
			GetDpartName = "<font color=#FF0000>일</font>"
		elseif Fdpart=2 then
			GetDpartName = "월"
		elseif Fdpart=3 then
			GetDpartName = "화"
		elseif Fdpart=4 then
			GetDpartName = "수"
		elseif Fdpart=5 then
			GetDpartName = "목"
		elseif Fdpart=6 then
			GetDpartName = "금"
		elseif Fdpart=7 then
			GetDpartName = "<font color=#0000FF>토</font>"
		else
			GetDpartName = ""
		end if
	end function

	Public function JumunMethodName()
		if Cstr(Faccountdiv) = 7 then
			JumunMethodName = "무통장"
		elseif Cstr(Faccountdiv) = 14 then
			JumunMethodName = "편의점결제"
		elseif Cstr(Faccountdiv) = 100 then
			JumunMethodName = "신용"
		elseif Cstr(Faccountdiv) = 110 then
			JumunMethodName = "OK+신용"
		elseif Cstr(Faccountdiv) = 30 then
			JumunMethodName = "포인트"
		elseif Cstr(Faccountdiv) = 50 then
			JumunMethodName = "입점몰"
		elseif Cstr(Faccountdiv) = 80 then
			JumunMethodName = "All@"
		elseif Cstr(Faccountdiv) = 90 then
			JumunMethodName = "상품권"
		elseif Cstr(Faccountdiv) = 400 then
			JumunMethodName = "휴대폰"
		elseif Cstr(Faccountdiv) = 20 then
			JumunMethodName = "실시간"
		elseif Cstr(Faccountdiv) = 550 then
			JumunMethodName = "기프팅"
		elseif Cstr(Faccountdiv) = 560 then
			JumunMethodName = "기프티콘"		
		end if
	end function

	Public function Itemgubun()
		if Faccountdiv="7" then
			Itemgubun="01"
		elseif Faccountdiv="100" then
			Itemgubun="02"
		elseif Faccountdiv="30" then
			Itemgubun="03"
		elseif Faccountdiv="50" then
			Itemgubun="04"
		elseif Faccountdiv="80" then
			Itemgubun="05"
		elseif Faccountdiv="90" then
			Itemgubun="06"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

class CJumunMasterItem
	public FMasterItemList()
    public Fselltotal
    public Fseldate
    public Fsellcnt
	public maxt
	public maxc
	public FResultCount
    public FItemCount
	public FItemID
	public FItemName

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CJumunMaster
	public FMasterItemList()
	public maxt
	public maxc
	public maxa
	public maxb
	public maxt2
	public maxc2
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FRectRegStart
	public FRectRegEnd
    public FRectItemid
	public FCurrPage
	public Fsitename
	public FRectFromDate
	public FRectToDate
	public FRectIpkumDiv4
    public FRectDesignerID
    public FItemCount
	public FItemID
	public FItemName
	public FItemimgsmall
	public FTotalFavoriteCount
	public FSubtotal
    public Fsellcnt
	public Ftotalmoney
	public FTotalsellcnt
	public FPurchasetype
	public FRectPurchasetype
	public FRectMwDiv
    public FMtotalmoney
	public FMtotalsellcnt
    public FNtotalmoney
	public FNtotalsellcnt
    public FBtotalmoney
	public FBtotalsellcnt

	public FRectJoinMallNotInclude
	public FRectExtMallNotInclude
	public FRectPointNotInclude
	public FRectMinusNotInclude
	public FRectReturnNotInclude
	public FRectSearchType
	public FManTotalMoney
	public FManTotalCount
	public FWoManTotalMoney
	public FWoManTotalCount
	public FRectToDateTime
	public FRectckpointsearch
	public FRectOrderSerial
	public FRectDispY
	public FRectSellY
	public FRectMalltype
	public FRectOrdertype
	public FTotalPrice
	public FTotalBuyPrice
	public FTotalEA
	public FRectCD1
	public FRectCD2
	public FRectCD3
	public FRectYYYY
	public FRectMM
	public FRectItemGubun
	public FRectSearchDate
	public FRectOldJumun
	public FRectDelNoSearch
	public FRectDateType
	Public FRectSort
	public FRectOrderBy
	Public FRectCnt
	Public FRectToDateGubun
    public FRectTenDeliverExists
	public FRectRdsite
	public FRectChannelDiv
	public FRectDispCate

	Private Sub Class_Initialize()
		redim FMasterItemList(0)

		FCurrPage = 1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	function MaxVal(a,b)
		if (CDbl(a)> CDbl(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function

	public Sub getLectureMeaChul
		dim i,sqlStr
		maxt = -1
    	maxc = -1

		sqlStr = "select l.mastercode, sum(d.itemcost*d.itemno) as sumtotal,"
		sqlStr = sqlStr + " sum(d.itemno) as sellcnt, v.cnt as lcount"
		if FRectOldJumun="on" then
			sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m,"
			sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d,"
		else
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
		end if
		sqlStr = sqlStr + " [db_contents].[dbo].tbl_lecture_item l"
		sqlStr = sqlStr + " left join (select mastercode, count(idx) as cnt"
		sqlStr = sqlStr + " from [db_contents].[dbo].tbl_lecture_item"
		If FRectDesignerID <> "" Then
			sqlStr = sqlStr + " where lecturerid='" + FRectDesignerID + "'"
		End If
		sqlStr = sqlStr + " group by mastercode) as v on l.mastercode=v.mastercode"

		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.regdate>'2004-05-01'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.ipkumdiv>3"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.itemid=l.linkitemid"
		If FRectDesignerID <> "" Then
			sqlStr = sqlStr + " and l.lecturerid='" + FRectDesignerID + "'"
		End If
		sqlStr = sqlStr + " group by l.mastercode, v.cnt"
		sqlStr = sqlStr + " order by l.mastercode desc"
'response.write sqlStr
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
        redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList

			    FMasterItemList(i).Fsitename = rsget("mastercode")
			    FMasterItemList(i).Fsocname = rsget("lcount")
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end Sub

	public Sub GetLecturerMonthMeaChul
		dim i,sqlStr
		maxt = -1
    	maxc = -1

		sqlStr = "select l.mastercode, sum(d.itemcost*d.itemno) as sumtotal," + vbcrlf
		sqlStr = sqlStr + " sum(d.itemno) as sellcnt, l.lecturer" + vbcrlf
		if FRectOldJumun="on" then
			sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m,"
			sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d,"
		else
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
		end if
		sqlStr = sqlStr + " [db_contents].[dbo].tbl_lecture_item l" + vbcrlf
		sqlStr = sqlStr + " where m.orderserial=d.orderserial" + vbcrlf
		sqlStr = sqlStr + " and d.itemid=l.linkitemid" + vbcrlf
		sqlStr = sqlStr + " and l.mastercode='" + Cstr(FRectFromDate) + "'" + vbcrlf
		sqlStr = sqlStr + " and m.cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + " and m.ipkumdiv>3" + vbcrlf
		sqlStr = sqlStr + " and d.cancelyn<>'Y'" + vbcrlf
		sqlStr = sqlStr + " and d.itemid<>0" + vbcrlf
		sqlStr = sqlStr + " group by l.mastercode, l.lecturer" + vbcrlf
		If FRectSort = "name" Then
		sqlStr = sqlStr + " order by l.lecturer asc"
		ElseIf FRectSort = "tcnt" Then
		sqlStr = sqlStr + " order by sellcnt desc"
		Else
		sqlStr = sqlStr + " order by sumtotal desc"
		End If

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
        redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList

			   FMasterItemList(i).Fsitename = rsget("mastercode")
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")
				FMasterItemList(i).Flecturer = rsget("lecturer")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end Sub

	public Sub GetLectureMonthUserReport
		dim i,sqlStr
		maxt = -1
    	maxc = -1

		sqlStr = "select T.tcnt, count(T.tcnt) as gcnt" + vbcrlf
		sqlStr = sqlStr + " from (select userid,count(userid) as tcnt" + vbcrlf
		if FRectOldJumun="on" then
			sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m,"
			sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d,"
		else
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
		end if
		sqlStr = sqlStr + " [db_contents].[dbo].tbl_lecture_item l" + vbcrlf
		sqlStr = sqlStr + " where m.orderserial=d.orderserial" + vbcrlf
		sqlStr = sqlStr + " and convert(varchar(7),m.regdate,20)<='" + Cstr(FRectFromDate) + "'" + vbcrlf
		sqlStr = sqlStr + " and m.cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + " and m.ipkumdiv>3" + vbcrlf
		sqlStr = sqlStr + " and d.cancelyn<>'Y'" + vbcrlf
		sqlStr = sqlStr + " and d.itemid<>0" + vbcrlf
		sqlStr = sqlStr + " and d.itemid=l.linkitemid" + vbcrlf
		sqlStr = sqlStr + " and m.userid <> ''" + vbcrlf
		sqlStr = sqlStr + " group by userid" + vbcrlf
		sqlStr = sqlStr + " ) as T" + vbcrlf
		sqlStr = sqlStr + " group by T.tcnt" + vbcrlf
		sqlStr = sqlStr + " order by T.tcnt asc"

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
        redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList

			   FMasterItemList(i).Fsitename = rsget("tcnt")
				FMasterItemList(i).Fsellcnt = rsget("gcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end Sub

	public Sub GetLectureCountUserID
		dim i,sqlStr
		maxt = -1
    	maxc = -1

		sqlStr = sqlStr + "select top " + Cstr(FRectCnt) + " userid,count(userid) as tcnt" + vbcrlf
		if FRectOldJumun="on" then
			sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m,"
			sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d,"
		else
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
		end if
		sqlStr = sqlStr + " [db_contents].[dbo].tbl_lecture_item l" + vbcrlf
		sqlStr = sqlStr + " where m.orderserial=d.orderserial" + vbcrlf
		sqlStr = sqlStr + " and convert(varchar(7),m.regdate,20)<='" + Cstr(FRectFromDate) + "'" + vbcrlf
		sqlStr = sqlStr + " and m.cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + " and m.ipkumdiv>3" + vbcrlf
		sqlStr = sqlStr + " and d.cancelyn<>'Y'" + vbcrlf
		sqlStr = sqlStr + " and d.itemid<>0" + vbcrlf
		sqlStr = sqlStr + " and d.itemid=l.linkitemid" + vbcrlf
		sqlStr = sqlStr + " and m.userid <> ''" + vbcrlf
		sqlStr = sqlStr + " group by userid" + vbcrlf
		sqlStr = sqlStr + " order by count(userid) desc"
'response.write sqlStr
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
        redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList

			   FMasterItemList(i).Fsitename = rsget("userid")
				FMasterItemList(i).Fsellcnt = rsget("tcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end Sub

	public Sub GetMonthlyLectureStuffReport
		dim i,sqlStr
		maxt = -1
    	maxc = -1

		sqlStr = "select convert(varchar(7),m.regdate,20) as yyyymm, sum(d.itemcost*d.itemno) as sumtotal," + vbcrlf
		sqlStr = sqlStr + " sum(d.itemno) as sellcnt" + vbcrlf
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m," + vbcrlf
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where m.orderserial=d.orderserial" + vbcrlf
		sqlStr = sqlStr + " and m.cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + " and m.ipkumdiv>3" + vbcrlf
		sqlStr = sqlStr + " and d.cancelyn<>'Y'" + vbcrlf
		sqlStr = sqlStr + " and d.itemid<>0" + vbcrlf
		sqlStr = sqlStr + " and d.itemid=i.itemid"  + vbcrlf
		sqlStr = sqlStr + " and i.itemdiv='88'"  + vbcrlf
		sqlStr = sqlStr + " group by convert(varchar(7),m.regdate,20)" + vbcrlf
		sqlStr = sqlStr + " order by convert(varchar(7),m.regdate,20) desc"
'response.write sqlStr
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
        redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = rsget("yyyymm")
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end Sub

	public Sub GetMonthlySeminarRoomReport
		dim i,sqlStr
		maxt = -1
    	maxc = -1

		sqlStr = "select convert(varchar(7),regdate,20) as yyyymm," + vbcrlf
		sqlStr = sqlStr + " sum(realsum) as sumtotal, count(orderno) as sellcnt" + vbcrlf
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master" + vbcrlf
		sqlStr = sqlStr + " where shopid='cafe003'" + vbcrlf
		sqlStr = sqlStr + " and regdate>'2004-05-01'" + vbcrlf
		sqlStr = sqlStr + " and cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + " group by convert(varchar(7),regdate,20)" + vbcrlf
		sqlStr = sqlStr + " order by convert(varchar(7),regdate,20) desc"

'response.write sqlStr
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
        redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = rsget("yyyymm")
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end Sub

	public Sub GetDailySeminarRoomReport
		dim i,sqlStr
		maxt = -1
    	maxc = -1

		sqlStr = "select convert(varchar(10),regdate,20) as yyyymm," + vbcrlf
		sqlStr = sqlStr + " sum(realsum) as sumtotal, count(orderno) as sellcnt" + vbcrlf
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master" + vbcrlf
		sqlStr = sqlStr + " where shopid='cafe003'" + vbcrlf
		sqlStr = sqlStr + " and convert(varchar(7),regdate,20)='" + FRectYYYY + "'" + vbcrlf
		sqlStr = sqlStr + " and cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + " group by convert(varchar(10),regdate,20)" + vbcrlf
		sqlStr = sqlStr + " order by convert(varchar(10),regdate,20) desc"

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
        redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = rsget("yyyymm")
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end Sub

	public sub SearchSellrePortDesum()
    	Dim sql, sqltmp, i

    	maxt = -1
    	maxc = -1

		sql = "select d.makerid, sum(d.itemcost*d.itemno) as sumtotal, sum(d.buycash*d.itemno) as buytotal,"
		sql = sql + " sum(d.itemno) as sellcnt from [db_order].[dbo].tbl_order_master m,"
		sql = sql + " [db_order].[dbo].tbl_order_detail d"
		sql = sql + " where m.orderserial = d.orderserial"
		sql = sql + " and (m.regdate >= '" & FRectFromDate & "') and (m.regdate < '" & FRectToDate & "')"
		sql = sql + " and d.itemid <> 0"
		sql = sql + " and m.jumundiv<>9"
		sql = sql + " and m.cancelyn = 'N'"
		sql = sql + " and d.cancelyn <> 'Y'"
		sql = sql + " and m.ipkumdiv>=4"
        sql = sql + " Group by d.makerid"
        sql = sql + " order by sumtotal "

		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount
	    redim preserve FMasterItemList(FResultCount)
		do until rsget.eof

			set FMasterItemList(i) = new CDesignerJumunList
			FMasterItemList(i).Fmakerid     = rsget("makerid")
			FMasterItemList(i).Fselltotal   = rsget("sumtotal")
			FMasterItemList(i).Fbuytotal    = rsget("buytotal")
			FMasterItemList(i).Fsellcnt     = rsget("sellcnt")

			if Not IsNull(FMasterItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
			end if

			rsget.MoveNext
			i = i + 1
		loop

		rsget.close
	end sub

	public sub SearchSellrePort()
        Dim sql, sqltmp, i

    	maxt = -1
    	maxc = -1

        if (FRectDesignerID<>"") then
		    sql = "select convert(varchar(10),m.regdate,21) as yyyymmdd, "
		else
		    sql = "select d.makerid, "
	    end if

		sql = sql + " sum(d.itemcost*d.itemno) as sumtotal, sum(d.buycash*d.itemno) as buytotal,"
		sql = sql + " sum(d.itemno) as sellcnt, count(distinct d.itemid) as itemcount, count(distinct m.orderserial) as ordercount "
		
		if FRectOldJumun="on" then
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m"
			sql = sql + " 	Join [db_log].[dbo].tbl_old_order_detail_2003 d"
			sql = sql + " 	on m.orderserial=d.orderserial"
		else
			sql = sql + " from [db_order].[dbo].tbl_order_master m "
			sql = sql + " 	Join [db_order].[dbo].tbl_order_detail d"
			sql = sql + " 	on m.orderserial=d.orderserial"
		end if
		sql = sql + " inner join [db_partner].[dbo].[tbl_partner] as p on d.makerid = p.id "
		
		sql = sql + " where (m.regdate >= '" & FRectFromDate & "') and (m.regdate < '" & FRectToDate & "')"
		sql = sql + " and d.itemid <> 0"
		sql = sql + " and m.jumundiv<>9"
		sql = sql + " and m.cancelyn = 'N'"
		sql = sql + " and d.cancelyn <> 'Y'"
		sql = sql + " and m.ipkumdiv>=4"
		
		If FPurchasetype <> "" Then
			sql = sql + " and p.purchasetype = '" & FPurchasetype &"'"
		End IF
		
		if (FRectChannelDiv<>"") then
			if FRectChannelDiv="web" then
				sql = sql & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") and m.accountdiv<>'50' "
			elseif FRectChannelDiv="jaehu" then
				sql = sql & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") "
			elseif FRectChannelDiv="mjaehu" then
				sql = sql & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") "
			elseif FRectChannelDiv="mobile" then
				sql = sql & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") and m.accountdiv<>'50' "
			elseif FRectChannelDiv="ipjum" then
				sql = sql & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") "
			end if
		end if
		
		if (FRectDesignerID<>"") then
		    sql = sql + " and d.makerid='" & FRectDesignerID & "'"
		    sql = sql + " Group by convert(varchar(10),m.regdate,21)"
		    sql = sql + " order by yyyymmdd "
		else
            sql = sql + " Group by d.makerid"

            if FRectOrdertype = "totalprice" then
    			sql = sql + " order by sumtotal Desc"
    		elseif FRectOrdertype = "totalgain" then
    		    sql = sql + " order by sum(d.itemcost*d.itemno-d.buycash*d.itemno) Desc"
    		else
    			sql = sql + " order by sellcnt Desc"
    		end if
        end if

		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount
	    redim preserve FMasterItemList(FResultCount)
		do until rsget.eof
			set FMasterItemList(i) = new CDesignerJumunList
			if (FRectDesignerID<>"") then
			    FMasterItemList(i).FDate     = rsget("yyyymmdd")
			else
			    FMasterItemList(i).Fmakerid     = rsget("makerid")
			end if
			FMasterItemList(i).Fselltotal   = rsget("sumtotal")
			FMasterItemList(i).Fbuytotal    = rsget("buytotal")

			FMasterItemList(i).Fitemcount   = rsget("itemcount")
			FMasterItemList(i).Fsellcnt     = rsget("sellcnt")
            FMasterItemList(i).Fordercount  = rsget("ordercount")

			if Not IsNull(FMasterItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
			end if

			rsget.MoveNext
			i = i + 1
		loop

		rsget.close
	end sub

	public sub SearchMallSellrePort3()
		Dim sql
		sql = "select sum(subtotalprice) as sumtotal,"
		sql = sql + " count(orderserial) as sellcnt"
		sql = sql + " from [db_order].[dbo].tbl_order_master"
		sql = sql + " where (regdate >= '" & FRectFromDate & "') and (regdate<'" & FRectToDate & "')"
		sql = sql + " and cancelyn = 'N'"
		sql = sql + " and ipkumdiv>=4"
		sql = sql + " and jumundiv<>9"
		if FRectSearchType=1 then
			sql = sql + " and sitename='10x10'"
			sql = sql + " and ((rdsite is Null) or (rdsite=''))"
		elseif FRectSearchType=2 then
			sql = sql + " and ((sitename='10x10' and rdsite<>'') or (sitename<>'10x10'))"
			sql = sql + " and accountdiv<>'30'"
			sql = sql + " and accountdiv<>'50'"
		elseif FRectSearchType=3 then
			sql = sql + " and accountdiv='50'"
		elseif FRectSearchType=4 then
			sql = sql + " and accountdiv='30'"
		end if

		rsget.Open sql,dbget,1
		FMtotalmoney = rsget("sumtotal")
		FMtotalsellcnt = rsget("sellcnt")
		if IsNull(FMtotalmoney) then
			FMtotalmoney =0
		end if

		if IsNull(FMtotalsellcnt) then
			FMtotalsellcnt =0
		end if
		rsget.close

	end sub

	public sub SearchMallSellrePort2()
   		Dim sql, i ,orderMasterDB
		dim wheredetail

		if FRectExtMallNotInclude<>"" then
			wheredetail = " and jumundiv<>'5'"
		end if

		if FRectPointNotInclude<>"" then
			wheredetail = wheredetail + " and accountdiv<>'30'"
		end if

		if FRectMinusNotInclude<>"" then
			wheredetail = wheredetail + " and jumundiv<>'9'"
		end if

		IF FRectOldJumun<>"" Then
			orderMasterDB =" [db_log].[dbo].tbl_old_order_master_2003 "
		Else
			orderMasterDB =" [db_order].[dbo].tbl_order_master "
		End IF

		sql = "select sum(subtotalprice) as sumtotal,"
		sql = sql + " count(orderserial) as sellcnt"
		sql = sql + " from "& orderMasterDB
		sql = sql + " where (regdate >= '" & FRectFromDate & "') and (regdate<'" & FRectToDate & "')"
		sql = sql + " and cancelyn = 'N'"
		sql = sql + " and ipkumdiv>=4"
		sql = sql + wheredetail

		rsget.Open sql,dbget,1
		FMtotalmoney = rsget("sumtotal")
		FMtotalsellcnt = rsget("sellcnt")
		rsget.close

   		maxt = -1
   		maxc = -1

		sql = "select sum(subtotalprice) as sumtotal,"
		sql = sql + " count(orderserial) as sellcnt, (sitename + IsNull(rdsite,'')) as sitename"
		sql = sql + " from "& orderMasterDB
		sql = sql + " where (regdate >= '" & FRectFromDate & "') and (regdate<'" & FRectToDate & "')"
		sql = sql + " and cancelyn = 'N'"
		sql = sql + " and ipkumdiv>=4"
		sql = sql + wheredetail
        sql = sql + " Group by (sitename + IsNull(rdsite,''))"
        sql = sql + " order by sumtotal Desc"

		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount
	    redim preserve FMasterItemList(FResultCount)
		do until rsget.eof
			set FMasterItemList(i) = new CDesignerJumunList
		    FMasterItemList(i).Fsitename = rsget("sitename")
			FMasterItemList(i).Fselltotal = rsget("sumtotal")
			FMasterItemList(i).Fsellcnt = rsget("sellcnt")


			if Not IsNull(FMasterItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
			end if

			rsget.MoveNext
			i = i + 1
		loop

		rsget.close
	end sub

	public sub mwdivsellsum()
   		Dim sql, i

   		maxt = -1
   		maxc = -1

		sql = "select "
		sql = sql + " sum(case when d.itemno>0 then d.itemcost*d.itemno else 0 end) as sumtotal, "
		sql = sql + " sum(case when d.itemno>0 then d.buycash*d.itemno else 0 end) as buytotal, sum(d.itemno) as sellcnt,"
		sql = sql + " sum(case when d.itemno<0 then d.itemcost*d.itemno else 0 end) as minustotal, "
		sql = sql + " sum(case when d.itemno<0 then d.buycash*d.itemno else 0 end) as minusbuytotal, "
		sql = sql + " sum(case when d.itemno>0 then d.itemno else 0 end ) as sellcnt,"
		sql = sql + " sum(case when d.itemno<0 then d.itemno else 0 end ) as minuscount,"
		sql = sql + " i.mwdiv"
		if FRectOldJumun="on" then
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m,"
			sql = sql + " [db_log].[dbo].tbl_old_order_detail_2003 d,"
		else
			sql = sql + " from [db_order].[dbo].tbl_order_master m,"
			sql = sql + " [db_order].[dbo].tbl_order_detail d,"
		end if

		sql = sql + " [db_item].[dbo].tbl_item i"
		sql = sql + " where m.orderserial=d.orderserial"
		sql = sql + " and m.regdate>'" & FRectFromDate & "' and m.regdate<'" & FRectToDate & "'"
		sql = sql + " and m.ipkumdiv>3"
		sql = sql + " and d.itemid<>0"
		if FRectExtMallNotInclude<>"" then
			sql = sql + " and m.jumundiv<>'5'"
		end if
		sql = sql + " and m.cancelyn='N'"
		sql = sql + " and d.itemid=i.itemid"
		sql = sql + " and d.cancelyn<>'Y'"
		sql = sql + " group by i.mwdiv"

		rsget.Open sql,dbget,1
		rsget.close

		rsget.Open sql,dbget,1

		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		FMtotalmoney = 0
		FMtotalsellcnt = 0

		do until rsget.eof
			set FMasterItemList(i) = new CDesignerJumunList
			if rsget("mwdiv") = "W" then
			    FMasterItemList(i).Fsitename = "특정"
			elseif rsget("mwdiv") = "M" then
			    FMasterItemList(i).Fsitename = "매입"
			elseif rsget("mwdiv") = "U" then
			    FMasterItemList(i).Fsitename = "업체"
			elseif rsget("mwdiv") = "B" then
			    FMasterItemList(i).Fsitename = "배송료"
			else
			    FMasterItemList(i).Fsitename = rsget("mwdiv")
			end if
			FMasterItemList(i).Fselltotal = rsget("sumtotal")
			FMasterItemList(i).Fbuytotal  = rsget("buytotal")
			FMasterItemList(i).Fsellcnt = rsget("sellcnt")

            FMasterItemList(i).Fminustotal = rsget("minustotal")
            FMasterItemList(i).Fminusbuytotal = rsget("minusbuytotal")
            FMasterItemList(i).Fminuscount = rsget("minuscount")

			FMtotalmoney = Cdbl(FMtotalmoney) + Cdbl(rsget("sumtotal"))
			FMtotalsellcnt = Cdbl(FMtotalsellcnt) + Cdbl(rsget("sellcnt"))

			if Not IsNull(FMasterItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
			end if

			rsget.MoveNext
			i = i + 1
		loop

		rsget.close
	end sub

	public sub SearchMallSellrePort4()
		Dim sql, i
		maxt = -1
   		maxc = -1
        
        sql = " SET Transaction Isolation Level Read Uncommitted " & vbCrLf
        sql = sql + "  "
        
		sql = "select convert(varchar(7),m.regdate,20) as yyyymm, sum(m.subtotalprice) as sumtotal, count(m.idx) as sellcnt"
		if FRectOldJumun="on" then
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m"
		else
			sql = sql + " from [db_order].[dbo].tbl_order_master m"
		end if
		sql = sql + " where m.regdate>'" + FRectFromDate + "'"

		if FRectSearchType="curr" then
			sql = sql + " and day(m.regdate)<=day(getdate())"
		else
			'sql = sql + " and day(m.regdate)<=day(getdate())"
		end if

		sql = sql + " and ipkumdiv>3"
		sql = sql + " and cancelyn='N'"
		sql = sql + " and beadaldiv<>90"

		if FRectJoinMallNotInclude<>"on" then
			sql = sql + " and m.sitename in ('10x10','tingmart')"
		end if

		if FRectExtMallNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'50'"
		end if

		if FRectPointNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'30'"
		end if

		if FRectReturnNotInclude="on" then
			sql = sql + " and m.jumundiv not in ('6','9')"
		end if

		sql = sql + " group by  convert(varchar(7),m.regdate,20)"
		sql = sql + " order by  convert(varchar(7),m.regdate,20) desc"
'response.write sql
        rsget.CursorLocation = 3
		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
			set FMasterItemList(i) = new CDesignerJumunList
		    FMasterItemList(i).Fsitename = rsget("yyyymm")
			FMasterItemList(i).Fselltotal = rsget("sumtotal")
			FMasterItemList(i).Fsellcnt = rsget("sellcnt")

			if Not IsNull(FMasterItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
			end if

			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end sub

	public sub SearchMallSellrePort5()
		Dim sql, i
		maxt = -1
   		maxc = -1
        
        sql = " SET Transaction Isolation Level Read Uncommitted " & vbCrLf
        sql = sql + "  "
        
		sql = sql + " select convert(varchar(10),m.regdate,20) as yyyymmdd, datepart(w,m.regdate) as dpart," + vbcrlf
		sql = sql + " sum(case when jumundiv='9' then 0 else m.subtotalprice end ) as sumtotal," + vbcrlf
		sql = sql + " sum(case when jumundiv='9' then 0 else m.totalsum end ) as totalsum," + vbcrlf
		sql = sql + " sum(case when jumundiv='9' then 0  else 1 end ) as sellcnt," + vbcrlf
		sql = sql + " sum(case when jumundiv='9' then m.subtotalprice else 0 end ) as minustotal," + vbcrlf
		sql = sql + " sum(case when jumundiv='9' then 1  else 0 end ) as minuscount, " + vbcrlf
		sql = sql + " sum(case when jumundiv='9' then 0  else miletotalprice end ) as miletotalprice, " + vbcrlf
		sql = sql + " IsNull(sum(case when jumundiv='9' then 0  else spendmembership end ),0) as spendmembership, " + vbcrlf
		sql = sql + " sum(case when jumundiv='9' then 0  else tencardspend end ) as tencardspend, " + vbcrlf
		sql = sql + " IsNull(sum(case when jumundiv='9' then 0  else allatdiscountprice end ),0) as allatdiscountprice, " + vbcrlf
		sql = sql + " sum(d.buytotal) as buytotal , sum(case when d.tenitemcount>0 then 1 else 0 end) as tenbeasongcount"
		if FRectOldJumun="on" then
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m"
		    sql = sql + "   left join ("
			sql = sql + "       select m.orderserial, sum(d.buycash*d.itemno) as buytotal ,sum(case when d.isupchebeasong='N' then 1 else 0 end) as tenitemcount"
			sql = sql + "       from [db_log].[dbo].tbl_old_order_master_2003 m,"
			sql = sql + "       [db_log].[dbo].tbl_old_order_detail_2003 d"
			sql = sql + "       where m.orderserial=d.orderserial "
			sql = sql + "       and m.regdate>='" + CStr(FRectFromDate) + "'" + vbcrlf
		    sql = sql + "       and m.regdate<'" + CStr(FRectToDate) + "'" + vbcrlf
		    sql = sql + "       and m.ipkumdiv>3" + vbcrlf
		    sql = sql + "       and m.cancelyn='N'" + vbcrlf
		    sql = sql + "       and m.jumundiv<>'9'"
			sql = sql + "       and d.itemid<>0 "
			sql = sql + "       and d.cancelyn<>'Y'"
			sql = sql + "       group by m.orderserial"
			sql = sql + "       ) d on m.orderserial=d.orderserial"
		else
			sql = sql + " from [db_order].[dbo].tbl_order_master m" + vbcrlf
			sql = sql + "   left join ("
			sql = sql + "       select m.orderserial, sum(d.buycash*d.itemno) as buytotal ,sum(case when d.isupchebeasong='N' then 1 else 0 end) as tenitemcount"
			sql = sql + "       from [db_order].[dbo].tbl_order_master m,"
			sql = sql + "       [db_order].[dbo].tbl_order_detail d"
			sql = sql + "       where m.orderserial=d.orderserial "
			sql = sql + "       and m.regdate>='" + CStr(FRectFromDate) + "'" + vbcrlf
		    sql = sql + "       and m.regdate<'" + CStr(FRectToDate) + "'" + vbcrlf
		    sql = sql + "       and m.ipkumdiv>3" + vbcrlf
		    sql = sql + "       and m.cancelyn='N'" + vbcrlf
		    sql = sql + "       and m.jumundiv<>'9'"
			sql = sql + "       and d.itemid<>0 "
			sql = sql + "       and d.cancelyn<>'Y'"
			sql = sql + "       group by m.orderserial"
			sql = sql + "       ) d on m.orderserial=d.orderserial"
		end if
		sql = sql + " where m.regdate>='" + CStr(FRectFromDate) + "'" + vbcrlf
		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'" + vbcrlf
		sql = sql + " and m.ipkumdiv>3" + vbcrlf
		sql = sql + " and m.cancelyn='N'" + vbcrlf

		if FRectJoinMallNotInclude<>"on" then
			sql = sql + " and m.sitename='10x10'"
		end if

		if FRectExtMallNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'50'"
		end if

		if FRectPointNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'30'"
		end if

		if FRectRdsite="on" then
			sql = sql + " and m.rdsite in ('mobile','mobile_kakaotalk','mobile_nate','mobile_kakaotms') "
		end if

		sql = sql + " group by  convert(varchar(10),m.regdate,20), datepart(w,m.regdate)"
		sql = sql + " order by  convert(varchar(10),m.regdate,20) desc"
        
        rsget.CursorLocation = 3
		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
			set FMasterItemList(i) = new CDesignerJumunList
		    FMasterItemList(i).Fsitename = rsget("yyyymmdd")
			FMasterItemList(i).Fselltotal = rsget("sumtotal")
			FMasterItemList(i).Fbuytotal  = rsget("buytotal")
			FMasterItemList(i).Fsellcnt = rsget("sellcnt")
			FMasterItemList(i).Fdpart = rsget("dpart")

			FMasterItemList(i).Fminustotal = rsget("minustotal")
			FMasterItemList(i).Fminuscount = rsget("minuscount")

			FMasterItemList(i).Fmiletotalprice = rsget("miletotalprice")
			FMasterItemList(i).Fspendmembership = rsget("spendmembership")
			FMasterItemList(i).Ftencardspend = rsget("tencardspend")
			FMasterItemList(i).Fallatdiscountprice = rsget("allatdiscountprice")

            FMasterItemList(i).Ftenbeasongcount = rsget("tenbeasongcount")

			if IsNULL(FMasterItemList(i).Fallatdiscountprice) then FMasterItemList(i).Fallatdiscountprice=0
			if IsNULL(FMasterItemList(i).Ftencardspend) then FMasterItemList(i).Ftencardspend=0

			FMasterItemList(i).Fjumuntotalsum = rsget("totalsum")

			if Not IsNull(FMasterItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
			end if
            
            if IsNULL(FMasterItemList(i).Fmiletotalprice) then FMasterItemList(i).Fmiletotalprice=0
            if IsNULL(FMasterItemList(i).Fbuytotal) then FMasterItemList(i).Fbuytotal=0
			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end sub

	public sub SearchMallSellrePort6()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select "
		sql = sql + "	case "
		sql = sql + "	when (m.subtotalprice>=0) and (m.subtotalprice<10000) then '0~10000'"
		sql = sql + "	when (m.subtotalprice>=10000) and (m.subtotalprice<20000) then '10000~20000'"
		sql = sql + "	when (m.subtotalprice>=20000) and (m.subtotalprice<30000) then '20000~30000'"
		sql = sql + "	when (m.subtotalprice>=30000) and (m.subtotalprice<40000) then '30000~40000'"
		sql = sql + "	when (m.subtotalprice>=40000) and (m.subtotalprice<50000) then '40000~50000'"
		sql = sql + "	when (m.subtotalprice>=50000) and (m.subtotalprice<60000) then '50000~60000'"
		sql = sql + "	when (m.subtotalprice>=60000) and (m.subtotalprice<70000) then '60000~70000'"
		sql = sql + "	when (m.subtotalprice>=70000) and (m.subtotalprice<80000) then '70000~80000'"
		sql = sql + "	when (m.subtotalprice>=80000) and (m.subtotalprice<90000) then '80000~90000'"
		sql = sql + "	when (m.subtotalprice>=90000) and (m.subtotalprice<100000) then '90000~100000'"
		sql = sql + "	when (m.subtotalprice>=100000) and (m.subtotalprice<110000) then 'A100000~110000'"
		sql = sql + "	when (m.subtotalprice>=110000) and (m.subtotalprice<120000) then 'A110000~120000'"
		sql = sql + "	when (m.subtotalprice>=120000) and (m.subtotalprice<130000) then 'A120000~130000'"
		sql = sql + "	when (m.subtotalprice>=130000) and (m.subtotalprice<140000) then 'A130000~140000'"
		sql = sql + "	when (m.subtotalprice>=140000) and (m.subtotalprice<150000) then 'A140000~150000'"
		sql = sql + "	when (m.subtotalprice>=150000) and (m.subtotalprice<160000) then 'A150000~160000'"
		sql = sql + "	when (m.subtotalprice>=160000) and (m.subtotalprice<170000) then 'A160000~170000'"
		sql = sql + "	when (m.subtotalprice>=170000) and (m.subtotalprice<180000) then 'A170000~180000'"
		sql = sql + "	when (m.subtotalprice>=180000) and (m.subtotalprice<190000) then 'A180000~190000'"
		sql = sql + "	when (m.subtotalprice>=190000) and (m.subtotalprice<200000) then 'A190000~200000'"
		sql = sql + "	when (m.subtotalprice>=200000) and (m.subtotalprice<250000) then 'A200000~250000'"
		sql = sql + "	when (m.subtotalprice>=250000) and (m.subtotalprice<300000) then 'A250000~300000'"
		sql = sql + "	when (m.subtotalprice>=300000) and (m.subtotalprice<350000) then 'A300000~350000'"
		sql = sql + "	when (m.subtotalprice>=350000) and (m.subtotalprice<400000) then 'A350000~400000'"
		sql = sql + "	when (m.subtotalprice>=400000) and (m.subtotalprice<450000) then 'A400000~450000'"
		sql = sql + "	when (m.subtotalprice>=450000) and (m.subtotalprice<500000) then 'A450000~500000'"
		sql = sql + "	when (m.subtotalprice>=500000) and (m.subtotalprice<600000) then 'A500000~600000'"
		sql = sql + "	when (m.subtotalprice>=600000) and (m.subtotalprice<700000) then 'A600000~700000'"
		sql = sql + "	when (m.subtotalprice>=700000) and (m.subtotalprice<800000) then 'A700000~800000'"
		sql = sql + "	when (m.subtotalprice>=800000) and (m.subtotalprice<900000) then 'A800000~900000'"
		sql = sql + "	when (m.subtotalprice>=900000) and (m.subtotalprice<1000000) then 'A900000~1000000'"
		sql = sql + "	when m.subtotalprice>=1000000 then 'B1000000~'"
		sql = sql + " end as pricegubun "
		sql = sql + ",count(m.idx) as cnt, sum(m.subtotalprice) as sumtotal"

		if FRectOldJumun="on" then
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m"
		else
			sql = sql + " from [db_order].[dbo].tbl_order_master m" + vbcrlf
		end if

        if (FRectTenDeliverExists<>"") then
            sql = sql + " left join (" + vbcrlf
            sql = sql + "  select distinct m.orderserial" + vbcrlf
            if FRectOldJumun="on" then
                sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m," + vbcrlf
                sql = sql + " [db_log].[dbo].tbl_old_order_detail_2003 d" + vbcrlf
            else
                sql = sql + " from [db_order].[dbo].tbl_order_master m," + vbcrlf
                sql = sql + " [db_order].[dbo].tbl_order_detail d" + vbcrlf
            end if
            sql = sql + " where m.regdate>='" + CStr(FRectFromDate) + "'" + vbcrlf
            sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'" + vbcrlf
            sql = sql + " and m.orderserial=d.orderserial" + vbcrlf
            sql = sql + " and m.ipkumdiv>3" + vbcrlf
            sql = sql + " and m.cancelyn='N'" + vbcrlf
            sql = sql + " and d.cancelyn<>'Y'" + vbcrlf
            sql = sql + " and d.itemid<>0" + vbcrlf
            sql = sql + " and d.isupchebeasong='N') T on m.orderserial=T.orderserial" + vbcrlf
        end if

		sql = sql + " where m.regdate>='" + CStr(FRectFromDate) + "'"
		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'"
		sql = sql + " and m.cancelyn='N'"
		sql = sql + " and m.ipkumdiv>3"
		if FRectJoinMallNotInclude<>"on" then
			sql = sql + " and m.sitename in ('10x10','tingmart')"
		end if

		if FRectExtMallNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'50'"
		end if

		if FRectPointNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'30'"
		end if

		if (FRectTenDeliverExists<>"") then
		    sql = sql + " and T.orderserial is not null"
		end if

		sql = sql + " group by case "
		sql = sql + "	when (m.subtotalprice>=0) and (m.subtotalprice<10000) then '0~10000'"
		sql = sql + "	when (m.subtotalprice>=10000) and (m.subtotalprice<20000) then '10000~20000'"
		sql = sql + "	when (m.subtotalprice>=20000) and (m.subtotalprice<30000) then '20000~30000'"
		sql = sql + "	when (m.subtotalprice>=30000) and (m.subtotalprice<40000) then '30000~40000'"
		sql = sql + "	when (m.subtotalprice>=40000) and (m.subtotalprice<50000) then '40000~50000'"
		sql = sql + "	when (m.subtotalprice>=50000) and (m.subtotalprice<60000) then '50000~60000'"
		sql = sql + "	when (m.subtotalprice>=60000) and (m.subtotalprice<70000) then '60000~70000'"
		sql = sql + "	when (m.subtotalprice>=70000) and (m.subtotalprice<80000) then '70000~80000'"
		sql = sql + "	when (m.subtotalprice>=80000) and (m.subtotalprice<90000) then '80000~90000'"
		sql = sql + "	when (m.subtotalprice>=90000) and (m.subtotalprice<100000) then '90000~100000'"
		sql = sql + "	when (m.subtotalprice>=100000) and (m.subtotalprice<110000) then 'A100000~110000'"
		sql = sql + "	when (m.subtotalprice>=110000) and (m.subtotalprice<120000) then 'A110000~120000'"
		sql = sql + "	when (m.subtotalprice>=120000) and (m.subtotalprice<130000) then 'A120000~130000'"
		sql = sql + "	when (m.subtotalprice>=130000) and (m.subtotalprice<140000) then 'A130000~140000'"
		sql = sql + "	when (m.subtotalprice>=140000) and (m.subtotalprice<150000) then 'A140000~150000'"
		sql = sql + "	when (m.subtotalprice>=150000) and (m.subtotalprice<160000) then 'A150000~160000'"
		sql = sql + "	when (m.subtotalprice>=160000) and (m.subtotalprice<170000) then 'A160000~170000'"
		sql = sql + "	when (m.subtotalprice>=170000) and (m.subtotalprice<180000) then 'A170000~180000'"
		sql = sql + "	when (m.subtotalprice>=180000) and (m.subtotalprice<190000) then 'A180000~190000'"
		sql = sql + "	when (m.subtotalprice>=190000) and (m.subtotalprice<200000) then 'A190000~200000'"
		sql = sql + "	when (m.subtotalprice>=200000) and (m.subtotalprice<250000) then 'A200000~250000'"
		sql = sql + "	when (m.subtotalprice>=250000) and (m.subtotalprice<300000) then 'A250000~300000'"
		sql = sql + "	when (m.subtotalprice>=300000) and (m.subtotalprice<350000) then 'A300000~350000'"
		sql = sql + "	when (m.subtotalprice>=350000) and (m.subtotalprice<400000) then 'A350000~400000'"
		sql = sql + "	when (m.subtotalprice>=400000) and (m.subtotalprice<450000) then 'A400000~450000'"
		sql = sql + "	when (m.subtotalprice>=450000) and (m.subtotalprice<500000) then 'A450000~500000'"
		sql = sql + "	when (m.subtotalprice>=500000) and (m.subtotalprice<600000) then 'A500000~600000'"
		sql = sql + "	when (m.subtotalprice>=600000) and (m.subtotalprice<700000) then 'A600000~700000'"
		sql = sql + "	when (m.subtotalprice>=700000) and (m.subtotalprice<800000) then 'A700000~800000'"
		sql = sql + "	when (m.subtotalprice>=800000) and (m.subtotalprice<900000) then 'A800000~900000'"
		sql = sql + "	when (m.subtotalprice>=900000) and (m.subtotalprice<1000000) then 'A900000~1000000'"
		sql = sql + "	when m.subtotalprice>=1000000 then 'B1000000~'"
		sql = sql + " end"
		sql = sql + " order by pricegubun"

		'response.Write sql
		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = rsget("pricegubun")
				FMasterItemList(i).Fsellcnt = rsget("cnt")
				FMasterItemList(i).Fselltotal = rsget("sumtotal")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if
				FTotalsellcnt = FTotalsellcnt + FMasterItemList(i).Fsellcnt
				Ftotalmoney = Ftotalmoney + FMasterItemList(i).Fselltotal
				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end sub

	public sub SearchMallSellrePortChannel()
		Dim sql, i,selcate, tblcate
		maxt = -1
   		maxc = -1
   		''[db_item].[dbo].tbl_item -> [db_item].[dbo].tbl_item_Category 로 수정 .. 특정쿼리 느림..?
   		if FRectCD1<>"" and FRectCD2<>"" Then
   			selcate = "code_small"
   			tblcate = "cate_small"
   		ElseIf FRectCD1<>"" then
   			selcate ="code_mid"
   			tblcate = "cate_mid"
   		else
   			selcate ="code_large"
   			tblcate = "cate_large"
   		end if
		sql = "select isnull(i."& selcate &",999) as itemgubun,isnull(c.code_nm,'기타')  as ItemGubunNm "
		If FRectToDateGubun="M" Then
		sql = sql + " ,convert(varchar(7),m.regdate,20) as yyyymmdd"
		Else
		sql = sql + " ,convert(varchar(10),m.regdate,20) as yyyymmdd"
		End If
		'sql = sql + " convert(varchar(7),m.regdate,20) as yyyymm,"
		sql = sql + " ,sum(d.itemno*d.itemcost) as sumtotal, sum(d.itemno*d.buycash) as buytotal"
		sql = sql + " ,count(d.itemno) as sellcnt"
		if FRectOldJumun="on" then
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m "
			sql = sql + " join [db_log].[dbo].tbl_old_order_detail_2003 d "
		else
			sql = sql + " from [db_order].[dbo].tbl_order_master m "
			sql = sql + " join [db_order].[dbo].tbl_order_detail d "
		end if
		sql = sql + " on m.orderserial=d.orderserial"
		sql = sql + " and m.ipkumdiv>3"
		sql = sql + " and d.itemid<>0"
		sql = sql + " and m.cancelyn='N'"
		sql = sql + " and d.cancelyn<>'Y'"
		sql = sql + " and m.jumundiv<>9"

		sql = sql + " LEFT JOIN [db_item].[dbo].tbl_item_Category i on d.itemid=i.itemid and i.code_div='D'"

		sql = sql + " LEFT JOIN db_item.dbo.tbl_"& tblcate &" c on c.code_large= i.code_large "

		IF FRectCD1<>"" Then
			sql = sql + " and c.code_mid = i.code_mid "
		End IF
		IF FRectCD2<>"" Then
			sql = sql + " and c.code_small = i.code_small "
		End if
		sql = sql + " WHERE m.regdate between '"& FRectFromDate &"' and '"& FRectToDate &"'"

		if FRectJoinMallNotInclude<>"on" then
			sql = sql + " and m.sitename ='10x10'"
		end if

		if FRectExtMallNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'50'"
		end if

''		if FRectPointNotInclude<>"on" then
''			sql = sql + " and m.accountdiv<>'30'"
''		end if

		IF FRectCD1<>"" Then
			sql = sql + " and i.code_large='"& FRectCD1 &"'"
		End IF
		IF FRectCD2<>"" Then
			sql = sql + " and i.code_mid='"& FRectCD2 &"'"
		End IF

		IF FRectToDateGubun="M" THEN
		sql = sql + " group by  i."& selcate &", c.code_nm ,convert(varchar(7),m.regdate,20)"
		sql = sql + " order by  yyyymmdd desc ,itemgubun asc "
		ELSE
		sql = sql + " group by  i."& selcate &", c.code_nm ,convert(varchar(10),m.regdate,20)"
		sql = sql + " order by  yyyymmdd desc ,itemgubun asc"

		END IF

		'response.write sql
		rsget.CursorLocation = adUseClient
		rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
			set FMasterItemList(i) = new CDesignerJumunList
		    FMasterItemList(i).Fitemgubun = rsget("itemgubun")
		    FMasterItemList(i).FItemGubunNm = rsget("ItemGubunNm")
		    FMasterItemList(i).Fsitename = rsget("yyyymmdd")
			FMasterItemList(i).Fselltotal = rsget("sumtotal")
			FMasterItemList(i).Fsellcnt = rsget("sellcnt")
			'FMasterItemList(i).Fdpart = rsget("dpart")

			FMasterItemList(i).Fbuytotal = rsget("buytotal")

			if Not IsNull(FMasterItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
			end if

			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end sub

	public sub SearchMallSellrePortMonthlyChannel()
		Dim sql, i,selcate, tblcate
		maxt = -1
   		maxc = -1
        ''[db_item].[dbo].tbl_item -> [db_item].[dbo].tbl_item_Category 로 수정 .. 특정쿼리 느림..?

   		if FRectCD1<>"" and FRectCD2<>"" Then
   			selcate = "code_small"
   			tblcate = "cate_small"
   		ElseIf FRectCD1<>"" then
   			selcate ="code_mid"
   			tblcate = "cate_mid"
   		else
   			selcate ="code_large"
   			tblcate = "cate_large"
   		end if

		sql = "select isnull(i."& selcate &",999) as itemgubun,isnull(c.code_nm,'기타')  as ItemGubunNm "
		''sql = sql + " ,convert(varchar(7),m.regdate,20) as yyyymm"  ''월별로만 검색되는듯..?
		sql = sql + " ,'' as yyyymm"
		sql = sql + " ,sum(d.itemno*d.itemcost) as sumtotal, sum(d.itemno*d.buycash) as buytotal"
		sql = sql + " ,count(d.itemno) as sellcnt"
		if FRectOldJumun="on" then
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m "
			sql = sql + " join [db_log].[dbo].tbl_old_order_detail_2003 d "
		else
			sql = sql + " from [db_order].[dbo].tbl_order_master m "
			sql = sql + " join [db_order].[dbo].tbl_order_detail d "
		end if
		sql = sql + " on m.orderserial=d.orderserial"
		sql = sql + " and m.ipkumdiv>3"
		sql = sql + " and d.itemid<>0"
		sql = sql + " and m.cancelyn='N'"
		sql = sql + " and d.cancelyn<>'Y'"
		sql = sql + " and m.jumundiv<>9"

		sql = sql + " left join [db_item].[dbo].tbl_item_Category i on d.itemid=i.itemid and i.code_div='D'"

		sql = sql + " Left JOIN db_item.dbo.tbl_"& tblcate &" c on c.code_large= i.code_large "

		IF FRectCD1<>"" Then
			sql = sql + " and c.code_mid = i.code_mid "
		End IF
		IF FRectCD2<>"" Then
			sql = sql + " and c.code_small = i.code_small "
		End if
		sql = sql + " WHERE m.regdate between '"& sDate &"' and '"& dateadd("m",1,sDate) &"'"

		if FRectJoinMallNotInclude<>"on" then
			sql = sql + " and m.sitename ='10x10'"
		end if

		if FRectExtMallNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'50'"
		end if

''		if FRectPointNotInclude<>"on" then
''			sql = sql + " and m.accountdiv<>'30'"
''		end if

		IF FRectCD1<>"" Then
			sql = sql + " and i.code_large='"& FRectCD1 &"'"
		End IF
		IF FRectCD2<>"" Then
			sql = sql + " and i.code_mid='"& FRectCD2 &"'"
		End IF
		sql = sql + " group by  i."& selcate &", c.code_nm " '',convert(varchar(7),m.regdate,20)"
		sql = sql + " order by  itemgubun asc ,yyyymm desc"

		rsget.CursorLocation = adUseClient
		rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fitemgubun = rsget("itemgubun")
			    FMasterItemList(i).FItemGubunNm = rsget("ItemGubunNm")
			    FMasterItemList(i).Fsitename = rsget("yyyymm")
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")

				FMasterItemList(i).Fbuytotal = rsget("buytotal")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)

					FTotalPrice = FTotalPrice + FMasterItemList(i).Fselltotal
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end sub

	public sub SearchMallSellTimerePortChannel()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select convert(varchar(10),m.regdate,20) as yyyymmdd,"
		sql = sql + " datepart(w,m.regdate) as dpart, sum(d.itemno*d.itemcost) as sumtotal,"
		sql = sql + " count(d.itemid) as sellcnt"
		sql = sql + " from [db_order].[dbo].tbl_order_master m"
		sql = sql + " ,[db_order].[dbo].tbl_order_detail d"
		sql = sql + " where m.regdate>='" + CStr(FRectFromDate) + "'"
		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'"
		sql = sql + " and m.orderserial=d.orderserial"
		sql = sql + " and d.itemid<>0"
		sql = sql + " and m.ipkumdiv>3"
		sql = sql + " and m.cancelyn='N'"
		sql = sql + " and d.cancelyn<>'Y'"
		sql = sql + " and m.jumundiv<>9"

		if FRectJoinMallNotInclude<>"on" then
			sql = sql + " and m.sitename in ('10x10','tingmart')"
		end if

		if FRectExtMallNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'50'"
		end if

		if FRectPointNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'30'"
		end if

		sql = sql + " group by  convert(varchar(10),m.regdate,20), datepart(w,m.regdate)"
		sql = sql + " order by  convert(varchar(10),m.regdate,20) desc"

		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
			set FMasterItemList(i) = new CDesignerJumunList

		    FMasterItemList(i).Fsitename = rsget("yyyymmdd")
			FMasterItemList(i).Fselltotal = rsget("sumtotal")
			FMasterItemList(i).Fsellcnt = rsget("sellcnt")
			FMasterItemList(i).Fdpart = rsget("dpart")

			if Not IsNull(FMasterItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
			end if

			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end sub

	public sub SearchMallSellTimerePortChannel1()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select convert(varchar(10),m.regdate,20) as yyyymmdd,"
		sql = sql + " datepart(w,m.regdate) as dpart, sum(d.itemno*d.itemcost) as sumtotal,"
		sql = sql + " count(d.itemid) as sellcnt"
		sql = sql + " from [db_order].[dbo].tbl_order_master m"
		sql = sql + " ,[db_order].[dbo].tbl_order_detail d"
		sql = sql + " where m.regdate>='" + CStr(FRectFromDate) + "'"
		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'"
		sql = sql + " and right(convert(varchar(19),m.regdate,120),8) >='00:00:00'"
		sql = sql + " and right(convert(varchar(19),m.regdate,120),8) < '" + Cstr(FRectToDateTime) + "'"
		sql = sql + " and m.orderserial=d.orderserial"
		sql = sql + " and d.itemid<>0"
		sql = sql + " and m.ipkumdiv>3"
		sql = sql + " and m.cancelyn='N'"
		sql = sql + " and d.cancelyn<>'Y'"
		sql = sql + " and m.jumundiv<>9"

		if FRectJoinMallNotInclude<>"on" then
			sql = sql + " and m.sitename in ('10x10','tingmart')"
		end if

		if FRectExtMallNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'50'"
		end if

		if FRectPointNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'30'"
		end if

		sql = sql + " group by  convert(varchar(10),m.regdate,20), datepart(w,m.regdate)"
		sql = sql + " order by  convert(varchar(10),m.regdate,20) desc"
'response.write sql
		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList

				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end sub

	public sub SearchMallSellTimerePortChannel2()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select IsNull(i.cate_large,'90') as itemgubun, convert(varchar(10),m.regdate,20) as yyyymmdd,"
		sql = sql + " datepart(w,m.regdate) as dpart, sum(d.itemno*d.itemcost) as sumtotal,"
		sql = sql + " count(m.idx) as sellcnt"
		sql = sql + " from [db_order].[dbo].tbl_order_master m"
		sql = sql + " ,[db_order].[dbo].tbl_order_detail d"
		sql = sql + " ,[db_item].[dbo].tbl_item i"

		sql = sql + " where m.regdate>='" + CStr(FRectFromDate) + "'"
		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'"
		sql = sql + " and right(convert(varchar(19),m.regdate,120),8) >='00:00:00'"
		sql = sql + " and right(convert(varchar(19),m.regdate,120),8) < '" + Cstr(FRectToDateTime) + "'"
		sql = sql + " and m.orderserial=d.orderserial"
		sql = sql + " and d.itemid=i.itemid"
		sql = sql + " and d.itemid<>0"
		sql = sql + " and m.ipkumdiv>3"
		sql = sql + " and m.cancelyn='N'"
		sql = sql + " and d.cancelyn<>'Y'"
		sql = sql + " and m.jumundiv<>9"

		if FRectJoinMallNotInclude<>"on" then
			sql = sql + " and m.sitename in ('10x10','tingmart')"
		end if

		if FRectExtMallNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'50'"
		end if

		if FRectPointNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'30'"
		end if

		sql = sql + " group by  i.cate_large, convert(varchar(10),m.regdate,20), datepart(w,m.regdate)"
		sql = sql + " order by  convert(varchar(10),m.regdate,20) desc, i.cate_large asc"

'response.write sql
		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fitemgubun = rsget("itemgubun")
			    FMasterItemList(i).Fsitename = rsget("yyyymmdd")
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")
				FMasterItemList(i).Fdpart = rsget("dpart")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end sub

	public sub SearchMallSellrePort_Week()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select year(m.regdate) as yyyy, DATEPART(ww,m.regdate) as weekdt , sum(m.subtotalprice) as sumtotal, count(m.idx) as sellcnt"
		sql = sql + " from [db_order].[dbo].tbl_order_master m"
		sql = sql + " where m.regdate>'2002-01-01'"

		if FRectSearchType="24" then
			sql = sql + " and datediff(ww,m.regdate,getdate())<24"
		elseif FRectSearchType="48" then
			sql = sql + " and datediff(ww,m.regdate,getdate())<48"
		else
			'sql = sql + " and day(m.regdate)<=day(getdate())"
		end if

		sql = sql + " and ipkumdiv>3"
		sql = sql + " and cancelyn='N'"

		if FRectJoinMallNotInclude<>"on" then
			sql = sql + " and m.sitename in ('10x10','tingmart')"
		end if

		if FRectExtMallNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'50'"
		end if

		if FRectPointNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'30'"
		end if

		sql = sql + " group by  year(m.regdate), DATEPART(ww,m.regdate)"
		sql = sql + " order by  year(m.regdate) desc, DATEPART(ww,m.regdate) desc"

		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = CStr(rsget("yyyy")) + "-" + CStr(rsget("weekdt"))
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end sub

	public sub SearchMallSellrePort()
    Dim sql, i

    maxt = -1
    maxc = -1

			sql = "select sum(subtotalprice) as sumtotal,"
			sql = sql + " count(orderserial) as sellcnt, sitename"
			sql = sql + " from [db_order].[dbo].tbl_order_master"
			sql = sql + " where (regdate >= '" & FRectFromDate & "') and (regdate<'" & FRectToDate & "')"
			sql = sql + " and cancelyn = 'N'"
			sql = sql + " and ipkumdiv>=4"
            sql = sql + " Group by sitename"
            sql = sql + " order by sumtotal Desc"

			rsget.Open sql,dbget,1
			FResultCount = rsget.RecordCount
		    redim preserve FMasterItemList(FResultCount)
			do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = rsget("sitename")
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
			loop
			rsget.close
	end sub

	public sub MemberBuySex()
		Dim sql, i

		sql = "select count(m.orderserial) as cnt, sum(subtotalprice) as sumprice,"
		sql = sql + " Left(Right(u.juminno,7),1) as sex"
		if FRectOldJumun="on" then
		    sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m, [db_user].[dbo].tbl_user_n u"
		ELSE
		    sql = sql + " from [db_order].[dbo].tbl_order_master m, [db_user].[dbo].tbl_user_n u"
	    END IF
		sql = sql + " where m.regdate >='" & FRectFromDate & "'"
        sql = sql + " and m.regdate < '" & FRectToDate & "'"
		sql = sql + " and m.sitename='10x10'"
		sql = sql + " and m.cancelyn='N'"
		sql = sql + " and m.userid=u.userid"
		sql = sql + " and m.userid <> ''"
		sql = sql + " and m.ipkumdiv>=4"
		sql = sql + " and m.jumundiv<9"
		sql = sql + " group by Left(Right(juminno,7),1)"

		rsget.Open sql,dbget,1
		do until rsget.Eof
			if rsget("sex")="1" then
				FManTotalMoney = rsget("sumprice")
				FManTotalCount = rsget("cnt")
			end if

			if rsget("sex")="2" then
				FWoManTotalMoney = rsget("sumprice")
				FWoManTotalCount = rsget("cnt")
			end if

			rsget.MoveNext
		loop
		rsget.close
	end sub

	public sub MemberBuyPercent()
    	Dim sql, i

			sql = "select count(orderserial) as cnt, sum(subtotalprice) as sumprice"
			if FRectOldJumun="on" then
				sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m"
			else
				sql = sql + " from [db_order].[dbo].tbl_order_master m"
			end if
			sql = sql + " where regdate >='" & FRectFromDate & "'"
            sql = sql + " and regdate < '" & FRectToDate & "'"
			sql = sql + " and sitename='10x10'"
			sql = sql + " and cancelyn='N'"
			sql = sql + " and ipkumdiv>=4"
			sql = sql + " and jumundiv<9"

			rsget.Open sql,dbget,1
			if not rsget.Eof then
						Ftotalmoney = rsget("sumprice")
						FTotalsellcnt = rsget("cnt")
			end if
			rsget.close

			if isNUll(Ftotalmoney) then Ftotalmoney =0
			if isNUll(FTotalsellcnt) then FTotalsellcnt =0

			sql = "select count(m.orderserial) as cnt, sum(subtotalprice) as sumprice"
			if FRectOldJumun="on" then
				sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m,"
			else
				sql = sql + " from [db_order].[dbo].tbl_order_master m,"
			end if
			sql = sql + " [db_user].[dbo].tbl_user_n u"
			sql = sql + " where m.regdate >='" & FRectFromDate & "'"
            sql = sql + " and m.regdate < '" & FRectToDate & "'"
			sql = sql + " and m.sitename='10x10'"
			sql = sql + " and m.cancelyn='N'"
			sql = sql + " and m.userid=u.userid"
			sql = sql + " and m.userid <> ''"
			sql = sql + " and m.ipkumdiv>=4"
			sql = sql + " and m.jumundiv<9"
			sql = sql + " and u.regdate >='" & FRectFromDate & "'"
			sql = sql + " and u.regdate < '" & FRectToDate & "'"


			rsget.Open sql,dbget,1
			if not rsget.Eof then
						FNtotalmoney = rsget("sumprice")
						FNtotalsellcnt = rsget("cnt")
			end if
			rsget.close

			if isNUll(FNtotalmoney) then FNtotalmoney =0
			if isNUll(FNtotalsellcnt) then FNtotalsellcnt =0

			sql = "select count(orderserial) as cnt, sum(subtotalprice) as sumprice"
			sql = sql + " from [db_order].[dbo].tbl_order_master"
			sql = sql + " where regdate >='" & FRectFromDate & "'"
            sql = sql + " and regdate < '" & FRectToDate & "'"
			sql = sql + " and sitename='10x10'"
			sql = sql + " and cancelyn='N'"
			sql = sql + " and userid = ''"
			sql = sql + " and ipkumdiv>=4"
			sql = sql + " and jumundiv<9"

			rsget.Open sql,dbget,1
			if not rsget.Eof then
				FBtotalmoney = rsget("sumprice")
				FBTotalsellcnt = rsget("cnt")
			end if
			rsget.close

			if isNUll(FBtotalmoney) then FBtotalmoney =0
			if isNUll(FBTotalsellcnt) then FBTotalsellcnt =0

          FMtotalmoney = Ftotalmoney - FBtotalmoney - FNtotalmoney
          FMtotalsellcnt = Ftotalsellcnt - FBtotalsellcnt - FNtotalsellcnt

	end sub

	public sub MemberBuyPercent2()
    	Dim sql, i

			sql = "select count(orderserial) as cnt, sum(subtotalprice) as sumprice"
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003"
			sql = sql + " where regdate >='" & FRectFromDate & "'"
            sql = sql + " and regdate < '" & FRectToDate & "'"
			sql = sql + " and sitename='10x10'"
			sql = sql + " and cancelyn='N'"
			sql = sql + " and ipkumdiv>=4"
			sql = sql + " and jumundiv<9"

			rsget.Open sql,dbget,1
						Ftotalmoney = rsget("sumprice")
						FTotalsellcnt = rsget("cnt")
			rsget.close

			sql = "select count(m.orderserial) as cnt, sum(subtotalprice) as sumprice"
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m, [db_user].[dbo].tbl_user_n u"
			sql = sql + " where m.regdate >='" & FRectFromDate & "'"
            sql = sql + " and m.regdate < '" & FRectToDate & "'"
			sql = sql + " and m.sitename='10x10'"
			sql = sql + " and m.cancelyn='N'"
			sql = sql + " and m.userid=u.userid"
			sql = sql + " and m.userid <> ''"
			sql = sql + " and m.ipkumdiv>=4"
			sql = sql + " and m.jumundiv<9"
			sql = sql + " and u.regdate >='" & FRectFromDate & "'"
			sql = sql + " and u.regdate < '" & FRectToDate & "'"

			rsget.Open sql,dbget,1
						FNtotalmoney = rsget("sumprice")
						FNtotalsellcnt = rsget("cnt")
			rsget.close

			sql = "select count(orderserial) as cnt, sum(subtotalprice) as sumprice"
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003"
			sql = sql + " where regdate >='" & FRectFromDate & "'"
            sql = sql + " and regdate < '" & FRectToDate & "'"
			sql = sql + " and sitename='10x10'"
			sql = sql + " and cancelyn='N'"
			sql = sql + " and userid = ''"
			sql = sql + " and ipkumdiv>=4"
			sql = sql + " and jumundiv<9"

			rsget.Open sql,dbget,1
				FBtotalmoney = rsget("sumprice")
				FBTotalsellcnt = rsget("cnt")
			rsget.close

			if isNUll(FBtotalmoney) then FBtotalmoney =0
			if isNUll(FBTotalsellcnt) then FBTotalsellcnt =0

          FMtotalmoney = Ftotalmoney - FBtotalmoney - FNtotalmoney
          FMtotalsellcnt = Ftotalsellcnt - FBtotalsellcnt - FNtotalsellcnt

	end sub

	public Sub SearchJumunListBybestseller()
		dim sqlStr
		dim wheredetail
		dim i

		wheredetail = ""

		sqlStr = "select top " + CStr(FPageSize)
		sqlStr = sqlStr + " d.itemid, d.buycash, d.itemcost, "
		sqlStr = sqlStr + " d.itemname, d.makerid, d.itemoptionname"
		sqlStr = sqlStr + " , sum(d.itemno) as sm ,sum(d.itemno*d.buycash)as sm2"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m "
		sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + "     on  m.orderserial=d.orderserial"
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + " and m.ipkumdiv>=4"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"
		if (FRectRegStart<>"") then
			sqlStr = sqlStr + " and m.regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			sqlStr = sqlStr + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectDesignerID<>"") then
		    sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
		end if

		if (FRectckpointsearch = "") then
		    sqlStr = sqlStr + " and m.accountdiv <> 30"
		end if
		sqlStr = sqlStr + " group by d.itemid, d.buycash, d.itemcost, d.itemname, d.makerid, d.itemoptionname"
		sqlStr = sqlStr + " order by sm2 Desc"
        
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		''rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).FItemSellprice       = rsget("sm2")
				FMasterItemList(i).FItemNo       = rsget("sm")
				FMasterItemList(i).FItemID       = rsget("itemid")
				FMasterItemList(i).FItemCost       = rsget("itemcost")
				FMasterItemList(i).FItemName     = db2html(rsget("itemname"))
				FMasterItemList(i).FItemOptionStr= db2html(rsget("itemoptionname"))
				FMasterItemList(i).FBuycash		= rsget("buycash")
				FMasterItemList(i).FMakerid		= rsget("makerid")
				rsget.movenext
				i=i+1
			loop
		rsget.Close
	end sub

	public Sub i_SearchJumunListBybestsellerDesc()
		dim sqlStr
		dim wheredetail
		dim i

		wheredetail = ""

		if (FRectRegStart<>"") then
			wheredetail = wheredetail + " and m.regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			wheredetail = wheredetail + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectDesignerID<>"") then
		wheredetail = wheredetail + " and d.makerid='" + FRectDesignerID + "'"
		end if

		if (FRectckpointsearch = "") then
		wheredetail = wheredetail + " and m.accountdiv <> 30"
		end if

		sqlStr = "select top " + CStr(FPageSize)
		sqlStr = sqlStr + " d.itemid, d.makerid, IsNull(T.sm,0) as sm, IsNull(T.buycash,0) as buycash, "
		sqlStr = sqlStr + " IsNull(T.itemcost,0) as itemcost, ii.itemname, T.optname"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item ii"
		sqlStr = sqlStr + " Left join "
		sqlStr = sqlStr + " (select "
		sqlStr = sqlStr + " sum(d.itemno) as sm, d.buycash, d.itemcost, d.itemid, d.itemname, d.itemoptionname"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, "
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.ipkumdiv>=4"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " group by d.itemid, d.buycash, d.itemcost, d.itemname,  d.itemoptionname"
		sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " on ii.itemid=T.itemid"

		sqlStr = sqlStr + " where ii.itemid<>0"
		if FRectDispY="on" then
			sqlStr = sqlStr + " and ii.dispyn='Y'"
		end if

		if FRectSellY="on" then
			sqlStr = sqlStr + " and ii.sellyn='Y'"
		end if

		sqlStr = sqlStr + " order by sm asc, ii.itemid desc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).FItemNo       = rsget("sm")
				FMasterItemList(i).FItemID       = rsget("itemid")
				FMasterItemList(i).FItemCost       = rsget("itemcost")
				FMasterItemList(i).FItemName     = rsget("itemname")
				FMasterItemList(i).FItemOptionStr= rsget("optname")
				FMasterItemList(i).FBuycash		= rsget("buycash")
				FMasterItemList(i).FMakerid		= rsget("makerid")
				rsget.movenext
				i=i+1
			loop
		rsget.Close
	end sub

	'//admin/report/Channelupchesellamount.asp
	public sub ChannelBrandSellrePort()
        Dim sql, sqltmp, i, sqlsearch

    	maxt = -1
    	maxc = -1

		if (FRectCD1<>"") then
		    sqlsearch = sqlsearch + " and i.cate_large='" + FRectCD1 + "'"
		end if
		if (Fsitename<>"") then
		    sqlsearch = sqlsearch + " and (m.rdsite='" + Fsitename + "' or m.sitename='" + Fsitename + "')"
		end if
		IF FRectDispCate<>"" THEN	
			sqlsearch = sqlsearch + " and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		END IF	
		If FRectPurchasetype <> "" Then
			sqlsearch = sqlsearch & " and p.purchasetype = '" & FRectPurchasetype &"'"
		End IF
		if (FRectMwDiv<>"") then
			sqlsearch = sqlsearch & " and d.omwdiv = '" & FRectMwDiv &"'"
		end if

		sql = "select sum(d.itemcost*d.itemno) as sumtotal, sum(d.buycash*d.itemno) as buytotal, "
		sql = sql + " d.makerid,"
		sql = sql + " sum(d.itemno) as sellcnt"
		sql = sql + " , c.vatinclude, c.maeipdiv, c.defaultmargine, c.defaultFreeBeasongLimit, c.defaultDeliverPay, c.defaultDeliveryType"
		sql = sql + " ,IsNull(p.M_margin,0) as M_margin, IsNull(p.W_margin,0) as W_margin, IsNull(p.U_margin,0) as U_margin"

		if FRectOldJumun="on" then
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m"
			sql = sql + " join [db_log].[dbo].tbl_old_order_detail_2003 d"
			sql = sql + " 	on m.orderserial = d.orderserial"
		else
			sql = sql + " from [db_order].[dbo].tbl_order_master m"
			sql = sql + " join [db_order].[dbo].tbl_order_detail d"
			sql = sql + " 	on m.orderserial = d.orderserial"
		end if

		sql = sql + " join [db_item].[dbo].tbl_item as i"
		sql = sql + " 	on d.itemid=i.itemid"
		sql = sql + " left join [db_user].[dbo].tbl_user_c c"
		sql = sql + " 	on d.makerid=c.userid"
		sql = sql + " left join [db_partner].[dbo].tbl_partner p"
		sql = sql + "  	on d.makerid=p.id"

		IF FRectDispCate<>"" THEN	'2014-02-27 정윤정 전시카테고리 검색 추가 
			sql = sql & " join db_item.dbo.tbl_display_cate_item as dc"
			sql = sql & " 	on d.itemid = dc.itemid"
		END IF

		sql = sql + " where d.itemid not in (0,100)"
		sql = sql + " and (m.regdate >= '" & FRectFromDate & "') "
		sql = sql + " and (m.regdate < '" & FRectToDate & "')"
		sql = sql + " and m.cancelyn = 'N'"
		sql = sql + " and m.jumundiv<>'9'"
		sql = sql + " and d.cancelyn <> 'Y'"
		sql = sql + " and m.ipkumdiv>=4"
		sql = sql + sqlsearch
        sql = sql + " Group by d.makerid"
		sql = sql + " , c.vatinclude, c.maeipdiv, c.defaultmargine, c.defaultFreeBeasongLimit, c.defaultDeliverPay, c.defaultDeliveryType"
		sql = sql + " ,IsNull(p.M_margin,0), IsNull(p.W_margin,0), IsNull(p.U_margin,0)"

		if FRectOrdertype = "totalprice" then
			sql = sql + " order by sumtotal Desc"
		elseif FRectOrdertype = "totalgain" then
		    sql = sql + " order by sum(d.itemcost*d.itemno-d.buycash*d.itemno) Desc"
		else
			sql = sql + " order by sellcnt Desc"
		end if

		'response.write sql
		rsget.Open sql,dbget,1

		FResultCount = rsget.RecordCount
        redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
			set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).Fmakerid = rsget("makerid")
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")
				FMasterItemList(i).Fbuytotal = rsget("buytotal")
	
				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if
	
				FTotalPrice = FTotalPrice + FMasterItemList(i).Fselltotal
				FTotalBuyPrice = FTotalBuyPrice + FMasterItemList(i).Fbuytotal
				FTotalEA = FTotalEA + FMasterItemList(i).Fsellcnt
				FMasterItemList(i).Fmaeipdiv       	= rsget("maeipdiv")
				FMasterItemList(i).Fdefaultmargine 	= rsget("defaultmargine")
				FMasterItemList(i).FM_margin 		 	= rsget("M_margin")
				FMasterItemList(i).FW_margin       	= rsget("W_margin")
				FMasterItemList(i).FU_margin       	= rsget("U_margin")
			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end sub

	public sub SearchCardOnline()
		Dim sql, i, vDBTable
		maxt = -1
   		maxc = -1
   		
   		If FRectOldJumun = "y" Then
   			vDBTable = "[db_log].[dbo].tbl_old_order_master_2003"
   		Else
   			vDBTable = "[db_order].[dbo].tbl_order_master"
   		End If

		sql = "select convert(varchar(10),m.regdate,20) as yyyymmdd, datepart(w,m.regdate) as dpart, "
		sql = sql + " sum(m.subtotalprice) as sumtotal, count(m.idx) as sellcnt, accountdiv"
		sql = sql + " from " + vDBTable + " m"
		sql = sql + " where m.regdate>='" + CStr(FRectFromDate) + "'"
		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'"

		sql = sql + " and ipkumdiv>3"
		sql = sql + " and cancelyn='N'"

		if FRectOrdertype="on" then
			sql = sql + " and m.sitename in ('10x10','tingmart')"
		end if

		sql = sql + " group by  convert(varchar(10),m.regdate,20), datepart(w,m.regdate),accountdiv"
		sql = sql + " order by  convert(varchar(10),m.regdate,20) desc"
''response.write sql
		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = rsget("yyyymmdd")
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")
				FMasterItemList(i).Fdpart = rsget("dpart")
				FMasterItemList(i).Faccountdiv = rsget("accountdiv")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end sub

	public sub SearchCardOnlineMonth()
		Dim sql, i, vDBTable
		maxt = -1
   		maxc = -1

   		If FRectOldJumun = "y" Then
   			vDBTable = "[db_log].[dbo].tbl_old_order_master_2003"
   		Else
   			vDBTable = "[db_order].[dbo].tbl_order_master"
   		End If

		sql = "select convert(varchar(7),m.regdate,20) as yyyymm, sum(m.subtotalprice) as sumtotal, count(m.idx) as sellcnt,accountdiv"
		sql = sql + " from " + vDBTable + " m"
'		sql = sql + " where m.regdate>='2002-10-01'"
'		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'"
		sql = sql + " where m.regdate>='" + CStr(FRectFromDate) + "'"
		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'"

		sql = sql + " and ipkumdiv>3"
		sql = sql + " and cancelyn='N'"

		if FRectOrdertype="on" then
			sql = sql + " and m.sitename in ('10x10','tingmart')"
		end if

		sql = sql + " group by  convert(varchar(7),m.regdate,20),accountdiv"
		sql = sql + " order by  convert(varchar(7),m.regdate,20) desc"
''response.write sql
		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = rsget("yyyymm")
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")
'				FMasterItemList(i).Fdpart = rsget("dpart")
				FMasterItemList(i).Faccountdiv = rsget("accountdiv")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end sub

	public sub OnlineBankingReport()
		Dim sql, i
		maxt = -1
   		maxc = -1
		maxa = -1
   		maxb = -1

		sql = "select convert(varchar(10),regdate,20) as yyyymmdd, datepart(w,regdate) as dpart," + vbcrlf
		sql = sql + " sum(subtotalprice) as sumtotal, count(idx) as sellcnt," + vbcrlf
		sql = sql + " sum(" + vbcrlf
		sql = sql + " case " + vbcrlf
		sql = sql + "	when (accountdiv='7') and (ipkumdiv='2') then subtotalprice" + vbcrlf
		sql = sql + "	else 0" + vbcrlf
		sql = sql + " end) as cash," + vbcrlf
		sql = sql + " sum(" + vbcrlf
		sql = sql + " case " + vbcrlf
		sql = sql + "	when (accountdiv='7') and (ipkumdiv='2') then 1" + vbcrlf
		sql = sql + "	else 0" + vbcrlf
		sql = sql + " end) as onlinecnt" + vbcrlf
		sql = sql + " from [db_order].[dbo].tbl_order_master" + vbcrlf
		sql = sql + " where regdate>='" + CStr(FRectFromDate) + "'" + vbcrlf
		sql = sql + " and regdate<'" + CStr(FRectToDate) + "'" + vbcrlf
		sql = sql + " and cancelyn='N'" + vbcrlf
		sql = sql + " and accountdiv='7'" + vbcrlf
'		sql = sql + " and ipkumdiv>1" + vbcrlf
		sql = sql + " group by  convert(varchar(10),regdate,20), datepart(w,regdate)" + vbcrlf
		sql = sql + " order by  convert(varchar(10),regdate,20) desc"

		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = rsget("yyyymmdd")
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")
				FMasterItemList(i).Fcash = rsget("cash")
				FMasterItemList(i).Fonlinecnt = rsget("onlinecnt")
				FMasterItemList(i).Fdpart = rsget("dpart")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
					maxa = MaxVal(maxa,FMasterItemList(i).Fcash)
					maxb = MaxVal(maxb,FMasterItemList(i).Fonlinecnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end sub

	public Sub CooperationJumunListBybestseller()
		dim sqlStr
		dim wheredetail
		dim i

		wheredetail = ""

		if (FRectRegStart<>"") then
			wheredetail = wheredetail + " and m.regdate >='" + CStr(FRectRegStart) + "'" + vbcrlf
		end if

		if (FRectRegEnd<>"") then
			wheredetail = wheredetail + " and m.regdate <'" + CStr(FRectRegEnd) + "'" + vbcrlf
		end if

		if (FRectDesignerID<>"") then
			if FRectDesignerID = "yahoo" or FRectDesignerID = "mym" or FRectDesignerID = "empas" then
				wheredetail = wheredetail + " and m.rdsite = '" + FRectDesignerID + "'" + vbcrlf
			else
				wheredetail = wheredetail + " and m.sitename = '" + FRectDesignerID + "'" + vbcrlf
			end if
		end if

		if FRectDispY="on" then
			wheredetail = wheredetail + " and i.dispyn='Y'" + vbcrlf
		end if

		if FRectSellY="on" then
			wheredetail = wheredetail + " and i.sellyn='Y'" + vbcrlf
		end if

		sqlStr = "select top 100  l.code_nm, sum(i.sellcash) as sellsum, sum(d.itemno) as sellcnt" + vbcrlf
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on d.itemid = i.itemid" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_large l on i.cate_large = l.code_large" + vbcrlf
		sqlStr = sqlStr + " where m.orderserial = d.orderserial" + vbcrlf
		sqlStr = sqlStr + " and m.ipkumdiv >= 4" + vbcrlf
		sqlStr = sqlStr + " and m.cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + " and d.cancelyn<>'Y'" + vbcrlf
		sqlStr = sqlStr + " and i.itemid <> 0" + vbcrlf
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " group by l.code_nm" + vbcrlf
		sqlStr = sqlStr + " order by sum(d.itemno) desc"
'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).Fcode_nm       = rsget("code_nm")
				FMasterItemList(i).Fsubtotalprice       = rsget("sellsum")
				FMasterItemList(i).Fitemno       = rsget("sellcnt")
				rsget.movenext
				i=i+1
			loop
		rsget.Close
	end sub

	public sub CaMallSellrePort()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select convert(varchar(10),m.regdate,20) as yyyymmdd," + vbcrlf
		sql = sql + " datepart(w,m.regdate) as dpart," + vbcrlf
		sql = sql + " sum(d.itemcost) as sumtotal," + vbcrlf
		sql = sql + " count(d.itemid) as sellcnt" + vbcrlf
		sql = sql + " from [db_contents].[dbo].tbl_camall_item c, [db_order].[dbo].tbl_order_detail d," + vbcrlf
		sql = sql + " [db_order].[dbo].tbl_order_master m" + vbcrlf
		sql = sql + " where c.itemid = d.itemid" + vbcrlf
		sql = sql + " and m.orderserial = d.orderserial" + vbcrlf
		sql = sql + " and m.regdate>='" + CStr(FRectFromDate) + "'"
		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'"
		sql = sql + " and m.ipkumdiv>3" + vbcrlf
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " and d.cancelyn<>'Y'" + vbcrlf
		sql = sql + " and m.accountdiv<>'30'" + vbcrlf
		if FRectCD1 <> "" then
		sql = sql + " and c.code_large='" + FRectCD1 + "'" + vbcrlf
		end if
		if FRectCD2 <> "" then
		sql = sql + " and c.code_mid='" + FRectCD2 + "'" + vbcrlf
		end if
		sql = sql + " group by convert(varchar(10),m.regdate,20), datepart(w,m.regdate)" + vbcrlf
		sql = sql + " order by convert(varchar(10),m.regdate,20) desc" + vbcrlf
'response.write sql
		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = rsget("yyyymmdd")
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")
				FMasterItemList(i).Fdpart = rsget("dpart")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end sub

	public sub MailGuMaeDayTotalReport()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select convert(varchar(10),m.regdate,20) as yyyymmdd," + vbcrlf
		sql = sql + " sum(m.subtotalprice) as sumtotal, count(m.orderserial) as sellcnt " + vbcrlf
		sql = sql + " from [db_order].[dbo].tbl_order_master m," + vbcrlf
		sql = sql + " [db_user].[dbo].tbl_user_n u" + vbcrlf
		sql = sql + " where m.userid = u.userid" + vbcrlf
		sql = sql + " and m.userid <> ''" + vbcrlf
		sql = sql + " and m.regdate >= '" + CStr(FRectFromDate) + "'" + vbcrlf
		sql = sql + " and m.regdate < '" + CStr(FRectToDate) + "'" + vbcrlf
		sql = sql + " group by convert(varchar(10),m.regdate,20)" + vbcrlf
		sql = sql + " order by convert(varchar(10),m.regdate,20) asc" + vbcrlf
'response.write sql
		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList

				FMasterItemList(i).FDate = rsget("yyyymmdd")
				FMasterItemList(i).FDayselltotal = rsget("sumtotal")
				FMasterItemList(i).FDaysellcnt = rsget("sellcnt")

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end sub

	public sub MailGuMaeReport()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select SUBSTRING(u.juminno,8,1) as sex, convert(varchar(10),m.regdate,20) as yyyymmdd," + vbcrlf
		sql = sql + " datepart(w,m.regdate) as dpart, sum(m.subtotalprice) as sumtotal, count(m.orderserial) as sellcnt " + vbcrlf
		sql = sql + " from [db_order].[dbo].tbl_order_master m," + vbcrlf
		sql = sql + " [db_user].[dbo].tbl_user_n u" + vbcrlf
		sql = sql + " where m.userid = u.userid" + vbcrlf
		sql = sql + " and m.userid <> ''" + vbcrlf
		sql = sql + " and m.regdate >= '" + CStr(FRectFromDate) + "'" + vbcrlf
		sql = sql + " and m.regdate < '" + CStr(FRectToDate) + "'" + vbcrlf
		sql = sql + " group by convert(varchar(10),m.regdate,20),SUBSTRING(u.juminno,8,1),datepart(w,m.regdate)" + vbcrlf
		sql = sql + " order by convert(varchar(10),m.regdate,20) asc,SUBSTRING(u.juminno,8,1) desc" + vbcrlf
'response.write sql
		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsex = rsget("sex")
				FMasterItemList(i).Fsitename = rsget("yyyymmdd")
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")
				FMasterItemList(i).Fdpart = rsget("dpart")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end sub

public FRectSellChannelDiv

	public sub SearchTimeSellrePort()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select datepart(hh,m.regdate) as dpart," + vbcrlf
		sql = sql + " sum(m.subtotalprice) as sumtotal, count(m.orderserial) as sellcnt" + vbcrlf
		sql = sql + " ,sum(Case When m.jumundiv=9 then m.subtotalprice else 0 end) as minustt " + vbcrlf
		sql = sql + " ,sum(Case When m.jumundiv=9 then 1 else 0 end) as minuscnt " + vbcrlf

		if FRectOldJumun="on" then
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m " + vbcrlf
		else
			sql = sql + " from [db_order].[dbo].tbl_order_master m " + vbcrlf
		end if

		sql = sql + " where m.regdate>='" + CStr(FRectRegStart) + "'"
		sql = sql + " and m.regdate<'" + CStr(FRectRegEnd) + "'"
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " and ipkumdiv>3" + vbcrlf
		sql = sql + " and sitename <>'way2way'" + vbcrlf

		if FRectRdsite="on" then
			sql = sql + " and m.rdsite in ('mobile','mobile_kakaotalk','mobile_nate','mobile_kakaotms') "
		end if
		if (FRectSellChannelDiv<>"") then   '변경 2015.05.27 정윤정
       		sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDxGroup(FRectSellChannelDiv)&")"
    	end if
		
' 	if (FRectChannelDiv<>"") then
'			if FRectChannelDiv="web" then
'				sql = sql & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") and m.accountdiv<>'50' "
'			elseif FRectChannelDiv="jaehu" then
'				sql = sql & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") "
'			elseif FRectChannelDiv="mjaehu" then
'				sql = sql & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") "
'			elseif FRectChannelDiv="mobile" then
'				sql = sql & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") and m.accountdiv<>'50' "
'			elseif FRectChannelDiv="ipjum" then
'				sql = sql & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") "
'			end if
'		end if

		sql = sql + " group by datepart(hh,m.regdate)" + vbcrlf
		sql = sql + " order by datepart(hh,m.regdate) asc"
 'response.write sql
		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList

				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")
				FMasterItemList(i).Fdpart = rsget("dpart")
				FMasterItemList(i).Fminustotal = rsget("minustt")
				FMasterItemList(i).Fminuscount = rsget("minuscnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end sub

	public sub SearchChannalDailySellRePort()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select convert(varchar(10),m.regdate,20) as yyyymmdd, datepart(w,m.regdate) as dpart," + vbcrlf
		sql = sql + " sum(d.itemno*d.itemcost) as sumtotal, count(m.idx) as sellcnt" + vbcrlf
		if FRectOldJumun="on" then
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m,"
			sql = sql + " [db_log].[dbo].tbl_old_order_detail_2003 d,"
		else
			sql = sql + " from [db_order].[dbo].tbl_order_master m, "
			sql = sql + " [db_order].[dbo].tbl_order_detail d,"
		end if

		sql = sql + " [db_item].[dbo].tbl_item i" + vbcrlf
		sql = sql + " where m.orderserial = d.orderserial" + vbcrlf
		sql = sql + " and d.itemid = i.itemid" + vbcrlf
		sql = sql + " and m.regdate>='" + CStr(FRectFromDate) + "'" + vbcrlf
		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'" + vbcrlf
		sql = sql + " and m.ipkumdiv>3" + vbcrlf
		sql = sql + " and m.cancelyn='N'" + vbcrlf
'		sql = sql + " and m.sitename <>'way2way'" + vbcrlf
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " and m.jumundiv<>9" + vbcrlf
		sql = sql + " and d.cancelyn<>'Y'" + vbcrlf
		sql = sql + " and i.cate_large='" + FRectCD1 + "'" + vbcrlf
		If FRectCD2 <> "" Then
		sql = sql + " and i.cate_mid='" + FRectCD2 + "'" + vbcrlf
		End If
		If FRectCD3 <> "" Then
		sql = sql + " and i.cate_small='" + FRectCD3 + "'" + vbcrlf
		End If
		sql = sql + " group by  convert(varchar(10),m.regdate,20), datepart(w,m.regdate)"
		sql = sql + " order by  yyyymmdd desc"
'response.write sql
		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = rsget("yyyymmdd")
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")
				FMasterItemList(i).Fdpart = rsget("dpart")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end sub

	public Sub SearchJumunListByupcheSelllist()
		dim sqlStr
		dim wheredetail
		dim i

		wheredetail = ""

		if (FRectRegStart<>"") then
			if (FRectDateType="ipkumil") then
				wheredetail = wheredetail + " and m.ipkumdate >='" + CStr(FRectRegStart) + "'"
			elseif (FRectDateType="beadal") then
				wheredetail = wheredetail + " and m.beadaldate >='" + CStr(FRectRegStart) + "'"
			else
				wheredetail = wheredetail + " and m.regdate >='" + CStr(FRectRegStart) + "'"
			end if
		end if

		if (FRectRegEnd<>"") then
			if (FRectDateType="ipkumil") then
				wheredetail = wheredetail + " and m.ipkumdate <'" + CStr(FRectRegEnd) + "'"
			elseif (FRectDateType="beadal") then
				wheredetail = wheredetail + " and m.beadaldate <'" + CStr(FRectRegEnd) + "'"
			else
				wheredetail = wheredetail + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
			end if
		end if

		if (FRectDelNoSearch<>"") then
			wheredetail = wheredetail + " and m.cancelyn ='N'"
		end if

		if (FRectIpkumDiv4<>"") then
			wheredetail = wheredetail + " and m.ipkumdiv>=4"
		end if

		if (FRectItemid<>"") then
		wheredetail = wheredetail + " and d.itemid=" + FRectItemid
		end if

		if (FRectDesignerID<>"") then
		wheredetail = wheredetail + " and d.makerid='" + FRectDesignerID + "'"
		end if

		sqlStr = "select top 500 "
		sqlStr = sqlStr + " d.itemid, d.itemcost, sum(d.itemno) as sm, d.itemname, d.itemoptionname"

		if FRectOldJumun="on" then
			sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m,"
			sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d"
		else
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		end if
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.ipkumdiv>1"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " group by d.itemid,  d.itemcost, d.itemname,  d.itemoptionname"
		sqlStr = sqlStr + " order by sm desc"

'response.write sqlStr
		rsget.PageSize = FPageSize
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage
			do until (i >= FResultCount)

				set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).FItemNo       = rsget("sm")
				FMasterItemList(i).FItemID       = rsget("itemid")
				FMasterItemList(i).FItemCost       = rsget("itemcost")
				FMasterItemList(i).FItemName     = db2html(rsget("itemname"))
				FMasterItemList(i).FItemOptionStr= db2html(rsget("itemoptionname"))

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public Sub GetWeeklySellCount()
		Dim sql, ix

		sql = "select convert(varchar(7),m.regdate,20) as yyyymm" + vbcrlf
		if FRectOldJumun="on" then
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m" + vbcrlf
		else
			sql = sql + " from [db_order].[dbo].tbl_order_master m" + vbcrlf
		end if
		sql = sql + " where m.regdate >= '" + CStr(FRectFromDate) + "'" + vbcrlf
		sql = sql + " and m.regdate < '" + CStr(FRectToDate) + "'" + vbcrlf
		sql = sql + " and m.ipkumdiv>3" + vbcrlf
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " group by convert(varchar(7),m.regdate,20)" + vbcrlf
		sql = sql + " order by yyyymm desc"

		rsget.Open sql,dbget,1

		FTotalCount = rsget.RecordCount

	    redim preserve FMasterItemList(FTotalCount)

		do until rsget.eof
				set FMasterItemList(ix) = new CDesignerJumunList
			    FMasterItemList(ix).FYYYYMMDDHHNNSS = rsget("yyyymm")
				rsget.MoveNext
				ix = ix + 1
		loop
		rsget.close

	end Sub

	public Sub GetWeeklySellReport()

		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select datepart(w,m.regdate) as dpart," + vbcrlf
		sql = sql + " sum(m.subtotalprice) as sumtotal," + vbcrlf
		sql = sql + " count(m.idx) as sellcnt," + vbcrlf
		sql = sql + " count(distinct convert(varchar(10),m.regdate,20)) as dpartcount" + vbcrlf

		if FRectOldJumun="on" then
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m" + vbcrlf
		else
			sql = sql + " from [db_order].[dbo].tbl_order_master m" + vbcrlf
		end if
		sql = sql + " where m.regdate >= '" + CStr(FRectFromDate) + "'" + vbcrlf
		sql = sql + " and m.regdate < '" + CStr(FRectToDate) + "'" + vbcrlf
		sql = sql + " and m.ipkumdiv>3" + vbcrlf
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " and m.jumundiv<>'9'" + vbcrlf
		sql = sql + " group by  datepart(w,m.regdate)" + vbcrlf
		sql = sql + " order by  dpart" + vbcrlf

		rsget.Open sql,dbget,1

		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")
				FMasterItemList(i).Fdpart = rsget("dpart")
				FMasterItemList(i).Fdpartcount = rsget("dpartcount")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close

	end Sub

	public Sub SearchBestsellerList()
		dim sqlStr
		dim i

        sqlStr = "select T.*, i.smallimage "
        sqlStr = sqlStr + " from ( "
		sqlStr = sqlStr + " select top " + CStr(FPageSize)
		sqlStr = sqlStr + " sum(d.itemno) as sm, d.buycash, d.itemcost, d.itemid, d.itemname, d.makerid, d.itemoptionname "
    	if FRectOldJumun="on" then
    		sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m,"
    		sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d"
    	else
    		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, "
    		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
    	end if

		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.ipkumdiv>3"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"

		if (FRectRegStart<>"") then
			sqlStr = sqlStr + " and m.regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			sqlStr = sqlStr + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectDesignerID<>"") then
			sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
		end if

		sqlStr = sqlStr + " group by d.itemid, d.buycash, d.itemcost, d.itemname, d.makerid, d.itemoptionname"

        if (FRectOrderBy="sumtotal") then
			sqlStr = sqlStr + " order by sum(d.itemno*d.itemcost) Desc"
		else
			sqlStr = sqlStr + " order by sm Desc"
		end if

        sqlStr = sqlStr + " ) as T"
        sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on T.itemid=i.itemid"

        if (FRectOrderBy="sumtotal") then
			sqlStr = sqlStr + " order by T.sm*T.itemcost Desc"
		else
			sqlStr = sqlStr + " order by T.sm Desc"
		end if

''response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).FItemNo       = rsget("sm")
				FMasterItemList(i).FItemID       = rsget("itemid")
				FMasterItemList(i).FItemCost       = rsget("itemcost")
				FMasterItemList(i).FItemName     = db2html(rsget("itemname"))
				FMasterItemList(i).FItemOptionStr= db2html(rsget("itemoptionname"))
				FMasterItemList(i).FBuycash		= rsget("buycash")
				FMasterItemList(i).FMakerid		= rsget("makerid")

				FMasterItemList(i).Fsmallimage  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FMasterItemList(i).FItemID) + "/" + rsget("smallimage")
				rsget.movenext
				i=i+1
			loop
		rsget.Close
	end sub

	public Sub SearchNewItemReport()
		dim sqlStr
		dim i

		sqlStr = " select cate_large,code_nm,count(A.itemid) as cnt "
		sqlStr = sqlStr & " from [db_item].[dbo].tbl_item A "
		sqlStr = sqlStr & " join db_item.[dbo].tbl_cate_large B "
		sqlStr = sqlStr & " 	on A.cate_large= b.code_large "
		sqlStr = sqlStr & " where cate_large < 90  "

		if (FRectRegStart<>"") then
			sqlStr = sqlStr + " and regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			sqlStr = sqlStr + " and regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectDesignerID<>"") then
			sqlStr = sqlStr + " and makerid = '" + FRectDesignerID + "'"
		end if

		sqlStr = sqlStr & " group by cate_large,code_nm "
		sqlStr = sqlStr + " order by cate_large asc" + vbcrlf

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
					 FMasterItemList(i).Fcate_large  = rsget("cate_large")
					 FMasterItemList(i).FCateName  = rsget("code_nm")
					 FMasterItemList(i).FTcnt      = rsget("cnt")
				rsget.movenext
				i=i+1
			loop
		rsget.Close
	end sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

class CReportMasterItem
 	public Fselldate
    public Fselltotal
    public Fsellcnt
	public Fitemid
	public maxt
	public maxc

	public function GetDpartName()
		if Fdpart=1 then
			GetDpartName = "<font color=#FF0000>일</font>"
		elseif Fdpart=2 then
			GetDpartName = "월"
		elseif Fdpart=3 then
			GetDpartName = "화"
		elseif Fdpart=4 then
			GetDpartName = "수"
		elseif Fdpart=5 then
			GetDpartName = "목"
		elseif Fdpart=6 then
			GetDpartName = "금"
		elseif Fdpart=7 then
			GetDpartName = "<font color=#0000FF>토</font>"
		else
			GetDpartName = ""
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

class CReportMaster
	public FMasterItemList()
	public maxt
	public maxc
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FRectRegStart
	public FRectRegEnd
    public FRectItemid
	public FCurrPage
	public FRectItemList
	public FRectSettle2

	Private Sub Class_Initialize()
		'redim preserve FMasterItemList(0)
		redim  FMasterItemList(0)

		FCurrPage = 1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub
	function MaxVal(a,b)
		if (CDbl(a)> CDbl(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function

	public sub SearchEachItemReport()
		Dim sql, i
		maxt = -1
   		maxc = -1

		FRectItemList = replace(FRectItemList,",","','")
		FRectItemList = "'" & FRectItemList & "'"

       if FRectSettle2 = "m" then
		sql = "select convert(varchar(7),m.regdate,20) as date, d.itemid,"
       elseif FRectSettle2 = "d" then
	    sql = "select convert(varchar(10),m.regdate,20) as date, d.itemid,"
	   end if

		sql = sql + " sum(m.subtotalprice) as sumtotal,"
		sql = sql + " sum(d.itemno) as sellcnt"
		sql = sql + " from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
		sql = sql + " where m.orderserial = d.orderserial"
		sql = sql + " and d.itemid in (" & Cstr(FRectItemList) & ")"
		sql = sql + " and m.regdate >='" & Cstr(FRectRegStart) & "'"
		sql = sql + " and m.regdate < '" & Cstr(FRectRegEnd) & "'"
		sql = sql + " and m.ipkumdiv>3"
		sql = sql + " and m.cancelyn='N'"
		sql = sql + " and d.cancelyn<>'Y'"

       if FRectSettle2 = "m" then
		sql = sql + " group by convert(varchar(7),m.regdate,20), d.itemid"
		sql = sql + " order by  d.itemid,convert(varchar(7),m.regdate,20) Asc"
       elseif FRectSettle2 = "d" then
		sql = sql + " group by convert(varchar(10),m.regdate,20), d.itemid"
		sql = sql + " order by  d.itemid,convert(varchar(10),m.regdate,20) Asc"
	   end if

'response.write sql
		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CReportMasterItem
			    FMasterItemList(i).Fselldate = rsget("date")
				FMasterItemList(i).Fselltotal = rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsget("sellcnt")
				FMasterItemList(i).Fitemid = rsget("itemid")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end sub
end class

class CBuyNumReport

	public Fsubtotalprice
	public Fitemno
	public Fcnt
	public Ftotalcnt
	public FRectRegStart
	public FRectRegEnd
	public FRectBuyNum
	public FRectYYYY

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public Sub FirstBuySellReport()
		dim sqlStr

	if FRectBuyNum <= 1 then
		sqlStr = "select count(m.userid) as onebuycnt, sum(m.subtotalprice) as tsum" + vbcrlf
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m" + vbcrlf
		sqlStr = sqlStr + " where m.userid <> ''" + vbcrlf
		sqlStr = sqlStr + " and m.ipkumdiv>=4" + vbcrlf
		sqlStr = sqlStr + " and m.cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + " and m.regdate >= '" + CStr(FRectRegStart) + "'" + vbcrlf
		sqlStr = sqlStr + " and m.regdate < '" + CStr(FRectRegEnd) + "'" + vbcrlf
		sqlStr = sqlStr + " and m.userid not in" + vbcrlf
		sqlStr = sqlStr + "(" + vbcrlf
		sqlStr = sqlStr + "	select userid from" + vbcrlf
		sqlStr = sqlStr + "	(" + vbcrlf
		sqlStr = sqlStr + "	select userid, count(*) as cnt" + vbcrlf
		sqlStr = sqlStr + "	from [db_order].[dbo].tbl_order_master" + vbcrlf
		sqlStr = sqlStr + "	 where userid <> ''" + vbcrlf
		sqlStr = sqlStr + "	 and ipkumdiv>=4" + vbcrlf
		sqlStr = sqlStr + "	 and cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + "	 and regdate >= '" + CStr(FRectRegStart) + "'" + vbcrlf
		sqlStr = sqlStr + "	 and regdate < '" + CStr(FRectRegEnd) + "'" + vbcrlf
		sqlStr = sqlStr + "	group by userid" + vbcrlf
		sqlStr = sqlStr + "	) t" + vbcrlf
		sqlStr = sqlStr + "	where t.cnt >= 2" + vbcrlf
		sqlStr = sqlStr + ")"

		rsget.Open sqlStr,dbget,1

		Fsubtotalprice       = rsget("tsum")
		Fitemno       = rsget("onebuycnt")

		rsget.Close
	else

		sqlStr = "select count(m.userid) as buycnt, sum(m.subtotalprice) as tsum" + vbcrlf
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m" + vbcrlf
		sqlStr = sqlStr + " where m.userid <> ''" + vbcrlf
		sqlStr = sqlStr + " and m.ipkumdiv>=4" + vbcrlf
		sqlStr = sqlStr + " and m.cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + " and m.regdate >= '" + CStr(FRectRegStart) + "'" + vbcrlf
		sqlStr = sqlStr + " and m.regdate < '" + CStr(FRectRegEnd) + "'" + vbcrlf
		sqlStr = sqlStr + " and m.userid in" + vbcrlf
		sqlStr = sqlStr + "(" + vbcrlf
		sqlStr = sqlStr + "	select userid from" + vbcrlf
		sqlStr = sqlStr + "	(" + vbcrlf
		sqlStr = sqlStr + "	select userid, count(*) as cnt" + vbcrlf
		sqlStr = sqlStr + "	from [db_order].[dbo].tbl_order_master" + vbcrlf
		sqlStr = sqlStr + "	 where userid <> ''" + vbcrlf
		sqlStr = sqlStr + "	 and ipkumdiv>=4" + vbcrlf
		sqlStr = sqlStr + "	 and cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + "	 and regdate >= '" + CStr(FRectRegStart) + "'" + vbcrlf
		sqlStr = sqlStr + "	 and regdate < '" + CStr(FRectRegEnd) + "'" + vbcrlf
		sqlStr = sqlStr + "	group by userid" + vbcrlf
		sqlStr = sqlStr + "	) t" + vbcrlf
		sqlStr = sqlStr + "	where t.cnt = " + FRectBuyNum + "" + vbcrlf
		sqlStr = sqlStr + ")"

		rsget.Open sqlStr,dbget,1

		Fsubtotalprice       = rsget("tsum")
		Fitemno       = rsget("buycnt")

		rsget.Close
	end if

		sqlStr = "select count(*) as cnt, sum(t.cnt) as cnt2 from" + vbcrlf
		sqlStr = sqlStr + " (" + vbcrlf
		sqlStr = sqlStr + " select userid, count(*) as cnt" + vbcrlf
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master" + vbcrlf
		sqlStr = sqlStr + " where userid <> ''" + vbcrlf
		sqlStr = sqlStr + " and ipkumdiv>=4" + vbcrlf
		sqlStr = sqlStr + " and cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + " and convert(varchar(4),regdate,20) ='" + Cstr(FRectYYYY) + "'" + vbcrlf
		sqlStr = sqlStr + " group by userid" + vbcrlf
		sqlStr = sqlStr + " ) t" + vbcrlf
		sqlStr = sqlStr + " where t.cnt >= 1"

		rsget.Open sqlStr,dbget,1

		Fcnt       = rsget("cnt")
		Ftotalcnt       = rsget("cnt2")

		rsget.Close

	end sub
end class


Class CMailzineItem
	public fgubun
	public fmailergubun
	public fmailer_key_maeching
	Public Fidx
	Public Ftitle
	Public Fstartdate
	Public Fenddate
	Public Freenddate
	Public Ftotalcnt
	public fsubject
	public fopencnt
	public fclickcnt
	public fsuccesscnt

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CMailzine
	public frectmailergubun
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FRectIdx

    Private Sub Class_Initialize()
		redim FItemList(0)
		FPagesize=20
		FCurrpage=1
		FTotalCount=0
		FResultcount=0
		FScrollCount=10
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	'//admin/mailopen/mail_list.asp
	Public Sub GetMailingList()
		dim sql,i, sqlsearch
		
		if frectmailergubun <> "" then
			sqlsearch = sqlsearch & " and mailergubun = '"&frectmailergubun&"'" & vbcrlf
		end if
		
		sql ="select count(idx) as cnt" & vbcrlf
		sql = sql + " From [db_log].[dbo].tbl_mailing_data with (readuncommitted)" & vbcrlf
		sql = sql + " where isusing='Y' " & sqlsearch
		
		'response.write sql &"<Br>"
		rsget.open sql,dbget,1
			FTotalcount =rsget("cnt")
		rsget.close
		
		if FTotalcount < 1 then exit Sub
 
		sql ="select top " + CStr(Fpagesize*FCurrpage)
		sql = sql + " idx, title, startdate, enddate, reenddate, totalcnt, gubun, mailergubun, mailer_key_maeching, subject" & vbcrlf
		sql = sql + " , successcnt, opencnt, clickcnt,mailergubun " & vbcrlf
		sql = sql + " from [db_log].[dbo].tbl_mailing_data with (readuncommitted)" & vbcrlf
		sql = sql + " where isusing='Y' " & sqlsearch
		sql = sql + " order by idx desc" & vbcrlf

		'response.write sql &"<Br>"
		rsget.pagesize=Fpagesize
		rsget.open sql,dbget,1

		FResultcount =rsget.Recordcount-((FCurrpage-1)*FPageSize)

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		if not rsget.eof then
		i=0
		redim preserve FItemList(FResultcount)

		rsget.absolutepage = FCurrPage
		do until rsget.EOF
			set FItemList(i) = new CMailzineItem
			FItemList(i).fsubject = rsget("subject")	'제목
			FItemList(i).fsuccesscnt = rsget("successcnt") '성공발송통수
			FItemList(i).fopencnt = rsget("opencnt") '오픈통수
			FItemList(i).fclickcnt = rsget("clickcnt") '클릭수


			FItemList(i).fgubun = rsget("gubun")
			FItemList(i).fmailergubun = rsget("mailergubun")
			FItemList(i).fmailer_key_maeching = rsget("mailer_key_maeching")
			FItemList(i).Fidx = rsget("idx")
			FItemList(i).Ftitle = db2html(rsget("title"))
			FItemList(i).Fstartdate = db2html(rsget("startdate"))
			FItemList(i).Fenddate = db2html(rsget("enddate"))
			FItemList(i).Freenddate = db2html(rsget("reenddate"))
			FItemList(i).Ftotalcnt = rsget("totalcnt")
			FItemList(i).fmailergubun = rsget("mailergubun")

			rsget.movenext
			i=i+1
		loop
		end if
		rsget.Close
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

End Class

Class CMailzineOne
	Public Ftitle
	public fgubun
	Public Fstartdate
	Public Fenddate
	Public Freenddate
	Public Ftotalcnt
	Public Frealcnt
	Public Frealpct
	Public Ffilteringcnt
	Public Ffilteringpct
	Public Fsuccesscnt
	Public Fsuccesspct
	Public Ffailcnt
	Public Ffailpct
	Public Fopencnt
	Public Fopenpct

	''2016-12-07 유태욱 추가
	public Fclickcnt
	public Fclickpct
	public fsubject
	
	Public Fnoopencnt
	Public Fnoopenpct
	public fisusing
	public fmailergubun
	public fmailer_key_maeching
	
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	'//admin/mailopen/mail_edit.asp
	Public Sub GetMailingOne(byval idx)
		dim sql

		sql ="select top 1" & vbcrlf
		sql = sql + " idx, title, startdate, enddate, reenddate, totalcnt, realcnt, realpct, filteringcnt, filteringpct" & vbcrlf
		sql = sql + " , successcnt, successpct, failcnt, failpct, opencnt, openpct, noopencnt, noopenpct, isusing, gubun" & vbcrlf
		sql = sql + " , mailergubun, mailer_key_maeching, clickcnt, clickpct" & vbcrlf
		sql = sql + " from [db_log].[dbo].tbl_mailing_data with (readuncommitted)" & vbcrlf
		sql = sql + " where idx=" + CStr(idx)
		
		'response.write sql &"<br>"
		rsget.open sql,dbget,1

		if not rsget.eof then
			fisusing = rsget("isusing")
			Ftitle = db2html(rsget("title"))
			fgubun = rsget("gubun")
			Fstartdate = db2html(rsget("startdate"))
			Fenddate = db2html(rsget("enddate"))
			Freenddate = db2html(rsget("reenddate"))
			Ftotalcnt = rsget("totalcnt")
			Frealcnt = rsget("realcnt")
			Frealpct = rsget("realpct")
			Ffilteringcnt = rsget("filteringcnt")
			Ffilteringpct = rsget("filteringpct")
			Fsuccesscnt = rsget("successcnt")
			Fsuccesspct = rsget("successpct")
			Ffailcnt = rsget("failcnt")
			Ffailpct = rsget("failpct")
			Fopencnt = rsget("opencnt")
			Fopenpct = rsget("openpct")
			Fnoopencnt = rsget("noopencnt")
			Fnoopenpct = rsget("noopenpct")
			fmailergubun = rsget("mailergubun")
			fmailer_key_maeching = rsget("mailer_key_maeching")

			Fclickcnt = rsget("clickcnt")
			Fclickpct = rsget("clickpct")
		end if
		rsget.Close
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
End Class

'// 회원등급별 매출 현황
Class CUserLevelSellItem
	Public FUserLevel
	public FSellTotal
	Public FSellCount
	Public FSellAvr
    Public Funiqcnt

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class CUserLevelSell
	public FItemList()
	Public FRectOld
	Public FRectSdate
	Public FRectEdate
	public FResultCount
    public FRectMinusInc
	public FRectMakerid
    
	Private Sub Class_Initialize()
		FResultCount = 0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	Public Sub GetLevelList()
		dim sql, ix

		sql =	"select " &_
				" (case" &_
				"	when accountdiv='50' or accountdiv='51' then '50'" &_
				"	when isnull(userid,'')='' then '99'" &_
				"	else userlevel end) as userlevel" &_
				"	, sum(subtotalprice+isnull(miletotalprice,0)) as totalsum " &_
				"	, count(distinct orderserial) as cnt " &_
				"	, (sum(subtotalprice+isnull(miletotalprice,0)) / count(distinct orderserial)) as avrPrice " &_
                "   , sum(CASE WHEN isNULL(userid,'')='' then 1 else 0 end) +count(distinct userid) as uniqcnt "

		if FRectOld="Y" then
			sql = sql & "from db_log.dbo.tbl_old_order_master_2003 "
		else
			sql = sql & "from db_order.[dbo].tbl_order_master "
		end if

		sql = sql &	"where cancelyn = 'N' " 
		sql = sql &	"	and jumundiv not in (6) " ''-- 교환주문제외
		if (FRectMinusInc="plus") then
		    sql = sql &	"	and jumundiv<>'9'"
		elseif (FRectMinusInc="minus") then
		    sql = sql &	"	and jumundiv='9'"
		end if
		    
		sql = sql &	"	and ipkumdiv>=4 " 
		sql = sql &	"	and convert(varchar(10),regdate,21) between '" & FRectSdate & "' and '" & FRectEdate & "' " &_
				" group by" &_
				" (case" &_
				"	when accountdiv='50' or accountdiv='51' then '50'" &_
				"	when isnull(userid,'')='' then '99'" &_
				"	else userlevel end)" &_
				" order by userlevel "

		'response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FResultCount = rsget.Recordcount
		redim preserve FItemList(FResultcount)

		if not rsget.eof then
			ix=0
			do until rsget.EOF
				set FItemList(ix) = new CUserLevelSellItem
					FItemList(ix).FUserLevel	= rsget("userlevel")
					FItemList(ix).FSellTotal	= rsget("totalsum")
					FItemList(ix).FSellCount	= rsget("cnt")
					FItemList(ix).FSellAvr	= rsget("avrPrice")
					FItemList(ix).Funiqcnt  = rsget("uniqcnt")
				rsget.MoveNext
				ix=ix+1
			loop
		end if
		rsget.Close
	End Sub

	'회원등급별 매출현황(상품 필터 추가)
	Public Sub GetLevelListWithDetail()
		dim sql, ix

		sql =	"select " &_
				" (case" &_
				"	when m.accountdiv='50' or m.accountdiv='51' then '50'" &_
				"	when isnull(m.userid,'')='' then '99'" &_
				"	else m.userlevel end) as userlevel" &_
				"	, sum(d.itemcost * d.itemno) as totalsum " &_
				"	, count(distinct m.orderserial) as cnt " &_
				"	, (sum(d.itemcost * d.itemno) / count(distinct m.orderserial)) as avrPrice " &_
                "   , sum(CASE WHEN isNULL(m.userid,'')='' then 1 else 0 end) +count(distinct m.userid) as uniqcnt "

		if FRectOld="Y" then
			sql = sql & "from db_log.dbo.tbl_old_order_master_2003 as m "
			sql = sql & " join db_log.dbo.tbl_old_order_detail_2003 as d "
			sql = sql & " 	on m.orderserial=d.orderserial "
		else
			sql = sql & "from db_order.[dbo].tbl_order_master as m "
			sql = sql & " join db_order.dbo.tbl_order_detail as d "
			sql = sql & " 	on m.orderserial=d.orderserial "
		end if

		sql = sql &	"where m.cancelyn = 'N' " 
		sql = sql &	"	and m.jumundiv not in (6) " ''-- 교환주문제외
		if (FRectMinusInc="plus") then
		    sql = sql &	"	and m.jumundiv<>'9'"
		elseif (FRectMinusInc="minus") then
		    sql = sql &	"	and m.jumundiv='9'"
		end if
		if FRectMakerid<>"" then
			sql = sql & " 	and d.makerid='" & FRectMakerid & "' "
		end if
		sql = sql &	"	and d.cancelyn<>'Y' "
		sql = sql &	"	and d.itemid not in (0,100) "
		sql = sql &	"	and m.ipkumdiv>=4 " 
		sql = sql &	"	and convert(varchar(10),m.regdate,21) between '" & FRectSdate & "' and '" & FRectEdate & "' " &_
				" group by" &_
				" (case" &_
				"	when m.accountdiv='50' or m.accountdiv='51' then '50'" &_
				"	when isnull(m.userid,'')='' then '99'" &_
				"	else m.userlevel end)" &_
				" order by userlevel "

		rsget.open sql,dbget,1

		FResultCount = rsget.Recordcount
		redim preserve FItemList(FResultcount)

		if not rsget.eof then
			ix=0
			do until rsget.EOF
				set FItemList(ix) = new CUserLevelSellItem
					FItemList(ix).FUserLevel	= rsget("userlevel")
					FItemList(ix).FSellTotal	= rsget("totalsum")
					FItemList(ix).FSellCount	= rsget("cnt")
					FItemList(ix).FSellAvr	= rsget("avrPrice")
					FItemList(ix).Funiqcnt  = rsget("uniqcnt")
				rsget.MoveNext
				ix=ix+1
			loop
		end if
		rsget.Close
	End Sub

End Class

'// 회원등급별/월별 현황
Class CUserLevelMonthItem
	public FAxisDate
	Public FOrange
	Public FYellow
	Public FGreen
	Public FBlue
	Public FSilver
	Public FGold
	Public FStaff
	Public FTotal
	public fVVIP
	public fFAMILY
	public fBIZ
	public fyyyymm
	public fuserlevelname
	public fuserlevelcount
	public fANDPushexistscount
	public fiOSexistscount
	public fANDALLY
	public fANDALLN
	public fiOSALLY
	public fiOSALLN
	public fANDPushY
	public fANDPushN
	public fANDPushYSmsY
	public fiOSPushY
	public fiOSPushN
	public fiOSPushYSmsY
	public femailokY
	public femailokN
	public fsmsokY
	public fsmsokN
	public fUserHOLD_count

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class CUserLevelMonth
	public FOneItem
	public FItemList()
	Public FRectSdate
	Public FRectEdate
	public FResultCount
	public FTotalCount
	public FRectyyyymm
	public tendb

	Private Sub Class_Initialize()
		FResultCount = 0
		FTotalCount =0

		IF application("Svr_Info")="Dev" THEN
			tendb="tendb."
		end if
	End Sub
	Private Sub Class_Terminate()
	End Sub

	Public Sub GetLevelList()
		dim sql, ix

		sql =	"select " &_
				"	yyyymm " &_
				"	,isnull(sum(Case userlevel When 5 Then usercount end),0) as [Orange] " &_
				"	,isnull(sum(Case userlevel When 0 Then usercount end),0) as [Yellow] " &_
				"	,isnull(sum(Case userlevel When 1 Then usercount end),0) as [Green] " &_
				"	,isnull(sum(Case userlevel When 2 Then usercount end),0) as [Blue] " &_
				"	,isnull(sum(Case userlevel When 3 Then usercount end),0) as [Silver] " &_
				"	,isnull(sum(Case userlevel When 4 Then usercount end),0) as [Gold] " &_
				"	,isnull(sum(Case userlevel When 6 Then usercount end),0) as [VVIP] " &_
				"	,isnull(sum(Case userlevel When 7 Then usercount end),0) as [Staff] " &_
				"	,isnull(sum(Case userlevel When 8 Then usercount end),0) as [FAMILY] " &_
				"	,isnull(sum(Case userlevel When 9 Then usercount end),0) as [BIZ] " &_
				"from db_log.dbo.tbl_monthly_userlevel " &_
				"where yyyymm = '"& FRectyyyymm &"'" &_
				"group by yyyymm " &_
				"order by yyyymm"

		'response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FResultCount = rsget.Recordcount
		redim preserve FItemList(FResultcount)

		if not rsget.eof then
			ix=0
			do until rsget.EOF
				set FItemList(ix) = new CUserLevelMonthItem

				FItemList(ix).FAxisDate	= rsget("yyyymm")
				FItemList(ix).FOrange	= rsget("Orange")
				FItemList(ix).FYellow	= rsget("Yellow")
				FItemList(ix).FGreen	= rsget("Green")
				FItemList(ix).FBlue		= rsget("Blue")
				FItemList(ix).FSilver	= rsget("Silver")
				FItemList(ix).FGold		= rsget("Gold")
				FItemList(ix).fVVIP		= rsget("VVIP")
				FItemList(ix).FStaff	= rsget("Staff")
				FItemList(ix).fFAMILY	= rsget("FAMILY")
				FItemList(ix).fBIZ	= rsget("BIZ")
				FItemList(ix).FTotal	= rsget("Orange")+rsget("Yellow")+rsget("Green")+rsget("Blue")+rsget("Silver")+rsget("Gold")+rsget("VVIP")+rsget("Staff")+rsget("BIZ")

				rsget.MoveNext
				ix=ix+1
			loop
		end if
		rsget.Close
	End Sub

	' /admin/report/userLevelMonthReport.asp
	Public Sub GetLevelagreeList()
		dim sqlstr, ix, sqlsearch, sqlsearchPart

		if FRectyyyymm="" then exit Sub

		if FRectyyyymm<>"" then
			sqlsearchPart = sqlsearchPart & " and l.yyyymm='"& FRectyyyymm &"'" & vbcrlf
			sqlsearch = sqlsearch & " and datediff(mm,u.regdate,'"& FRectyyyymm &"-01')>=0 "
		end if

		'회원 푸시 모수 취합(최근사용 기기)
		sqlstr = "select T.userid, T.appKey, T.lastpushyn " & vbcrlf
		sqlstr = sqlstr & " into #tmppush" & vbcrlf
		sqlstr = sqlstr & " from ( " & vbcrlf
		sqlstr = sqlstr & " 	select r.userid, r.appKey, isNull(r.pushyn,'N') as lastpushyn, ROW_NUMBER() over(partition by r.userid order by r.lastupdate desc) as num " & vbcrlf
		sqlstr = sqlstr & " 	from "& tendb &"db_contents.dbo.tbl_app_regInfo as r with(noLock) " & vbcrlf
		sqlstr = sqlstr & " 	where isNull(r.userid,'')<>'' " & vbcrlf
		sqlstr = sqlstr & " 		and r.isusing='Y' " & vbcrlf
		sqlstr = sqlstr & " ) as T " & vbcrlf
		sqlstr = sqlstr & " where T.num=1 " & vbcrlf
		'response.write sqlstr & "<Br>"
		db3_dbget.execute sqlstr

		'결과 산출
		sqlstr = " select '"& FRectyyyymm &"' as yyyymm" & vbcrlf
		sqlstr = sqlstr & " ,(case when isnull(L.userlevel,0) = 0 then 'WHITE'" & vbcrlf
		sqlstr = sqlstr & " 	when isnull(L.userlevel,0) = 1 then 'RED'" & vbcrlf
		sqlstr = sqlstr & " 	when isnull(L.userlevel,0) = 2 then 'VIP'" & vbcrlf
		sqlstr = sqlstr & " 	when isnull(L.userlevel,0) = 3 then 'VIP GOLD'" & vbcrlf
		sqlstr = sqlstr & " 	when isnull(L.userlevel,0) = 4 then 'VVIP'" & vbcrlf
		sqlstr = sqlstr & " 	when isnull(L.userlevel,0) = 7 then 'Staff'" & vbcrlf
		sqlstr = sqlstr & " 	when isnull(L.userlevel,0) = 8 then 'Family'" & vbcrlf
		sqlstr = sqlstr & " 	when isnull(L.userlevel,0) = 9 then 'BIZ'" & vbcrlf
		sqlstr = sqlstr & " 	end) as userlevelname" & vbcrlf
		sqlstr = sqlstr & " , count(u.userid) as userlevelcount" & vbcrlf
		sqlstr = sqlstr & " , sum(case when p.appkey=6 then 1 else 0 end) as 'ANDPushexistscount'" & vbcrlf
		sqlstr = sqlstr & " , sum(case when p.appkey=5 then 1 else 0 end) as 'iOSexistscount'" & vbcrlf
		sqlstr = sqlstr & " , sum(case when p.appkey=6 and p.lastpushyn='Y' and u.emailok ='Y' and isNull(u.smsok,'N')='Y' then 1 else 0 end) as 'ANDALLY'" & vbcrlf
		sqlstr = sqlstr & " , sum(case when p.appkey=6 and p.lastpushyn<>'Y' and u.emailok ='N' and isNull(u.smsok,'N')='N' then 1 else 0 end) as 'ANDALLN'" & vbcrlf
		sqlstr = sqlstr & " , sum(case when p.appkey=5 and p.lastpushyn='Y' and u.emailok ='Y' and isNull(u.smsok,'N')='Y' then 1 else 0 end) as 'iOSALLY'" & vbcrlf
		sqlstr = sqlstr & " , sum(case when p.appkey=5 and p.lastpushyn<>'Y' and u.emailok ='N' and isNull(u.smsok,'N')='N' then 1 else 0 end) as 'iOSALLN'" & vbcrlf
		sqlstr = sqlstr & " , sum(case when p.appkey=6 and p.lastpushyn='Y' then 1 else 0 end) as 'ANDPushY'" & vbcrlf
		sqlstr = sqlstr & " , sum(case when p.appkey=6 and p.lastpushyn<>'Y' then 1 else 0 end) as 'ANDPushN'" & vbcrlf
		sqlstr = sqlstr & " , sum(case when p.appkey=6 and p.lastpushyn='Y' and isNull(u.smsok,'N')='Y' then 1 else 0 end) as 'ANDPushYSmsY'" & vbcrlf
		sqlstr = sqlstr & " , sum(case when p.appkey=5 and p.lastpushyn='Y' then 1 else 0 end) as 'iOSPushY'" & vbcrlf
		sqlstr = sqlstr & " , sum(case when p.appkey=5 and p.lastpushyn<>'Y' then 1 else 0 end) as 'iOSPushN'" & vbcrlf
		sqlstr = sqlstr & " , sum(case when p.appkey=5 and p.lastpushyn='Y' and isNull(u.smsok,'N')='Y' then 1 else 0 end) as 'iOSPushYSmsY'" & vbcrlf
		sqlstr = sqlstr & " , sum(case when u.emailok ='Y' then 1 else 0 end) as emailokY" & vbcrlf
		sqlstr = sqlstr & " , sum(case when u.emailok ='N' then 1 else 0 end) as emailokN" & vbcrlf
		sqlstr = sqlstr & " , sum(case when isNull(u.smsok,'N')='Y' then 1 else 0 end) as smsokY" & vbcrlf
		sqlstr = sqlstr & " , sum(case when isNull(u.smsok,'N')='N' then 1 else 0 end) as smsokN" & vbcrlf
		sqlstr = sqlstr & " from "& tendb &"db_user.dbo.tbl_user_n u with (nolock)" & vbcrlf
		sqlstr = sqlstr & " left join "& tendb &"[db_my10x10].[dbo].[tbl_my_userlevel_data] l with (nolock)" & vbcrlf
		sqlstr = sqlstr & " 	on l.userid = u.userid" & vbcrlf
		sqlstr = sqlstr & " 		and 1=1 " & sqlsearchPart
		sqlstr = sqlstr & " left join #tmppush as p" & vbcrlf
		sqlstr = sqlstr & " 	on u.userid = p.userid" & vbcrlf
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " group by isnull(l.userlevel,0)" & vbcrlf
		sqlstr = sqlstr & " ,(case when isnull(L.userlevel,0) = 0 then 'WHITE'" & vbcrlf
		sqlstr = sqlstr & " 	when isnull(L.userlevel,0) = 1 then 'RED'" & vbcrlf
		sqlstr = sqlstr & " 	when isnull(L.userlevel,0) = 2 then 'VIP'" & vbcrlf
		sqlstr = sqlstr & " 	when isnull(L.userlevel,0) = 3 then 'VIP GOLD'" & vbcrlf
		sqlstr = sqlstr & " 	when isnull(L.userlevel,0) = 4 then 'VVIP'" & vbcrlf
		sqlstr = sqlstr & " 	when isnull(L.userlevel,0) = 7 then 'Staff'" & vbcrlf
		sqlstr = sqlstr & " 	when isnull(L.userlevel,0) = 8 then 'Family'" & vbcrlf
		sqlstr = sqlstr & " 	when isnull(L.userlevel,0) = 9 then 'BIZ'" & vbcrlf
		sqlstr = sqlstr & " 	end)" & vbcrlf
		sqlstr = sqlstr & " order by isnull(l.userlevel,0) asc" & vbcrlf

		'response.write sqlstr & "<Br>"
		db3_rsget.open sqlstr,db3_dbget,1

		FResultCount = db3_rsget.Recordcount
		redim preserve FItemList(FResultcount)

		if not db3_rsget.eof then
			ix=0
			do until db3_rsget.EOF
				set FItemList(ix) = new CUserLevelMonthItem

				FItemList(ix).fyyyymm	= db3_rsget("yyyymm")
				FItemList(ix).fuserlevelname	= db3_rsget("userlevelname")
				FItemList(ix).fuserlevelcount	= db3_rsget("userlevelcount")
				FItemList(ix).fANDPushexistscount	= db3_rsget("ANDPushexistscount")
				FItemList(ix).fiOSexistscount	= db3_rsget("iOSexistscount")
				FItemList(ix).fANDALLY	= db3_rsget("ANDALLY")
				FItemList(ix).fANDALLN	= db3_rsget("ANDALLN")
				FItemList(ix).fiOSALLY	= db3_rsget("iOSALLY")
				FItemList(ix).fiOSALLN	= db3_rsget("iOSALLN")
				FItemList(ix).fANDPushY	= db3_rsget("ANDPushY")
				FItemList(ix).fANDPushN	= db3_rsget("ANDPushN")
				FItemList(ix).fANDPushYSmsY	= db3_rsget("ANDPushYSmsY")
				FItemList(ix).fiOSPushY	= db3_rsget("iOSPushY")
				FItemList(ix).fiOSPushN	= db3_rsget("iOSPushN")
				FItemList(ix).fiOSPushYSmsY	= db3_rsget("iOSPushYSmsY")
				FItemList(ix).femailokY	= db3_rsget("emailokY")
				FItemList(ix).femailokN	= db3_rsget("emailokN")
				FItemList(ix).fsmsokY	= db3_rsget("smsokY")
				FItemList(ix).fsmsokN	= db3_rsget("smsokN")

				db3_rsget.MoveNext
				ix=ix+1
			loop
		end if
		db3_rsget.Close
	End Sub

	' /admin/report/userLevelMonthReport.asp
	Public Sub GetNonMemeberPushAgreeList()
		dim sqlstr, ix, sqlsearch, sqlsearchPart

		if FRectyyyymm="" then exit Sub

		'비회원 푸시 모수 취합
		sqlstr = "	select 'Non Member' as userlevelname " & vbcrlf
		sqlstr = sqlstr & " 	,count(*) as userlevelcount " & vbcrlf
		sqlstr = sqlstr & " 	, sum(case when appkey=6 then 1 else 0 end) as ANDPushexistscount " & vbcrlf
		sqlstr = sqlstr & " 	, sum(case when appkey=5 then 1 else 0 end) as iOSexistscount " & vbcrlf
		sqlstr = sqlstr & " 	, sum(case when appkey=6 and isNull(pushyn,'N')='Y' then 1 else 0 end) as ANDPushY " & vbcrlf
		sqlstr = sqlstr & " 	, sum(case when appkey=6 and isNull(pushyn,'N')<>'Y' then 1 else 0 end) as ANDPushN " & vbcrlf
		sqlstr = sqlstr & " 	, sum(case when appkey=5 and isNull(pushyn,'N')='Y' then 1 else 0 end) as iOSPushY " & vbcrlf
		sqlstr = sqlstr & " 	, sum(case when appkey=5 and isNull(pushyn,'N')<>'Y' then 1 else 0 end) as iOSPushN " & vbcrlf
		sqlstr = sqlstr & " from "& tendb &"db_contents.dbo.tbl_app_regInfo as r with(noLock)  " & vbcrlf
		sqlstr = sqlstr & " where isNull(r.userid,'')=''  " & vbcrlf
		sqlstr = sqlstr & " 	and r.isusing='Y'  " & vbcrlf
		sqlstr = sqlstr & "		and datediff(mm,r.regdate,'"& FRectyyyymm &"-01')>=0 "
		db3_rsget.open sqlstr,db3_dbget,1

		FResultCount = db3_rsget.Recordcount
		redim preserve FItemList(FResultcount)

		if not db3_rsget.eof then
			ix=0
			do until db3_rsget.EOF
				set FItemList(ix) = new CUserLevelMonthItem

				FItemList(ix).fuserlevelname	= db3_rsget("userlevelname")
				FItemList(ix).fuserlevelcount	= db3_rsget("userlevelcount")
				FItemList(ix).fANDPushexistscount	= db3_rsget("ANDPushexistscount")
				FItemList(ix).fiOSexistscount	= db3_rsget("iOSexistscount")
				FItemList(ix).fANDPushY	= db3_rsget("ANDPushY")
				FItemList(ix).fANDPushN	= db3_rsget("ANDPushN")
				FItemList(ix).fiOSPushY	= db3_rsget("iOSPushY")
				FItemList(ix).fiOSPushN	= db3_rsget("iOSPushN")

				db3_rsget.MoveNext
				ix=ix+1
			loop
		end if
		db3_rsget.Close
	End Sub

	' /admin/report/userLevelMonthReport.asp
	public sub GetUserHOLD_count()
		dim SqlStr 

		sqlStr = "EXEC [db_user_Hold].[dbo].[usp_SCM_UserHOLD_count]"
		
		'Response.write sqlStr &"<br>"
		'//간단한 쿼리에는 이거
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.recordcount
		fresultcount = rsget.recordcount
		
		if not rsget.EOF then
			set FOneItem = new CUserLevelMonthItem	
			
            FOneItem.fUserHOLD_count = rsget("cnt")      
            					
		end if
		rsget.close
	end sub	

End Class
%>
