<%

class CDesignerJumunList

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

	 public Fsex
	 public Fselltotal2
	 public Fsellcnt2

	 public Fcash
	 public Fonlinecnt

	 public FSocname
	 public Fdpart
	 public Fdpartcount
	 public Fitemgubun

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
	Public Fitemserial_large
	Public FCateName
	Public FTcnt

	Public Fmiletotalprice
	Public Fspendmembership
	Public Ftencardspend
	Public Fallatdiscountprice
	public Fjumuntotalsum

	public Fbuytotal
	public Ftenbeasongcount

	public function GetChannelName()
		GetChannelName = "CH" + Fitemgubun
	end function

	public function GetChannelName_Kor()
		if Fitemgubun="10" then
			GetChannelName_Kor = "디자인문구/개인소품"
		elseif Fitemgubun="15" then
			GetChannelName_Kor = "인테리어/리빙데코"
		elseif Fitemgubun="20" then
			GetChannelName_Kor = "취미/여가"
		elseif Fitemgubun="25" then
			GetChannelName_Kor = "주방/욕실/바디"
		elseif Fitemgubun="30" then
			GetChannelName_Kor = "패션/의류"
		elseif Fitemgubun="32" then
			GetChannelName_Kor = "패션/잡화"
		elseif Fitemgubun="35" then
			GetChannelName_Kor = "쥬얼리"
		elseif Fitemgubun="40" then
			GetChannelName_Kor = "키덜트/얼리/취미"
		elseif Fitemgubun="45" then
			GetChannelName_Kor = "선물"
		elseif Fitemgubun="50" then
			GetChannelName_Kor = "플라워"
		elseif Fitemgubun="98" then
			GetChannelName_Kor = "기타"
		else
			GetChannelName_Kor = Fitemgubun
		end if
	end Function

	public function GetCAName()
		if Fitemserial_large="10" then
			GetCAName = "디자인문구/개인소품"
		elseif Fitemserial_large="15" then
			GetCAName = "인테리어/리빙데코"
		elseif Fitemserial_large="20" then
			GetCAName = "취미/여가"
		elseif Fitemserial_large="25" then
			GetCAName = "주방/욕실/생활"
		elseif Fitemserial_large="30" then
			GetCAName = "패션/잡화"
		elseif Fitemserial_large="35" then
			GetCAName = "쥬얼리"
		elseif Fitemserial_large="40" then
			GetCAName = "키덜트/얼리"
		elseif Fitemserial_large="45" then
			GetCAName = "선물"
		elseif Fitemserial_large="50" then
			GetCAName = "플라워"
		elseif Fitemserial_large="98" then
			GetCAName = "기타"
		else
			GetCAName = Fitemserial_large
		end if
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
		elseif Cstr(Faccountdiv) = 30 then
			JumunMethodName = "포인트"
		elseif Cstr(Faccountdiv) = 50 then
			JumunMethodName = "입점몰"
		elseif Cstr(Faccountdiv) = 80 then
			JumunMethodName = "All@"
		elseif Cstr(Faccountdiv) = 90 then
			JumunMethodName = "상품권"
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


    public FMtotalmoney
	public FMtotalsellcnt
    public FNtotalmoney
	public FNtotalsellcnt
    public FBtotalmoney
	public FBtotalsellcnt

	public FRectJoinMallNotInclude
	public FRectExtMallNotInclude
	public FRectPointNotInclude
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

	public FRectOldJumun
	public FRectDelNoSearch
	public FRectDateType
	Public FRectSort
	public FRectOrderBy
	Public FRectCnt
	Public FRectToDateGubun
    public FRectTenDeliverExists
    
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
		if (CLng(a)> CLng(b)) then
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
		db3_rsget.Open sqlStr, db3_dbget, 1

		FResultCount = db3_rsget.RecordCount
        redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList

			    FMasterItemList(i).Fsitename = db3_rsget("mastercode")
			    FMasterItemList(i).Fsocname = db3_rsget("lcount")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
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

		db3_rsget.Open sqlStr, db3_dbget, 1

		FResultCount = db3_rsget.RecordCount
        redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList

			   FMasterItemList(i).Fsitename = db3_rsget("mastercode")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")
				FMasterItemList(i).Flecturer = db3_rsget("lecturer")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
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

		db3_rsget.Open sqlStr, db3_dbget, 1

		FResultCount = db3_rsget.RecordCount
        redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList

			   FMasterItemList(i).Fsitename = db3_rsget("tcnt")
				FMasterItemList(i).Fsellcnt = db3_rsget("gcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
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
		db3_rsget.Open sqlStr, db3_dbget, 1

		FResultCount = db3_rsget.RecordCount
        redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList

			   FMasterItemList(i).Fsitename = db3_rsget("userid")
				FMasterItemList(i).Fsellcnt = db3_rsget("tcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
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
		db3_rsget.Open sqlStr, db3_dbget, 1

		FResultCount = db3_rsget.RecordCount
        redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = db3_rsget("yyyymm")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
	end Sub

	public Sub GetMonthlySeminarRoomReport
		dim i,sqlStr
		maxt = -1
    	maxc = -1

'		sqlStr = "select convert(varchar(7),m.regdate,20) as yyyymm," + vbcrlf
'		sqlStr = sqlStr + " sum(d.sellprice*d.itemno) as sumtotal, sum(d.itemno) as sellcnt" + vbcrlf
'		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m," + vbcrlf
'		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d" + vbcrlf
'		sqlStr = sqlStr + " where d.masteridx = m.idx" + vbcrlf
'		sqlStr = sqlStr + " and m.shopid='cafe003'" + vbcrlf
'		sqlStr = sqlStr + " and m.regdate>'2004-05-01'" + vbcrlf
'		sqlStr = sqlStr + " and m.cancelyn='N'" + vbcrlf
'		sqlStr = sqlStr + " and d.cancelyn<>'Y'" + vbcrlf
'		sqlStr = sqlStr + " group by convert(varchar(7),m.regdate,20)" + vbcrlf
'		sqlStr = sqlStr + " order by convert(varchar(7),m.regdate,20) desc"

		sqlStr = "select convert(varchar(7),regdate,20) as yyyymm," + vbcrlf
		sqlStr = sqlStr + " sum(realsum) as sumtotal, count(orderno) as sellcnt" + vbcrlf
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master" + vbcrlf
		sqlStr = sqlStr + " where shopid='cafe003'" + vbcrlf
		sqlStr = sqlStr + " and regdate>'2004-05-01'" + vbcrlf
		sqlStr = sqlStr + " and cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + " group by convert(varchar(7),regdate,20)" + vbcrlf
		sqlStr = sqlStr + " order by convert(varchar(7),regdate,20) desc"
'response.write sqlStr
'dbget.close()	:	response.End
		db3_rsget.Open sqlStr, db3_dbget, 1

		FResultCount = db3_rsget.RecordCount
        redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = db3_rsget("yyyymm")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
	end Sub

	public Sub GetDailySeminarRoomReport
		dim i,sqlStr
		maxt = -1
    	maxc = -1

'		sqlStr = "select convert(varchar(10),m.regdate,20) as yyyymm," + vbcrlf
'		sqlStr = sqlStr + " sum(d.sellprice*d.itemno) as sumtotal, sum(d.itemno) as sellcnt" + vbcrlf
'		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m," + vbcrlf
'		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d" + vbcrlf
'		sqlStr = sqlStr + " where d.masteridx = m.idx" + vbcrlf
'		sqlStr = sqlStr + " and m.shopid='cafe003'" + vbcrlf
'		sqlStr = sqlStr + " and convert(varchar(7),m.regdate,20)='" + FRectYYYY + "'" + vbcrlf
'		sqlStr = sqlStr + " and m.cancelyn='N'" + vbcrlf
'		sqlStr = sqlStr + " and d.cancelyn<>'Y'" + vbcrlf
'		sqlStr = sqlStr + " group by convert(varchar(10),m.regdate,20)" + vbcrlf
'		sqlStr = sqlStr + " order by convert(varchar(10),m.regdate,20) desc"

		sqlStr = "select convert(varchar(10),regdate,20) as yyyymm," + vbcrlf
		sqlStr = sqlStr + " sum(realsum) as sumtotal, count(orderno) as sellcnt" + vbcrlf
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master" + vbcrlf
		sqlStr = sqlStr + " where shopid='cafe003'" + vbcrlf
		sqlStr = sqlStr + " and convert(varchar(7),regdate,20)='" + FRectYYYY + "'" + vbcrlf
		sqlStr = sqlStr + " and cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + " group by convert(varchar(10),regdate,20)" + vbcrlf
		sqlStr = sqlStr + " order by convert(varchar(10),regdate,20) desc"

		db3_rsget.Open sqlStr, db3_dbget, 1

		FResultCount = db3_rsget.RecordCount
        redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = db3_rsget("yyyymm")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
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

		db3_rsget.Open sql,db3_dbget,1

		FResultCount = db3_rsget.RecordCount


	    redim preserve FMasterItemList(FResultCount)


		do until db3_rsget.eof

			set FMasterItemList(i) = new CDesignerJumunList
			FMasterItemList(i).Fmakerid     = db3_rsget("makerid")
			FMasterItemList(i).Fselltotal   = db3_rsget("sumtotal")
			FMasterItemList(i).Fbuytotal    = db3_rsget("buytotal")
			FMasterItemList(i).Fsellcnt     = db3_rsget("sellcnt")

			if Not IsNull(FMasterItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
			end if

			db3_rsget.MoveNext
			i = i + 1
		loop

		db3_rsget.close
	end sub

	public sub SearchSellrePort()

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
        
        if FRectOrdertype = "totalprice" then
			sql = sql + " order by sumtotal Desc"
		elseif FRectOrdertype = "totalgain" then
		    sql = sql + " order by sum(d.itemcost*d.itemno-d.buycash*d.itemno) Desc"
		else
			sql = sql + " order by sellcnt Desc"
		end if

		db3_rsget.Open sql,db3_dbget,1

		FResultCount = db3_rsget.RecordCount


	    redim preserve FMasterItemList(FResultCount)


		do until db3_rsget.eof

			set FMasterItemList(i) = new CDesignerJumunList
			FMasterItemList(i).Fmakerid     = db3_rsget("makerid")
			FMasterItemList(i).Fselltotal   = db3_rsget("sumtotal")
			FMasterItemList(i).Fbuytotal    = db3_rsget("buytotal")
			FMasterItemList(i).Fsellcnt     = db3_rsget("sellcnt")

			if Not IsNull(FMasterItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
			end if

			db3_rsget.MoveNext
			i = i + 1
		loop

		db3_rsget.close
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

		db3_rsget.Open sql,db3_dbget,1
		FMtotalmoney = db3_rsget("sumtotal")
		FMtotalsellcnt = db3_rsget("sellcnt")
		if IsNull(FMtotalmoney) then
			FMtotalmoney =0
		end if

		if IsNull(FMtotalsellcnt) then
			FMtotalsellcnt =0
		end if
		db3_rsget.close

	end sub

	public sub SearchMallSellrePort2()
   		Dim sql, i
		dim wheredetail

		if FRectExtMallNotInclude<>"" then
			wheredetail = " and jumundiv<>'5'"
		end if

		if FRectPointNotInclude<>"" then
			wheredetail = wheredetail + " and accountdiv<>'30'"
		end if

		sql = "select sum(subtotalprice) as sumtotal,"
		sql = sql + " count(orderserial) as sellcnt"
		sql = sql + " from [db_order].[dbo].tbl_order_master"
		sql = sql + " where (regdate >= '" & FRectFromDate & "') and (regdate<'" & FRectToDate & "')"
		sql = sql + " and cancelyn = 'N'"
		sql = sql + " and ipkumdiv>=4"
		sql = sql + wheredetail

		db3_rsget.Open sql,db3_dbget,1
		FMtotalmoney = db3_rsget("sumtotal")
		FMtotalsellcnt = db3_rsget("sellcnt")
		db3_rsget.close

   		maxt = -1
   		maxc = -1


		''#################################################
		''데이타.
		''#################################################

		sql = "select sum(subtotalprice) as sumtotal,"
		sql = sql + " count(orderserial) as sellcnt, (sitename + IsNull(rdsite,'')) as sitename"
		sql = sql + " from [db_order].[dbo].tbl_order_master"
		sql = sql + " where (regdate >= '" & FRectFromDate & "') and (regdate<'" & FRectToDate & "')"
		sql = sql + " and cancelyn = 'N'"
		sql = sql + " and ipkumdiv>=4"
		sql = sql + wheredetail
        sql = sql + " Group by (sitename + IsNull(rdsite,''))"
        sql = sql + " order by sumtotal Desc"


		db3_rsget.Open sql,db3_dbget,1

		FResultCount = db3_rsget.RecordCount


	    redim preserve FMasterItemList(FResultCount)


		do until db3_rsget.eof

				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = db3_rsget("sitename")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")


				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop

		db3_rsget.close
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

		db3_rsget.Open sql,db3_dbget,1
		db3_rsget.close

		db3_rsget.Open sql,db3_dbget,1

		FResultCount = db3_rsget.RecordCount


	    redim preserve FMasterItemList(FResultCount)

		FMtotalmoney = 0
		FMtotalsellcnt = 0

		do until db3_rsget.eof
			set FMasterItemList(i) = new CDesignerJumunList
			if db3_rsget("mwdiv") = "W" then
			    FMasterItemList(i).Fsitename = "특정"
			elseif db3_rsget("mwdiv") = "M" then
			    FMasterItemList(i).Fsitename = "매입"
			elseif db3_rsget("mwdiv") = "U" then
			    FMasterItemList(i).Fsitename = "업체"
			elseif db3_rsget("mwdiv") = "B" then
			    FMasterItemList(i).Fsitename = "배송료"
			else
			    FMasterItemList(i).Fsitename = db3_rsget("mwdiv")
			end if
			FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
			FMasterItemList(i).Fbuytotal  = db3_rsget("buytotal")
			FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")
            
            FMasterItemList(i).Fminustotal = db3_rsget("minustotal")
            FMasterItemList(i).Fminusbuytotal = db3_rsget("minusbuytotal")
            FMasterItemList(i).Fminuscount = db3_rsget("minuscount")
            
			FMtotalmoney = Cdbl(FMtotalmoney) + Cdbl(db3_rsget("sumtotal"))
			FMtotalsellcnt = Cdbl(FMtotalsellcnt) + Cdbl(db3_rsget("sellcnt"))

			if Not IsNull(FMasterItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
			end if

			db3_rsget.MoveNext
			i = i + 1
		loop

		db3_rsget.close
	end sub

	public sub SearchMallSellrePort4()
		Dim sql, i
		maxt = -1
   		maxc = -1

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

		if FRectJoinMallNotInclude<>"on" then
			sql = sql + " and m.sitename in ('10x10','tingmart')"
		end if

		if FRectExtMallNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'50'"
		end if

		if FRectPointNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'30'"
		end if

		sql = sql + " group by  convert(varchar(7),m.regdate,20)"
		sql = sql + " order by  convert(varchar(7),m.regdate,20) desc"
'response.write sql
		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = db3_rsget("yyyymm")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
	end sub

	public sub SearchMallSellrePort5()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select convert(varchar(10),m.regdate,20) as yyyymmdd, datepart(w,m.regdate) as dpart," + vbcrlf
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

		sql = sql + " group by  convert(varchar(10),m.regdate,20), datepart(w,m.regdate)"
		sql = sql + " order by  convert(varchar(10),m.regdate,20) desc"

		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = db3_rsget("yyyymmdd")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fbuytotal  = db3_rsget("buytotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")
				FMasterItemList(i).Fdpart = db3_rsget("dpart")

				FMasterItemList(i).Fminustotal = db3_rsget("minustotal")
				FMasterItemList(i).Fminuscount = db3_rsget("minuscount")

				FMasterItemList(i).Fmiletotalprice = db3_rsget("miletotalprice")
				FMasterItemList(i).Fspendmembership = db3_rsget("spendmembership")
				FMasterItemList(i).Ftencardspend = db3_rsget("tencardspend")
				FMasterItemList(i).Fallatdiscountprice = db3_rsget("allatdiscountprice")
                
                FMasterItemList(i).Ftenbeasongcount = db3_rsget("tenbeasongcount")
                
				if IsNULL(FMasterItemList(i).Fallatdiscountprice) then FMasterItemList(i).Fallatdiscountprice=0
				if IsNULL(FMasterItemList(i).Ftencardspend) then FMasterItemList(i).Ftencardspend=0

				FMasterItemList(i).Fjumuntotalsum = db3_rsget("totalsum")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
	end sub

	public sub SearchMallSellrePort6()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select pricegubun ="
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
		sql = sql + "	when m.subtotalprice>=150000 then 'A150000~'"
		sql = sql + " end"
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
		sql = sql + "	when m.subtotalprice>=150000 then 'A150000~'"
		sql = sql + " end"
		sql = sql + " order by m.pricegubun"

		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = db3_rsget("pricegubun")
				FMasterItemList(i).Fsellcnt = db3_rsget("cnt")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if
				FTotalsellcnt = FTotalsellcnt + FMasterItemList(i).Fsellcnt
				Ftotalmoney = Ftotalmoney + FMasterItemList(i).Fselltotal
				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
	end sub

	public sub SearchMallSellrePortChannel()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select itemserial_large as itemgubun,"
		If FRectToDateGubun="M" Then
		sql = sql + " convert(varchar(7),m.regdate,20) as yyyymmdd,"
		Else
		sql = sql + " convert(varchar(10),m.regdate,20) as yyyymmdd,"
		End If
		sql = sql + " sum(d.itemno*d.itemcost) as sumtotal,sum(d.itemno*d.buycash) as buytotal,"
		sql = sql + " count(d.itemno) as sellcnt"
		if FRectOldJumun="on" then
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m,"
			sql = sql + " [db_log].[dbo].tbl_old_order_detail_2003 d,"
		else
			sql = sql + " from [db_order].[dbo].tbl_order_master m, "
			sql = sql + " [db_order].[dbo].tbl_order_detail d,"
		end if
		sql = sql + " [db_item].[dbo].tbl_item i"

		sql = sql + " where m.regdate>='" + CStr(FRectFromDate) + "'"
		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'"
		sql = sql + " and m.orderserial=d.orderserial"
		sql = sql + " and d.itemid=i.itemid"
		sql = sql + " and i.itemid<>0"
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
		If FRectToDateGubun="M" Then
		sql = sql + " group by  i.itemserial_large, convert(varchar(7),m.regdate,20)"
		sql = sql + " order by  convert(varchar(7),m.regdate,20) desc, i.itemgubun asc"
		Else
		sql = sql + " group by  i.itemserial_large, convert(varchar(10),m.regdate,20)"
		sql = sql + " order by  convert(varchar(10),m.regdate,20) desc, i.itemgubun asc"
		End If
'response.write sql
		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fitemgubun = db3_rsget("itemgubun")
			    FMasterItemList(i).Fsitename = db3_rsget("yyyymmdd")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")
				'FMasterItemList(i).Fdpart = db3_rsget("dpart")

				FMasterItemList(i).Fbuytotal = db3_rsget("buytotal")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
	end sub

	public sub SearchMallSellrePortMonthlyChannel()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select i.itemserial_large as itemgubun, convert(varchar(7),m.regdate,20) as yyyymm,"
		sql = sql + " sum(d.itemno*d.itemcost) as sumtotal, sum(d.itemno*d.buycash) as buytotal,"
		sql = sql + " count(d.itemno) as sellcnt"
		if FRectOldJumun="on" then
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m,"
			sql = sql + " [db_log].[dbo].tbl_old_order_detail_2003 d"
		else
			sql = sql + " from [db_order].[dbo].tbl_order_master m, "
			sql = sql + " [db_order].[dbo].tbl_order_detail d"
		end if

		sql = sql + " left join [db_item].[dbo].tbl_item i on d.itemid=i.itemid"

		sql = sql + " where Year(m.regdate)='" + FRectYYYY + "'"
		sql = sql + " and Month(m.regdate)='" + FRectMM + "'"
		sql = sql + " and m.orderserial=d.orderserial"
		sql = sql + " and m.ipkumdiv>3"
		sql = sql + " and d.itemid<>0"
		sql = sql + " and m.cancelyn='N'"
		sql = sql + " and d.cancelyn<>'Y'"
		sql = sql + " and m.jumundiv<>9"

		if FRectJoinMallNotInclude<>"on" then
			sql = sql + " and m.sitename ='10x10'"
		end if

		if FRectExtMallNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'50'"
		end if

		if FRectPointNotInclude<>"on" then
			sql = sql + " and m.accountdiv<>'30'"
		end if

		sql = sql + " group by  i.itemserial_large, convert(varchar(7),m.regdate,20)"
		sql = sql + " order by  convert(varchar(7),m.regdate,20) desc, i.itemserial_large asc"

'response.write sql
		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fitemgubun = db3_rsget("itemgubun")
			    FMasterItemList(i).Fsitename = db3_rsget("yyyymm")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")

				FMasterItemList(i).Fbuytotal = db3_rsget("buytotal")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)

					FTotalPrice = FTotalPrice + FMasterItemList(i).Fselltotal
				end if


				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
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

		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
			set FMasterItemList(i) = new CDesignerJumunList

		    FMasterItemList(i).Fsitename = db3_rsget("yyyymmdd")
			FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
			FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")
			FMasterItemList(i).Fdpart = db3_rsget("dpart")

			if Not IsNull(FMasterItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
			end if

			db3_rsget.MoveNext
			i = i + 1
		loop
		db3_rsget.close
	end sub

	public sub SearchMallSellTimerePortChannel1()
		Dim sql, i
		maxt = -1
   		maxc = -1

'###############################################################################
'특정시간까지 판매량
'###############################################################################
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
		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList

				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
	end sub

	public sub SearchMallSellTimerePortChannel2()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select IsNull(i.itemserial_large,'90') as itemgubun, convert(varchar(10),m.regdate,20) as yyyymmdd,"
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

		sql = sql + " group by  i.itemserial_large, convert(varchar(10),m.regdate,20), datepart(w,m.regdate)"
		sql = sql + " order by  convert(varchar(10),m.regdate,20) desc, i.itemserial_large asc"

'response.write sql
		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fitemgubun = db3_rsget("itemgubun")
			    FMasterItemList(i).Fsitename = db3_rsget("yyyymmdd")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")
				FMasterItemList(i).Fdpart = db3_rsget("dpart")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
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

		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = CStr(db3_rsget("yyyy")) + "-" + CStr(db3_rsget("weekdt"))
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
	end sub


	public sub SearchMallSellrePort()

    Dim sql, i

    maxt = -1
    maxc = -1


		''#################################################
		''데이타.
		''#################################################

			sql = "select sum(subtotalprice) as sumtotal,"
			sql = sql + " count(orderserial) as sellcnt, sitename"
			sql = sql + " from [db_order].[dbo].tbl_order_master"
			sql = sql + " where (regdate >= '" & FRectFromDate & "') and (regdate<'" & FRectToDate & "')"
			sql = sql + " and cancelyn = 'N'"
			sql = sql + " and ipkumdiv>=4"
            sql = sql + " Group by sitename"
            sql = sql + " order by sumtotal Desc"


			db3_rsget.Open sql,db3_dbget,1

			FResultCount = db3_rsget.RecordCount


		    redim preserve FMasterItemList(FResultCount)


			do until db3_rsget.eof

				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = db3_rsget("sitename")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")


				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
			loop

			db3_rsget.close

	end sub

	public sub MemberBuySex()
		Dim sql, i
		sql = "select count(m.orderserial) as cnt, sum(subtotalprice) as sumprice,"
		sql = sql + " Left(Right(u.juminno,7),1) as sex"
		sql = sql + " from [db_order].[dbo].tbl_order_master m, [db_user].[dbo].tbl_user_n u"
		sql = sql + " where m.regdate >='" & FRectFromDate & "'"
        sql = sql + " and m.regdate < '" & FRectToDate & "'"
		sql = sql + " and m.sitename='10x10'"
		sql = sql + " and m.cancelyn='N'"
		sql = sql + " and m.userid=u.userid"
		sql = sql + " and m.userid <> ''"
		sql = sql + " and m.ipkumdiv>=4"
		sql = sql + " and m.jumundiv<9"
		sql = sql + " group by Left(Right(juminno,7),1)"

		db3_rsget.Open sql,db3_dbget,1
		do until db3_rsget.Eof
			if db3_rsget("sex")="1" then
				FManTotalMoney = db3_rsget("sumprice")
				FManTotalCount = db3_rsget("cnt")
			end if

			if db3_rsget("sex")="2" then
				FWoManTotalMoney = db3_rsget("sumprice")
				FWoManTotalCount = db3_rsget("cnt")
			end if

			db3_rsget.MoveNext
		loop

		db3_rsget.close


	end sub

	public sub MemberBuyPercent()

    	Dim sql, i

		''#################################################
		''총데이타
		''#################################################

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

			db3_rsget.Open sql,db3_dbget,1
			if not db3_rsget.Eof then
						Ftotalmoney = db3_rsget("sumprice")
						FTotalsellcnt = db3_rsget("cnt")
			end if
			db3_rsget.close

			if isNUll(Ftotalmoney) then Ftotalmoney =0
			if isNUll(FTotalsellcnt) then FTotalsellcnt =0

		''#################################################
		''신규회원데이타
		''#################################################

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


			db3_rsget.Open sql,db3_dbget,1
			if not db3_rsget.Eof then
						FNtotalmoney = db3_rsget("sumprice")
						FNtotalsellcnt = db3_rsget("cnt")
			end if
			db3_rsget.close

			if isNUll(FNtotalmoney) then FNtotalmoney =0
			if isNUll(FNtotalsellcnt) then FNtotalsellcnt =0

		''#################################################
		''비회원데이타
		''#################################################

			sql = "select count(orderserial) as cnt, sum(subtotalprice) as sumprice"
			sql = sql + " from [db_order].[dbo].tbl_order_master"
			sql = sql + " where regdate >='" & FRectFromDate & "'"
            sql = sql + " and regdate < '" & FRectToDate & "'"
			sql = sql + " and sitename='10x10'"
			sql = sql + " and cancelyn='N'"
			sql = sql + " and userid = ''"
			sql = sql + " and ipkumdiv>=4"
			sql = sql + " and jumundiv<9"

			db3_rsget.Open sql,db3_dbget,1
			if not db3_rsget.Eof then
				FBtotalmoney = db3_rsget("sumprice")
				FBTotalsellcnt = db3_rsget("cnt")
			end if
			db3_rsget.close

			if isNUll(FBtotalmoney) then FBtotalmoney =0
			if isNUll(FBTotalsellcnt) then FBTotalsellcnt =0

          FMtotalmoney = Ftotalmoney - FBtotalmoney - FNtotalmoney
          FMtotalsellcnt = Ftotalsellcnt - FBtotalsellcnt - FNtotalsellcnt

	end sub

	public sub MemberBuyPercent2()

    	Dim sql, i

		''#################################################
		''총데이타
		''#################################################

			sql = "select count(orderserial) as cnt, sum(subtotalprice) as sumprice"
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003"
			sql = sql + " where regdate >='" & FRectFromDate & "'"
            sql = sql + " and regdate < '" & FRectToDate & "'"
			sql = sql + " and sitename='10x10'"
			sql = sql + " and cancelyn='N'"
			sql = sql + " and ipkumdiv>=4"
			sql = sql + " and jumundiv<9"

			db3_rsget.Open sql,db3_dbget,1
						Ftotalmoney = db3_rsget("sumprice")
						FTotalsellcnt = db3_rsget("cnt")
			db3_rsget.close

		''#################################################
		''신규회원데이타
		''#################################################

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

			db3_rsget.Open sql,db3_dbget,1
						FNtotalmoney = db3_rsget("sumprice")
						FNtotalsellcnt = db3_rsget("cnt")
			db3_rsget.close

		''#################################################
		''비회원데이타
		''#################################################

			sql = "select count(orderserial) as cnt, sum(subtotalprice) as sumprice"
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003"
			sql = sql + " where regdate >='" & FRectFromDate & "'"
            sql = sql + " and regdate < '" & FRectToDate & "'"
			sql = sql + " and sitename='10x10'"
			sql = sql + " and cancelyn='N'"
			sql = sql + " and userid = ''"
			sql = sql + " and ipkumdiv>=4"
			sql = sql + " and jumundiv<9"

			db3_rsget.Open sql,db3_dbget,1
				FBtotalmoney = db3_rsget("sumprice")
				FBTotalsellcnt = db3_rsget("cnt")
			db3_rsget.close

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

		'if FRectDispY="on" then
		'	wheredetail = wheredetail + " and i.dispyn='Y'"
		'end if

		'if FRectSellY="on" then
		'	wheredetail = wheredetail + " and i.sellyn='Y'"
		'end if


		''#################################################
		''데이타.
		''#################################################


		sqlStr = "select top " + CStr(FPageSize)
		sqlStr = sqlStr + " sum(d.itemno) as sm ,sum(d.itemno*d.buycash)as sm2, d.buycash, d.itemcost, d.itemid,"
		sqlStr = sqlStr + " d.itemname, d.makerid, d.itemoptionname"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.ipkumdiv>=4"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " group by d.itemid, d.buycash, d.itemcost, d.itemname, d.makerid, d.itemoptionname"

		sqlStr = sqlStr + " order by sum(d.itemno*d.buycash) Desc"

		db3_rsget.Open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).FItemSellprice       = db3_rsget("sm2")
				FMasterItemList(i).FItemNo       = db3_rsget("sm")
				FMasterItemList(i).FItemID       = db3_rsget("itemid")
				FMasterItemList(i).FItemCost       = db3_rsget("itemcost")
				FMasterItemList(i).FItemName     = db2html(db3_rsget("itemname"))
				FMasterItemList(i).FItemOptionStr= db2html(db3_rsget("itemoptionname"))
				FMasterItemList(i).FBuycash		= db3_rsget("buycash")
				FMasterItemList(i).FMakerid		= db3_rsget("makerid")
				db3_rsget.movenext
				i=i+1
			loop
		db3_rsget.Close
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


		''#################################################
		''데이타.
		''#################################################


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


		db3_rsget.Open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).FItemNo       = db3_rsget("sm")
				FMasterItemList(i).FItemID       = db3_rsget("itemid")
				FMasterItemList(i).FItemCost       = db3_rsget("itemcost")
				FMasterItemList(i).FItemName     = db3_rsget("itemname")
				FMasterItemList(i).FItemOptionStr= db3_rsget("optname")
				FMasterItemList(i).FBuycash		= db3_rsget("buycash")
				FMasterItemList(i).FMakerid		= db3_rsget("makerid")
				db3_rsget.movenext
				i=i+1
			loop
		db3_rsget.Close
	end sub



	public sub ChannelBrandSellrePort()

        Dim sql, sqltmp, i, wheredetail

    	maxt = -1
    	maxc = -1

		if (FRectCD1<>"") then
		    wheredetail = wheredetail + " and i.itemserial_large='" + FRectCD1 + "'"
		end if


			sql = "select sum(d.itemcost*d.itemno) as sumtotal, sum(d.buycash*d.itemno) as buytotal, "
			sql = sql + " d.makerid," 
			sql = sql + " sum(d.itemno) as sellcnt"
		if FRectOldJumun="on" then
			sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m,"
			sql = sql + " [db_log].[dbo].tbl_old_order_detail_2003 d"
		else
			sql = sql + " from [db_order].[dbo].tbl_order_master m,"
			sql = sql + " [db_order].[dbo].tbl_order_detail d"
		end if
			sql = sql + " 	,[db_item].[dbo].tbl_item as i "
			sql = sql + " where m.orderserial = d.orderserial "
			sql = sql + " and d.itemid=i.itemid "
			sql = sql + " and (m.regdate >= '" & FRectFromDate & "') "
			sql = sql + " and (m.regdate < '" & FRectToDate & "')"
			sql = sql + " and d.itemid <> 0"
			sql = sql + " and m.cancelyn = 'N'"
			sql = sql + " and m.jumundiv<>9"
			sql = sql + " and d.cancelyn <> 'Y'"
			sql = sql + " and m.ipkumdiv>=4"
			sql = sql + wheredetail
            sql = sql + " Group by d.makerid" 
			if FRectOrdertype = "totalprice" then
				sql = sql + " order by sumtotal Desc"
			elseif FRectOrdertype = "totalgain" then
			    sql = sql + " order by sum(d.itemcost*d.itemno-d.buycash*d.itemno) Desc"
			else
				sql = sql + " order by sellcnt Desc"
			end if

			db3_rsget.Open sql,db3_dbget,1

				FResultCount = db3_rsget.RecordCount
		        redim preserve FMasterItemList(FResultCount)

				do until db3_rsget.eof

					set FMasterItemList(i) = new CDesignerJumunList
					FMasterItemList(i).Fmakerid = db3_rsget("makerid")
					FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
					FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")
					
					FMasterItemList(i).Fbuytotal = db3_rsget("buytotal")

					if Not IsNull(FMasterItemList(i).Fselltotal) then
						maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
						maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
					end if
					
					FTotalPrice = FTotalPrice + FMasterItemList(i).Fselltotal
					FTotalBuyPrice = FTotalBuyPrice + FMasterItemList(i).Fbuytotal
					
					FTotalEA = FTotalEA + FMasterItemList(i).Fsellcnt
					db3_rsget.MoveNext
					i = i + 1
				loop

			db3_rsget.close

	end sub

	public sub SearchCardOnline()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select convert(varchar(10),m.regdate,20) as yyyymmdd, datepart(w,m.regdate) as dpart, "
		sql = sql + " sum(m.subtotalprice) as sumtotal, count(m.idx) as sellcnt, accountdiv"
		sql = sql + " from [db_order].[dbo].tbl_order_master m"
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
		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = db3_rsget("yyyymmdd")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")
				FMasterItemList(i).Fdpart = db3_rsget("dpart")
				FMasterItemList(i).Faccountdiv = db3_rsget("accountdiv")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
	end sub

	public sub SearchCardOnlineMonth()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select convert(varchar(7),m.regdate,20) as yyyymm, sum(m.subtotalprice) as sumtotal, count(m.idx) as sellcnt,accountdiv"
		sql = sql + " from [db_order].[dbo].tbl_order_master m"
		sql = sql + " where m.regdate>='2002-10-01'"
'		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'"

		sql = sql + " and ipkumdiv>3"
		sql = sql + " and cancelyn='N'"

		if FRectOrdertype="on" then
			sql = sql + " and m.sitename in ('10x10','tingmart')"
		end if

		sql = sql + " group by  convert(varchar(7),m.regdate,20),accountdiv"
		sql = sql + " order by  convert(varchar(7),m.regdate,20) desc"
''response.write sql
		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = db3_rsget("yyyymm")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")
'				FMasterItemList(i).Fdpart = db3_rsget("dpart")
				FMasterItemList(i).Faccountdiv = db3_rsget("accountdiv")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
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

		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = db3_rsget("yyyymmdd")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")
				FMasterItemList(i).Fcash = db3_rsget("cash")
				FMasterItemList(i).Fonlinecnt = db3_rsget("onlinecnt")
				FMasterItemList(i).Fdpart = db3_rsget("dpart")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
					maxa = MaxVal(maxa,FMasterItemList(i).Fcash)
					maxb = MaxVal(maxb,FMasterItemList(i).Fonlinecnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
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

		''#################################################
		''데이타.
		''#################################################


		sqlStr = "select top 100  l.code_nm, sum(i.sellcash) as sellsum, sum(d.itemno) as sellcnt" + vbcrlf
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on d.itemid = i.itemid" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_large l on i.itemserial_large = l.code_large" + vbcrlf
		sqlStr = sqlStr + " where m.orderserial = d.orderserial" + vbcrlf
		sqlStr = sqlStr + " and m.ipkumdiv >= 4" + vbcrlf
		sqlStr = sqlStr + " and m.cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + " and d.cancelyn<>'Y'" + vbcrlf
		sqlStr = sqlStr + " and i.itemid <> 0" + vbcrlf
		sqlStr = sqlStr + wheredetail
		sqlStr = sqlStr + " group by l.code_nm" + vbcrlf
		sqlStr = sqlStr + " order by sum(d.itemno) desc"
'response.write sqlStr
		db3_rsget.Open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).Fcode_nm       = db3_rsget("code_nm")
				FMasterItemList(i).Fsubtotalprice       = db3_rsget("sellsum")
				FMasterItemList(i).Fitemno       = db3_rsget("sellcnt")
				db3_rsget.movenext
				i=i+1
			loop
		db3_rsget.Close
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
		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = db3_rsget("yyyymmdd")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")
				FMasterItemList(i).Fdpart = db3_rsget("dpart")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
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
		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList

				FMasterItemList(i).FDate = db3_rsget("yyyymmdd")
				FMasterItemList(i).FDayselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).FDaysellcnt = db3_rsget("sellcnt")

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
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
		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsex = db3_rsget("sex")
				FMasterItemList(i).Fsitename = db3_rsget("yyyymmdd")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")
				FMasterItemList(i).Fdpart = db3_rsget("dpart")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
	end sub

	public sub SearchTimeSellrePort()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select datepart(hh,m.regdate) as dpart," + vbcrlf
		sql = sql + " sum(m.subtotalprice) as sumtotal, count(m.orderserial) as sellcnt" + vbcrlf
		sql = sql + " from [db_order].[dbo].tbl_order_master m" + vbcrlf
		sql = sql + " where m.regdate>='" + CStr(FRectRegStart) + "'"
		sql = sql + " and m.regdate<'" + CStr(FRectRegEnd) + "'"
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " and ipkumdiv>3" + vbcrlf
		sql = sql + " and sitename <>'way2way'" + vbcrlf
		sql = sql + " group by datepart(hh,m.regdate)" + vbcrlf
		sql = sql + " order by datepart(hh,m.regdate) asc"
'response.write sql
		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList

				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")
				FMasterItemList(i).Fdpart = db3_rsget("dpart")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
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
		sql = sql + " and i.itemserial_large='" + FRectCD1 + "'" + vbcrlf
		If FRectCD2 <> "" Then
		sql = sql + " and i.itemserial_mid='" + FRectCD2 + "'" + vbcrlf
		End If
		If FRectCD3 <> "" Then
		sql = sql + " and i.itemserial_small='" + FRectCD3 + "'" + vbcrlf
		End If
		sql = sql + " group by  convert(varchar(10),m.regdate,20), datepart(w,m.regdate)"
		sql = sql + " order by  convert(varchar(10),m.regdate,20) desc"
'response.write sql
		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
			    FMasterItemList(i).Fsitename = db3_rsget("yyyymmdd")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")
				FMasterItemList(i).Fdpart = db3_rsget("dpart")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
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



		''#################################################
		''데이타.
		''#################################################


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

		db3_rsget.PageSize = FPageSize
		db3_rsget.Open sqlStr,db3_dbget,1

		FTotalCount = db3_rsget.RecordCount
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FMasterItemList(FResultCount)

		if not db3_rsget.EOF then
			db3_rsget.absolutepage = FCurrPage
			do until (i >= FResultCount)

				set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).FItemNo       = db3_rsget("sm")
				FMasterItemList(i).FItemID       = db3_rsget("itemid")
				FMasterItemList(i).FItemCost       = db3_rsget("itemcost")
				FMasterItemList(i).FItemName     = db2html(db3_rsget("itemname"))
				FMasterItemList(i).FItemOptionStr= db2html(db3_rsget("itemoptionname"))

				db3_rsget.movenext
				i=i+1
			loop
		end if
		db3_rsget.Close
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

		db3_rsget.Open sql,db3_dbget,1

		FTotalCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FTotalCount)

		do until db3_rsget.eof
				set FMasterItemList(ix) = new CDesignerJumunList
			    FMasterItemList(ix).FYYYYMMDDHHNNSS = db3_rsget("yyyymm")
				db3_rsget.MoveNext
				ix = ix + 1
		loop
		db3_rsget.close

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

		db3_rsget.Open sql,db3_dbget,1

		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")
				FMasterItemList(i).Fdpart = db3_rsget("dpart")
				FMasterItemList(i).Fdpartcount = db3_rsget("dpartcount")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close

	end Sub

	public Sub SearchBestsellerList()
		dim sqlStr
		dim i

		''#################################################
		''데이타.
		''#################################################

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
		db3_rsget.Open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
				FMasterItemList(i).FItemNo       = db3_rsget("sm")
				FMasterItemList(i).FItemID       = db3_rsget("itemid")
				FMasterItemList(i).FItemCost       = db3_rsget("itemcost")
				FMasterItemList(i).FItemName     = db2html(db3_rsget("itemname"))
				FMasterItemList(i).FItemOptionStr= db2html(db3_rsget("itemoptionname"))
				FMasterItemList(i).FBuycash		= db3_rsget("buycash")
				FMasterItemList(i).FMakerid		= db3_rsget("makerid")

				FMasterItemList(i).Fsmallimage  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FMasterItemList(i).FItemID) + "/" + db3_rsget("smallimage")
				db3_rsget.movenext
				i=i+1
			loop
		db3_rsget.Close
	end sub

	public Sub SearchNewItemReport()
		dim sqlStr
		dim i

		''#################################################
		''데이타.
		''#################################################
		
		
		sqlStr = " select itemserial_large,code_nm,count(A.itemid) as cnt "  
		sqlStr = sqlStr & " from [db_item].[dbo].tbl_item A " 
		sqlStr = sqlStr & " join db_item.[dbo].tbl_item_large B " 
		sqlStr = sqlStr & " 	on A.itemserial_large= b.code_large "  
		sqlStr = sqlStr & " where itemserial_large < 90  " 

		
		if (FRectRegStart<>"") then
			sqlStr = sqlStr + " and regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			sqlStr = sqlStr + " and regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectDesignerID<>"") then
			sqlStr = sqlStr + " and makerid = '" + FRectDesignerID + "'"
		end if

		sqlStr = sqlStr & " group by itemserial_large,code_nm "
		sqlStr = sqlStr + " order by itemserial_large asc" + vbcrlf


		db3_rsget.Open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CDesignerJumunList
					 FMasterItemList(i).Fitemserial_large  = db3_rsget("itemserial_large")
					 FMasterItemList(i).FCateName  = db3_rsget("code_nm")
					 FMasterItemList(i).FTcnt      = db3_rsget("cnt")
				db3_rsget.movenext
				i=i+1
			loop
		db3_rsget.Close
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
		if (CLng(a)> CLng(b)) then
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
		db3_rsget.Open sql,db3_dbget,1
		FResultCount = db3_rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until db3_rsget.eof
				set FMasterItemList(i) = new CReportMasterItem
			    FMasterItemList(i).Fselldate = db3_rsget("date")
				FMasterItemList(i).Fselltotal = db3_rsget("sumtotal")
				FMasterItemList(i).Fsellcnt = db3_rsget("sellcnt")
				FMasterItemList(i).Fitemid = db3_rsget("itemid")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
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
		''#################################################
		'' 첫 구매 데이타.
		''#################################################

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

		db3_rsget.Open sqlStr,db3_dbget,1

		Fsubtotalprice       = db3_rsget("tsum")
		Fitemno       = db3_rsget("onebuycnt")

		db3_rsget.Close
	else

		''#################################################
		'' 여러번 구매 데이타.
		''#################################################

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

		db3_rsget.Open sqlStr,db3_dbget,1

		Fsubtotalprice       = db3_rsget("tsum")
		Fitemno       = db3_rsget("buycnt")

		db3_rsget.Close
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

		db3_rsget.Open sqlStr,db3_dbget,1

		Fcnt       = db3_rsget("cnt")
		Ftotalcnt       = db3_rsget("cnt2")

		db3_rsget.Close

	end sub
end class

Class CMailzineItem

	Public Fidx
	Public Ftitle
	Public Fstartdate
	Public Fenddate
	Public Freenddate
	Public Ftotalcnt

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class



Class CMailzine

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

	Public Sub GetMailingList()

	dim sql,i

		sql ="select count(idx) as cnt" + vbcrlf
		sql = sql + " From [db_log].[dbo].tbl_mailing_data" + vbcrlf
		sql = sql + " where isusing='Y'" + vbcrlf

		db3_rsget.open sql,db3_dbget,1

		FTotalcount =db3_rsget("cnt")

		db3_rsget.close

		sql ="select top " + CStr(Fpagesize*FCurrpage) + " idx,title,startdate,enddate,reenddate,totalcnt" + vbcrlf
		sql = sql + " from [db_log].[dbo].tbl_mailing_data" + vbcrlf
		sql = sql + " where isusing='Y'" + vbcrlf
		sql = sql + " order by idx desc" + vbcrlf

		db3_rsget.pagesize=Fpagesize
		'response.write sql
		db3_rsget.open sql,db3_dbget,1

		FResultcount =db3_rsget.Recordcount-((FCurrpage-1)*FPageSize)

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		if not db3_rsget.eof then
		i=0
		redim preserve FItemList(FResultcount)

		db3_rsget.absolutepage = FCurrPage
		do until db3_rsget.EOF
			set FItemList(i) = new CMailzineItem

			FItemList(i).Fidx = db3_rsget("idx")
			FItemList(i).Ftitle = db3_rsget("title")
			FItemList(i).Fstartdate = db3_rsget("startdate")
			FItemList(i).Fenddate = db3_rsget("enddate")
			FItemList(i).Freenddate = db3_rsget("reenddate")
			FItemList(i).Ftotalcnt = db3_rsget("totalcnt")

			db3_rsget.movenext
			i=i+1

		loop
		end if
		db3_rsget.Close

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
	Public Fnoopencnt
	Public Fnoopenpct

   Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	Public Sub GetMailingOne(byval idx)

	dim sql

		sql ="select top 1 * from [db_log].[dbo].tbl_mailing_data" + vbcrlf
		sql = sql + " where idx=" + CStr(idx)

		db3_rsget.open sql,db3_dbget,1

		if not db3_rsget.eof then

			Ftitle = db3_rsget("title")
			fgubun = db3_rsget("gubun")
			Fstartdate = db3_rsget("startdate")
			Fenddate = db3_rsget("enddate")
			Freenddate = db3_rsget("reenddate")
			Ftotalcnt = db3_rsget("totalcnt")
			Frealcnt = db3_rsget("realcnt")
			Frealpct = db3_rsget("realpct")
			Ffilteringcnt = db3_rsget("filteringcnt")
			Ffilteringpct = db3_rsget("filteringpct")
			Fsuccesscnt = db3_rsget("successcnt")
			Fsuccesspct = db3_rsget("successpct")
			Ffailcnt = db3_rsget("failcnt")
			Ffailpct = db3_rsget("failpct")
			Fopencnt = db3_rsget("opencnt")
			Fopenpct = db3_rsget("openpct")
			Fnoopencnt = db3_rsget("noopencnt")
			Fnoopenpct = db3_rsget("noopenpct")

		end if
		db3_rsget.Close

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
%>
