<%

class CAcademyDiyBaljuMaster
	public FBaljuID
	public FBaljudate
	public Fdifferencekey
	public Fworkgroup

	public FCount
	public Fsongjanginputed

	public Fsongjangcnt

	public FsongjangDiv

	public Fcancelcnt
    public Fdelay0chulgocnt
    public Fdelay1chulgocnt
    public Fdelay2chulgocnt
    public Fdelay3chulgocnt

    public Fbaljutype
	Public FextSiteName

	public Fitemno

	Public Function GetExtSiteName()
		Select Case FextSiteName
			Case "10x10"
				GetExtSiteName = "텐바이텐"
			Case "cjmall"
				GetExtSiteName = "CJ몰"
			Case "interpark"
				GetExtSiteName = "인터파크"
			Case "lotteCom"
				GetExtSiteName = "롯데닷컴"
			Case "lotteimall"
				GetExtSiteName = "롯데i몰"
			Case "etcExtSite"
				GetExtSiteName = "기타제휴몰"
			Case Else
				GetExtSiteName = FextSiteName
		End Select
	end Function

	public function getBaljuTypeName()
		if IsNULL(Fbaljutype) then Exit function

        if (Fbaljutype="D") then
            getBaljuTypeName = "DAS"
        elseif (FsongjangDiv="S") then
            getBaljuTypeName = "단품"
        end if
    end function

    public function getDeliverName()
        if IsNULL(FsongjangDiv) then Exit function

		if (FsongjangDiv="2") then
			getDeliverName = "현대"
		elseif (FsongjangDiv="24") then
			getDeliverName = "사가와"
		elseif (FsongjangDiv="4") then
			getDeliverName = "CJ택배"
		elseif (FsongjangDiv="90") then
			getDeliverName = "EMS"
		elseif (FsongjangDiv="8") then
			getDeliverName = "우체국"
		end if
	end function

	public function GetTotalChulgoCount()
		GetTotalChulgoCount = Fdelay0chulgocnt + Fdelay1chulgocnt + Fdelay2chulgocnt + Fdelay3chulgocnt
	end function

	public function GetTenMiChulgoCount()
		GetTenMiChulgoCount = Fsongjangcnt - GetTotalChulgoCount - Fcancelcnt
	end function

	Private Sub Class_Initialize()
		'
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
end class

Class CAcademyDiyBaljuItem
	public Forderserial
	public Fjumundiv
	public Fuserid
	public Faccountname
	public Faccountdiv
	public Faccountno
	public Ftotalmileage
	public Ftotalsum
	public Fipkumdiv
	public Fipkumdate
	public Fregdate
	public Fbeadaldiv
	public Fbeadaldate
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
	public Fdeliverno
	public Fsitename
	public Fpaygatetid
	public Fdiscountrate
	public Fsubtotalprice
	public Fresultmsg
	public Frduserid
	public Fmilelogid
	public Fmiletotalprice
	public Fjungsanflag
	public Freqzipaddr
	public Fauthcode


	public Ftenbeaexists
    public FDlvcountryCode

	public function IpkumDivColor()
		if FjumunDiv="9" then
			IpkumDivColor = "#FF0000"
		else
			if Fipkumdiv="0" then
				IpkumDivColor="#FF0000"
			elseif Fipkumdiv="1" then
				IpkumDivColor="#FF0000"
			elseif Fipkumdiv="2" then
				IpkumDivColor="#000000"
			elseif Fipkumdiv="3" then
				IpkumDivColor="#000000"
			elseif Fipkumdiv="4" then
				IpkumDivColor="#0000FF"
			elseif Fipkumdiv="5" then
				IpkumDivColor="#444400"
			elseif Fipkumdiv="6" then
				IpkumDivColor="#FFFF00"
			elseif Fipkumdiv="7" then
				IpkumDivColor="#004444"
			elseif Fipkumdiv="8" then
				IpkumDivColor="#FF00FF"
			end if
		end if
	end function

	public function SiteNameColor()
		if Fsitename="uto" then
			SiteNameColor = "#55AA22"
		elseif Fsitename="cara" then
			SiteNameColor = "#225555"
		elseif Fsitename="emoden" then
			SiteNameColor = "#992255"
		elseif Fsitename="netian" then
			SiteNameColor = "#AA22AA"
		elseif Fsitename="miclub" then
			SiteNameColor = "#22AA22"
		else
			SiteNameColor = "#000000"
		end if
	end function

	public function SubTotalColor()
		if FSubtotalPrice<0 then
			SubTotalColor = "#DD3333"
		''elseif FSubtotalPrice>50000 then
			''	SubTotalColor = "#33AAAA"
		else
			SubTotalColor = "#000000"
		end if
	end function

	public function CancelYnColor()
		if FCancelYn="D" then
			CancelYnColor = "#FF0000"
		elseif UCase(FCancelYn)="Y" then
			CancelYnColor = "#FF0000"
		elseif FCancelYn="N" then
			CancelYnColor = "#000000"
		end if
	end function

	public function CancelYnName()
		if FCancelYn="D" then
			CancelYnName = "삭제"
		elseif UCase(FCancelYn)="Y" then
			CancelYnName = "취소"
		elseif FCancelYn="N" then
			CancelYnName = "정상"
		end if
	end function

	Public function JumunMethodName()
		if Faccountdiv="7" then
			JumunMethodName="무통장"
		elseif Faccountdiv="100" then
			JumunMethodName="신용카드"
		elseif Faccountdiv="20" then
			JumunMethodName="실시간이체"
		elseif Faccountdiv="30" then
			JumunMethodName="포인트"
		elseif Faccountdiv="50" then
			JumunMethodName="외부몰"
		elseif Faccountdiv="80" then
			JumunMethodName="All@카드"
		elseif Faccountdiv="90" then
			JumunMethodName="상품권결제"
		elseif Faccountdiv="110" then
			JumunMethodName="OK+신용"
		elseif Faccountdiv="400" then
			JumunMethodName="휴대폰"
		end if
	end function

	Public function IpkumDivName()
		if FjumunDiv="9" then
			IpkumDivName = "마이너스"
		else
			if Fipkumdiv="0" then
				IpkumDivName="주문대기"
			elseif Fipkumdiv="1" then
				IpkumDivName="주문실패"
			elseif Fipkumdiv="2" then
				IpkumDivName="주문접수"
			elseif Fipkumdiv="3" then
				IpkumDivName="주문접수"
			elseif Fipkumdiv="4" then
				IpkumDivName="결제완료"
			elseif Fipkumdiv="5" then
				IpkumDivName="주문통보"
			elseif Fipkumdiv="6" then
				IpkumDivName="상품준비"
			elseif Fipkumdiv="7" then
				IpkumDivName="일부출고"
			elseif Fipkumdiv="8" then
				IpkumDivName="상품출고"
			end if
		end if
	end function

	Private Sub Class_Initialize()
		'
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub

end Class

Class CAcademyDiyBalju
	public FItemList()

	public FLastQuery

	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FTotalCount
	public FSubTotalsum
	public FAvgTotalsum

	public FRectRegStart

	public FStartdate
	public FEndDate
	public FMaxcount

	public FBaljumasterList()
	public property Get resultBaljucount()
		resultBaljucount = ubound(FBaljumasterList)
	end property

	public Sub GetBaljuItemListProc()

		dim sqlStr,i,tmp

		response.write "시스템팀 문의"
		response.end

		'======================================================================
		''총 갯수. 총금액
		sqlStr = "selec11t count(m.orderserial) as cnt, sum(m.subtotalprice) as subtotal , avg(m.subtotalprice) as avgtotal " & vbcrlf
		sqlStr = sqlStr & "from [db_academy].[dbo].tbl_academy_order_master m " & vbcrlf
		sqlStr = sqlStr & "where 1 = 1 " & vbcrlf
		sqlStr = sqlStr + " and m.sitename <> 'academy' "
		sqlStr = sqlStr & "	and m.cancelyn = 'N' " & vbcrlf
		sqlStr = sqlStr & "	and m.ipkumdiv > 3 " & vbcrlf
		sqlStr = sqlStr & "	and m.baljudate is NULL " & vbcrlf
		sqlStr = sqlStr & "	and m.jumundiv <> 9 " & vbcrlf					'마이너스 주문 제외
		sqlStr = sqlStr & "	and m.jumundiv <> 4 " & vbcrlf					'티켓 주문 제외
		sqlStr = sqlStr & "	and m.jumundiv <> 6 " & vbcrlf					'교환 주문 제외
		sqlStr = sqlStr & "	and m.jumundiv <> 7 " & vbcrlf					'7:현장수령
		sqlStr = sqlStr & "	and m.regdate>'" & FRectRegStart & "' " & vbcrlf
		''response.write sqlStr

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")

			FSubtotalsum = rsACADEMYget("subtotal")
			FAvgTotalsum = rsACADEMYget("avgtotal")

			if IsNull(FSubtotalsum) then FSubtotalsum=0
			if IsNull(FAvgTotalsum) then FAvgTotalsum=0
		rsACADEMYget.Close


		'======================================================================
		'데이타
		sqlStr = "exec [db_academy].[dbo].[usp_ACA_DIY_MakeBaljuList] " + CStr(FPageSize) + ", '" + CStr(FRectRegStart) + "' "

		response.write sqlStr
		''response.end

		FLastQuery = sqlStr


		rsACADEMYget.CursorLocation = 3
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr, dbACADEMYget, 3, 1
		'rsACADEMYget.Open sqlStr, dbACADEMYget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

        FtotalPage =  CLng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		do until rsACADEMYget.eof
			set FItemList(i) = new CTenBaljuItem
			FItemList(i).Forderserial 	= rsACADEMYget("orderserial")
			FItemList(i).Fjumundiv	  	= rsACADEMYget("jumundiv")
			FItemList(i).Fuserid		= rsACADEMYget("userid")
			FItemList(i).Faccountdiv	= trim(rsACADEMYget("accountdiv"))
			FItemList(i).Ftotalsum		= rsACADEMYget("totalsum")
			FItemList(i).Fipkumdiv		= rsACADEMYget("ipkumdiv")
			FItemList(i).Fregdate		= rsACADEMYget("regdate")
			FItemList(i).Fcancelyn		= rsACADEMYget("cancelyn")
			FItemList(i).Fbuyname		= db2Html(rsACADEMYget("buyname"))
			FItemList(i).Freqname		= db2Html(rsACADEMYget("reqname"))
			FItemList(i).Fsitename		= rsACADEMYget("sitename")
			FItemList(i).Fsubtotalprice	= rsACADEMYget("subtotalprice")

            if (rsACADEMYget("TenbeaCnt")=0) then
				FItemList(i).Ftenbeaexists = false
			else
				FItemList(i).Ftenbeaexists = true
			end if

            FItemList(i).FDlvcountryCode = rsACADEMYget("DlvcountryCode")

			rsACADEMYget.movenext
			i=i+1
		loop
		rsACADEMYget.Close
	end Sub

	public sub getAcademyDiyBaljumaster
		dim sqlStr,i

		if (FStartdate<>"") and (FEnddate<>"") then
			sqlStr = "select top " + CStr(FMaxcount) + " m.id,m.baljudate,m.songjanginputed, m.differencekey, m.workgroup, m.songjangdiv, m.baljutype, IsNull(m.extSiteName, '') as extSiteName, count(d.orderserial) as cnt"
			sqlStr = sqlStr + " ,sum(case when baljusongjangno is null then 0 else 1 end) as songjangcnt "
			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.baljuflag=1) then 1 else 0 end) cancelcnt" + VbCrlf
			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<1) and (datediff(d,m.baljudate,d.chulgodate)>=0) then 1 else 0 end) delay0chulgocnt" + VbCrlf
			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<2) and (datediff(d,m.baljudate,d.chulgodate)>=1) then 1 else 0 end) delay1chulgocnt" + VbCrlf
			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<3) and (datediff(d,m.baljudate,d.chulgodate)>=2) then 1 else 0 end) delay2chulgocnt" + VbCrlf
			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)>=3) then 1 else 0 end) delay3chulgocnt" + VbCrlf
			sqlStr = sqlStr + " , IsNull(T.itemno,0) as itemno "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	[db_academy].[dbo].tbl_academy_baljumaster m "
			sqlStr = sqlStr + " 	join [db_academy].[dbo].tbl_academy_baljudetail d "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		m.id=d.baljuid "
			sqlStr = sqlStr + " 	left join ( "
			sqlStr = sqlStr + " 		select "
			sqlStr = sqlStr + " 			m.id as id2 "
			sqlStr = sqlStr + " 			, IsNull(sum(dd.itemno),0) as itemno "
			sqlStr = sqlStr + " 		from "
			sqlStr = sqlStr + " 			[db_academy].[dbo].tbl_academy_baljumaster m "
			sqlStr = sqlStr + " 			join [db_academy].[dbo].tbl_academy_baljudetail d "
			sqlStr = sqlStr + " 			on "
			sqlStr = sqlStr + " 				m.id=d.baljuid "
			sqlStr = sqlStr + " 			join [db_academy].[dbo].tbl_academy_order_detail dd "
			sqlStr = sqlStr + " 			on "
			sqlStr = sqlStr + " 				1 = 1 "
			sqlStr = sqlStr + " 				and dd.orderserial = d.orderserial "
			sqlStr = sqlStr + " 				AND dd.itemid <> 0 "
			sqlStr = sqlStr + " 				AND dd.cancelyn <> 'Y' "
			sqlStr = sqlStr + " 				AND (dd.isupchebeasong <> 'Y' OR m.songjangdiv = '90') "
			sqlStr = sqlStr + " 		where "
			sqlStr = sqlStr + " 			1 = 1 "
			sqlStr = sqlStr + " 			and m.baljudate >= '" + FStartdate + "' "
			sqlStr = sqlStr + " 			and m.baljudate < '" + FEnddate + "' "
			sqlStr = sqlStr + " 		group by "
			sqlStr = sqlStr + " 			m.id "
			sqlStr = sqlStr + " 	) T "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		T.id2 = m.id "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " and m.baljudate>='" + FStartdate + "'"
			sqlStr = sqlStr + " and m.baljudate<'" + FEnddate + "'"
			sqlStr = sqlStr + " group by m.id,m.baljudate,m.songjanginputed, m.differencekey, m.workgroup, m.songjangdiv, m.baljutype, IsNull(m.extSiteName, ''), IsNull(T.itemno,0) "
			sqlStr = sqlStr + " order by m.id desc"
		else
			sqlStr = "select top " + CStr(10) + " m.id,m.baljudate,m.songjanginputed, m.differencekey, m.workgroup, m.songjangdiv, m.baljutype, IsNull(m.extSiteName, '') as extSiteName, count(d.orderserial) as cnt, 0 as songjangcnt"
			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.baljuflag=1) then 1 else 0 end) cancelcnt" + VbCrlf
			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<1) and (datediff(d,m.baljudate,d.chulgodate)>=0) then 1 else 0 end) delay0chulgocnt" + VbCrlf
			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<2) and (datediff(d,m.baljudate,d.chulgodate)>=1) then 1 else 0 end) delay1chulgocnt" + VbCrlf
			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<3) and (datediff(d,m.baljudate,d.chulgodate)>=2) then 1 else 0 end) delay2chulgocnt" + VbCrlf
			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)>=3) then 1 else 0 end) delay3chulgocnt" + VbCrlf
			sqlStr = sqlStr + " , 0 as itemno"
			sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_baljumaster	m,"
			sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_baljudetail d"
			sqlStr = sqlStr + " where m.id=d.baljuid"
			sqlStr = sqlStr + " group by m.id,m.baljudate,m.songjanginputed, m.differencekey, m.workgroup, m.songjangdiv, m.baljutype, IsNull(m.extSiteName, '')"
			sqlStr = sqlStr + " order by id desc"
		end if

		''response.write sqlStr
		''response.end

		''rsACADEMYget.Open sqlStr,dbget,1
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly


		redim preserve FBaljumasterList(rsACADEMYget.RecordCount)
		i=0
		do until rsACADEMYget.Eof
			set FBaljumasterList(i) = new CAcademyDiyBaljuMaster
			FBaljumasterList(i).FBaljuID = rsACADEMYget("id")
			FBaljumasterList(i).FBaljudate = rsACADEMYget("baljudate")
			FBaljumasterList(i).FCount = rsACADEMYget("cnt")
			FBaljumasterList(i).Fsongjangcnt = rsACADEMYget("songjangcnt")
			FBaljumasterList(i).Fsongjanginputed = rsACADEMYget("songjanginputed")

			FBaljumasterList(i).Fdifferencekey = rsACADEMYget("differencekey")
			FBaljumasterList(i).Fworkgroup = rsACADEMYget("workgroup")
			FBaljumasterList(i).FsongjangDiv = rsACADEMYget("songjangdiv")

			FBaljumasterList(i).Fcancelcnt = rsACADEMYget("cancelcnt")
			FBaljumasterList(i).Fdelay0chulgocnt = rsACADEMYget("delay0chulgocnt")
			FBaljumasterList(i).Fdelay1chulgocnt = rsACADEMYget("delay1chulgocnt")
			FBaljumasterList(i).Fdelay2chulgocnt = rsACADEMYget("delay2chulgocnt")
			FBaljumasterList(i).Fdelay3chulgocnt = rsACADEMYget("delay3chulgocnt")

			FBaljumasterList(i).Fbaljutype = rsACADEMYget("baljutype")

			FBaljumasterList(i).FextSiteName = rsACADEMYget("extSiteName")

			FBaljumasterList(i).Fitemno = rsACADEMYget("itemno")

			i=i+1
			rsACADEMYget.MoveNext
		loop
		rsACADEMYget.close
	end sub

	Private Sub Class_Initialize()
		redim FItemList(0)
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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
end Class

%>
