<%
'###########################################################
' Description :  텐바이텐 매출통계
' History : 2007.12.06 한용민 생성
'           2008.03.13 허진원 - 고객등급별 상세통계 추가
'###########################################################

class Cmaechul_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public forderdate		'주문일
	public fipkumdate		'입금일
	public fcanceldate		'취소일
	public fjumundiv		'주문구분
	public faccountdiv		'결제구분
	public fsitename		'사이트구분
	public frdsite			'헤더사이트
	public ftotalsum		'총금액
	public ftotalcount		'총건수
	public fsubtotalprice	'실금액
	public ftotalbuysum		'매입가
	public fspendScoupon	'쿠폰
	public fspendMileage	'마일리지
	public fdiscountEtc		'기타할인
	public fspendIcoupon	'상품쿠폰
	public ftendeliverCount	'텐바이텐 배송수
	public ftendeliversum	'택배비(ftendeliverCount*2500원)
	public fsunsuik			'순수익
	public fmagin			'마진율
	public fuserlevel		'회원등급명
end class

class Cmaechul_list
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public flist

	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	public FRectStartdate
	public FRectEndDate
	public frectdatecancle
	public frectbancancle
	public frectaccountdiv
	public frectsitename
	public frectipkumdatesucc
	public fmonth
	public fmonthday
	public FRectInc3pl          '' 3pl매출 포함여부
	public FRectItemID
	public FRectSDate
	public FRectEDate
	public FNaItemName
	public FArrJust1Day
	public FMakerid
	public FArrEpNotMakerid
    public FArrEpNotItemid
    public FRectrdsellsum 
    public FRectNotIncOutmall
    
	public function fmaechul_graph		'월별 통계그래프용
	dim i , sql

	sql = "select"
		if dateview1 = "yes" then
			sql = sql & " convert(varchar(7),orderdate) as orderdate,"
		elseif dateview1 = "no" then
			sql = sql & " convert(varchar(7),ipkumdate) as orderdate,"			
		end if
			
		if frectdatecancle <> "" then
			sql = sql & " canceldate,"
		end if				
	sql = sql & " sum(totalcount) as totalcount,"	
	sql = sql & " sum(subtotalprice) as subtotalprice,"	
	sql = sql & " (sum(subtotalprice)-(sum(totalbuysum)+sum(tendeliverCount*" & chkIIF(date()>="2019-01-01","2500","2000") & "))) as sunsuik"
	sql = sql & " from db_datamart.dbo.tbl_mkt_daily_totalsale" 
	sql = sql & " where 1=1"
	
		if frectsitename <> "" then
			sql = sql & " and sitename = '" & frectsitename & "'"
		end if	
		if frectaccountdiv <> "" then
			sql = sql & " and accountdiv = '" & frectaccountdiv & "'"
		end if	
		if dateview1 = "yes" then
			sql = sql & " and convert(varchar(4),orderdate) = '" & FRectStartdate & "'" 
		elseif dateview1 = "no" then
			sql = sql & " and convert(varchar(4),ipkumdate) = '" & FRectStartdate & "'" 			
		end if	
		if frectdatecancle <> "" then
			sql = sql & " and canceldate is not null"
		end if
		if frectbancancle = "1" then				 
		elseif frectbancancle = "2" then
			sql = sql & " and jumundiv = '9'" 
		else
			sql = sql & " and jumundiv <> '9'"
		end if
		if frectipkumdatesucc = "" then
			sql = sql & " and ipkumdate is not null" 
		end if				 
	sql = sql & " group by"
		if dateview1 = "yes" then
			sql = sql & " convert(varchar(7),orderdate)"
		elseif dateview1 = "no" then
			sql = sql & " convert(varchar(7),ipkumdate)"		
		end if
			
		if frectdatecancle <> "" then
			sql = sql & " ,canceldate"
		end if
			
	sql = sql & " having sum(totalsum) is not null" 
	sql = sql & " order by"
		if dateview1 = "yes" then
			sql = sql & " convert(varchar(7),orderdate)"
		elseif dateview1 = "no" then
			sql = sql & " convert(varchar(7),ipkumdate)"		
		end if		
			
	db3_rsget.open sql,db3_dbget,1	
	'response.write sql&"<br>"
	
	FTotalCount = db3_rsget.recordcount
	redim flist(FTotalCount)
	i = 0
	if not db3_rsget.eof then
		do until db3_rsget.eof
			set flist(i) = new Cmaechul_oneitem
			
				flist(i).forderdate = db3_rsget("orderdate")			
				flist(i).ftotalcount = db3_rsget("totalcount")
				flist(i).fsunsuik = db3_rsget("sunsuik")
				flist(i).fsubtotalprice = db3_rsget("subtotalprice")									
		db3_rsget.movenext
		i = i + 1
		loop	
	end if
		
	db3_rsget.close		
	end function	


	public function fmaechul_graph_new		'월별 통계그래프용
	dim i , sql

	sql = "select"
		if dateview1 = "yes" then
			If fmonthday = "m" Then
				sql = sql & " convert(varchar(7),s.orderdate) as orderdate,"
			Else
				sql = sql & " convert(varchar(10),s.orderdate) as orderdate,"
			End If
		elseif dateview1 = "no" then
			sql = sql & " convert(varchar(7),s.ipkumdate) as orderdate,"			
		end if
			
		if frectdatecancle <> "" then
			sql = sql & " s.canceldate,"
		end if				
	sql = sql & " sum(s.totalcount) as totalcount,"	
	sql = sql & " sum(s.subtotalprice) as subtotalprice,"	
	sql = sql & " (sum(s.subtotalprice)-(sum(s.totalbuysum)+sum(s.tendeliverCount*" & chkIIF(date()>="2019-01-01","2500","2000") & "))) as sunsuik"
	sql = sql & " from db_datamart.dbo.tbl_mkt_daily_totalsale s" 
	sql = sql & "       left join db_partner.dbo.tbl_partner p"
	sql = sql & "       on s.sitename=p.id "
	sql = sql & " where 1=1"
	    if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sql = sql & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sql = sql & " and isNULL(p.tplcompanyid,'')=''"
	    end if
	    
		if frectsitename <> "" then
			sql = sql & " and s.sitename = '" & frectsitename & "'"
		end if	
		if frectaccountdiv <> "" then
			sql = sql & " and s.accountdiv = '" & frectaccountdiv & "'"
		end if	
		if dateview1 = "yes" then
			If fmonthday = "m" Then
				sql = sql & " and convert(varchar(4),s.orderdate) >= '" & FRectStartdate & "'" 
			Else
				sql = sql & " and convert(varchar(4),s.orderdate) >= '" & FRectStartdate & "' and right(convert(varchar(7),s.orderdate),2) = '" & fmonth & "'" 
			End If
		elseif dateview1 = "no" then
			sql = sql & " and convert(varchar(4),s.ipkumdate) >= '" & FRectStartdate & "'" 			
		end if	
		if frectdatecancle <> "" then
			sql = sql & " and s.canceldate is not null"
		end if
		if frectbancancle = "1" then				 
		elseif frectbancancle = "2" then
			sql = sql & " and s.jumundiv = '9'" 
		else
			sql = sql & " and s.jumundiv <> '9'"
		end if
		if frectipkumdatesucc = "" then
			sql = sql & " and s.ipkumdate is not null" 
		end if				 
	sql = sql & " group by"
		if dateview1 = "yes" then
			If fmonthday = "m" Then
				sql = sql & " convert(varchar(7),s.orderdate)"
			Else
				sql = sql & " convert(varchar(10),s.orderdate)"
			End If
		elseif dateview1 = "no" then
			sql = sql & " convert(varchar(7),s.ipkumdate)"		
		end if
			
		if frectdatecancle <> "" then
			sql = sql & " ,canceldate"
		end if
			
	sql = sql & " having sum(s.totalsum) is not null" 
	sql = sql & " order by"
		if dateview1 = "yes" then
			If fmonthday = "m" Then
				sql = sql & " convert(varchar(7),s.orderdate)"
			Else
				sql = sql & " convert(varchar(10),s.orderdate)"
			End If
		elseif dateview1 = "no" then
			sql = sql & " convert(varchar(7),s.ipkumdate)"		
		end if		
			
	db3_rsget.open sql,db3_dbget,1	
	'response.write sql&"<br>"
	
	FTotalCount = db3_rsget.recordcount
	redim flist(FTotalCount)
	i = 0
	if not db3_rsget.eof then
		do until db3_rsget.eof
			set flist(i) = new Cmaechul_oneitem
			
				flist(i).forderdate = db3_rsget("orderdate")			
				flist(i).ftotalcount = db3_rsget("totalcount")
				flist(i).fsunsuik = db3_rsget("sunsuik")
				flist(i).fsubtotalprice = db3_rsget("subtotalprice")									
		db3_rsget.movenext
		i = i + 1
		loop	
	end if
		
	db3_rsget.close		
	end function	


	Public Function fnNaverMaechulByItem
		Dim sqlStr, i, addSql, orderbysql

		sqlStr = "select itemname,makerid from [db_analyze_data_raw].dbo.tbl_item where itemid = '" & FRectItemID & "'"
		rsAnalget.CursorLocation = adUseClient
		rsAnalget.Open sqlStr,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.eof Then
			FNaItemName = rsAnalget("itemname")
			FMakerid    = rsAnalget("makerid")
		End If
		rsAnalget.close
		
		if (FMakerid<>"") then
		    sqlStr = "select top 10 makerid,mallgubun,isusing,regdate,lastupdate,regid,updateid from [DBAPPWISH].db_outmall.dbo.tbl_EpShop_not_in_makerid where makerid='"&FMakerid&"'" '' and mallgubun='naverep'
		    rsAnalget.CursorLocation = adUseClient
		    rsAnalget.Open sqlStr,dbAnalget,adOpenForwardOnly,adLockReadOnly
		    If not rsAnalget.eof Then
    			FArrEpNotMakerid = rsAnalget.getRows()
    		End If
    		rsAnalget.close
        end if
    
        sqlStr = "select top 10 itemid,mallgubun,isusing,regdate,lastupdate,regid,updateid from [DBAPPWISH].db_outmall.dbo.tbl_EpShop_not_in_itemid where itemid="&FRectItemID  ''mallgubun='naverEP'"
        rsAnalget.CursorLocation = adUseClient
	    rsAnalget.Open sqlStr,dbAnalget,adOpenForwardOnly,adLockReadOnly
	    If not rsAnalget.eof Then
			FArrEpNotItemid = rsAnalget.getRows()
		End If
		rsAnalget.close

		sqlStr = ""
		sqlStr = sqlStr & "declare @sdate datetime " & vbCrLf
		sqlStr = sqlStr & "declare @edate datetime " & vbCrLf

		sqlStr = sqlStr & "set @sdate = '" & FRectSDate & " 00:00:00' " & vbCrLf
		sqlStr = sqlStr & "set @edate = '" & FRectEDate & " 23:59:59' " & vbCrLf

		sqlStr = sqlStr & "select " & vbCrLf
		sqlStr = sqlStr & "	N.itemid, DT.solar_date as yyyymm, DT.weekname, " & vbCrLf
		sqlStr = sqlStr & "	(CASE WHEN N.myrank>=1000 THEN 101 ELSE N.myrank END) myrank, isNULL(T.sellcnt,0) sellcnt ,isNULL(T.[NP_DAUM],0) [NP_DAUM_sellcNT] " & vbCrLf
		sqlStr = sqlStr & "from [db_analyze_etc].[dbo].[LunarToSolar] DT " & vbCrLf
		sqlStr = sqlStr & "left join [db_analyze_etc].[dbo].[tbl_naver_low_master] N on DT.solar_date=convert(Varchar(10),N.regdate,21) and N.itemid='" & FRectItemID & "' " & vbCrLf
		sqlStr = sqlStr & "left join  " & vbCrLf
		sqlStr = sqlStr & "( " & vbCrLf
		sqlStr = sqlStr & "			select " & vbCrLf
		sqlStr = sqlStr & "				itemid,convert(varchar(10),regdate,21) as yyyymm "
		
		
		if (FRectrdsellsum="I") then
		    ''상품수량
    		sqlStr = sqlStr & "				, sum(d.itemno) sellcnt " & vbCrLf
    		sqlStr = sqlStr & "				, sum(CASE WHEN T.gubun is Not NULL THEN d.itemno END) as [NP_DAUM] " & vbCrLf
        else
            ''주문건수
		    sqlStr = sqlStr & "				, count(distinct m.orderserial) sellcnt " & vbCrLf
		    sqlStr = sqlStr & "				, count(distinct (CASE WHEN T.gubun is Not NULL   THEN m.orderserial END)) as [NP_DAUM] " & vbCrLf
		end if
		
		
		sqlStr = sqlStr & "			from [db_analyze_data_raw].dbo.tbl_order_master m " & vbCrLf
		sqlStr = sqlStr & "			join [db_analyze_data_raw].dbo.tbl_order_detail d on m.orderserial=d.orderserial " & vbCrLf
		sqlStr = sqlStr & "			left join [db_analyze_etc].[dbo].[tbl_Outmall_RdsiteGubun_NVDAUM] T on m.rdsite=T.rdsite " & vbCrLf
		sqlStr = sqlStr & "			where " & vbCrLf
		sqlStr = sqlStr & "				m.regdate >= @sdate and m.regdate < @edate " & vbCrLf
		sqlStr = sqlStr & "				and m.cancelyn='N' " & vbCrLf
		sqlStr = sqlStr & "				and m.beadaldiv not in (90)  " & vbCrLf '-- 3pl 제외
		if (FRectNotIncOutmall<>"") then
		    sqlStr = sqlStr & "				and m.beadaldiv not in (50,80)" & vbCrLf  '' 입점몰 제외 2017/11/12
	    end if
		sqlStr = sqlStr & "				and m.jumundiv not in (6,9)  " & vbCrLf '-- 교환/반품 제외
		sqlStr = sqlStr & "				and d.cancelyn<>'Y' " & vbCrLf
		sqlStr = sqlStr & "				and d.itemid = '" & FRectItemID & "' " & vbCrLf
		sqlStr = sqlStr & "			group by itemid,convert(varchar(10),regdate,21) " & vbCrLf
		sqlStr = sqlStr & ") as T on DT.solar_date=T.yyyymm and T.itemid='" & FRectItemID & "' " & vbCrLf
		sqlStr = sqlStr & "where solar_date >= @sdate and solar_date <= @edate " & vbCrLf
		sqlStr = sqlStr & "order by yyyymm asc "

		rsAnalget.CursorLocation = adUseClient
		rsAnalget.Open sqlStr,dbAnalget,adOpenForwardOnly,adLockReadOnly

		FTotalCount = rsAnalget.recordcount

		If not rsAnalget.eof Then
			fnNaverMaechulByItem = rsAnalget.getRows()
		End If
		rsAnalget.Close
	End Function
	
	
	Public Function fnItemSellcashHistory
		Dim sqlStr, i, addSql, orderbysql

		sqlStr = "select JustDate from [db_sitemaster].[dbo].[tbl_just1day] "
		sqlStr = sqlStr & "where itemid = '" & FRectItemID & "' and JustDate between '" & FRectSDate & "' and '" & FRectEDate & "'"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly
		If not db3_rsget.eof Then
			FArrJust1Day = db3_rsget.getRows()
		End If
		db3_rsget.Close
		

		sqlStr = ""
		sqlStr = sqlStr & "declare @sdate datetime " & vbCrLf
		sqlStr = sqlStr & "declare @edate datetime " & vbCrLf

		sqlStr = sqlStr & "set @sdate = '" & FRectSDate & " 00:00:00' " & vbCrLf
		sqlStr = sqlStr & "set @edate = '" & FRectEDate & " 23:59:59' " & vbCrLf

		sqlStr = sqlStr & "select " & vbCrLf
		sqlStr = sqlStr & "	DT.BaseDate as yyyymm, DT.weekname, isNULL(T.sellcash,-1) sellcnt " & vbCrLf
		sqlStr = sqlStr & "from [db_datamart].[dbo].[tbl_DumpDate] DT " & vbCrLf
		sqlStr = sqlStr & "left join  " & vbCrLf
		sqlStr = sqlStr & "( " & vbCrLf
		sqlStr = sqlStr & "		select convert(varchar(10),A.regdate,120) as yyyymm, A.sellcash from " & vbCrLf
		sqlStr = sqlStr & "		( " & vbCrLf
		sqlStr = sqlStr & "			select regdate, sellcash, RANK() Over (partition by convert(varchar(10),regdate,120) order by regdate desc) as LastRank " & vbCrLf
		sqlStr = sqlStr & "			from db_log.dbo.tbl_iteminfo_history where itemid='" & FRectItemID & "' " & vbCrLf
		sqlStr = sqlStr & "		) as A " & vbCrLf
		sqlStr = sqlStr & "		where A.LastRank=1" & vbCrLf
		sqlStr = sqlStr & ") as T on DT.BaseDate=T.yyyymm " & vbCrLf
		sqlStr = sqlStr & "where BaseDate >= @sdate and BaseDate <= @edate " & vbCrLf
		sqlStr = sqlStr & "order by yyyymm asc "

		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly

		FTotalCount = db3_rsget.recordcount

		If not db3_rsget.eof Then
			fnItemSellcashHistory = db3_rsget.getRows()
		End If
		db3_rsget.Close
	End Function
	

	Public Function fnCouponMasterList
		Dim sqlStr, i, addSql, orderbysql

		sqlStr = sqlStr & "declare @sdate datetime " & vbCrLf
		sqlStr = sqlStr & "declare @edate datetime " & vbCrLf
		sqlStr = sqlStr & "set @sdate = '" & FRectSDate & " 00:00:00' " & vbCrLf
		sqlStr = sqlStr & "set @edate = '" & FRectEDate & " 23:59:59' " & vbCrLf
		sqlStr = sqlStr & "select " & vbCrLf
		sqlStr = sqlStr & "	m.couponvalue, isNull(m.couponname,'') as couponname, " & vbCrLf
		sqlStr = sqlStr & "	convert(varchar(10),m.startdate,120) as startdate, convert(varchar(10),m.expiredate,120) as expiredate " & vbCrLf
		sqlStr = sqlStr & "from [db_user].[dbo].tbl_user_coupon_master m " & vbCrLf
		sqlStr = sqlStr & "where m.startdate >= @sdate and m.startdate <= @edate " & vbCrLf
		sqlStr = sqlStr & "and datediff(d,startdate,expiredate)<16" & vbCrLf
		sqlStr = sqlStr & "order by m.startdate asc "
		
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

		FTotalCount = rsget.recordcount

		If not rsget.eof Then
			fnCouponMasterList = rsget.getRows()
		End If
		rsget.Close
	End Function
	

end class

Function GraphFile(vGraph)
	Dim vFileName
	Select Case vGraph
		Case "1"
			vFileName = "MSLine.swf"
		Case "2"
			vFileName = "MSColumn3D.swf"
		Case Else
			vFileName = "MSLine.swf"
	End Select
	GraphFile = vFileName
End Function


Function GraphFile2(vGraph)
	Dim vFileName
	Select Case vGraph
		Case "1"
			vFileName = "msline"
		Case "2"
			vFileName = "mscolumn3d"
		Case Else
			vFileName = "msline"
	End Select
	GraphFile2 = vFileName
End Function
%>
