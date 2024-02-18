<%

Sub SelectBoxEvent(byval selectedId)
   dim tmp_str,query1
   %><select name="eventid">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select idx, eventname from [db_contents].[dbo].tbl_event_master"
   query1 = query1 & " order by idx Desc"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("idx")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("idx")&"' "&tmp_str&">"&rsget("eventname")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub

class CReportMasterItemList

 	public Fselldate
	public Fsellcnt
	public Fselltotal
	public Fbuytotal
	public Fitemid
    public FitemOption

	public Fitem
	public FTotalPrice
	public FTotalEa

	public Fsmallimage

	Public FEventname
	Public FEventIdx
	Public FEventBanImage
	Public FEndday
	Public FStartDay

    public Fsellcnt_mobile
    Public Fsellcnt_PC
    Public Fsellcnt_outmall
    Public Fsellcnt_3PL
    Public Fsellcnt_App
    Public Fsellsum_mobile
    Public Fsellsum_PC
    Public Fsellsum_outmall
    Public Fsellsum_3PL
    Public Fsellsum_App
    Public Fbuysum_mobile
    Public Fbuysum_PC
    Public Fbuysum_outmall
    Public Fbuysum_3PL
    Public Fbuysum_App
	public fTotalreducedprice
	Public freducedprice_Mobile
	Public freducedprice_PC
	Public freducedprice_Outmall
	Public freducedprice_3PL
    Public FwishCnt

	public FYYYYMMDD
	public Fmakerid

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class

class CReportMaster

	public FMasterItemList()
	public maxt
	public maxc

	Public FRectItemID
	Public FRectItemOption
	public Fstartday
	public Fendday
	public FResultCount
	public FTotalCount
	public FRectStart
	public FRectEnd
	public FRectEventid
	public FRectMakerid
	Public FRectGubun

	public FRectOldJumun
	Public FRectCateNo
	public FRectDispCate
	Public FRectEvtKind
	Public FRectEvtType
	Public FRectReportType
    public FRectSort

	Public MasterTbl
	Public DetailTbl


	Public FTotalNo
	Public FTotalCost

	public FTotreducedprice
    public FTotSell
    public FTotBuy
    public FTotCnt

    public FTotCnt_m
    public FTotCnt_p
    public FTotCnt_o
    public FTotCnt_3

    public FTotSell_m
    public FTotSell_p
    public FTotSell_o
    public FTotSell_3

    public FTotBuy_m
    public FTotBuy_p
    public FTotBuy_o
    public FTotBuy_3

    public FTotreducedprice_m
    public FTotreducedprice_p
    public FTotreducedprice_o
    public FTotreducedprice_3

	Private Sub Class_Initialize()

		if FRectOldJumun="on" then
			MasterTbl = " 	[db_log].[dbo].tbl_old_order_detail_2003 "
			DetailTbl = " 	[db_log].[dbo].tbl_old_order_master_2003 "
		else
			MasterTbl = " 	[db_order].[dbo].tbl_order_master "
			DetailTbl = " 	[db_order].[dbo].tbl_order_detail "
		end if

		if FRectEvtKind="" then
			FRectEvtKind=1
		end if
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

    ''데이타 마트에서 가져옴.
    Public Sub GetEventStatisticsDataMart()
        dim strSQL,i

        strSQL = " SELECT E.evt_code,E.evt_name,D.evt_bannerimg, D.evt_bannerimg2010, E.evt_startdate,E.evt_enddate , " &_
                " sum(isnull(S.totSellCnt,0)) as TotalSellCount , " &_
                " isnull(sum(S.totSellSum),0) as TotalCost" &_
                " ,isnull(sum(S.totBuySum),0) as TotalBuyCost" &_
                " ,isNull(sum(S.totreducedSum),0)  as Totalreducedprice "&_
				" ,isNull(sum(S.sellCnt_mobile),0)  as sellcnt_mobile "&_
                " ,isNull(sum(S.sellCnt_PC),0) as sellcnt_PC "&_
                " ,isNull(sum(S.sellCnt_outmall),0)  as sellcnt_outmall "&_
                " ,isNull(sum(S.sellCnt_3PL),0)  as sellcnt_3PL "&_
                " ,isNull(sum(S.sellSum_mobile),0) as sellSum_mobile "&_
                " ,isNull(sum(S.sellSum_PC),0)  as sellSum_PC "&_
                " ,isNull(sum(S.sellSum_outmall),0)  as sellSum_outmall "&_
                " ,isNull(sum(S.sellSum_3PL),0)  as sellSum_3PL "&_
                " ,isNull(sum(S.buysum_mobile),0)  as buysum_mobile "&_
                " ,isNull(sum(S.buysum_PC),0)  as buysum_PC "&_
                " ,isNull(sum(S.buysum_outmall),0)  as buysum_outmall "&_
                " ,isNull(sum(S.buysum_3PL),0)  as buysum_3PL "&_
				" ,isNull(sum(S.reducedprice_Mobile),0)  as reducedprice_Mobile "&_
				" ,isNull(sum(S.reducedprice_PC),0)  as reducedprice_PC "&_
				" ,isNull(sum(S.reducedprice_Outmall),0)  as reducedprice_Outmall "&_
				" ,isNull(sum(S.reducedprice_3PL),0)  as reducedprice_3PL "&_
                " FROM db_event.[dbo].tbl_event E with (nolock)" &_
                " JOIN db_event.[dbo].tbl_event_display D with (nolock)" &_
                " 	on E.evt_code = D.evt_code " &_
                " 	and E.evt_using='Y' " &_
                " 	and E.evt_state>=5 " &_
                " 	and E.evt_enddate>='" & FRectStart & "' " &_
                " 	and E.evt_startdate<'" & FRectEnd & "' "

		if FRectEvtKind<>"" then
			strSQL = strSQL & " 	and E.evt_kind='" & FRectEvtKind & "' "
		'else
		'	strSQL = strSQL & " 	and E.evt_kind='1' "
		end if

		if FRectEvtType<>"" then strSQL  = strSQL & " and  E.evt_type=" & FRectEvtType

        strSQL = strSQL & " left Join [DBDATAMART].db_datamart.dbo.tbl_mkt_daily_event_sell_summary S with (nolock)" &_
                " 	on E.evt_code=S.evt_code"

        IF ReportType="s" then
	        strSQL = strSQL & " 	and yyyymmdd>='" & FRectStart & "' "
	        strSQL = strSQL & " 	and yyyymmdd<'" & FRectEnd & "' "
        END IF

        strSQL = strSQL & " WHERE 1=1"

		If FRectCateNo<>"" then
			strSQL = strSQL & " 	and D.evt_category='" & FRectCateNo & "' "
		end if
		If FRectDispCate <>  "" THEN
			strSQL = strSQL & " 	and D.evt_dispCate like '" & FRectDispCate & "%' "
		END IF
		if FRectEventid <> "" then
			strSQL = strSQL & "		and E.evt_code ='" & FRectEventid & "'"
		end if

        strSQL = strSQL & " GROUP BY E.evt_code,E.evt_name,D.evt_bannerimg, D.evt_bannerimg2010, E.evt_startdate,E.evt_enddate "

        if FRectSort = "ED" then
         strSQL = strSQL &        " order by E.evt_code desc, TotalCost desc "
        elseif FRectSort = "EA" then
         strSQL = strSQL &        " order by E.evt_code ASC, TotalCost desc "
        elseif FRectSort = "TPD" then
           strSQL = strSQL &      " order by  (isnull(sum(S.totSellSum),0)-isnull(sum(S.totBuySum),0) ) desc"
        elseif FRectSort = "TPA" then
           strSQL = strSQL &      " order by  (isnull(sum(S.totSellSum),0)-isnull(sum(S.totBuySum),0)) asc"
        elseif FRectSort = "TRD" then
           strSQL = strSQL &      " order by  isNull(sum(S.totreducedSum),0) desc"
        elseif FRectSort = "TRA" then
           strSQL = strSQL &      " order by  isNull(sum(S.totreducedSum),0) asc"
        elseif FRectSort = "MMD" then
            strSQL = strSQL &      " order by  sellsum_mobile desc "
        elseif FRectSort = "MMA" then
            strSQL = strSQL &      " order by  sellsum_mobile asc"
        elseif FRectSort = "MRD" then
            strSQL = strSQL &      " order by  isNull(sum(S.reducedprice_Mobile),0) desc "
        elseif FRectSort = "MRA" then
            strSQL = strSQL &      " order by  isNull(sum(S.reducedprice_Mobile),0) asc"
        elseif FRectSort = "MPD" then
            strSQL = strSQL &      " order by (isNull(sum(S.sellSum_mobile),0)-isNull(sum(S.buysum_mobile),0)) desc"
        elseif FRectSort = "MPA" then
            strSQL = strSQL &      " order by (isNull(sum(S.sellSum_mobile),0)-isNull(sum(S.buysum_mobile),0)) asc  "
        elseif FRectSort = "WMD" then
            strSQL = strSQL &      " order by  sellsum_PC desc "
        elseif FRectSort = "WMA" then
            strSQL = strSQL &      " order by  sellsum_PC asc "
        elseif FRectSort = "WRD" then
            strSQL = strSQL &      " order by  isNull(sum(S.reducedprice_PC),0) desc "
        elseif FRectSort = "WRA" then
            strSQL = strSQL &      " order by  isNull(sum(S.reducedprice_PC),0) asc "
        elseif FRectSort = "WPD" then
            strSQL = strSQL &      " order by  (isNull(sum(S.sellSum_PC),0)-isNull(sum(S.buysum_PC),0)) desc "
        elseif FRectSort = "WPA" then
            strSQL = strSQL &      " order by  (isNull(sum(S.sellSum_PC),0)-isNull(sum(S.buysum_PC),0)) asc "
        elseif FRectSort = "BMD" then
            strSQL = strSQL &      " order by sellSum_outmall desc "
        elseif FRectSort = "BMA" then
            strSQL = strSQL &      " order by  sellSum_outmall asc "
        elseif FRectSort = "BRD" then
            strSQL = strSQL &      " order by isNull(sum(S.reducedprice_Outmall),0) desc "
        elseif FRectSort = "BRA" then
            strSQL = strSQL &      " order by isNull(sum(S.reducedprice_Outmall),0) asc "
        elseif FRectSort = "BPD" then
            strSQL = strSQL &      " order by  (isNull(sum(S.sellSum_outmall),0)-isNull(sum(S.buysum_outmall),0)) desc"
        elseif FRectSort = "BPA" then
            strSQL = strSQL &      " order by  (isNull(sum(S.sellSum_outmall),0)-isNull(sum(S.buysum_outmall),0)) asc"
        elseif FRectSort = "3MD" then
            strSQL = strSQL &      " order by  sellsum_3PL desc "
        elseif FRectSort = "3MA" then
            strSQL = strSQL &      " order by  sellsum_3PL asc "
        elseif FRectSort = "3RD" then
            strSQL = strSQL &      " order by  isNull(sum(S.reducedprice_3PL),0) desc "
        elseif FRectSort = "3RA" then
            strSQL = strSQL &      " order by  isNull(sum(S.reducedprice_3PL),0) asc "
        elseif FRectSort = "3PD" then
            strSQL = strSQL &      " order by   (isNull(sum(S.sellSum_3PL),0)-isNull(sum(S.buysum_3PL),0)) desc "
        elseif FRectSort = "3PA" then
            strSQL = strSQL &      " order by   (isNull(sum(S.sellSum_3PL),0)-isNull(sum(S.buysum_3PL),0)) asc"
        elseif FRectSort = "TMA" then
           strSQL = strSQL &      " order by TotalCost asc "
        else
           strSQL = strSQL &      " order by TotalCost desc "
        end if

		'response.write strSQL & "<br>"
 		'response.End
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalCount = rsget.RecordCount
        FTotSell = 0
        FTotBuy  = 0
        FTotCnt  = 0
		FTotreducedprice = 0

        FTotCnt_m = 0
        FTotCnt_p= 0
        FTotCnt_o = 0
        FTotCnt_3 = 0

        FTotSell_m= 0
        FTotSell_p= 0
        FTotSell_o= 0
        FTotSell_3= 0

        FTotBuy_m = 0
        FTotBuy_p = 0
        FTotBuy_o = 0
        FTotBuy_3 = 0

        FTotreducedprice_m = 0
        FTotreducedprice_p = 0
        FTotreducedprice_o = 0
        FTotreducedprice_3 = 0

		redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
			set FMasterItemList(i) = new CReportMasterItemList

			FMasterItemList(i).FEventIdx = rsget("evt_code")
			if Not(rsget("evt_bannerimg")="" or isNull(rsget("evt_bannerimg"))) then
				FMasterItemList(i).FEventBanImage = rsget("evt_bannerimg")
			else
				FMasterItemList(i).FEventBanImage = rsget("evt_bannerimg2010")
			end if
			FMasterItemList(i).FEventName = db2html(rsget("evt_name"))
			FMasterItemList(i).FStartDay = left(rsget("evt_startdate"),10)
			FMasterItemList(i).FEndDay = left(rsget("evt_enddate"),10)

			FMasterItemList(i).Fselltotal = rsget("TotalCost")
			FMasterItemList(i).Fbuytotal = rsget("TotalBuyCost") '공급가 추가 2014-03-28 정윤정
			FMasterItemList(i).Fsellcnt = rsget("TotalSellCount")
			FMasterItemList(i).fTotalreducedprice = rsget("Totalreducedprice")

			FMasterItemList(i).Fsellcnt_mobile  = rsget("sellcnt_mobile")
			FMasterItemList(i).Fsellcnt_PC      = rsget("sellcnt_PC")
			FMasterItemList(i).Fsellcnt_outmall = rsget("sellcnt_outmall")
			FMasterItemList(i).Fsellcnt_3PL     = rsget("sellcnt_3PL")

			FMasterItemList(i).Fsellsum_mobile  = rsget("sellsum_mobile")
			FMasterItemList(i).Fsellsum_PC      = rsget("sellsum_PC")
			FMasterItemList(i).Fsellsum_outmall = rsget("sellsum_outmall")
			FMasterItemList(i).Fsellsum_3PL     = rsget("sellsum_3PL")

			FMasterItemList(i).Fbuysum_mobile   = rsget("buysum_mobile")
			FMasterItemList(i).Fbuysum_PC       = rsget("buysum_PC")
			FMasterItemList(i).Fbuysum_outmall  = rsget("buysum_outmall")
			FMasterItemList(i).Fbuysum_3PL      = rsget("buysum_3PL")

			FMasterItemList(i).freducedprice_Mobile   = rsget("reducedprice_Mobile")
			FMasterItemList(i).freducedprice_PC       = rsget("reducedprice_PC")
			FMasterItemList(i).freducedprice_Outmall  = rsget("reducedprice_Outmall")
			FMasterItemList(i).freducedprice_3PL      = rsget("reducedprice_3PL")

            FTotSell = FTotSell + FMasterItemList(i).Fselltotal
            FTotBuy  = FTotBuy  + FMasterItemList(i).Fbuytotal
            FTotCnt  = FTotCnt + FMasterItemList(i).Fsellcnt
			FTotreducedprice  = FTotreducedprice + FMasterItemList(i).fTotalreducedprice

            FTotCnt_m = FTotCnt_m + FMasterItemList(i).Fsellcnt_mobile
            FTotCnt_p = FTotCnt_p + FMasterItemList(i).Fsellcnt_PC
            FTotCnt_o = FTotCnt_o + FMasterItemList(i).Fsellcnt_outmall
            FTotCnt_3 = FTotCnt_3 + FMasterItemList(i).Fsellcnt_3PL

            FTotSell_m = FTotSell_m + FMasterItemList(i).Fsellsum_mobile
            FTotSell_p = FTotSell_p + FMasterItemList(i).Fsellsum_PC
            FTotSell_o = FTotSell_o + FMasterItemList(i).Fsellsum_outmall
            FTotSell_3 = FTotSell_3 + FMasterItemList(i).Fsellsum_3PL

            FTotBuy_m = FTotBuy_m + FMasterItemList(i).Fbuysum_mobile
            FTotBuy_p = FTotBuy_p + FMasterItemList(i).Fbuysum_PC
            FTotBuy_o = FTotBuy_o + FMasterItemList(i).Fbuysum_outmall
            FTotBuy_3 = FTotBuy_3 + FMasterItemList(i).Fbuysum_3PL

            FTotreducedprice_m = FTotreducedprice_m + FMasterItemList(i).freducedprice_Mobile
            FTotreducedprice_p = FTotreducedprice_p + FMasterItemList(i).freducedprice_PC
            FTotreducedprice_o = FTotreducedprice_o + FMasterItemList(i).freducedprice_Outmall
            FTotreducedprice_3 = FTotreducedprice_3 + FMasterItemList(i).freducedprice_3PL

		rsget.MoveNext
		i = i + 1
		loop

		rsget.close
    End Sub

	'// 기간별 전체 이벤트 통계
	Public Sub GetEventStatisticsAll()
	    dim strSQL,i
	    '' OM : 서머리 작성 후 Join

	    strSQL = " SELECT EM.evt_code,EM.evt_name,EM.evt_bannerimg, EM.evt_startdate,EM.evt_enddate, sum(isnull(OM.TotalNo,0)) as TotalSellCount ,sum(isnull(OM.totalCost,0)) as TotalCost	" &_
				" FROM (	" &_

				" 	SELECT e1.evt_code,e1.evt_name,d.evt_bannerimg,e1.evt_startdate,e1.evt_enddate 	" &_
				" 	FROM db_event.[dbo].tbl_event e1	" &_
				" 	JOIN db_event.[dbo].tbl_event_display d	" &_
				" 		on e1.evt_code = d.evt_code	" &_
				" 	WHERE e1.evt_startdate<'" & FRectEnd & "' 	" &_
				"   and e1.evt_enddate>='" & FRectStart & "' " &_
				" 	and evt_kind='" & FRectEvtKind & "' and evt_state>=7 and evt_using='Y'	"
				if FRectCateNo <> "" then
				strSQL = strSQL + "" &_
				"	and d.evt_category ='" & FRectCateNo & "'"
				end if
				if FRectEventid <> "" then
				strSQL = strSQL + "" &_
				"	and e1.evt_code ='" & FRectEventid & "'"
				end if
				strSQL = strSQL + "" &_
				" 	) AS EM	" &_
				" JOIN db_event.dbo.tbl_eventitem ET	" &_
				" ON EM.evt_code = ET.evt_code	" &_
				" LEFT JOIN ( 	" &_
				" 	SELECT 0 as itemid,0 as TotalNo,0 as totalCost	" &_
				" 	) as OM 	" &_
				" ON ET.itemid=OM.itemid 	" &_
				" GROUP BY  EM.evt_code,EM.evt_name,EM.evt_bannerimg,EM.evt_startdate,EM.evt_enddate	" &_
				" ORDER BY sum(OM.totalCost) DESC,sum(OM.TotalNo) DESC, EM.evt_code DESC "

		rsget.open strSQL,dbget

		FResultCount = rsget.RecordCount


		redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
			set FMasterItemList(i) = new CReportMasterItemList

			FMasterItemList(i).FEventIdx = rsget("evt_code")
			FMasterItemList(i).FEventBanImage = rsget("evt_bannerimg")
			FMasterItemList(i).FEventName = db2html(rsget("evt_name"))
			FMasterItemList(i).FStartDay = left(rsget("evt_startdate"),10)
			FMasterItemList(i).FEndDay = left(rsget("evt_enddate"),10)
			FMasterItemList(i).Fselltotal = rsget("TotalCost")
			FMasterItemList(i).Fsellcnt = rsget("TotalSellCount")

		rsget.MoveNext
		i = i + 1
		loop

		rsget.close

	end Sub


	Public Sub OLD_GetEventStatisticsAll()
		dim strSQL,i

'		strSQL = " SELECT EM.evt_code,EM.evt_name,EM.evt_bannerimg, EM.evt_startdate,EM.evt_enddate, sum(isnull(OM.TotalNo,0)) as TotalSellCount ,sum(isnull(OM.totalCost,0)) as TotalCost	" &_
'				" FROM (	" &_
'
'				" 	SELECT e1.evt_code,e1.evt_name,d.evt_bannerimg,e1.evt_startdate,e1.evt_enddate 	" &_
'				" 	FROM db_event.[dbo].tbl_event e1	" &_
'				" 	JOIN db_event.[dbo].tbl_event_display d	" &_
'				" 		on e1.evt_code = d.evt_code	" &_
'		        " 	WHERE e1.evt_startdate<'" & FRectEnd & "' 	" &_
'		        "   and e1.evt_enddate>='" & FRectStart & "' " &_
'				" 	and evt_kind='" & FRectEvtKind & "' and evt_state>=5 and evt_using='Y'	"
'				if FRectCateNo <> "" then
'				strSQL = strSQL + "" &_
'				"	and d.evt_category ='" & FRectCateNo & "'"
'				end if
'				if FRectEventid <> "" then
'				strSQL = strSQL + "" &_
'				"	and e1.evt_code ='" & FRectEventid & "'"
'				end if
'				strSQL = strSQL + "" &_
'				" 	) AS EM	" &_
'				" JOIN db_event.dbo.tbl_eventitem ET	" &_
'				" ON EM.evt_code = ET.evt_code	" &_
'				" LEFT JOIN ( 	" &_
'				" 	SELECT d.itemid,Sum(itemno) as TotalNo,sum(d.itemno*d.itemcost) as totalCost	" &_
'				" 	FROM " & MasterTbl & " m 	" &_
'				" 	JOIN " & DetailTbl & " d 	" &_
'				" 	    On m.orderserial=d.orderserial 	" &_
'				" 		and m.ipkumdiv >=4 and m.cancelyn='N'	" &_
'				" 		and d.cancelyn<>'Y' and d.itemid<>0 	" &_
'				" 		and m.regdate between '" & FRectStart & "'  and '" & FRectEnd & "' 	" &_
'				" 	GROUP BY d.itemid	" &_
'
'				" 	) as OM 	" &_
'				" ON ET.itemid=OM.itemid 	" &_
'				" GROUP BY  EM.evt_code,EM.evt_name,EM.evt_bannerimg,EM.evt_startdate,EM.evt_enddate	" &_
'				" ORDER BY sum(OM.totalCost) DESC,sum(OM.TotalNo) DESC "
		strSQL =" SELECT E.evt_code,E.evt_name,D.evt_bannerimg,E.evt_startdate,E.evt_enddate " &_
				" ,sum(isnull(OD.itemNo,0)) as TotalSellCount,isnull(sum(OD.itemcost*OD.itemNo),0) as TotalCost " &_
				" FROM db_event.[dbo].tbl_event E  " &_
				" JOIN db_event.[dbo].tbl_event_display D " &_
				" 	on E.evt_code = D.evt_code " &_
				" 	and E.evt_startdate<'" & FRectEnd & "' and E.evt_enddate>='" & FRectStart & "' " &_
				" 	and E.evt_kind='" & FRectEvtKind & "' and E.evt_state>=5 and E.evt_using='Y' " &_
				" JOIN db_event.dbo.tbl_eventitem ET " &_
				" 	ON E.evt_code = ET.evt_code  " &_
				" JOIN " & DetailTbl & " OD  " &_
				" 	ON ET.itemid = OD.itemid  " &_
				" JOIN " & MasterTbl & " OM  " &_
				" 	On OM.orderserial=OD.orderserial and OM.ipkumdiv >=4  " &_
				" 	and OM.cancelyn='N' and OD.cancelyn<>'Y' and OD.itemid<>0  " &_
				" 	and OM.RegDate between E.evt_startdate and dateadd(day,1,E.evt_enddate) " &_
				" WHERE 1=1 "

				IF FRectEventid <> "" THEN
					strSQL = strSQL & "	and E.Evt_Code ='" & FRectEventid & "'"
				END IF

				IF FRectCateNo <> "" then
					strSQL = strSQL & "	and D.evt_category ='" & FRectCateNo & "'"
				END IF
				IF FRectReportType="s" THEN
					strSQL = strSQL & "	and OM.RegDate between '" & FRectStart & "' and dateadd(day,1,'" & FRectEnd & "') "
				END IF

				strSQL = strSQL &_
				" GROUP BY E.evt_code,E.evt_name,D.evt_bannerimg,E.evt_startdate,E.evt_enddate " &_
				" ORDER BY isnull(sum(OD.itemcost*OD.itemNo),0) desc"

		'response.write strSQL
		'dbget.close()	:	response.End
		rsget.open strSQL,dbget



		FResultCount = rsget.RecordCount


		redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
			set FMasterItemList(i) = new CReportMasterItemList

			FMasterItemList(i).FEventIdx = rsget("evt_code")
			FMasterItemList(i).FEventBanImage = rsget("evt_bannerimg")
			FMasterItemList(i).FEventName = db2html(rsget("evt_name"))
			FMasterItemList(i).FStartDay = left(rsget("evt_startdate"),10)
			FMasterItemList(i).FEndDay = left(rsget("evt_enddate"),10)
			FMasterItemList(i).Fselltotal = rsget("TotalCost")
			FMasterItemList(i).Fsellcnt = rsget("TotalSellCount")

		rsget.MoveNext
		i = i + 1
		loop

		rsget.close

	end Sub


	Public Sub GetEventStatisticsAllSelectedTerm()
		dim strSQL,i


		strSQL =" SELECT T.evt_code,T.evt_name , T.evt_bannerimg , T.evt_startdate , T.evt_enddate , SUM(T.TotalSellCount) as TotalSellCount , SUM(T.TotalCost) as TotalCost "&_
				" FROM ( "

				IF (datediff("m",Sdate,now()) <= 6 or datediff("m",Edate,now()) <= 6) Then
					strSQL= strSQL &_
					" 	SELECT E.evt_code,E.evt_name,D.evt_bannerimg,E.evt_startdate,E.evt_enddate "&_
					" 	,sum(isnull(OD.itemNo,0)) as TotalSellCount , isnull(sum(OD.itemcost*OD.itemNo),0) as TotalCost "&_
					" 	FROM db_event.[dbo].tbl_event E "&_
					" 	left JOIN db_event.[dbo].tbl_event_display D "&_
					" 		on E.evt_code = D.evt_code "&_
					" 	JOIN db_event.dbo.tbl_eventitem ET "&_
					" 		ON E.evt_code = ET.evt_code "&_
					" 	JOIN [db_order].[dbo].tbl_order_detail OD "&_
					" 		ON ET.itemid = OD.itemid "&_
					" 	JOIN [db_order].[dbo].tbl_order_master OM "&_
					" 		On OM.orderserial=OD.orderserial and OM.ipkumdiv >=4 "&_
					" 	and OM.cancelyn='N' and OD.cancelyn<>'Y' and OD.itemid<>0 "&_
					" 	and OM.RegDate between E.evt_startdate and dateadd(day,1,E.evt_enddate) "&_
					" 	WHERE E.evt_using='Y' and E.evt_state>=5 "&_
					" 		and E.evt_startdate<'" & FRectEnd & "' and E.evt_enddate>='" & FRectStart & "' and E.evt_kind='" & FRectEvtKind & "' "

						IF FRectEventid <> "" THEN
							strSQL = strSQL & "	and E.Evt_Code ='" & FRectEventid & "'"
						END IF

						IF FRectCateNo <> "" then
							strSQL = strSQL & "	and D.evt_category ='" & FRectCateNo & "'"
						END IF
						IF FRectReportType="s" THEN
							strSQL = strSQL & "	and OM.RegDate between '" & FRectStart & "' and dateadd(day,1,'" & FRectEnd & "') "
						END IF

					strSQL = strSQL &_
					" 	GROUP BY E.evt_code,E.evt_name,D.evt_bannerimg,E.evt_startdate,E.evt_enddate "
				END IF

				IF not ((datediff("m",Sdate,now()) <= 6 and datediff("m",Edate,now()) <= 6) or (datediff("m",Sdate,now()) > 6 and datediff("m",Edate,now()) > 6)) Then
					strSQL = strSQL & " 	UNION ALL "
				End IF

				IF (datediff("m",Sdate,now()) > 6 or datediff("m",Edate,now()) > 6) Then
					strSQL = strSQL &_
					" 	SELECT E.evt_code,E.evt_name,D.evt_bannerimg,E.evt_startdate,E.evt_enddate "&_
					" 	,sum(isnull(OD.itemNo,0)) as TotalSellCount , isnull(sum(OD.itemcost*OD.itemNo),0) as TotalCost "&_
					" 	FROM db_event.[dbo].tbl_event E "&_
					" 	left JOIN db_event.[dbo].tbl_event_display D "&_
					" 		on E.evt_code = D.evt_code "&_
					" 	JOIN db_event.dbo.tbl_eventitem ET "&_
					" 		ON E.evt_code = ET.evt_code "&_
					" 	JOIN db_log.[dbo].tbl_old_order_detail_2003 OD "&_
					" 		ON ET.itemid = OD.itemid "&_
					" 	JOIN db_log.[dbo].tbl_old_order_master_2003 OM "&_
					" 		On OM.orderserial=OD.orderserial and OM.ipkumdiv >=4 "&_
					" 	and OM.cancelyn='N' and OD.cancelyn<>'Y' and OD.itemid<>0 "&_
					" 	and OM.RegDate between E.evt_startdate and dateadd(day,1,E.evt_enddate) "&_
					" 	WHERE E.evt_using='Y' and E.evt_state>=5 "&_
					" 		and E.evt_startdate<'" & FRectEnd & "' and E.evt_enddate>='" & FRectStart & "' and E.evt_kind='" & FRectEvtKind & "' "

						IF FRectEventid <> "" THEN
							strSQL = strSQL & "	and E.Evt_Code ='" & FRectEventid & "'"
						END IF

						IF FRectCateNo <> "" then
							strSQL = strSQL & "	and D.evt_category ='" & FRectCateNo & "'"
						END IF
						IF FRectReportType="s" THEN
							strSQL = strSQL & "	and OM.RegDate between '" & FRectStart & "' and dateadd(day,1,'" & FRectEnd & "') "
						END IF
					strSQL = strSQL &_
					" 	GROUP BY E.evt_code,E.evt_name,D.evt_bannerimg,E.evt_startdate,E.evt_enddate "
				End IF

				strSQL = strSQL &_
				" ) AS T "&_
				" GROUP BY T.evt_code,T.evt_name , T.evt_bannerimg , T.evt_startdate , T.evt_enddate "&_
				" ORDER BY isnull(sum(T.TotalCost),0) desc  "

		response.write strSQL
		dbget.close()	:	response.End
		'rsget.open strSQL,dbget



		FResultCount = rsget.RecordCount


		redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
			set FMasterItemList(i) = new CReportMasterItemList

			FMasterItemList(i).FEventIdx = rsget("evt_code")
			FMasterItemList(i).FEventBanImage = rsget("evt_bannerimg")
			FMasterItemList(i).FEventName = db2html(rsget("evt_name"))
			FMasterItemList(i).FStartDay = left(rsget("evt_startdate"),10)
			FMasterItemList(i).FEndDay = left(rsget("evt_enddate"),10)
			FMasterItemList(i).Fselltotal = rsget("TotalCost")
			FMasterItemList(i).Fsellcnt = rsget("TotalSellCount")

		rsget.MoveNext
		i = i + 1
		loop

		rsget.close

	end Sub


    ' //  날짜별 이벤트 상품 판매 통계_DataMart New 2016.01.27
	Public Sub GetEventStatisticsByDateDataMart_New
        dim strSQL,i

        strSQL = " SELECT S.yyyymmdd, " &_
                " sum(isnull(S.totSellCnt,0)) as TotalSellCount , " &_
                " isnull(sum(S.totSellSum),0) as TotalCost" &_
                " ,isnull(sum(S.totBuySum),0) as TotalBuyCost" &_
                " ,isNull(sum(S.sellCnt_mobile),0)  as sellcnt_mobile "&_
                " ,isNull(sum(S.sellCnt_PC),0) as sellcnt_PC "&_
                " ,isNull(sum(S.sellCnt_outmall),0)  as sellcnt_outmall "&_
                " ,isNull(sum(S.sellCnt_3PL),0)  as sellcnt_3PL "&_
                " ,isNull(sum(S.sellSum_mobile),0) as sellSum_mobile "&_
                " ,isNull(sum(S.sellSum_PC),0)  as sellSum_PC "&_
                " ,isNull(sum(S.sellSum_outmall),0)  as sellSum_outmall "&_
                " ,isNull(sum(S.sellSum_3PL),0)  as sellSum_3PL "&_
                " ,isNull(sum(S.buysum_mobile),0)  as buysum_mobile "&_
                " ,isNull(sum(S.buysum_PC),0)  as buysum_PC "&_
                " ,isNull(sum(S.buysum_outmall),0)  as buysum_outmall "&_
                " ,isNull(sum(S.buysum_3PL),0)  as buysum_3PL "&_
                " FROM db_event.[dbo].tbl_event E " &_
                "   JOIN db_event.[dbo].tbl_event_display D " &_
                " 	    on E.evt_code = D.evt_code " &_
                " 	        and E.evt_using='Y' " &_
                " 	        and E.evt_state>=5 "
		if FRectEvtKind<>"" then
			strSQL = strSQL & " 	and E.evt_kind='" & FRectEvtKind & "' "
		else
			strSQL = strSQL & " 	and E.evt_kind='1' "
		end if

        strSQL = strSQL & " left Join [DBDATAMART].db_datamart.dbo.tbl_mkt_daily_event_sell_summary S" &_
                " 	on E.evt_code=S.evt_code"
	    strSQL = strSQL & " 	and S.yyyymmdd>='" & FRectStart & "' "
	    strSQL = strSQL & " 	and S.yyyymmdd<'" & FRectEnd & "' "

        strSQL = strSQL & " WHERE 1=1"

		If FRectCateNo<>"" then
			strSQL = strSQL & " 	and D.evt_category='" & FRectCateNo & "' "
		end if
		If FRectDispCate <>  "" THEN
			strSQL = strSQL & " 	and D.evt_dispCate='" & FRectDispCate & "' "
		END IF
		if FRectEventid <> "" then
			strSQL = strSQL & "		and E.evt_code ='" & FRectEventid & "'"
		end if

        strSQL = strSQL & " GROUP BY S.yyyymmdd "
        strSQL = strSQL & " order by S.yyyymmdd asc "

        rsget.open strSQL,dbget

		FResultCount = rsget.RecordCount
        FTotSell = 0
        FTotBuy  = 0
        FTotCnt  = 0

        FTotCnt_m = 0
        FTotCnt_p= 0
        FTotCnt_o = 0
        FTotCnt_3 = 0

        FTotSell_m= 0
        FTotSell_p= 0
        FTotSell_o= 0
        FTotSell_3= 0

        FTotBuy_m = 0
        FTotBuy_p = 0
        FTotBuy_o = 0
        FTotBuy_3 = 0


		redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
			set FMasterItemList(i) = new CReportMasterItemList
			FMasterItemList(i).FYYYYMMDD = left(rsget("yyyymmdd"),10)

			FMasterItemList(i).Fselltotal = rsget("TotalCost")
			FMasterItemList(i).Fbuytotal = rsget("TotalBuyCost") '공급가 추가 2014-03-28 정윤정
			FMasterItemList(i).Fsellcnt = rsget("TotalSellCount")

			FMasterItemList(i).Fsellcnt_mobile  = rsget("sellcnt_mobile")
			FMasterItemList(i).Fsellcnt_PC      = rsget("sellcnt_PC")
			FMasterItemList(i).Fsellcnt_outmall = rsget("sellcnt_outmall")
			FMasterItemList(i).Fsellcnt_3PL     = rsget("sellcnt_3PL")

			FMasterItemList(i).Fsellsum_mobile  = rsget("sellsum_mobile")
			FMasterItemList(i).Fsellsum_PC      = rsget("sellsum_PC")
			FMasterItemList(i).Fsellsum_outmall = rsget("sellsum_outmall")
			FMasterItemList(i).Fsellsum_3PL     = rsget("sellsum_3PL")

			FMasterItemList(i).Fbuysum_mobile   = rsget("buysum_mobile")
			FMasterItemList(i).Fbuysum_PC       = rsget("buysum_PC")
			FMasterItemList(i).Fbuysum_outmall  = rsget("buysum_outmall")
			FMasterItemList(i).Fbuysum_3PL      = rsget("buysum_3PL")

            FTotSell = FTotSell + FMasterItemList(i).Fselltotal
            FTotBuy  = FTotBuy  + FMasterItemList(i).Fbuytotal
            FTotCnt  = FTotCnt + FMasterItemList(i).Fsellcnt

            FTotCnt_m = FTotCnt_m + FMasterItemList(i).Fsellcnt_mobile
            FTotCnt_p = FTotCnt_p + FMasterItemList(i).Fsellcnt_PC
            FTotCnt_o = FTotCnt_o + FMasterItemList(i).Fsellcnt_outmall
            FTotCnt_3 = FTotCnt_3 + FMasterItemList(i).Fsellcnt_3PL

            FTotSell_m = FTotSell_m + FMasterItemList(i).Fsellsum_mobile
            FTotSell_p = FTotSell_p + FMasterItemList(i).Fsellsum_PC
            FTotSell_o = FTotSell_o + FMasterItemList(i).Fsellsum_outmall
            FTotSell_3 = FTotSell_3 + FMasterItemList(i).Fsellsum_3PL

            FTotBuy_m = FTotBuy_m + FMasterItemList(i).Fbuysum_mobile
            FTotBuy_p = FTotBuy_p + FMasterItemList(i).Fbuysum_PC
            FTotBuy_o = FTotBuy_o + FMasterItemList(i).Fbuysum_outmall
            FTotBuy_3 = FTotBuy_3 + FMasterItemList(i).Fbuysum_3PL
		rsget.MoveNext
		i = i + 1
		loop

		rsget.close
    End Sub

    ' //  날짜별 이벤트 상품 판매 통계_DataMart
	Public Sub GetEventStatisticsByDateDataMart
		dim strSQL ,i

        strSQL = " select yyyymmdd as dates, totsellCnt as TotalNo, totsellSum as  TotalCost "
        strSQL = strSQL + " from  [DBDATAMART].db_datamart.dbo.tbl_mkt_daily_event_sell_summary "
        				IF FRectEventid <> "" THEN
        strSQL = strSQL + " where evt_code=" & FRectEventid & " "
       					 END IF
        strSQL = strSQL + " and yyyymmdd>='" & FRectStart & "' "
        strSQL = strSQL + " and yyyymmdd<'" & FRectEnd & "' "
        strSQL = strSQL + " order by Dates "

		rsget.Open strSQL,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FMasterItemList(FResultCount)

		if not rsget.eof then
			do until rsget.eof
				set FMasterItemList(i) = new CReportMasterItemList
				FMasterItemList(i).Fselldate 	= rsget("Dates")
				FMasterItemList(i).Fselltotal = rsget("totalCost")
				FMasterItemList(i).Fsellcnt 	= rsget("TotalNo")
				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxc = MaxVal(maxc,FMasterItemList(i).Fselltotal)
				end if

			rsget.MoveNext
			i = i + 1
			loop
		end if
		rsget.close

    end Sub

	' //  날짜별 이벤트 상품 판매 통계
	Public Sub GetEventStatisticsByDate
		dim strSQL ,i

		strSQL = " SELECT  convert(varchar(10),m.regdate,20) as Dates , Sum(itemno) as TotalNo,sum(d.itemno*d.itemcost) as TotalCost  " &VbCRLF
		strSQL = strSQL& " FROM " & MasterTbl & " m   " &VbCRLF
		strSQL = strSQL& " JOIN " & DetailTbl & " d   " &VbCRLF
		strSQL = strSQL& " 	ON m.orderserial=d.orderserial and m.ipkumdiv >=4 and m.cancelyn='N' and d.cancelyn<>'Y' and d.itemid<>0   " &VbCRLF
		strSQL = strSQL& " 	and m.regdate between '" & FRectStart & "' and dateadd(day,1,'" & FRectEnd & "')   " &VbCRLF
		
		if FRectEventid<>"" then
			strSQL = strSQL& " JOIN db_event.dbo.tbl_event e   " &VbCRLF
			strSQL = strSQL& " 	ON e.evt_code='" & FRectEventid & "' " &VbCRLF
			strSQL = strSQL& " 		AND m.regdate between e.evt_startdate and dateadd(day,1,e.evt_enddate) " &VbCRLF
		end if
		
		strSQL = strSQL& " WHERE 1=1"
		strSQL = strSQL& " and d.itemid in ( "&VbCRLF
		IF FRectItemID<>"" then
			strSQL = strSQL & FRectItemID &VbCRLF
		ELSE
			strSQL = strSQL & " SELECT itemid FROM db_event.[dbo].tbl_eventitem WHERE evt_code='" & FRectEventid & "'" &VbCRLF
		END IF
		strSQL = strSQL & ") " &VbCRLF

		if (FRectItemID<>"") and (FRectItemOption<>"") then
		    strSQL = strSQL& " and d.itemoption='"&FRectItemOption&"'"
		end if
		strSQL = strSQL & " GROUP BY convert(varchar(10),m.regdate,20) order by Dates "
		''response.write strSQL
		rsget.Open strSQL,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FMasterItemList(FResultCount)

		if not rsget.eof then
			do until rsget.eof
				set FMasterItemList(i) = new CReportMasterItemList
				FMasterItemList(i).Fselldate 	= rsget("Dates")
				FMasterItemList(i).Fselltotal = rsget("totalCost")
				FMasterItemList(i).Fsellcnt 	= rsget("TotalNo")
				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxc = MaxVal(maxc,FMasterItemList(i).Fselltotal)
				end if

			rsget.MoveNext
			i = i + 1
			loop
		end if
		rsget.close

	end Sub

	' //  상품별 이벤트 상품 판매 통계
	Public Sub GetEventStatisticsByItemID
		dim strSQL
		response.write "사용중지 : GetEventStatisticsByItemIDDataMart() 사용할것"
		dbget.close()	:	response.end

		'' strSQL = " SELECT T.itemid , isnull(OM.TotalNO,0) as TotalNO , isnull(OM.TotalCost,0) as TotalCost, smallimage " & VbCRLF
		'' strSQL = strSQL & " FROM db_event.[dbo].tbl_eventitem T " & VbCRLF
		'' strSQL = strSQL & " Left JOIN ( " & VbCRLF
		'' strSQL = strSQL & " 	SELECT d.itemid,Sum(itemno) as TotalNo,sum(d.itemno*d.itemcost) as totalCost  " & VbCRLF
		'' strSQL = strSQL & " 	FROM [db_order].[dbo].tbl_order_master m  " & VbCRLF
		'' strSQL = strSQL & " 	JOIN [db_order].[dbo].tbl_order_detail d  " & VbCRLF
		'' strSQL = strSQL & " 		ON m.orderserial=d.orderserial and m.ipkumdiv >=4 and m.cancelyn='N' and d.cancelyn<>'Y' and d.itemid<>0  " & VbCRLF
		'' strSQL = strSQL & " 		and m.regdate between '" & FRectStart & "' and dateadd(day,1,'" & FRectEnd & "')  " & VbCRLF
		'' strSQL = strSQL & "   JOIN db_event.[dbo].tbl_eventitem e" & VbCRLF
		'' strSQL = strSQL & " 		ON e.evt_code = '" & FRectEventid & "'" & VbCRLF
		'' strSQL = strSQL & " 		and d.itemid=e.itemid " & VbCRLF
		'' strSQL = strSQL & " 	GROUP BY d.itemid  " & VbCRLF
		'' strSQL = strSQL & " 	) as OM  " & VbCRLF
		'' strSQL = strSQL & " 	on T.ItemID = OM.ItemID " & VbCRLF
		'' strSQL = strSQL & " Left Outer Join db_item.dbo.tbl_item AS i On T.itemid = i.itemid "
		'' strSQL = strSQL & " WHERE T.evt_code = '" & FRectEventid & "'" & VbCRLF
		'' strSQL = strSQL & " order by TotalNO desc , T.itemid desc"
		'' rsget.Open strSQL,dbget,1

		strSQL = " exec [db_event].[dbo].[usp_Ten_Event_StatisticsByItemID] '" + CStr(FRectStart) + "', '" + CStr(FRectEnd) + "', " + CStr(FRectEventid) + ", '" & FRectMakerid & "' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly

		''response.write "수정중" & strSQL ''데이타 마트로 이전할것?
		''dbget.close()	:	response.end

		FResultCount = rsget.RecordCount

		redim preserve FMasterItemList(FResultCount)

		if not rsget.eof then
			do until rsget.eof
				set FMasterItemList(i) = new CReportMasterItemList
				FMasterItemList(i).Fitemid 	= rsget("itemid")
				FMasterItemList(i).Fselltotal = rsget("TotalCost")
				FMasterItemList(i).Fsellcnt 	= rsget("TotalNO")
				FMasterItemList(i).Fsmallimage = rsget("smallimage")
				FMasterItemList(i).Fmakerid = rsget("makerid")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxc = MaxVal(maxc,FMasterItemList(i).Fselltotal)
				end if


			rsget.MoveNext
			i = i + 1
			loop
		end if
		rsget.close

	end Sub

	' //  상품별 이벤트 상품 판매 통계(데이타마트)
	Public Sub GetEventStatisticsByItemIDDataMart
		dim strSQL

		strSQL = " exec [db_datamart].[dbo].[usp_Ten_Event_StatisticsByItemID] '" + CStr(FRectStart) + "', '" + CStr(FRectEnd) + "', " + CStr(FRectEventid) + ", '" & FRectMakerid & "', '" & FRectSort & "' "
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open strSQL, db3_dbget, adOpenForwardOnly

		''response.write "수정중" & strSQL ''데이타 마트로 이전할것?
		''db3_dbget.close()	:	response.end

		FResultCount = db3_rsget.RecordCount

		redim preserve FMasterItemList(FResultCount)

		if not db3_rsget.eof then
			do until db3_rsget.eof
				set FMasterItemList(i) = new CReportMasterItemList
				FMasterItemList(i).Fitemid 	= db3_rsget("itemid")
				FMasterItemList(i).Fselltotal = db3_rsget("TotalCost")
				FMasterItemList(i).Fsellcnt 	= db3_rsget("TotalNO")
				FMasterItemList(i).Fsmallimage = db3_rsget("smallimage")
				FMasterItemList(i).Fmakerid = db3_rsget("makerid")

				FMasterItemList(i).Fsellcnt_PC		= db3_rsget("PcTotalNO")
				FMasterItemList(i).Fsellsum_PC		= db3_rsget("PcTotalCost")
				FMasterItemList(i).Fsellcnt_mobile	= db3_rsget("MobTotalNO")
				FMasterItemList(i).Fsellsum_mobile	= db3_rsget("MobTotalCost")
				FMasterItemList(i).Fsellcnt_App		= db3_rsget("AppTotalNO")
				FMasterItemList(i).Fsellsum_App		= db3_rsget("AppTotalCost")
				FMasterItemList(i).Fsellcnt_outmall	= db3_rsget("ExtTotalNO")
				FMasterItemList(i).Fsellsum_outmall	= db3_rsget("ExtTotalCost")
				FMasterItemList(i).FwishCnt			= db3_rsget("wishCount")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxc = MaxVal(maxc,FMasterItemList(i).Fselltotal)
				end if


			db3_rsget.MoveNext
			i = i + 1
			loop
		end if
		db3_rsget.close

	end Sub

	' //  브랜드별 이벤트 상품 판매 통계
	Public Sub GetEventStatisticsByMakerID
		dim strSQL

		response.write "사용중지 : GetEventStatisticsByMakerIDDataMart() 사용할것"
		dbget.close()	:	response.end

		strSQL = " exec [db_event].[dbo].[usp_Ten_Event_StatisticsByMakerID] '" + CStr(FRectStart) + "', '" + CStr(FRectEnd) + "', " + CStr(FRectEventid) + " "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly

		FResultCount = rsget.RecordCount

		redim preserve FMasterItemList(FResultCount)

		if not rsget.eof then
			do until rsget.eof
				set FMasterItemList(i) = new CReportMasterItemList
				FMasterItemList(i).Fselltotal = rsget("TotalCost")
				FMasterItemList(i).Fsellcnt 	= rsget("TotalNO")
				FMasterItemList(i).Fmakerid = rsget("makerid")

			rsget.MoveNext
			i = i + 1
			loop
		end if
		rsget.close

	end Sub

	' //  브랜드별 이벤트 상품 판매 통계(데이타마트)
	Public Sub GetEventStatisticsByMakerIDDataMart
		dim strSQL

		strSQL = " exec [db_datamart].[dbo].[usp_Ten_Event_StatisticsByMakerID] '" + CStr(FRectStart) + "', '" + CStr(FRectEnd) + "', " + CStr(FRectEventid) + " "
		''response.write strSQL
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open strSQL, db3_dbget, adOpenForwardOnly

		FResultCount = db3_rsget.RecordCount

		redim preserve FMasterItemList(FResultCount)

		if not db3_rsget.eof then
			do until db3_rsget.eof
				set FMasterItemList(i) = new CReportMasterItemList
				FMasterItemList(i).Fselltotal = db3_rsget("TotalCost")
				FMasterItemList(i).Fsellcnt 	= db3_rsget("TotalNO")
				FMasterItemList(i).Fmakerid = db3_rsget("makerid")

			db3_rsget.MoveNext
			i = i + 1
			loop
		end if
		db3_rsget.close

	end Sub

    ' //  옵션별 이벤트 상품 판매 통계 / 2013/10/14
	Public Sub GetEventStatisticsByItemOption
		dim strSQL

		response.write "사용중지 : GetEventStatisticsByItemOptionDataMart() 사용할것"
		dbget.close()	:	response.end

		strSQL = " SELECT T.itemid ,OM.itemoption , isnull(OM.TotalNO,0) as TotalNO , isnull(OM.TotalCost,0) as TotalCost " & VbCRLF
		strSQL = strSQL & " FROM db_event.[dbo].tbl_eventitem T " & VbCRLF
		strSQL = strSQL & " Left JOIN ( " & VbCRLF
		strSQL = strSQL & " 	SELECT d.itemid,d.itemoption,Sum(itemno) as TotalNo,sum(d.itemno*d.itemcost) as totalCost  " & VbCRLF
		strSQL = strSQL & " 	FROM [db_order].[dbo].tbl_order_master m  " & VbCRLF
		strSQL = strSQL & " 	JOIN [db_order].[dbo].tbl_order_detail d  " & VbCRLF
		strSQL = strSQL & " 		ON m.orderserial=d.orderserial and m.ipkumdiv >=4 and m.cancelyn='N' and d.cancelyn<>'Y' and d.itemid<>0  " & VbCRLF
		strSQL = strSQL & " 		and m.regdate between '" & FRectStart & "' and dateadd(day,1,'" & FRectEnd & "')  " & VbCRLF
		strSQL = strSQL & "   JOIN db_event.[dbo].tbl_eventitem e" & VbCRLF
		strSQL = strSQL & " 		ON e.evt_code = '" & FRectEventid & "'" & VbCRLF
		strSQL = strSQL & " 		and d.itemid=e.itemid " & VbCRLF
		strSQL = strSQL & " 	GROUP BY d.itemid,d.itemoption  " & VbCRLF
		strSQL = strSQL & " 	) as OM  " & VbCRLF
		strSQL = strSQL & " 	on T.ItemID = OM.ItemID " & VbCRLF
		strSQL = strSQL & " WHERE T.evt_code = '" & FRectEventid & "'" & VbCRLF
		strSQL = strSQL & " order by TotalNO desc , T.itemid desc"



		rsget.Open strSQL,dbget,1


		FResultCount = rsget.RecordCount

		redim preserve FMasterItemList(FResultCount)

		if not rsget.eof then
			do until rsget.eof
				set FMasterItemList(i) = new CReportMasterItemList
				FMasterItemList(i).Fitemid 	= rsget("itemid")
				FMasterItemList(i).Fitemoption 	= rsget("itemoption")
				FMasterItemList(i).Fselltotal = rsget("TotalCost")
				FMasterItemList(i).Fsellcnt 	= rsget("TotalNO")
				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxc = MaxVal(maxc,FMasterItemList(i).Fselltotal)
				end if


			rsget.MoveNext
			i = i + 1
			loop
		end if
		rsget.close

	end Sub

    ' //  옵션별 이벤트 상품 판매 통계 / 2013/10/14(데이타마트)
	Public Sub GetEventStatisticsByItemOptionDataMart
		dim strSQL

		strSQL = " exec [db_datamart].[dbo].[usp_Ten_Event_StatisticsByItemOption] '" + CStr(FRectStart) + "', '" + CStr(FRectEnd) + "', " + CStr(FRectEventid) + " "
		''response.write strSQL
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open strSQL, db3_dbget, adOpenForwardOnly

		FResultCount = db3_rsget.RecordCount

		redim preserve FMasterItemList(FResultCount)

		if not db3_rsget.eof then
			do until db3_rsget.eof
				set FMasterItemList(i) = new CReportMasterItemList
				FMasterItemList(i).Fitemid 	= db3_rsget("itemid")
				FMasterItemList(i).Fitemoption 	= db3_rsget("itemoption")
				FMasterItemList(i).Fselltotal = db3_rsget("TotalCost")
				FMasterItemList(i).Fsellcnt 	= db3_rsget("TotalNO")
				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxc = MaxVal(maxc,FMasterItemList(i).Fselltotal)
				end if

			db3_rsget.MoveNext
			i = i + 1
			loop
		end if
		db3_rsget.close

	end Sub

    '// 상품별&날짜별 상품 판매 총합_DataMart
	Public Sub GetEventStatisticsTotalDataMart()

		dim strSQL
        strSQL = " select IsNULL(sum(totsellCnt),0) as TotalNo, IsNULL(sum(totsellSum),0) as  TotalCost " &_
                " from  [DBDATAMART].db_datamart.dbo.tbl_mkt_daily_event_sell_summary " &_
                " where evt_code=" & FRectEventid & " " &_
                " and yyyymmdd>='" & FRectStart & "' " &_
                " and yyyymmdd<'" & FRectEnd & "' "

		'response.write strSQL
		rsget.Open strSQL,dbget,1

		If not rsget.eof then
			FTotalNo = rsget("TotalNo")
			FTotalCost = rsget("TotalCost")
		end if

		rsget.close
	End Sub

	'// 상품별&날짜별 상품 판매 총합
	Public Sub GetEventStatisticsTotal()

		dim strSQL

		strSQL = " SELECT  isnull(Sum(itemno),0) as TotalNo, isnull(sum(d.itemno*d.itemcost),0) as TotalCost  " &_
					" FROM " & MasterTbl & " m   " &_
					" JOIN " & DetailTbl & " d   " &_
					" 	ON m.orderserial=d.orderserial and m.ipkumdiv >=4 and m.cancelyn='N' and d.cancelyn<>'Y' and d.itemid<>0   " &_
					" 	and m.regdate between '" & FRectStart & "' and dateadd(day,1,'" & FRectEnd & "')   " &_
					" WHERE d.itemid in ("
					IF FRectItemID<>"" then
						strSQL = strSQL &	FRectItemID
					ELSE
						strSQL = strSQL &_
						" SELECT itemid FROM db_event.[dbo].tbl_eventitem WHERE evt_code='" & FRectEventid & "'"
					END IF
					strSQL = strSQL &")"
		'response.write strSQL
		rsget.Open strSQL,dbget,1

		If not rsget.eof then
			FTotalNo = rsget("TotalNo")
			FTotalCost = rsget("TotalCost")
		end if

		rsget.close
	End Sub


	Public Sub  SearchJointEventReport()
		Dim sql, i
   		maxc = -1


if FRectEventid <> "" then

'####################################################################
'데이터
'####################################################################

		sql = "select convert(varchar(10),om.regdate,20) as dates,count(om.orderserial)as cnt,sum(od.itemcost) as totalmoney,d.itemid" + vbcrlf
		sql = sql + " from [db_event].[dbo].tbl_eventItem d," + vbcrlf
		if FRectOldJumun="on" then
			sql = sql + " [db_log].[dbo].tbl_old_order_detail_2003 od,[db_log].[dbo].tbl_old_order_master_2003 om" + vbcrlf
		else
			sql = sql + " [db_order].[dbo].tbl_order_detail od,[db_order].[dbo].tbl_order_master om" + vbcrlf
		end if
		sql = sql + " where om.orderserial = od.orderserial" + vbcrlf
		sql = sql + " and om.regdate >= '" + Cstr(FRectStart) + "'" + vbcrlf
		sql = sql + " and om.regdate < '" + Cstr(FRectEnd) + "'" + vbcrlf
		sql = sql + " and d.evt_code in (" + FRectEventid + ")" + vbcrlf
		sql = sql + " and d.itemid = od.itemid" + vbcrlf
		sql = sql + " and om.ipkumdiv >=4" + vbcrlf
		sql = sql + " and om.cancelyn='N'" + vbcrlf
		sql = sql + " and od.cancelyn<>'Y'" + vbcrlf

'		sql = sql + " and om.jumundiv <> 5"
		sql = sql + " group by convert(varchar(10),om.regdate,20),d.itemid" + vbcrlf
		sql = sql + " order by convert(varchar(10),om.regdate,20)"

'--------------- // 구이벤트 구문 ----------------
'		sql = "select convert(varchar(10),om.regdate,20) as dates,count(om.orderserial)as cnt,sum(od.itemcost) as totalmoney,d.itemid" + vbcrlf
'		sql = sql + " from [db_contents].[dbo].tbl_event_detail d," + vbcrlf
'		if FRectOldJumun="on" then
'			sql = sql + " [db_log].[dbo].tbl_old_order_detail_2003 od,[db_log].[dbo].tbl_old_order_master_2003 om" + vbcrlf
'		else
'			sql = sql + " [db_order].[dbo].tbl_order_detail od,[db_order].[dbo].tbl_order_master om" + vbcrlf
'		end if
'		sql = sql + " where om.orderserial = od.orderserial" + vbcrlf
'		sql = sql + " and om.regdate >= '" + Cstr(FRectStart) + "'" + vbcrlf
'		sql = sql + " and om.regdate < '" + Cstr(FRectEnd) + "'" + vbcrlf
'		sql = sql + " and d.masteridx in (" + FRectEventid + ")" + vbcrlf
'		sql = sql + " and d.itemid = od.itemid" + vbcrlf
'		sql = sql + " and om.ipkumdiv >=4" + vbcrlf
'		sql = sql + " and om.cancelyn='N'" + vbcrlf
'		sql = sql + " and od.cancelyn<>'Y'" + vbcrlf
'		sql = sql + " and om.jumundiv <> 5"
'		sql = sql + " group by convert(varchar(10),om.regdate,20),d.itemid" + vbcrlf
'		sql = sql + " order by convert(varchar(10),om.regdate,20)"
'--------------- 구이벤트 구문 // ----------------

'response.write sql
		rsget.Open sql,dbget,1

		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CReportMasterItemList
			    FMasterItemList(i).Fselldate = rsget("dates")
				FMasterItemList(i).Fselltotal = rsget("totalmoney")
				FMasterItemList(i).Fsellcnt = rsget("cnt")
				FMasterItemList(i).Fitemid = rsget("itemid")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
end if
	End Sub



end class

class CReportTotal
	public FMasterItemList()
	public Fstartday
	public Fendday
	public FResultCount
	public FRectRegStart
	public FRectRegEnd
	public FRectEventid
	public FRectOldJumun



	Private Sub Class_Initialize()
		'redim preserve FMasterItemList(0)
		redim  FMasterItemList(0)

		FResultCount = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public sub SearchEventReportTotal()
	dim sql,i
	if FRectEventid <> "" then
		sql = " SELECT d.itemid ,sum(od.itemno)as cnt,sum(od.itemcost*od.itemno) as totalmoney " &_
					" FROM [db_contents].[dbo].tbl_event_detail d "
					if FRectOldJumun="on" then
					sql = sql + "" &_
					" JOIN [db_log].[dbo].tbl_old_order_detail_2003 od " &_
					" 	ON d.itemid = od.itemid " &_
					" JOIN [db_log].[dbo].tbl_old_order_master_2003 om "
					else
					sql = sql + "" &_
					" JOIN [db_order].[dbo].tbl_order_detail od " &_
					" 	ON d.itemid = od.itemid " &_
					" JOIN [db_order].[dbo].tbl_order_master om "
					end if
					sql = sql + "" &_
					" 	ON om.orderserial = od.orderserial " &_
					" 	and om.regdate >= '" + Cstr(Fstartday) + "' and om.regdate < '" + Cstr(Fendday) + "' " &_
					" 	and om.ipkumdiv >=4 " &_
					" 	and om.cancelyn='N' " &_
					" 	and od.cancelyn<>'Y' " &_
					" WHERE d.masteridx=" + FRectEventid + "" &_
					" GROUP BY d.itemid " &_
					" ORDER BY d.itemid "

		'response.write sql
		rsget.Open sql,dbget,1


		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CReportMasterItemList


				FMasterItemList(i).FTotalPrice = rsget("totalmoney")
				FMasterItemList(i).FTotalEa = rsget("cnt")
				FMasterItemList(i).Fitem = rsget("itemid")

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end if
	end Sub

	public sub SearchJointEventReportTotal()
	dim sql,i
	if FRectEventid <> "" then
		sql = "select count(om.orderserial)as cnt,sum(od.itemcost) as totalmoney,d.itemid"
		sql = sql + " from [db_event].[dbo].tbl_eventItem d,"
		if FRectOldJumun="on" then
			sql = sql + " [db_log].[dbo].tbl_old_order_detail_2003 od,[db_log].[dbo].tbl_old_order_master_2003 om" + vbcrlf
		else
			sql = sql + " [db_order].[dbo].tbl_order_detail od,[db_order].[dbo].tbl_order_master om" + vbcrlf
		end if
		sql = sql + " where om.orderserial = od.orderserial"
		sql = sql + " and om.regdate >= '" + Cstr(Fstartday) + "'" + vbcrlf
		sql = sql + " and om.regdate < '" + Cstr(Fendday) + "'" + vbcrlf
		sql = sql + " and d.evt_code in (" + FRectEventid + ")" + vbcrlf
		sql = sql + " and d.itemid = od.itemid"
		sql = sql + " and om.ipkumdiv >=4" + vbcrlf
		sql = sql + " and om.cancelyn='N'" + vbcrlf
		sql = sql + " and od.cancelyn<>'Y'" + vbcrlf
		sql = sql + " group by d.itemid"
		sql = sql + " order by d.itemid"

'---------------- // 구 이벤트 구문 ----------------
'		sql = "select count(om.orderserial)as cnt,sum(od.itemcost) as totalmoney,d.itemid"
'		sql = sql + " from [db_contents].[dbo].tbl_event_detail d,"
'		if FRectOldJumun="on" then
'			sql = sql + " [db_log].[dbo].tbl_old_order_detail_2003 od,[db_log].[dbo].tbl_old_order_master_2003 om" + vbcrlf
'		else
'			sql = sql + " [db_order].[dbo].tbl_order_detail od,[db_order].[dbo].tbl_order_master om" + vbcrlf
'		end if
'		sql = sql + " where om.orderserial = od.orderserial"
'		sql = sql + " and om.regdate >= '" + Cstr(Fstartday) + "'" + vbcrlf
'		sql = sql + " and om.regdate < '" + Cstr(Fendday) + "'" + vbcrlf
'		sql = sql + " and d.masteridx in (" + FRectEventid + ")" + vbcrlf
'		sql = sql + " and d.itemid = od.itemid"
'		sql = sql + " and om.ipkumdiv >=4" + vbcrlf
'		sql = sql + " and om.cancelyn='N'" + vbcrlf
'		sql = sql + " and od.cancelyn<>'Y'" + vbcrlf
'		sql = sql + " group by d.itemid"
'		sql = sql + " order by d.itemid"
'---------------- 구 이벤트 구문 // ----------------

		rsget.Open sql,dbget,1

		FResultCount = rsget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
				set FMasterItemList(i) = new CReportMasterItemList

				FMasterItemList(i).FTotalPrice = rsget("totalmoney")
				FMasterItemList(i).FTotalEa = rsget("cnt")
				FMasterItemList(i).Fitem = rsget("itemid")

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end if
	end Sub

end class
%>
