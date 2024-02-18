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
	public FSellPlusTotal
	public FSellPlusCnt

	public FTotalPrice
	public FTotalEa

 	public FItemid
 	public Fitemname
 	public Fmakerid
 	public FSmallImage
 	public Fsellcash
 	public ForgPrice
 	public Fcdl
 	public Fcdm
 	public Fcds
 	public FSellyn
 	public FSaleYn
 	public FLimitYn
 	public FLimitNo
 	public FLimitSold
 	public FItemCouponYn
 	public FItemCouponValue
 	public FItemCouponType

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class

class CReportMaster

	public FMasterItemList()
	
	Public FRectItemID
	public FResultCount
	public FRectStart
	public FRectEnd
	Public FRectCateNo
	
	Public FTotalNo
	Public FTotalCost
	
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	
	'// 기간별 전체 할인 통계 
	Public Sub GetSaleStatisticsAll()
		dim strSQL,i
		
		strSQL =" SELECT I.Itemid, I.itemname, I.makerid, I.SmallImage,I.sellcash,I.orgPrice " &_
				" 	,I.cate_Large,I.cate_mid,I.cate_Small " &_
				" 	,I.Sellyn,I.SailYn as SaleYn,I.LimitYn,I.LimitNo,I.LimitSold " &_
				" 	,I.ItemCouponYn,I.ItemCouponValue,I.ItemCouponType " &_
				" 	,O.TotalSellCount, O.TotalCost " &_
				" from db_item.dbo.tbl_item as I " &_
				" 	Join ( " &_
				" 		select o1.itemid " &_
				" 			,isnull(sum(o2.TotalSellCount),0) as TotalSellCount " &_
				" 			,isnull(sum(o2.TotalCost),0) as TotalCost " &_
				" 		from " &_
				" 			(select m1.orderserial, d1.itemid " &_
				" 				from db_order.dbo.tbl_order_master as m1 " &_
				" 					join db_order.dbo.tbl_order_detail as d1 " &_
				" 						on m1.orderserial=d1.orderserial " &_
				" 					Join db_item.dbo.tbl_PlusSaleLinkitemList as p1 " &_
				" 						on d1.itemid=p1.plusSaleLinkItemid " &_
				" 				where m1.regdate between '" & FRectStart & "' and dateadd(day,1,'" & FRectEnd & "') " &_
				" 					and m1.ipkumdiv>=4 and m1.cancelyn='N' and d1.cancelyn<>'Y' and d1.itemid<>0  " &_
				" 				group by m1.orderserial, d1.itemid " &_
				" 			) as o1 " &_
				" 			Join " &_
				" 			(select m2.orderserial, d2.itemid " &_
				" 					,sum(isnull(d2.itemNo,0)) as TotalSellCount " &_
				" 					,isnull(sum(d2.itemcost*d2.itemNo),0) as TotalCost " &_
				" 				from db_order.dbo.tbl_order_master as m2 " &_
				" 					join db_order.dbo.tbl_order_detail as d2 " &_
				" 						on m2.orderserial=d2.orderserial " &_
				" 				where d2.issailitem='P' " &_
				" 					and m2.regdate between '" & FRectStart & "' and dateadd(day,1,'" & FRectEnd & "') " &_
				" 					and m2.ipkumdiv>=4 and m2.cancelyn='N' and d2.cancelyn<>'Y' and d2.itemid<>0  " &_
				" 				group by m2.orderserial, d2.itemid " &_
				" 			) as o2 " &_
				" 			on o1.orderserial=o2.orderserial " &_
				" 		group by o1.itemid " &_
				" 	) O " &_
				" 	on O.itemid=I.itemid " &_
				" WHERE 1=1 "

				IF FRectItemid <> "" THEN 
					strSQL = strSQL & "	and I.itemid ='" & FRectItemid & "'"
				END IF
				IF FRectCateNo <> "" then 
					strSQL = strSQL & "	and I.cate_large ='" & FRectCateNo & "'"
				END IF

				strSQL = strSQL &_
				" order by O.TotalCost desc "

		rsget.open strSQL,dbget

		FResultCount = rsget.RecordCount

		redim preserve FMasterItemList(FResultCount)

		do until rsget.eof
			set FMasterItemList(i) = new CReportMasterItemList

			FMasterItemList(i).FItemid			= rsget("plusSaleLinkItemid")
			FMasterItemList(i).Fitemname		= rsget("itemname")
			FMasterItemList(i).Fmakerid			= rsget("makerid")
			FMasterItemList(i).FSmallImage		= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FMasterItemList(i).FItemID) + "/" + rsget("smallimage")
			FMasterItemList(i).Fsellcash		= rsget("sellcash")
			FMasterItemList(i).ForgPrice		= rsget("orgPrice")
			FMasterItemList(i).Fcdl				= rsget("cate_Large")
			FMasterItemList(i).Fcdm				= rsget("cate_mid")
			FMasterItemList(i).Fcds				= rsget("cate_Small")
			FMasterItemList(i).FSellyn			= rsget("Sellyn")
			FMasterItemList(i).FSaleYn			= rsget("SaleYn")
			FMasterItemList(i).FLimitYn			= rsget("LimitYn")
			FMasterItemList(i).FLimitNo			= rsget("LimitNo")
			FMasterItemList(i).FLimitSold		= rsget("LimitSold")
			FMasterItemList(i).FItemCouponYn	= rsget("ItemCouponYn")
			FMasterItemList(i).FItemCouponValue	= rsget("ItemCouponValue")
			FMasterItemList(i).FItemCouponType	= rsget("ItemCouponType")
			FMasterItemList(i).Fselltotal		= rsget("TotalCost")
			FMasterItemList(i).Fsellcnt			= rsget("TotalSellCount")

		rsget.MoveNext
		i = i + 1
		loop

		rsget.close
		
	end Sub
	
	' //  날짜별 할인상품 판매 통계 
	Public Sub GetSaleStatisticsByDate
		dim strSQL ,i
		
		strSQL = " SELECT  convert(varchar(10),m.regdate,20) as Dates " &_
				" 	,Sum(Case When d.plusSaleDiscount>0 then itemno else 0 End) as PlusSaleNo " &_
				" 	,Sum(Case When d.plusSaleDiscount>0 then (d.itemno*d.itemcost) else 0 End) as PlusSaleCost " &_
				" 	,Sum(itemno) as TotalNo,sum(d.itemno*d.itemcost) as totalCost " &_
				" from db_order.dbo.tbl_order_master as m " &_
				" 	Join db_order.dbo.tbl_order_detail as d " &_
				" 		on m.orderserial=d.orderserial " &_
				" 	Join db_item.dbo.tbl_item as i " &_
				" 		on d.itemid=i.itemid " &_
				" where m.ipkumdiv >=4 and m.cancelyn='N' and d.cancelyn<>'Y' and d.itemid<>0 " &_
				" 	and m.regdate between '" & FRectStart & "' and dateadd(day,1,'" & FRectEnd & "') " &_
				" 	and d.itemid in (select distinct plusSaleItemid from db_item.dbo.tbl_PlusSaleLinkitemList) "

				IF FRectCateNo <> "" then 
					strSQL = strSQL & "	and i.cate_large ='" & FRectCateNo & "'"
				END IF

				strSQL = strSQL &_
				" GROUP BY convert(varchar(10),m.regdate,20) order by Dates "

		rsget.Open strSQL,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FMasterItemList(FResultCount)
		
		if not rsget.eof then
			do until rsget.eof
				set FMasterItemList(i) = new CReportMasterItemList
				FMasterItemList(i).Fselldate 		= rsget("Dates")
				FMasterItemList(i).FSellPlusTotal	= rsget("PlusSaleCost")
				FMasterItemList(i).FSellPlusCnt		= rsget("PlusSaleNo")
				FMasterItemList(i).FSellTotal		= rsget("totalCost")
				FMasterItemList(i).FSellCnt			= rsget("TotalNo")

			rsget.MoveNext
			i = i + 1
			loop
		end if
		rsget.close
	
	end Sub
	
	' //  상품별 할인상품 판매 통계 
	Public Sub GetSaleStatisticsByItemID
		dim strSQL 
		
		strSQL = " select i.itemid, i.itemname, I.makerid, I.SmallImage " &_
				" 	, isNull(P.PlusSaleNo,0) as PlusSaleNo, isNull(P.PlusSaleCost,0) as PlusSaleCost, TotalNo, totalCost " &_
				" from db_item.dbo.tbl_item as i " &_
				" 	Join ( " &_
				" 		select d.itemid " &_
				" 			,Sum(Case When d.plusSaleDiscount>0 then itemno End) as PlusSaleNo " &_
				" 			,Sum(Case When d.plusSaleDiscount>0 then (d.itemno*d.itemcost) End) as PlusSaleCost " &_
				" 			,Sum(itemno) as TotalNo,sum(d.itemno*d.itemcost) as totalCost " &_
				" 		from db_order.dbo.tbl_order_master as m " &_
				" 			Join db_order.dbo.tbl_order_detail as d " &_
				" 				on m.orderserial=d.orderserial " &_
				" 		where m.ipkumdiv >=4 and m.cancelyn='N' and d.cancelyn<>'Y' and d.itemid<>0 " &_
				" 			and m.regdate between '" & FRectStart & "' and dateadd(day,1,'" & FRectEnd & "') " &_
				" 			and itemid in (select distinct plusSaleItemid from db_item.dbo.tbl_PlusSaleLinkitemList) " &_
				" 		group by d.itemid ) as P " &_
				" 	on i.itemid=P.itemid " &_
				" where 1=1 "
'				" where P.PlusSaleNo is not Null "		'빈상품 제거

				IF FRectCateNo <> "" then 
					strSQL = strSQL & "	and i.cate_large ='" & FRectCateNo & "'"
				END IF

				strSQL = strSQL &_
				" order by P.PlusSaleCost desc, P.PlusSaleNo desc "

		rsget.Open strSQL,dbget,1
		

		FResultCount = rsget.RecordCount

		redim preserve FMasterItemList(FResultCount)
		
		if not rsget.eof then
			do until rsget.eof
				set FMasterItemList(i) = new CReportMasterItemList
				FMasterItemList(i).Fitemid			= rsget("itemid")
				FMasterItemList(i).Fitemname		= rsget("itemname")
				FMasterItemList(i).Fmakerid			= rsget("makerid")
				FMasterItemList(i).FSmallImage		= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FMasterItemList(i).FItemID) + "/" + rsget("smallimage")
				FMasterItemList(i).FSellPlusTotal	= rsget("PlusSaleCost")
				FMasterItemList(i).FSellPlusCnt		= rsget("PlusSaleNo")
				FMasterItemList(i).FSellTotal		= rsget("totalCost")
				FMasterItemList(i).FSellCnt			= rsget("TotalNo")

			rsget.MoveNext
			i = i + 1
			loop
		end if
		rsget.close
	end Sub

end class

%>
