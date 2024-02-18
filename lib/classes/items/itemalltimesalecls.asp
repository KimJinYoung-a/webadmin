<%
'####################################################
' Description : 상시할인관리
' History : 2018.01.22
'####################################################
Class ClsAllTimeSale

public FPSize
public FCPage
public FTotCnt

public FRectMakerid
public FRectSale 
public FRectItemid
public FRectdispcate
public FRectcouponyn
public FRectSort
public FRectinvalidmargin

public Function fnGetItemList
  dim strSql, strSqlAdd
	dim iSPageNo 
	dim FRectSort
	iSPageNo =  FPSize*(FCPage-1)  
	
	strSqlAdd=""
	if  FRectMakerid <> "" then
		strSqlAdd = strSqlAdd & " and i.makerid = '"&FRectMakerid&"'"
  end if
 
  if FRectdispcate <> "" then
 		strSqlAdd = strSqlAdd & " and dci.catecode like '"&FRectdispcate&"%'"
	end if
	
	if (FRectItemid <> "") then
      if right(trim(FRectItemid),1)="," then
      	FRectItemid = Replace(FRectItemid,",,",",")
        strSqlAdd = strSqlAdd & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
      else
				FRectItemid = Replace(FRectItemid,",,",",")
        strSqlAdd = strSqlAdd & " and i.itemid in (" + FRectItemid + ")"
      end if
  end if
  
   if FRectcouponyn <> "" then
   	 strSqlAdd = strSqlAdd & " and i.itemcouponyn='"&FRectcouponyn&"'"
  end if
   
   if FRectSale = "1" then '상시할인
   	 strSqlAdd = strSqlAdd & " and  i.sailyn ='Y' and st.sale_code is null   "
   elseif 	  FRectSale = "2" then '이벤트할인
   	 strSqlAdd = strSqlAdd & " and  i.sailyn ='Y' and st.sale_code is not null and ( st.orgsailyn ='N' or  st.orgsailyn is null ) "
   elseif 	  FRectSale = "3" then '상시+ 이벤트할인
   	 strSqlAdd = strSqlAdd & " and  i.sailyn ='Y' and st.sale_code is not null and   st.orgsailyn ='Y'   " 
   elseif FRectSale ="9" then '할인안함
   	strSqlAdd = strSqlAdd & " and  i.sailyn ='N' "
   end if

	if FRectInvalidMargin="Y" then
		strSqlAdd = strSqlAdd & " and (case when isnull(i.sailprice,0)-isnull(i.sailsuplycash,0)>0 and isnull(i.sailprice,0)>0"
		strSqlAdd = strSqlAdd & " then ((isnull(i.sailprice,0)-isnull(i.sailsuplycash,0))/isnull(i.sailprice,0)*100)"
		strSqlAdd = strSqlAdd & " else -1 end) < 0"
		'strSqlAdd = strSqlAdd & " and ( ( i.sailprice-i.sailsuplycash)/i.sailprice*100)  < 0  "
	end if
        
	strSql = " select count(i.itemid) as totcnt "
	strSql = strSql & " FROM 	db_item.dbo.tbl_item as i with (nolock)"
	strSql = strSql & " left outer join db_item.dbo.tbl_display_cate_item as dci with (nolock) on i.itemid = dci.itemid and dci.isdefault ='Y' "
	'strSql = strSql & "	left outer join db_item.dbo.tbl_display_cate as dc on dci.catecode = dci.catecode " 
	strSql = strSql & " left outer join "
	strSql = strSql & " 	  (	select s.sale_code, si.itemid , si.orgsailprice, si.orgsailsuplycash , si.orgsailyn "
	strSql = strSql & " 			from db_event.dbo.tbl_sale as s with (nolock)"
	strSql = strSql & " 			inner join 	db_event.dbo.tbl_saleitem as si with (nolock) on s.sale_code = si.sale_code"
	strSql = strSql & " 			where s.sale_status = 6  "
	strSql = strSql & " 				and si.saleItem_status = 6  "
	strSql = strSql & " 				and s.sale_using =1  "
	strSql = strSql & " 				and s.sale_startdate<=convert(varchar(10),getdate(),121) and s.sale_enddate >=convert(varchar(10),getdate(),121) " 
	strSql = strSql & " 	  ) as st on i.itemid = st.itemid "
	strSql = strSql & "  where  i.isusing ='Y' " 
	strSql = strSql & strSqlAdd
 rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		End if 
	rsget.close
 
	strSql = "  select i.itemid, i.makerid, i.itemname, i.smallimage ,i.sailyn,i.sellcash, i.buycash,i.orgprice, i.orgsuplycash, i.sailprice, i.sailsuplycash, i.mwdiv,i.limityn,i.limitno, i.limitsold,i.isusing   "
	strSql = strSql & " ,i.itemCouponyn, i.itemCoupontype, i.itemCouponvalue"
	strSql = strSql & " ,  Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail with (nolock) Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "
	strSql = strSql & " , st.sale_code, st.itemid as atItemid , st.orgsailprice as atorgsp, st.orgsailsuplycash as atorgsc, st.orgsailyn as atorgsyn"
	strSql = strSql & ", i.lastupdate "
	strSql = strSql & "  from db_item.dbo.tbl_item as i "
	strSql = strSql & " left outer join db_item.dbo.tbl_display_cate_item as dci with (nolock) on i.itemid = dci.itemid and dci.isdefault ='Y' "
	strSql = strSql & "   left outer join "
	strSql = strSql & " 	  (	select s.sale_code, si.itemid , si.orgsailprice, si.orgsailsuplycash , si.orgsailyn "
	strSql = strSql & " 			from db_event.dbo.tbl_sale as s with (nolock)"
	strSql = strSql & " 			inner join 	db_event.dbo.tbl_saleitem as si with (nolock) on s.sale_code = si.sale_code"
	strSql = strSql & " 			where s.sale_status = 6  "
	strSql = strSql & " 				and si.saleItem_status = 6  "
	strSql = strSql & " 				and s.sale_using =1  "
	strSql = strSql & " 				and s.sale_startdate<=convert(varchar(10),getdate(),121) and s.sale_enddate >=convert(varchar(10),getdate(),121) " 
	strSql = strSql & " 	  ) as st on i.itemid = st.itemid "
	strSql = strSql & "  where i.isusing ='Y' and i.itemid <> 0    "
  strSql = strSql & strSqlAdd
	strSql = strSql & " 		order by "
	if FRectSort = "2" then
	strSql = strSql & " i.itemid desc "
else
	strSql = strSql & " i.lastupdate desc, i.itemid desc "
end if
  strSql = strSql & "    offset "&iSPageNo&" rows "
	strSql = strSql & "  fetch next "&FPSize&"  rows only   "
	
	rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			fnGetItemList = rsget.getRows()
		End if 
	rsget.close
	
End Function

End Class
%>