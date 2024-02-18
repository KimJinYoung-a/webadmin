<%
'###########################################################
' Description : 통계 클래스
' History : 2017.3.27 정윤정 생성 -추후 키바나로 옮길 예정 
'###########################################################

class cStaticList

 public FRectItemid
 public FRectStartdate
 public FRectEndDate
 public FTotalCount
 
	public Function fnGetOrderWishList
		dim strSql
		
		 if FRectItemid ="" then Exit Function
		 	
		strSql = "select  itemid "
		strSql = strSql & " into #tmpitem "
		strSql = strSql & " from db_analyze_data_raw.dbo.tbl_item "
		strSql = strSql & " where itemid in ("&FRectItemid&")"
		dbanalget.Execute strSql	
		
		strSql = "	select top 200 t.itemid,  isNull(count(distinct m.orderserial),0)  as ordercnt "
		strSql = strSql & " , 	(select  isNull(count(distinct f.userid),0) as usercnt "
		strSql = strSql & "  			from dbdatamart.db_my10x10.dbo.tbl_myfavorite as f where  f.itemid = t.itemid and f.regdate >='"&FRectStartdate&"' and f.regdate <'"&FRectEndDate&"') "
		strSql = strSql & "  from #tmpitem as t "
 		strSql = strSql & " left outer join   db_analyze_data_raw.dbo.tbl_order_detail as d on t.itemid = d.itemid "
 		strSql = strSql & " left outer join  db_analyze_data_raw.dbo.tbl_order_master as m  "
  	strSql = strSql & " on m.orderserial = d.orderserial "
 	 	strSql = strSql & " and m.ipkumdiv>3  "
		strSql = strSql & " and m.cancelyn='N' "
		strSql = strSql & " and d.cancelyn<>'Y' "
		strSql = strSql & " and m.regdate >= '"&FRectStartdate&"'  and m.regdate < '"&FRectEndDate&"'   "
		strSql = strSql & " and ( beadaldiv='7' or beadaldiv='8' ) 		 "
		strSql = strSql & " group by t.itemid "
		strSql = strSql & " order by t.itemid " 		
		rsAnalget.open strSql,dbAnalget,1						
	
		If Not (rsAnalget.Eof or rsAnalget.bof ) Then			 
			fnGetOrderWishList = rsAnalget.getRows()
		end if
		
		rsAnalget.close
		
	End Function
 
End class

%>