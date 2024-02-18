<%
Class CMain
public FRectMakerid
public FOrderDate2
public FOnSellCnt2
public FOnBuyCash2
public FOnMaechul2
public FOffSellCnt2
public FOffBuyCash2
public FOffMaechul2
public FOrderDate3
public FOnSellCnt3
public FOnBuyCash3
public FOnMaechul3
public FOffSellCnt3
public FOffBuyCash3
public FOffMaechul3
public FOrderDate1
public FOnSellCnt1
public FOnBuyCash1
public FOnMaechul1
public FOffSellCnt1
public FOffBuyCash1
public FOffMaechul1


public FShopOrderDate1
public FShopOrderDate2
public FShopOrderDate3
public FShopSellCnt1
public FShopSellCnt2
public FShopSellCnt3
public FShopSumTotal1
public FShopSumTotal2
public FShopSumTotal3

public Fitemqanotfinish
public Fnowcsnofincnt
public Fone2onenofincnt
public Fcsnofincnt
public Fcscancelnofincnt
public Feventotfinish

public Flogisnotconfirmcnt
public Flogisnotsendcnt
public FtmpsoldoutItemCnt

public Foffshopnotconfirmcnt
public Foffshopnotsendcnt

public FRectDefDate

public FoffmibaljuCount
public FoffmibeaCount

public function getDlvTrackMifinCount
	Dim sqlStr
	getDlvTrackMifinCount = 0
	sqlStr = "exec [db_order].[dbo].[usp_Ten_Delivery_Trace_BrandView_GetFakeList_CNT_Main]  '"&FRectMakerid&"'"

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		getDlvTrackMifinCount = rsget("cnt")
	rsget.Close

end function

public Function fnGetMainQnA
	dim sqlstr

	'미처리 상품문의
	sqlstr = "select count(*) as cnt" & vbcrlf
	sqlstr = sqlstr & " from [db_cs].[dbo].tbl_my_item_qna p WITH(NOLOCK)" & vbcrlf
	sqlstr = sqlstr & " left join [db_item].[dbo].tbl_item i WITH(NOLOCK)" & vbcrlf
	sqlstr = sqlstr & " 	on p.itemid=i.itemid" & vbcrlf
	sqlstr = sqlstr & " where p.itemid<>0 and i.makerid = '"& FRectMakerid &"'" & vbcrlf
	sqlstr = sqlstr & " and p.replydate is null" & vbcrlf
	sqlstr = sqlstr & " and p.isusing ='Y'" & vbcrlf
	sqlstr = sqlstr & " and p.id >= 400000" & vbcrlf
	'sqlstr = "select count(id) as cnt"
	'sqlstr = sqlstr + " from [db_cs].[dbo].tbl_my_item_qna"
	'sqlstr = sqlstr + " where makerid='" + FRectMakerid + "'"
	'sqlstr = sqlstr + " and isusing='Y'"
	'sqlstr = sqlstr + " and replyuser=''"
	'sqlstr = sqlstr + " and id >= 400000 "

	'response.write sqlstr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		Fitemqanotfinish = rsget("cnt")
	rsget.Close

	'상품준비중 주문취소접수 미처리 CS
	sqlStr = "exec [db_cs].[dbo].sp_Ten_upcheCsCancelCount '" + CStr(session("ssBctID")) +  "'"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
  		Fcscancelnofincnt = rsget("cnt")
	rsget.Close

	'미처리 CS(업체긴급문의)
	sqlStr = "exec [db_cs].[dbo].sp_Ten_upcheNowCsCount '" + CStr(session("ssBctID")) +  "'"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
  		Fnowcsnofincnt = rsget("cnt")
	rsget.Close

	'미처리 고객문의
	sqlStr = "exec [db_cs].[dbo].sp_Ten_UpcheOne2OneCount '" + CStr(session("ssBctID")) +  "'"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
  		Fone2onenofincnt = rsget("cnt")
	rsget.Close

	'미처리 CS
	sqlStr = "exec [db_cs].[dbo].sp_Ten_upcheCsCount '" + CStr(session("ssBctID")) +  "'"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
  		Fcsnofincnt = rsget("cnt")
	rsget.Close


	'미처리 사은품
	sqlstr = "select count(id) as cnt"
	sqlstr = sqlstr + " from [db_sitemaster].[dbo].tbl_etc_songjang w"
	sqlstr = sqlstr + " where w.delivermakerid='" + FRectMakerid + "' and w.deleteyn='N' and ((w.songjangno is NULL) or (w.songjangno='')) and w.isupchebeasong='Y' and datediff(d,reqdeliverdate,getdate())>=-2"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		Feventotfinish = rsget("cnt")
	rsget.Close

End Function

public Function fnUpchebeasongExists
dim sqlStr
	sqlStr = "select IsNULL(sum(smKeyValInt),0) as cnt from db_partner.dbo.tbl_partner_summaryInfo WITH(NOLOCK)"
	sqlStr = sqlStr + " where makerid='" +FRectMakerid + "'"
	sqlStr = sqlStr + " and smKeyName in ('UDTT','UDFX','UD0','UD2','UD3')"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		fnUpchebeasongExists = rsget("cnt")>0
	rsget.Close

End Function

public Function fnGetMainLogics
dim sqlstr
	sqlstr = "select  count(idx) as cnt" + VbCrlf
	sqlstr = sqlstr + " from [db_storage].[dbo].tbl_ordersheet_master WITH(NOLOCK)" + VbCrlf
	sqlstr = sqlstr + " where targetid='" + FRectMakerid + "'" + VbCrlf
	sqlstr = sqlstr + " and baljuid='10x10'" + VbCrlf
	sqlstr = sqlstr + " and statecd='0'" + VbCrlf
	sqlstr = sqlstr + " and deldt is null" + VbCrlf

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		Flogisnotconfirmcnt = rsget("cnt")
	rsget.Close


	sqlstr = "select  count(idx) as cnt" + VbCrlf
	sqlstr = sqlstr + " from [db_storage].[dbo].tbl_ordersheet_master WITH(NOLOCK)" + VbCrlf
	sqlstr = sqlstr + " where targetid='" + FRectMakerid+ "'" + VbCrlf
	sqlstr = sqlstr + " and baljuid='10x10'" + VbCrlf
	sqlstr = sqlstr + " and statecd='1'" + VbCrlf
	sqlstr = sqlstr + " and deldt is null" + VbCrlf

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		Flogisnotsendcnt = rsget("cnt")
	rsget.Close

	''일시품절상품갯수
	sqlstr = "select  count(i.itemid) as cnt" + VbCrlf
	sqlstr = sqlstr + " from db_item.dbo.tbl_item i WITH(NOLOCK)" + VbCrlf
	sqlstr = sqlstr + "     left join [db_item].[dbo].tbl_item_option v WITH(NOLOCK) on i.itemid=v.itemid "
	sqlstr = sqlstr + "     left join [db_item].[dbo].tbl_item_option_stock ot WITH(NOLOCK)"
	sqlstr = sqlstr + "         on ot.itemgubun='10'  "
	sqlstr = sqlstr + "         and ot.itemid=i.itemid "
	sqlstr = sqlstr + "         and ot.itemoption=IsNULL(v.itemoption,'0000') "
	sqlstr = sqlstr + " where i.makerid='" + FRectMakerid+ "'" + VbCrlf
	sqlstr = sqlstr + " and i.sellyn='S'" + VbCrlf
	sqlstr = sqlstr + " and i.mwdiv in ('M','W')" + VbCrlf
	sqlstr = sqlstr + " and i.isusing='Y'"                      ''확인
	sqlstr = sqlstr + " and i.danjongyn in ('S','N')"
	sqlstr = sqlstr + " and ot.stockreipgodate is NULL"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FtmpsoldoutItemCnt = rsget("cnt")
	rsget.Close

End Function

public Function fnGetMainShopOrder
dim sqlstr
	sqlStr = " select count(*) as cnt" + VbCrlf
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_master WITH(NOLOCK)" + VbCrlf
	sqlStr = sqlStr + " where deleteyn='N'" + VbCrlf
	sqlStr = sqlStr + " and ipchulmoveidx is null" + VbCrlf
	sqlStr = sqlStr + " and statecd=-2" + VbCrlf
	sqlStr = sqlStr + " and scheduledate>=DATEADD(MM,-1,GETDATE())" + VbCrlf
	sqlStr = sqlStr + " and scheduledate<DATEADD(MM,1,GETDATE())" + VbCrlf
	sqlStr = sqlStr + " and chargeid='"&FRectMakerid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		Foffshopnotconfirmcnt = rsget("cnt")
	rsget.Close

	sqlStr = " select count(*) as cnt" + VbCrlf
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_master WITH(NOLOCK)" + VbCrlf
	sqlStr = sqlStr + " where deleteyn='N'" + VbCrlf
	sqlStr = sqlStr + " and ipchulmoveidx is null" + VbCrlf
	sqlStr = sqlStr + " and statecd=-1" + VbCrlf
	sqlStr = sqlStr + " and scheduledate>=DATEADD(MM,-1,GETDATE())" + VbCrlf
	sqlStr = sqlStr + " and scheduledate<DATEADD(MM,1,GETDATE())" + VbCrlf
	sqlStr = sqlStr + " and chargeid='"&FRectMakerid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		Foffshopnotsendcnt = rsget("cnt")
	rsget.Close

End Function

public FItemCnt1
public FItemCnt2
public Function fnGetMainItem
	dim strSql
		strSql ="[db_Item].[dbo].sp_Ten_partnerA_getItemSummary('"&FRectMakerid&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			 FItemCnt1 = rsget("cnt1")
			 FItemCnt2 = rsget("cnt2")
		END IF
		rsget.close
End Function

public Function fnGetMainOrder
	dim strSql , arrList, intLoop
	 IF (application("Svr_Info")	= "Dev") then
	 	strSql ="[db_Order].[dbo].sp_Ten_partnerA_getOrderSummary('"&FRectDefDate&"','"&FRectMakerid&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			arrList = rsget.getRows()
		END IF
		rsget.close
	else
		strSql ="[db_datamart].[dbo].[sp_Ten_partnerA_getOrderMonthlySummary]('"&FRectDefDate&"','"&FRectMakerid&"')"
		db3_rsget.Open strSql, db3_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (db3_rsget.EOF OR db3_rsget.BOF) THEN
			arrList = db3_rsget.getRows()
		END IF
		db3_rsget.close
	end if
		if isArray(arrList)	 THEN
			for intLoop = 0 To UBound(arrList,2)
			 	if arrList(0,intLoop) ="preY" THEN
			 		 FOrderDate2 =  arrList(1,intLoop)
			 		 FOnSellCnt2 =  arrList(2,intLoop)
			 		 FOnBuyCash2 =  arrList(3,intLoop)
			 		 FOnMaechul2 =  arrList(4,intLoop)
			 		 FOffSellCnt2 =  arrList(5,intLoop)
			 		 FOffBuyCash2 =  arrList(6,intLoop)
			 		 FOffMaechul2 =  arrList(7,intLoop)
				elseif arrList(0,intLoop) ="preM" THEN
					FOrderDate3 =  arrList(1,intLoop)
			 		 FOnSellCnt3 =  arrList(2,intLoop)
			 		 FOnBuyCash3 =  arrList(3,intLoop)
			 		 FOnMaechul3 =  arrList(4,intLoop)
			 		 FOffSellCnt3 =  arrList(5,intLoop)
			 		 FOffBuyCash3 =  arrList(6,intLoop)
			 		 FOffMaechul3 =  arrList(7,intLoop)
		  	else
		  		FOrderDate1 =  arrList(1,intLoop)
			 		 FOnSellCnt1 =  arrList(2,intLoop)
			 		 FOnBuyCash1 =  arrList(3,intLoop)
			 		 FOnMaechul1 =  arrList(4,intLoop)
			 		 FOffSellCnt1 =  arrList(5,intLoop)
			 		 FOffBuyCash1 =  arrList(6,intLoop)
			 		 FOffMaechul1 =  arrList(7,intLoop)
			 	end if
			Next
	  end if
End Function



'''오프라인 매장배송 관련
public Function fnISOffDlvBrand
dim sqlstr
sqlStr = "select count(*) as CNT from db_shop.dbo.tbl_shop_designer WITH(NOLOCK) where makerid='"&FRectMakerid&"'"
sqlStr = sqlStr & " and defaultbeasongdiv=2"
rsget.CursorLocation = adUseClient
rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
If Not rsget.Eof then
    fnISOffDlvBrand = rsget("CNT")>0
End IF
rsget.Close
End Function

public Function fnGetMainShopInfo
dim sqlstr
   sqlStr = "[db_shop].[dbo].[sp_Ten_Shop_Upche_MibaljuMibeasong_Count] ('"&FRectMakerid&"')"
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
        FoffmibaljuCount = rsget("MiBaljuCnt")
        FoffmibeaCount   = rsget("MiBeasongCnt")
   	END IF
		rsget.close
End Function

'업무문의
	public Function fnGetQnAMain
		dim strSql
		strSql = " SELECT top 5 A.idx,  A.title, A.regdate, isnull(A.replyuser,'') as replyn, isNull(B.username,'') AS worker    "&VBCRLF
		strSql = strSql & " from [db_board].[dbo].tbl_upche_qna AS A WITH(NOLOCK)"&VBCRLF
		strSql = strSql & "  Left JOIN db_partner.dbo.tbl_user_tenbyten AS B WITH(NOLOCK) ON A.workerid = B.userid"&VBCRLF
		strSql = strSql & " where A.isusing = 'Y' and A.userid ='"&FRectMakerid&"'"
		strSql = strSql & " order by idx desc "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		if not rsget.eof Then
			fnGetQnAMain = rsget.getRows()
		end if
		rsget.close
  End Function

' 업무문의 카운트	' 2020.03.13 한용민 생성
public Function fnGetQnAcount
	dim strSql, QnAcount
	QnAcount = 0

	strSql = " SELECT count(a.idx) as cnt"&VBCRLF
	strSql = strSql & " from [db_board].[dbo].tbl_upche_qna AS A WITH (readuncommitted)" & VBCRLF
	strSql = strSql & " where A.isusing = 'Y' and A.userid ='"&FRectMakerid&"'" & VBCRLF
	strSql = strSql & " and datediff(day,a.regdate,getdate())<8" & VBCRLF

	'response.write strSql & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.eof Then
		QnAcount = rsget("cnt")
	end if
	rsget.close

	fnGetQnAcount = QnAcount
End Function

  '//공지사항
	public Function fnGetMainNotice()
		dim strSql
		strSql = " select top 10  board_idx, name, title, writeday,fixnotics "&VBCRLF
		strSql = strSql & ",(select count(comidx) from db_board.dbo.tbl_partnerA_notice_comment WITH(NOLOCK) where isusing =1 and board_idx = T.board_idx) as comCnt "&VBCRLF
		strSql = strSql&" from ( "&VBCRLF
		strSql = strSql&"   select board_idx, name, title, writeday ,fixnotics from ( select top 7 board_idx, name, title, writeday ,fixnotics "&VBCRLF
		strSql = strSql&" 	from [db_board].[dbo].tbl_designer_notice as n WITH(NOLOCK) "&VBCRLF
		strSql = strSql&" 			left outer join db_partner.dbo.tbl_partner_dispcate as p WITH(NOLOCK) "&VBCRLF
		strSql = strSql&"					on n.dispcate1 = p.catecode  and p.makerid ='"&FRectMakerID&"'"&VBCRLF
		strSql = strSql&"  	where deleteyn ='N' and fixnotics='Y' and( (fixsdate <=getdate() and fixedate >=getdate() ) or (fixsdate is Null and fixedate is null))  "&VBCRLF
		strSql = strSql&" 	and (n.dispcate1 is null or n.dispcate1 ='' or (n.dispcate1 is not null and p.catecode is not null) )"&VBCRLF
		strSql = strSql&"		order by board_idx desc ) as F "&VBCRLF
		strSql = strSql&" 	union all  "&VBCRLF
		strSql = strSql&"	select board_idx, name, title, writeday ,fixnotics"&VBCRLF
		strSql = strSql&" 	from [db_board].[dbo].tbl_designer_notice as n WITH(NOLOCK) "&VBCRLF
		strSql = strSql&" 			left outer join db_partner.dbo.tbl_partner_dispcate as p WITH(NOLOCK) "&VBCRLF
		strSql = strSql&"					on n.dispcate1 = p.catecode  and p.makerid ='"&FRectMakerID&"'"&VBCRLF
		strSql = strSql&" 	where deleteyn ='N' and fixnotics<>'Y'  "&VBCRLF
		strSql = strSql&" 	and (n.dispcate1 is null or n.dispcate1 ='' or (n.dispcate1 is not null and p.catecode is not null) )"&VBCRLF
		strSql = strSql&" ) as T order by fixnotics desc, board_idx desc"
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetMainNotice = rsget.getRows()
		END IF
		rsget.close
	End Function

	'//팝업공지
	public Function fnGetMainPopup()
	dim strSql
	strSql = " select board_idx, title, content "&VBCRLF
    strSql = strSql & " from [db_board].[dbo].tbl_designer_notice WITH(NOLOCK)"&VBCRLF
    strSql = strSql & " where ispopup ='Y' and popsdate <=getdate() and popedate>=getdate() and deleteyn ='N' "&VBCRLF
    strSql = strSql & " order by board_idx desc "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetMainPopup = rsget.getRows()
		END IF
		rsget.close
	End Function

End Class
%>
