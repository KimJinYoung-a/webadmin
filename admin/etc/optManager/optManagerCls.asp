<%
Class COptManagerItem
	Public Fitemid
	Public FItemoption
	Public FMakerid	
	Public FItemname
	Public FMallid
	Public FNewCode
	Public FOptionname
	Public FMallItemname
	Public FSellcash	
	Public FOptaddprice
	Public FMallCash

	Public FIdx
	Public FSmallImage 
	Public FRegdate
	Public FLastUpdate
	Public FOrgPrice
	Public FBuycash 
	Public FSellYn
	Public FSaleyn
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public FDeliverytype
	Public FCateMapCnt 
	Public FNewitemname
	Public FMallRegdate
	Public FMallLastUpdate
	Public FMallGoodNo
	Public FMallPrice
	Public FMallSellYn
	Public FRegUserid
	Public FMallStatCd 
	Public FRctSellCNT
	Public FAccFailCNT
	Public FLastErrStr
	Public FDivcode 
	Public FOptlimitno
	Public FOptlimitsold
	Public FOptsellyn
	Public FSafecode
	Public FIsvat 
	Public FInfoDiv
	Public FSafeCertGbnCd 
	Public FDeliveryCd
	Public FDeliveryAddrCd
	Public FBrandcd
	Public FItemdiv
	Public FDefaultfreeBeasongLimit 
End Class

Class COptManager
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public FRectMallid
	Public FRectNotMallid
	Public FRectMakerid
	Public FRectItemid
	Public FRectIsReged

	Public FRectCDL				
	Public FRectCDM				
	Public FRectCDS				
	Public FRectItemName			
	Public FRectSellYn				
	Public FRectLimitYn			
	Public FRectSailYn				
	Public FRectonlyValidMargin	
	Public FRectMallgoodno		
	Public FRectMatchCate			
	Public FRectPrdDivMatch		
	Public FRectIsMadeHand			
	Public FRectIsOption			
                       
	Public FRectExtNotReg			
	Public FRectExpensive10x10		
	Public FRectdiffPrc			
	Public FRectMallYes10x10No	
	Public FRectMallNo10x10Yes	
	Public FRectExtSellYn			
	Public FRectInfoDiv			
	Public FRectFailCntOverExcept	
	Public FRectFailCntExists		
	Public FRectReqEdit			
	Public FRectOrdType

	Private Sub Class_Initialize()
		Redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function

	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function

	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function

	Public Function getoOptManagerItemList()
		Dim i, sqlStr, addSql

		'상품코드 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            End If
        End If

		sqlStr = ""
		sqlStr = sqlStr & " EXEC db_etcmall.[dbo].[sp_optmanager_Cnt] '"&FRectMallid&"', '"&FRectMakerid&"', '"&FRectItemid&"', '"&FRectCDL&"', '"&FRectCDM&"', '"&FRectCDS&"', '"&FRectNotMallid&"' "
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		If FTotalCount < 1 then exit Function

		sqlStr = ""
		sqlStr = sqlStr & " EXEC db_etcmall.[dbo].[sp_optmanager_List] '"&FRectMallid&"' , '" & FCurrPage * FPageSize & "', '"&FRectMakerid&"', '"&FRectItemid&"', '"&FRectCDL&"', '"&FRectCDM&"', '"&FRectCDS&"', '"&FRectNotMallid&"' "
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
        rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		rsget.absolutepage = FCurrPage
		IF not rsget.EOF THEN
			getoOptManagerItemList = rsget.getRows()
		End IF
		rsget.Close
	End Function
End Class
%>