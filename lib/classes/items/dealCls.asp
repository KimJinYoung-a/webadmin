<%
Class CDealPrdItem
	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

	Public FDealCode
	Public FDealItemid
	Public FMasterItemCode
	Public FViewDIV
	Public FStartDate
	Public FEndDate
	Public FMasterSellCash
	Public FMasterDiscountRate
	Public FPricesDash
	Public FSailsDash
	Public FisJust1day

End Class

Class DealCls

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()
	End Sub

	dim Prd
	dim FResultCount
	dim FADD

	Public Function IsImageBasic(ByVal ItemID, ByVal BasicImage)
	    IsImageBasic = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(ItemID) + "/" + BasicImage
	End Function

	Public Sub GetIDealInfo(ByVal iid)


		dim strSQL, vIsTest
		IF application("Svr_Info") = "Dev" THEN
			'vIsTest = "test"
		Else
			vIsTest = ""
		End If

		strSQL = "execute [db_event].[dbo].sp_Ten_Dealinfo_New @vItemID ='" & CStr(iid) & "'"

		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		set Prd = new CDealPrdItem

		if  not rsget.EOF  then

			FResultCount = 1
			rsget.Movefirst
				Prd.FDealCode    			= rsget("idx")	'딜 관리코드
				Prd.FDealItemid				= rsget("dealitemid")	'딜 상품 코드
				Prd.FMasterItemCode		= rsget("masteritemcode")	'딜 대표 상품 코드
				Prd.FViewDIV					= rsget("viewdiv")	'전시 구분(상시/기간제)
				Prd.FStartDate				= rsget("startdate")	'시작일
				Prd.FEndDate					= rsget("enddate")	'종료일
				Prd.FMasterSellCash 		= rsget("mastersellcash")	'대표가격
				Prd.FMasterDiscountRate	= rsget("masterdiscountrate")	'대표 할인율
				Prd.FPricesDash 			= rsget("pricesdash")	'가격 물결표시
				Prd.FSailsDash 				= rsget("sailsdash")	'할인율 물결표시
				Prd.FisJust1day 				= rsget("isJust1day")	'저스트 원데이 진행 여부
		else
			FResultCount = 0
		end if

		rsget.close

	End Sub

    Public function GetDealItemList(byval masteridx)
        dim strSQL
        strSQL = "exec [db_event].[dbo].[sp_Ten_DealItemList] " & masteridx & ""
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF) then
			GetDealItemList = rsget.getRows
		end if
		rsget.Close
	End Function

    Public function GetDealItemEvalList(byval masteridx)
        dim strSQL
        strSQL = "exec [db_event].[dbo].[sp_Ten_DealItemEvalList] " & masteridx & ""
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF) then
			GetDealItemEvalList = rsget.getRows
		end if
		rsget.Close
	End Function

    Public function GetDealItemQNAList(byval masteridx)
        dim strSQL
        strSQL = "exec [db_event].[dbo].[sp_Ten_DealItemQNAList] " & masteridx & ""
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF) then
			GetDealItemQNAList = rsget.getRows
		end if
		rsget.Close
	End Function

    Public function GetDealItemCouponList(byval masteridx)
        dim strSQL
        strSQL = "exec [db_event].[dbo].[sp_Ten_DealItemCoponList] " & masteridx & ""
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF) then
			GetDealItemCouponList = rsget.getRows
		end if
		rsget.Close
	End Function
End Class
%>