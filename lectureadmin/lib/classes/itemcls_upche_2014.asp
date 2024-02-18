<%
'#################################################### 
' Description :  업체어드민 상품 관련
' History : 2014.03.18 정윤정  생성 
'####################################################
 

Class CItem
public FTotCnt
public FSPageNo
public FEPageNo
public FPageSize
public FCurrPage

public FRectMakerid
public FRectItemid
public FRectItemname
public FRectDispCate
public FRectSellyn
public FRectlimityn
public FRectSort
public FSellCash
public FItemCouponYN
public Fitemcoupontype
public Fitemcouponvalue 
public FRectCheckEX

public FRectStartDate
public FRectEndDate
public FRectReqType
public FRectIsFinish
public FRectSortDiv

	'업체배송 상품 리스트(텐배제외)
	'/designer/itemmaster/upche_item_requestmodify.asp
		public Function fnGetItemUpcheBaesongList
		Dim strSql
		 
			strSql ="[db_academy].[dbo].sp_Fingers_item_onlyUpchebaesongListCnt('"&FRectMakerid&"','"&FRectItemid&"','"&FRectItemname&"','"&FRectDispCate&"','"&FRectSellyn&"','"&FRectlimityn&"','"&FRectCheckEX&"')"
			rsACADEMYget.Open strSql, dbACADEMYget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsACADEMYget.EOF OR rsACADEMYget.BOF) THEN
				FTotCnt = rsACADEMYget(0)
			END IF
			rsACADEMYget.close


			IF FTotCnt > 0 THEN
			FSPageNo = (FPageSize*(FCurrPage-1)) + 1
			FEPageNo = FPageSize*FCurrPage

			strSql ="[db_academy].[dbo].sp_Fingers_item_onlyUpchebaesongList('"&FRectMakerid&"','"&FRectItemid&"','"&FRectItemname&"','"&FRectDispCate&"','"&FRectSellyn&"','"&FRectlimityn&"','"&FRectSort&"','"&FRectCheckEX&"',"&FSPageNo&","&FEPageNo&")"
			rsACADEMYget.Open strSql, dbACADEMYget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsACADEMYget.EOF OR rsACADEMYget.BOF) THEN
				fnGetItemUpcheBaesongList = rsACADEMYget.getRows()
			END IF
			rsACADEMYget.close
			END IF
	End Function
	
	

	'//업체배송 상품수정요청 승인상품리스트
	public Function fnGetItemEditRequestList
		Dim strSql

			strSql ="[db_academy].[dbo].sp_Fingers_item_EditReqListCnt('"&FRectMakerid&"','"&FRectItemid&"','"&FRectItemname&"','"&FRectDispCate&"','"&FRectStartDate&"','"&FRectEndDate&"','"&FRectReqType&"','"&FRectIsFinish&"')"
			rsACADEMYget.Open strSql, dbACADEMYget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsACADEMYget.EOF OR rsACADEMYget.BOF) THEN
				FTotCnt = rsACADEMYget(0)
			END IF
			rsACADEMYget.close

			IF FTotCnt > 0 THEN
			FSPageNo = (FPageSize*(FCurrPage-1)) + 1
			FEPageNo = FPageSize*FCurrPage

			strSql ="[db_academy].[dbo].sp_Fingers_item_EditReqList('"&FRectMakerid&"','"&FRectItemid&"','"&FRectItemname&"','"&FRectDispCate&"','"&FRectStartDate&"','"&FRectEndDate&"','"&FRectReqType&"','"&FRectIsFinish&"','"&FRectSortDiv&"',"&FSPageNo&","&FEPageNo&")"
			rsACADEMYget.Open strSql, dbACADEMYget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			''response.write strSql
			IF Not (rsACADEMYget.EOF OR rsACADEMYget.BOF) THEN
				fnGetItemEditRequestList = rsACADEMYget.getRows()
			END IF
			rsACADEMYget.close
			END IF
	End Function
	
End Class


class CUpCheItemEdit
	public FItemList()

	public FResultCount
	public FTotalCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	
	public FTotCnt
	public FSPageNo
	public FEPageNo
	
	public FRectMakerid 
	public FRectItemname
	public FRectDispCate
	public FRectSellyn
	public FRectlimityn
	public FRectSort
	public FSellCash
	public FItemCouponYN
	public Fitemcoupontype
	public Fitemcouponvalue 
	public FRectIsFinish
	
	public FRectDesignerID
	public FRectItemId
	public FRectNotFinish

	public FRectOrderDesc
	public FRectTenBeasongOnly


	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

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
	
	'//업체배송 상품수정요청 결과 리스트
		public Function fnGetItemEditResultList
		Dim strSql
		 
			strSql ="[db_academy].[dbo].sp_Fingers_item_UpcheEditReqListCnt('"&FRectMakerid&"','"&FRectItemid&"','"&FRectItemname&"','"&FRectDispCate&"','"&FRectSellyn&"','"&FRectlimityn&"','"&FRectIsFinish&"')"
			rsACADEMYget.Open strSql, dbACADEMYget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsACADEMYget.EOF OR rsACADEMYget.BOF) THEN
				FTotCnt = rsACADEMYget(0)
			END IF
			rsACADEMYget.close

			IF FTotCnt > 0 THEN
			FSPageNo = (FPageSize*(FCurrPage-1)) + 1
			FEPageNo = FPageSize*FCurrPage

			strSql ="[db_academy].[dbo].sp_Fingers_item_UpcheEditReqList('"&FRectMakerid&"','"&FRectItemid&"','"&FRectItemname&"','"&FRectDispCate&"','"&FRectSellyn&"','"&FRectlimityn&"','"&FRectIsFinish&"','"&FRectSort&"',"&FSPageNo&","&FEPageNo&")"
			rsACADEMYget.Open strSql, dbACADEMYget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsACADEMYget.EOF OR rsACADEMYget.BOF) THEN
				fnGetItemEditResultList = rsACADEMYget.getRows()
			END IF
			rsACADEMYget.close
			END IF
	End Function
end Class



Function fnGetReqStatus(ByVal isFinish)
 	IF isFinish = "N" THEN
 		fnGetReqStatus = "승인대기"
 	ELSEIF isFinish = "D" THEN
 		fnGetReqStatus = "<font color=red>반려</font>"
	ELSEIF isFinish ="Y" THEN
		fnGetReqStatus = "<font color=blue>승인</font>"
	END IF
End Function
%>