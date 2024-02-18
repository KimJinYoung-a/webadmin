<%
'####################################################
' Description : 상품 히스토리 클래스
' History : 2020.08.17 이상구 생성
'####################################################

Class CItemHistoryItem
    dim Fregdate
    dim FItemID
    dim Fitemoption
    dim Fsellcash
    dim Fbuycash
    dim Fsellyn
    dim Flimityn
    dim Fitemrackcode
    dim Fmwdiv
    dim Fdeliverytype
    dim Fitemname
    dim Foptioncnt
    dim Fcitemcouponidx
    dim Fbrandname
    public flimitno
    public flimitsold

	Private Sub Class_Initialize()
        ''
	End Sub

	Private Sub Class_Terminate()
        ''
	End Sub
End Class

class CItemHistory
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

    public FRectItemId

    public Sub getItemHistoryList()
        dim sqlStr, i, addSql

        if (FRectItemId <> "") then
            addSql = addSql & " and itemid = " & FRectItemId
        end if

        sqlStr = ""
        sqlStr = sqlStr & " select top " & FPageSize & " h.* "
        sqlStr = sqlStr & " from [db_log].[dbo].[tbl_iteminfo_history] h with (nolock)"
        sqlStr = sqlStr & " where 1=1 "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by regdate desc "

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		Ftotalcount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
    		do until rsget.eof
    			set FItemList(i) = new CItemHistoryItem

    			FItemList(i).Fregdate 		= rsget("regdate")
                FItemList(i).Fitemid 		= rsget("itemid")
                FItemList(i).Fitemoption 	= rsget("itemoption")
                FItemList(i).Fsellcash 		= rsget("sellcash")
                FItemList(i).Fbuycash 		= rsget("buycash")
                FItemList(i).Fsellyn 		= rsget("sellyn")
                FItemList(i).Flimityn 		= rsget("limityn")
                FItemList(i).Fitemrackcode 	= rsget("itemrackcode")
                FItemList(i).Fmwdiv 		= rsget("mwdiv")
                FItemList(i).Fdeliverytype 	= rsget("deliverytype")
                FItemList(i).Fitemname 		= rsget("itemname")
                FItemList(i).Foptioncnt 	= rsget("optioncnt")
                FItemList(i).Fcitemcouponidx = rsget("citemcouponidx")
                FItemList(i).Fbrandname 	= rsget("brandname")
                FItemList(i).flimitno 	= rsget("limitno")
                FItemList(i).flimitsold 	= rsget("limitsold")

    		    rsget.movenext
    			i=i+1
    		loop
		end if
		rsget.close
    End Sub

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 10
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

%>
