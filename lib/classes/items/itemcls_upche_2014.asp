<%
'#################################################### 
' Description :  ��ü���� ��ǰ ����
' History : 2014.03.18 ������  ���� 
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
	
	'��ü��� ��ǰ ����Ʈ(�ٹ�����)
	'/designer/itemmaster/upche_item_requestmodify.asp
		public Function fnGetItemUpcheBaesongList
		Dim strSql
		 
			strSql ="[db_item].[dbo].sp_Ten_item_onlyUpchebaesongListCnt('"&FRectMakerid&"','"&FRectItemid&"','"&FRectItemname&"','"&FRectDispCate&"','"&FRectSellyn&"','"&FRectlimityn&"','"&FRectCheckEX&"')"
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				FTotCnt = rsget(0)
			END IF
			rsget.close

			IF FTotCnt > 0 THEN
			FSPageNo = (FPageSize*(FCurrPage-1)) + 1
			FEPageNo = FPageSize*FCurrPage

			strSql ="[db_item].[dbo].sp_Ten_item_onlyUpchebaesongList('"&FRectMakerid&"','"&FRectItemid&"','"&FRectItemname&"','"&FRectDispCate&"','"&FRectSellyn&"','"&FRectlimityn&"','"&FRectSort&"','"&FRectCheckEX&"',"&FSPageNo&","&FEPageNo&")"
		 
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				fnGetItemUpcheBaesongList = rsget.getRows()
			END IF
			rsget.close
			END IF
	End Function
End Class


%>