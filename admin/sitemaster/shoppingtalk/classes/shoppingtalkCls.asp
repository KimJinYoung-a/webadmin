<%
'###########################################################
' Description : GIFT TALK class
' Hieditor : 강준구 생성
'			 2022.07.08 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################

class CShoppingTalkItem
	public FItemID
	public FItemName
	public FSellcash
	public FOrgPrice
	public FMakerID
	public FBrandName
	public FBrandName_kor
	public FBrandLogo
	public FMakerName
	public FcdL
	public FcdM
	public FcdS
	public FCateName
	public FImageBasic
	public FImageList
	public FImageList120
	public FImageSmall
	public FImageBasicIcon
	public FImageIcon1
	public FImageIcon2
	public FTalkIdx
	public FUserID
	public FTheme
	public FKeyword
	public FItem
	public FContents
	public FUseYN
	public FRegdate
	public FCommCnt
	public FIsNewComm
	public FIdx
	public FTag
	public FDepth
	public FCode
	public FCodename
	public FSortNo

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


Class CShoppingTalk
	public FItemList()
	public FOneItem
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectGubun
	public FRectIdx
	public FRectTalkIdx
	public FRectItemId
	public FRectUserId
	public FRectTheme
	public FRectUseYN
	public FRectGoodBad
	public FRectContents
	public FRectKeyword
	public FRectOnlyCount
	public FRectDiv
	public FRectSort
	public FRectDepth
	public FRectCode
	
    	    
	Private Sub Class_Initialize()
		redim preserve FItemList(0)
		FCurrPage =1
		FPageSize = 10
		FResultCount = 0
		FScrollCount = 10
		FTotalCount = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	'####### talk 리스트 -->
	public Function fnShoppingTalkList
		Dim strSql, i, vKey

		If FRectKeyword <> "" Then
			FRectKeyword = Replace(FRectKeyword," ","")
			For i = LBound(Split(FRectKeyword,",")) To UBound(Split(FRectKeyword,","))
				vKey = vKey & "''" & Split(FRectKeyword,",")(i) & "'',"
			Next
			If vKey <> "" Then
				If Right(vKey,1) = "," Then
					FRectKeyword = Left(vKey,Len(vKey)-1)
				End IF
			End If
		End If
		FResultCount = 0
		
			strSql = "EXECUTE [db_board].[dbo].[sp_Ten_GiftTalk_List_Count] '" & FpageSize & "', '" & FRectUserId & "', '" & FRectItemId & "', '" & FRectTheme & "', '" & FRectKeyword & "', '" & FRectUseYN & "', '" & FRectDiv & "'"
			'response.write strSql
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
			rsget.Open strSql,dbget,1
				FTotalCount = rsget(0)
				FTotalPage	= rsget(1)
			rsget.close
			
			If FRectTalkIdx <> "" Then
				FTotalCount = 1
				FTotalPage = 1
			End IF
		

		If FTotalCount > 0 AND FRectOnlyCount = "" Then
			If FRectSort = "" Then
				FRectSort = "t.talk_idx DESC"
			Else
				If FRectSort = "1" Then
					FRectSort = "t.talk_idx DESC"
				ElseIf FRectSort = "2" Then
					FRectSort = "t.view_cnt DESC, t.talk_idx DESC"
				ElseIf FRectSort = "3" Then
					FRectSort = "t.comm_cnt DESC, t.talk_idx DESC"
				Else
					FRectSort = "t.talk_idx DESC"
				End If
			End IF
	
			strSql = "EXECUTE [db_board].[dbo].[sp_Ten_GiftTalk_List] '" & (FpageSize*FCurrPage) & "', '" & FRectTalkIdx & "', '" & FRectUserId & "', '" & FRectItemId & "', '" & FRectTheme & "', '" & FRectKeyword & "', '" & FRectUseYN & "', '" & FRectDiv & "', '" & FRectSort & "'"
			'response.write strSql
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
			rsget.pagesize = FPageSize
			rsget.Open strSql,dbget,1
			'response.write strSql

			FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
	        if (FResultCount<1) then FResultCount=0
			redim preserve FItemList(FResultCount)
			
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CShoppingTalkItem
	
					
					FItemList(i).FTalkIdx	= rsget("talk_idx")
					FItemList(i).FUserID	= rsget("userid")
					FItemList(i).FTheme		= rsget("theme")
					'FItemList(i).FKeyword	= rsget("keyword")
					FItemList(i).FItem		= rsget("item")
					'FItemList(i).FImageBasic	= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("basicimage")
					
					FItemList(i).FContents	= db2html(rsget("contents"))
					FItemList(i).FUseYN		= rsget("useyn")
					FItemList(i).FRegdate	= rsget("regdate")
					FItemList(i).FCommCnt	= rsget("comm_cnt")
					FItemList(i).FIsNewComm	= rsget("isnewcomm")
					FItemList(i).FTag		= rsget("tag")

					i=i+1
					rsget.moveNext
				loop
			end if
			rsget.close
		End If
	End Function

	'####### 나의 아이템 리스트 -->
	public Function fnShoppingTalkMyItemList
		Dim strSql, i
		
		strSql = "EXECUTE [db_board].[dbo].[sp_Ten_ShoppingTalk_MyItemList_Count] '" & FpageSize & "', '" & FRectUserId & "'"
		'response.write strSql
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql,dbget,1
			FTotalCount = rsget(0)
			FTotalPage	= rsget(1)
		rsget.close
		

		If FTotalCount > 0 Then
			strSql = "EXECUTE [db_board].[dbo].[sp_Ten_ShoppingTalk_MyItemList] '" & FpageSize & "', '" & FRectUserId & "'"
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
			rsget.pagesize = FPageSize
			rsget.Open strSql,dbget,1
			'response.write strSql

			FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
	        if (FResultCount<1) then FResultCount=0
			redim preserve FItemList(FResultCount)
			
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CShoppingTalkItem
	
					FItemList(i).FItemID		= rsget("itemid")
					FItemList(i).FItemName		= db2html(rsget("itemname"))
					FItemList(i).FImageBasic	= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("basicimage")
					FItemList(i).FImageIcon1	= "http://webimage.10x10.co.kr/image/icon1/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("icon1image")
					FItemList(i).FBrandName		= db2html(rsget("brandname"))
					FItemList(i).FMakerID		= rsget("makerid")
					
					i=i+1
					rsget.moveNext
				loop
			end if
			rsget.close
		End If
	End Function
	
	
	'####### talk comment 리스트 -->
	public Function fnShoppingTalkCommList
		Dim strSql, i
		
			strSql = "EXECUTE [db_board].[dbo].[sp_Ten_ShoppingTalk_CommList_Count] '" & FpageSize & "', '" & FRectTalkIdx & "', '" & FRectUserId & "', '" & FRectUseYN & "'"
			'response.write strSql
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
			rsget.Open strSql,dbget,1
				FTotalCount = rsget(0)
				FTotalPage	= rsget(1)
			rsget.close
			

		If FTotalCount > 0 Then
			strSql = "EXECUTE [db_board].[dbo].[sp_Ten_ShoppingTalk_CommList] '" & (FpageSize*FCurrPage) & "', '" & FRectTalkIdx & "', '" & FRectUserId & "', '" & FRectUseYN & "'"
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
			rsget.pagesize = FPageSize
			rsget.Open strSql,dbget,1
			'response.write strSql

			FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
	        if (FResultCount<1) then FResultCount=0
			redim preserve FItemList(FResultCount)
			
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CShoppingTalkItem
	
					
					FItemList(i).FIdx		= rsget("idx")
					FItemList(i).FUserID	= rsget("userid")
					FItemList(i).FContents	= db2html(rsget("contents"))
					FItemList(i).FUseYN		= rsget("useyn")
					FItemList(i).FRegdate	= rsget("regdate")

					i=i+1
					rsget.moveNext
				loop
			end if
			rsget.close
		End If
	End Function
	
	
	public Function fnShoppingTalkCodeList
		Dim strSql, i, subSql
		
		If FRectCode <> "" Then
			subSql = " AND Left(code," & Len(FRectCode) & ") = '" & FRectCode & "' "
		End If
		
		strSql = "SELECT depth, code, codename, sortno, useyn FROM [db_board].[dbo].[tbl_shopping_talk_keywordcode] WHERE depth = '" & FRectDepth & "' " & subSql & " ORDER BY sortno ASC, code ASC"
		rsget.Open strSql,dbget,1

		if  not rsget.EOF  then
			fnShoppingTalkCodeList = rsget.getRows()
		end if
		rsget.close
		
	End Function
	
	
	public Function fnShoppingTalkCodeDetail
		Dim strSql, i, subSql

		strSql = "SELECT depth, code, codename, sortno, useyn FROM [db_board].[dbo].[tbl_shopping_talk_keywordcode] WHERE code = '" & FRectCode & "'"
		rsget.Open strSql,dbget,1

		if  not rsget.EOF  then
			set FOneItem = new CShoppingTalkItem
			FOneItem.FDepth = rsget("depth")
			FOneItem.FCode = rsget("code")
			FOneItem.FCodename = rsget("codename")
			FOneItem.FSortNo = rsget("sortno")
			FOneItem.FUseYN = rsget("useyn")

		end if
		rsget.close
		
	End Function
	
	
	
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class


Function keywordSelectBox(key1, key2)
	Dim i, vBody, vQuery
	vBody = vBody & "	<option value="""">상황 선택하기</option>" & vbCrLf
	
	vQuery = "SELECT code, codename FROM [db_board].[dbo].[tbl_shopping_talk_keywordcode] WHERE depth = '1' ORDER BY sortno ASC, code ASC"
	rsget.Open vQuery,dbget,1
	
	Do Until rsget.Eof
		vBody = vBody & "	<option value=""" & rsget("code") & """ " & CHKIIF(key1=rsget("code"),"selected","") & ">" & rsget("codename") & "</option>" & vbCrLf
		rsget.MoveNext
	Loop
	rsget.close()

	Response.Write vBody
End Function
%>