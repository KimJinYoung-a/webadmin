<%
'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3
'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

Class ClsDiaryItem


	Public FIdx
	Public FYear
	Public FItemname
	Public FDiaryType
	Public FItemid
	public FBasicImg
	Public FListimg
	Public FIconImg
	Public FIsusing

	public FgiftYn
	public FonlyYearYn
	public FhitYn


	public Function StrDiaryTypeName()

		if FDiaryType="illust" then
			StrDiaryTypeName="일러스트"
		elseif FDiaryType = "photo" then
			StrDiaryTypeName="포토/명화"
		elseif FDiaryType = "simple" then
			StrDiaryTypeName="기능/심플"
		elseif FDiaryType = "system" then
			StrDiaryTypeName="시스템"
		end if
	End Function

	Public Function getBasicImgUrl()
		getBasicImgUrl = "http://webimage.10x10.co.kr/diary_collection/" & FYear & "/basic/" & FBasicImg
	End Function
	Public Function getListImgUrl()
		getListImgUrl = "http://webimage.10x10.co.kr/diary_collection/" & FYear & "/list/" & FListimg
	End Function
	Public Function getIconImgUrl()
		getIconImgUrl = "http://webimage.10x10.co.kr/diary_collection/" & FYear & "/icon/" & FIconImg
	End Function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

End Class

Class clsDiaryContensItem
	public FYear
	public ConIdx
	public ConfImg
	public ConTTxt

	Public Function getContImgUrl()
		getContImgUrl = "http://webimage.10x10.co.kr/diary_collection/" & FYear & "/Cont/" & ConfImg
	End Function

End Class

CLASS clsDiaryInfoItem
	public FYear
	public FInfoidx
	public FinfoGubun
	public Finfoname
	public Finfoimg
	public FinfoPageCnt

	public Function getInfoImgUrl()
		getInfoImgUrl = "http://webimage.10x10.co.kr/diary_collection/" & FYear & "/info/" & Finfoimg
	End Function
End Class

CLASS clsDiaryLinkedItem

	public FItemid
	public FItemName

	public Function getInfoImgUrl()
		getInfoImgUrl = "http://webimage.10x10.co.kr/diary_collection/" & FYear & "/info/" & Finfoimg
	End Function
End Class

Class ClsDiary



	Private Sub Class_Initialize()
		FYearUse = "2008"
	End Sub

	Private Sub Class_Terminate()

	End Sub


	public FYearUse
	public DiaryPrd
	Public FItemList()
	Public FTotalCount
	Public FTotalPage
	Public FResultCount


	Public FDiaryType
	Public FCurrPage
	Public FPageSize
	Public FScrollCount
	Public FOrderType


	'// 다이어리 리스트
	Public Sub GetDiaryList()
		dim strSQL

		strSQL =" SELECT count(*) as Totalcnt ,CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ")  as TotalPage "&_
				" FROM [db_diary_collection].[dbo].[tbl_diary_master] " &_
				" WHERE yearuse ='" & FYearUse & "'"

			if FDiaryType<>"" then
				strSQL = strSQL + " and diarytype='" & CStr(FDiaryType) & "'"
			end if

			rsget.open strSQL,dbget,1

			if not rsget.eof then
				FTotalCount=rsget("Totalcnt")
				FTotalPage = rsget("TotalPage")
			end if

			rsget.close


		strSQL =" SELECT TOP " & FPageSize &_
				" m.idx, m.yearuse ,m.diaryType, m.itemid, m.isusing, i.itemname ,m.list_img, m.icon_img" &_
				" FROM [db_diary_collection].[dbo].tbl_diary_master m "&_
				" LEFT JOIN [db_item].[10x10].tbl_item i on m.itemid=i.itemid "&_
				" WHERE yearuse ='" & FYearUse & "'  and m.idx < "&_
				" 		(SELECT ISNULL(min(t.idx) ,99999999) FROM "&_
				" 			(SELECT top " & FPageSize*(FCurrPage-1) & " idx FROM db_diary_collection.dbo.tbl_diary_master WHERE yearuse ='" & FYearUse & "'"

			if FDiaryType<>"" then
				strSQL = strSQL + " 			and diaryType='" & CStr(FDiaryType) & "'"
			end if

				strSQL = strSQL + " 		ORDER BY idx desc ) as t ) "

			if FDiaryType<>"" then
				strSQL = strSQL + " and m.diaryType='" & CStr(FDiaryType) & "'"
			end if

				strSQl = strSQL & " ORDER BY idx desc "

			'response.write strSQL
			rsget.open strSQL,dbget,1



		If  not rsget.EOF  Then
			FResultCount = rsget.recordcount

			redim preserve FItemList(FResultCount)
			i=0

			Do Until rsget.eof
				set FItemList(i) = new ClsDiaryItem

					FItemList(i).FIdx				= rsget("idx")
					FItemList(i).FYear		= rsget("yearuse")
					FItemList(i).FItemname	= rsget("itemname")
					FItemList(i).FItemid		=	rsget("itemid")
					FItemList(i).FDiaryType			= rsget("diaryType")
					FItemList(i).FListimg		= rsget("list_img")
					FItemList(i).FIconImg		= rsget("icon_img")
					FItemList(i).FIsusing		= rsget("isusing")

				i=i+1
				rsget.Movenext

			Loop

		End If

		rsget.close


	End Sub
	''// 다이어리 상세
	Public Sub getDiaryItem(byval idx)

		dim strSQL
		strSQL =" SELECT idx ,yearuse, diaryType ,Itemid ,basic_img ,List_img ,icon_img ,isusing ,giftYn ,onlyYearYn ,hitYn " &_
				" FROM [db_diary_collection].dbo.tbl_diary_master " &_
				" WHERE idx =" & idx

		rsget.open strSQL,dbget,1
		if not rsget.eof then

			set DiaryPrd = new ClsDiaryItem

			DiaryPrd.FIdx = rsget("idx")
			DiaryPrd.FYear = rsget("yearuse")
			DiaryPrd.FDiaryType = rsget("diaryType")
			DiaryPrd.FItemid = rsget("Itemid")
			DiaryPrd.FBasicImg = rsget("basic_img")
			DiaryPrd.FListimg = rsget("list_img")
			DiaryPrd.FIconImg = rsget("icon_img")
			DiaryPrd.Fisusing = rsget("isusing")
			DiaryPrd.FgiftYn = rsget("giftYn")
			DiaryPrd.FonlyYearYn = rsget("onlyYearYn")
			DiaryPrd.FhitYn = rsget("hitYn")

		end if
		rsget.close

	End Sub
	'// 다이어리 상세설명
	Public Function getDiaryContens(byval idx)
		dim strSQL,i
		strSQL =" execute db_diary_collection.dbo.ten_diary_contents @idx='" & idx & "'"

		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenForwardOnly
		rsget.LockType = adLockReadOnly

		rsget.open strSQL,dbget,1

		FResultCount = rsget.recordcount

		if not rsget.eof then

			redim preserve FItemList(FResultCount)
			i=0
			do until rsget.eof
				set FItemList(i) = new clsDiaryContensItem
				FItemList(i).FYear = FYearUse
				FItemList(i).ConIdx = rsget("cont_idx")
				FItemList(i).ConfImg = rsget("cont_file")
				FItemList(i).ConTTxt = rsget("cont_text")
				rsget.movenext
				i = i+1
			loop
		end if
		rsget.close

	End Function
	'// 다이어리 내지구성
	public Function getDiaryInfo(byval idx)
		dim strSQL,i

		strSQL =" execute db_diary_collection.dbo.ten_diary_info @idx='" & idx & "'"

		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenForwardOnly
		rsget.LockType = adLockReadOnly

		rsget.open strSQL,dbget

		FResultCount = rsget.recordcount

		if not rsget.eof then

			redim preserve FItemList(FResultCount)
			i=0
			do until rsget.eof
				set FItemList(i) = new clsDiaryInfoItem
				FItemList(i).FYear = FYearUse
				FItemList(i).FInfoidx = rsget("Info_idx")
				FItemList(i).FInfoGubun = rsget("info_gubun")
				FItemList(i).Finfoname = db2html(rsget("info_name"))
				FItemList(i).Finfoimg = rsget("info_img")
				FItemList(i).FinfoPageCnt = rsget("info_PageCnt")
				rsget.movenext
				i = i+1
			loop
		end if
		rsget.close


	End Function
	'// 다이어리 관련 상품

	public Function getDiaryLinkedItem(byval idx)
		dim strSQL,i

		strSQL =" execute db_diary_collection.dbo.ten_diary_Linkeditem @idx='" & idx & "'"

		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenForwardOnly
		rsget.LockType = adLockReadOnly

		rsget.open strSQL, dbget
		FResultCount = rsget.recordcount

		if not rsget.eof then
			redim preserve FItemList(FResultCount)
			i=0
			do until rsget.eof
				set FItemList(i) = new clsDiaryLinkedItem

				FItemList(i).FItemid = rsget("itemid")
				FItemList(i).FItemName = rsget("itemname")

				rsget.movenext
				i = i+1
			loop

		end if

		rsget.close
	End Function

	public Function getMdsPickList()
		dim strSQL

		strSQL =" SELECT top 10 diaryid,pickrank " &_
				" FROM db_diary_collection.dbo.tbl_diary_mdpick "&_
				" ORDER BY pickrank"

		rsget.open,strSQL,dbget,1

		if not rsget.eof then
			getMdsPickList = rsget.getRows()
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
		StartScrollPage = ((FCurrPage-1)\FScrollCount)*FScrollCount +1
	end Function


End Class

%>
