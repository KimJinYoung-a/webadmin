<%

Class ClsDiaryItem


	Public FIdx
	Public FItemname
	Public FType
	Public FItemid
	Public FListimg
	Public FIconImg
	Public FIsusing


	public Function StrDiaryTypeName()

		if FType="illust" then
			StrDiaryTypeName="일러스트"
		elseif FType = "photo" then
			StrDiaryTypeName="포토/명화"
		elseif FType = "simple" then
			StrDiaryTypeName="기능/심플"
		end if
	End Function
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

End Class


Class ClsDiary

	public FYearUse
	Public FItemList()
	Public FTotalCount
	Public FTotalPage
	Public FResultCount


	Public FType
	Public FCurrPage
	Public FPageSize
	Public FScrollCount
	Public FOrderType


	'// diary_reg_MainList.asp

	Public Sub GetDiaryList()
		dim strSQL

		strSQL =" SELECT count(*) as Totalcnt ,CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ")  as TotalPage "&_
				" FROM [db_diary_collection].[dbo].[tbl_diary_master] " &_
				" WHERE Year ='" & FYearUse & "'"

			if FType<>"" then
				strSQL = strSQL + " where type='" & CStr(FType) & "'"
			end if

			rsget.open strSQL,dbget,1

			if not rsget.eof then
				FTotalCount=rsget("Totalcnt")
				FTotalPage = rsget("TotalPage")
			end if

			rsget.close


		strSQL =" SELECT TOP " & FPageSize &_
				" m.idx, m.type, m.itemid, m.isusing, i.itemname ,m.list_img, m.icon_img" &_
				" FROM [db_diary_collection].[dbo].tbl_diary_master m "&_
				" LEFT JOIN [db_item].[10x10].tbl_item i on m.itemid=i.itemid "&_
				" WHERE m.idx > "&_
				" 		(SELECT ISNULL(max(t.idx) ,0) FROM "&_
				" 			(SELECT top " & FPageSize*(FCurrPage-1) & " idx FROM db_diary_collection.dbo.tbl_diary_master "

			if FType<>"" then
				strSQL = strSQL + " 			WHERE type='" & CStr(FType) & "'"
			end if

				strSQL = strSQL + " 		ORDER BY idx ) as t ) "

			if FType<>"" then
				strSQL = strSQL + " and m.type='" & CStr(FType) & "'"
			end if

			strSQL = strSQL + " ORDER BY m.idx	"

			response.write strSQL
			rsget.open strSQL,dbget,1



		If  not rsget.EOF  Then
			FResultCount = rsget.recordcount

			redim preserve FItemList(FResultCount)
			i=0

			Do Until rsget.eof
				set FItemList(i) = new ClsDiaryItem

					FItemList(i).FIdx				= rsget("idx")
					FItemList(i).FItemname	= rsget("itemname")
					FItemList(i).FItemid		=	rsget("itemid")
					FItemList(i).FType			= rsget("type")
					FItemList(i).FListimg		= "http://testimgstatic.10x10.co.kr/contents/diary/list/" & rsget("list_img")
					FItemList(i).FIconImg		= "http://testimgstatic.10x10.co.kr/contents/diary/icon/" & rsget("icon_img")
					FItemList(i).FIsusing		= rsget("isusing")

				i=i+1
				rsget.Movenext

			Loop

		End If

		rsget.close


	End Sub


	Private Sub Class_Initialize()

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
		StartScrollPage = ((FCurrPage-1)\FScrollCount)*FScrollCount +1
	end Function


End Class

%>
