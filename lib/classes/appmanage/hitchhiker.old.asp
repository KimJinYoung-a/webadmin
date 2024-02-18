<%
'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// APIs
'//

Dim IMSI_webImgUrl : IMSI_webImgUrl = "http://webimage.10x10.co.kr/"
Dim IsDev2RealImg : IsDev2RealImg = true

' 책 목록 얻기
Function API_GetBookList()
	API_GetBookList = GetBookListJSON(Factory.Create("Args"))
End Function

' 책 얻기
Function API_GetBook(sBookId)
	API_GetBook = GetBookJSON(Factory.Create("Args").SetArgs(Array( _
		"bookid", sBookId _
	)))
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// JSON Functions
'//

' 책 목록 얻기
Function GetBookListJSON(oArgs)
	Dim oBookListRS : Set oBookListRS = GetBookListRS(Factory.Create("Args").SetArgs(Array( _
		"pageSize", 10, _
		"pageNum", 1 _
	)))

	Dim oBookList : Set oBookList = Server.CreateObject("System.Collections.ArrayList")

	Do Until oBookListRS.EOF
		Dim sBookId : sBookId = oBookListRS("idx")
		Dim sVol : sVol = oBookListRS("vol")
		Dim sTopic : sTopic = oBookListRS("topic")
		Dim sThumbnail : sThumbnail = oBookListRS("thumbnail")
		Dim sPublishedDate : sPublishedDate = oBookListRS("publisheddate")

		Dim oBook : Set oBook = jsObject()
		oBook("bookid") = sBookId
		oBook("vol") = sVol
		oBook("topic") = sTopic
		oBook("thumbnail") = sThumbnail
		oBook("publisheddate") = sPublishedDate
		oBookList.Add oBook

		oBookListRS.MoveNext
	Loop

	GetBookListJSON = jsArray().RenderArray(oBookList.ToArray, 1, "")
End Function

Function GetBookJSON(oArgs)
	Dim oBookRS : Set oBookRS = GetBookRS(oArgs)
	Dim oBook : Set oBook = jsObject()
	oBook("bookid") = oBookRS("idx")
	oBook("topic") = oBookRS("topic")
	oBook("vol") = oBookRS("vol")
	oBook("publisheddate") = oBookRS("publisheddate")
	oBookRS.close ' -_-

	Dim oRS : Set oRS = GetPageListRS(oArgs)

	' tree 형태로 가공 - page 아래에 item
	Dim oPages : Set oPages = Server.CreateObject("System.Collections.ArrayList")
	Dim oCurrentPage : Set oCurrentPage = jsObject()
	Dim oCurrentPageItems

	Do Until oRS.EOF
		Dim sPageId : sPageId = oRS("pageid")
		Dim sImg : sImg = oRS("img")
		Dim nSeq : nSeq = oRS("seq")
		Dim sItemId : sItemId = oRS("itemid")
		Dim sItemName : sItemName = oRS("itemname")
		'Dim sItemImage : sItemImage = IMSI_webImgUrl & "/image/basic/" & GetImageSubFolderByItemid(oRS("itemid")) & "/" & oRS("basicimage")
		'Dim sItemImage : sItemImage = "http://testwebimage.10x10.co.kr/"' & "/image/basic/" & GetImageSubFolderByItemid(oRS("itemid")) & "/" & oRS("basicimage")
		'Dim sItemImage : sItemImage = IMSI_webImgUrl & "/image/basic/" & GetImageSubFolderByItemid(oRS("itemid")) & "/" & oRS("basicimage")
		Dim sItemImage : sItemImage = "http://webimage.10x10.co.kr/image/list/" & GetImageSubFolderByItemid(oRS("itemid")) & "/" & oRS("listimage")
		
		if (IsDev2RealImg) then
		    sImg = replace(sImg,"testwebimage","webimage")
	    end if
	    
		If oCurrentPage("pageid") <> sPageId Then
			Set oCurrentPage = jsObject()
			oCurrentPage("pageid") = sPageId
			oCurrentPage("img") = sImg
			oCurrentPage("seq") = nSeq
			oPages.Add oCurrentPage

			Set oCurrentPageItems = Server.CreateObject("System.Collections.ArrayList")
		End If

		' 페이지 관련 Item이 없을 수 있으므로 null 체크
		If Not IsNull(sItemId) Then
			Dim oItem : Set oItem = jsObject()
			oItem("itemid") = sItemId
			oItem("itemname") = sItemName
			oItem("itemimage") = sItemImage
			oCurrentPageItems.Add oItem

			oCurrentPage("items") = oCurrentPageItems.ToArray
		End If

		oRS.MoveNext
	Loop

	oBook("pages") = oPages.ToArray
	GetBookJSON = ToJSON(oBook)
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// Obj Functions
'//

' 페이지 목록 얻기
Function GetPageListObj(oArgs)
	Dim oRS : Set oRS = GetPageListRS(oArgs)

	' tree 형태로 가공 - page 아래에 item
	Dim oPages : Set oPages = Server.CreateObject("System.Collections.ArrayList")
	Dim oCurrentPage : Set oCurrentPage = Server.CreateObject("Scripting.Dictionary")

	Do Until oRS.EOF
		Dim sPageId : sPageId = oRS("pageid")
		Dim sImg : sImg = oRS("img")
		Dim nSeq : nSeq = oRS("seq")
		Dim sItemId : sItemId = oRS("itemid")
		Dim sItemName : sItemName = oRS("itemname")
		Dim sItemImage : sItemImage = IMSI_webImgUrl & "/image/list/" + GetImageSubFolderByItemid(oRS("itemid")) + "/" + oRS("listimage")

		If oCurrentPage.Item("pageid") <> sPageId Then
			Set oCurrentPage = Server.CreateObject("Scripting.Dictionary")
			oCurrentPage.Add "pageid", sPageId
			oCurrentPage.Add "img", sImg
			oCurrentPage.Add "seq", nSeq
			oCurrentPage.Add "items", Server.CreateObject("System.Collections.ArrayList")
			oPages.Add oCurrentPage
		End If

		If Not IsNull(sItemId) Then
			Dim oItem : Set oItem = Server.CreateObject("Scripting.Dictionary")
			oItem.Add "itemid", sItemId
			oItem.Add "itemname", sItemName
			oItem.Add "itemimage", sItemImage

			Dim oItems : Set oItems = oCurrentPage.Item("items")
			oItems.Add oItem
		End If

		oRS.MoveNext
	Loop

	Set GetPageListObj = oPages
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// RS Functions
'//

Function GetNumBooks(oArgs)
	Dim oRS : Set oRS = GetRsFromArgs(oArgs)

	GetNumBooks = GetTableRows(oRS, "db_contents.dbo.tbl_hhiker_book")
End Function

' 책 목록 얻기
Function GetBookListRS(oTempArgs)
	Dim oArgs : Set oArgs = Factory.Create("Args").SetArgs(Array( _
		"pageSize", 20, _
		"pageNum", 1 _
	))
	oArgs.SetArgs(oTempArgs)

	Dim oRS : Set oRS = GetRsFromArgs(oArgs)

	Dim nPageSize : nPageSize = oArgs.Item("pageSize")
	Dim nPageNum : nPageNum =  IF_(oArgs.Item("pageNum") > 1, oArgs.Item("pageNum"), 1)
	Dim nFirstIndex : nFirstIndex = (nPageNum - 1) * nPageSize
	Dim nLastIndex : nLastIndex = nPageNum * nPageSize + 1

	Dim sSQL : sSQL = "SELECT * FROM (SELECT ROW_NUMBER() OVER (ORDER BY idx DESC) AS ROWINDEX, * FROM db_contents.dbo.tbl_hhiker_book) AS SUB WHERE SUB.ROWINDEX > " & nFirstIndex & " AND SUB.ROWINDEX < " & nLastIndex

	oRS.open sSQL, dbget, 0

	Set GetBookListRS = oRS
End Function

' 책 얻기
Function GetBookRS(oArgs)
	Dim oRS : Set oRS = GetRsFromArgs(oArgs)

	Dim sSQL : sSQL = "SELECT BOOK.* FROM db_contents.dbo.tbl_hhiker_book AS BOOK WHERE BOOK.idx = " & oArgs.Item("bookid")

	oRS.open sSQL, dbget, 0

	Set GetBookRS = oRS
End Function

' 페이지 목록 얻기
Function GetPageListRS(oArgs)
	Dim oRS : Set oRS = GetRsFromArgs(oArgs)

	Dim sSQL : sSQL = "SELECT PAGE.idx AS pageid, PAGE.img, PAGE.seq, ITEM_REF.itemid, ITEM.itemname, ITEM.listimage" & _
						" FROM db_contents.dbo.tbl_hhiker_page AS PAGE" & _
						" LEFT OUTER JOIN db_contents.dbo.tbl_hhiker_item AS ITEM_REF" & _
						" ON ITEM_REF.pageid = PAGE.idx" & _
						" LEFT OUTER JOIN db_item.dbo.tbl_item AS ITEM" & _
						" ON ITEM.itemid = ITEM_REF.itemid" & _
						" WHERE PAGE.bookid = " & oArgs.Item("bookid") & _
						" ORDER BY PAGE.seq ASC"

	oRS.open sSQL, dbget, 0

	Set GetPageListRS = oRS
End Function

' 책 추가
Function AddBook(oArgs)
	Dim oRS : Set oRS = GetRsFromArgs(oArgs)
	Dim sVol : sVol = oArgs.Item("vol")
	Dim sTopic : sTopic = oArgs.Item("topic")
	Dim sThumbnail : sThumbnail = oArgs.Item("thumbnail")
	Dim sPublishedDate : sPublishedDate = oArgs.Item("publisheddate")

	Dim sSQL : sSQL = "INSERT INTO db_contents.dbo.tbl_hhiker_book (vol, topic, thumbnail, publisheddate) VALUES (" & sVol & ", '" & sTopic  &"', '" & sThumbnail & "', '" & sPublishedDate & "')"

	oRS.open sSQL, dbget, 0
End Function

' 책 삭제
Function DelBook(oArgs)
	Dim oRS : Set oRS = GetRsFromArgs(oArgs)
	Dim sBookId : sBookId = oArgs.Item("bookid")

	Dim sSQL : sSQL = "DELETE FROM db_contents.dbo.tbl_hhiker_item WHERE pageid IN ( SELECT idx FROM db_contents.dbo.tbl_hhiker_page WHERE bookid = " & sBookId & ");" &_
						"DELETE FROM db_contents.dbo.tbl_hhiker_page WHERE bookid = " & sBookId & ";" & _
						"DELETE FROM db_contents.dbo.tbl_hhiker_book WHERE idx = " & sBookId & ";"

	oRS.open sSQL, dbget, 0
End Function

' 페이지 추가
Function AddPage(oArgs)
	Dim oRS : Set oRS = GetRsFromArgs(oArgs)
	Dim sBookId : sBookId = oArgs.Item("bookid")
	Dim sImg : sImg = oArgs.Item("img")
	Dim nSeq : nSeq = 999

	Dim sSQL : sSQL = "INSERT INTO db_contents.dbo.tbl_hhiker_page (bookid, img, seq) VALUES (" & sBookId & ", '" & sImg  &"', " & nSeq & ")"

	oRS.open sSQL, dbget, 0
End Function

' 상품 추가
Function AddItem(oArgs)
	Dim oRS : Set oRS = GetRsFromArgs(oArgs)
	Dim sPageId : sPageId = oArgs.Item("pageid")
	Dim sItemIds : sItemIds = oArgs.Item("itemids")
	Dim aItemIds : aItemIds = Split(sItemIds, ",")

	Dim sItemId
	For Each sItemId In aItemIds
		sItemId = Trim(sItemId)
		Dim sSQL : sSQL = "INSERT INTO db_contents.dbo.tbl_hhiker_item (pageid, itemid) VALUES (" & sPageId & ", " & sItemId & ")"
		oRS.open sSQL, dbget, 0
	next
End Function
%>