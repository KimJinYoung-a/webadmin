<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Class CListItem
	Public FIdx
	Public FTitle
	Public FContents
	Public FRegdate

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class cList
	Public FItemList()
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public FRectIdx

	Public Sub getList()
		Dim sData, rst, i, objJson, iBody

		SET objJson = CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objJson.OPEN "GET", "http://localhost:17847/api/Values", false
			objJson.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objJson.Send()
			If objJson.Status = "200" Then
				iBody = BinaryToText(objJson.ResponseBody, "utf-8")
				Set rst = JSON.parse(iBody)
					FResultCount = rst.length
					Redim preserve FItemList(FResultCount)
					If rst.length > 0 Then
						For i = 0 to FResultCount - 1
							Set FItemList(i) = new CListItem
								FItemList(i).FIdx = rst.get(i).idx
								FItemList(i).FTitle = rst.get(i).title
								FItemList(i).FContents = rst.get(i).contents
								FItemList(i).FRegdate = rst.get(i).regdate
						Next
					End If
				Set rst = nothing
			End If
		SET objJson = nothing
	End Sub

	Private Sub Class_Initialize()
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
End Class
%>