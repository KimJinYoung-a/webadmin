<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
Response.CharSet = "euc-kr"

Dim i, vQuery, vBody, vAction, vIdx, vContents, vGubun, vContentsIdx, vSortNo, vUnitCount, vUnit, vArr
vIdx = requestCheckvar(request("idx"),15)
vAction = requestCheckvar(request("action"),10)
vContents = requestCheckvar(request("contents"),300)
vContents = Left(vContents, Len(vContents)-1)


'#######	[1] contents �Ѱ� �޾� �����۾��̵�, �̺�Ʈ���̵� �����ؼ� ���� ���� ����(�̸�����). ���� �ۼ�.	#######
For i = LBound(Split(vContents,",")) To UBound(Split(vContents,","))
	vGubun = Split(Split(vContents,",")(i),"$")(0)
	vContentsIdx = Split(Split(vContents,",")(i),"$")(1)

	vQuery = vQuery & "INSERT INTO [db_sitemaster].[dbo].[tbl_search_curator_unit](topidx, gubun, contentsidx, sortno) "
	vQuery = vQuery & "VALUES('" & vIdx & "', '" & vGubun & "', '" & vContentsIdx & "', '" & (i+1) & "'); "
Next


If i < 11 Then

	'#######	[2] unit DB�� �ٷ� ����.	#######
	If vQuery <> "" Then
		vQuery = "DELETE [db_sitemaster].[dbo].[tbl_search_curator_unit] WHERE topidx = '" & vIdx & "'; " & vQuery
		dbget.Execute vQuery
	End IF


	'#######	[3] Unit(��ǰ��, �̺�Ʈ�� ����) ��������	#######
	vQuery = ""
	vQuery = vQuery & "select e.evt_name, cu.gubun, cu.contentsidx, cu.sortno, e.evt_enddate from [db_sitemaster].[dbo].[tbl_search_curator_unit] as cu " & vbCrLf
	vQuery = vQuery & "inner join [db_event].[dbo].[tbl_event] as e on cu.contentsidx = e.evt_code " & vbCrLf
	vQuery = vQuery & "inner join [db_event].[dbo].[tbl_event_display] as ed on e.evt_code = ed.evt_code " & vbCrLf
	vQuery = vQuery & "where cu.topidx = '" & vIdx & "' and cu.gubun = 'event' " & vbCrLf
	vQuery = vQuery & "union all " & vbCrLf
	vQuery = vQuery & "select i.itemname, cu.gubun, cu.contentsidx, cu.sortno, getdate() from [db_sitemaster].[dbo].[tbl_search_curator_unit] as cu " & vbCrLf
	vQuery = vQuery & "inner join [db_item].[dbo].[tbl_item] as i on cu.contentsidx = i.itemid " & vbCrLf
	vQuery = vQuery & "inner join [db_item].[dbo].[tbl_item_contents] as ic on i.itemid = ic.itemid " & vbCrLf
	vQuery = vQuery & "where cu.topidx = '" & vIdx & "' and cu.gubun = 'item'"
	vQuery = vQuery & "order by sortno asc " & vbCrLf
	rsget.CursorLocation = adUseClient
	rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly
	If not rsget.Eof Then
		vArr = rsget.getRows()
	End If
	rsget.close


	If isArray(vArr) Then

		For i = 0 To UBound(vArr,2)
			
			'### �̸�(0), ����(1), ������idx(2), ���Ĺ�ȣ(3)
			vBody = vBody & "<li>" & vbCrLf
			vBody = vBody & "	<p class=""cell15 lt"">" & vArr(1,i) & "</p>" & vbCrLf
			vBody = vBody & "	<p class=""lt""><span class=""textOverflow"">"
			
			If vArr(1,i) = "event" AND date() > vArr(4,i) Then
				vBody = vBody & "<font color=red>[����]</font> "
			End If

			vBody = vBody & db2html(vArr(0,i)) & "</span></p>" & vbCrLf
			vBody = vBody & "	<p class=""cell05""><input type=""button"" class=""btn"" value=""����"" onClick=""jsUnitDelete('"&vArr(1,i)&"','"&vArr(2,i)&"');"" /></p>" & vbCrLf
			vBody = vBody & "	<input type=""hidden"" id=""sort"" name=""sort"" value="""&vSortNo&""">" & vbCrLf
			vBody = vBody & "	<input type=""hidden"" id=""svalue"" name=""svalue"" value="""&vArr(1,i)&"$"&vArr(2,i)&""">" & vbCrLf
			vBody = vBody & "</li>" & vbCrLf

		Next
	End If

	Response.Write vBody
Else

	Response.Write "10"

End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->