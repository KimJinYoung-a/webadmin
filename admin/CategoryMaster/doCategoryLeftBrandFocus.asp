<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode , cdl, makerid, sortNo, selIdx, arrIdx, arrSort
dim sqlStr, lp

mode = request("mode")
cdl = request("cdl")
makerid = request("makerid")
sortNo = request("sortNo")
selIdx = Trim(Request("selIdx"))

'// ��庰 �б� //
Select Case mode
	Case "add"
		'�߰�ó��
		sqlStr = "Insert into [db_sitemaster].[dbo].tbl_category_left_brand_rank " &_
				" (cdl, makerid) values  " &_
				" ('" & cdl & "'" &_
				" ,'" & makerid & "')"
	
		dbget.Execute(sqlStr)

	Case "del"
		'����ó��
		sqlStr = "Delete From [db_sitemaster].[dbo].tbl_category_left_brand_rank" &_
				" where idx in (" + selIdx + ")"
	
		dbget.Execute(sqlStr)

	Case "changeSort"
		'ǥ�ü��� �ϰ� ����
		if selIdx<>"" then
			arrIdx = split(selIdx,",")
			arrSort = split(SortNo,",")

			for lp=0 to ubound(arrIdx)
				if Not(arrSort(lp)="" or isNull(arrSort(lp))) then
					sqlStr = sqlStr & "Update [db_sitemaster].[dbo].tbl_category_left_brand_rank " &_
							" Set sortNo=" & arrSort(lp) &_
							" Where idx=" & arrIdx(lp) & ";" & vbCrLf
				end if
			next
			dbget.Execute(sqlStr)
		end if

end Select

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('���� �Ǿ����ϴ�.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->