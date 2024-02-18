<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doCategoryLeftBigchance.asp
' Discription : ī�װ� ������ ó�� ������
' History : 2008.03.31 ������ : ����
'           2008.07.25 ������ ���� : ��ǰ ���ļ��� �߰�
'###############################################

dim mode,cdl,cdm, itemid, selIdx, sortNo
dim i, refer, menupos
dim arrSortNo, arrIdx

menupos = request("menupos")
mode = request("mode")
cdl = request("cdl")
cdm = request("cdm")
itemid = trim(request("itemid"))
selIdx = Request("selIdx")
sortNo = Request("arrSort")

if right(itemid,1)="," then
	itemid = left(itemid,len(itemid)-1)
end if

dim sqlStr

'// ��庰 �б� //
Select Case mode
	Case "del"
		'���û�ǰ ����
		sqlStr = "delete from [db_sitemaster].[dbo].tbl_category_left_bigchance"
		sqlStr = sqlStr + " where idx in (" + selIdx + ")"
	
		rsget.Open sqlStr,dbget,1

		refer = request.ServerVariables("HTTP_REFERER")

	Case "sort"
		'���û�ǰ ���Ĺ�ȣ ����
		arrIdx = split(selIdx,",")
		arrSortNo = split(sortNo,",")

		sqlStr = ""
		for i=0 to ubound(arrIdx)
			sqlStr = sqlStr & "Update [db_sitemaster].[dbo].tbl_category_left_bigchance"
			sqlStr = sqlStr & " Set sortNo=" & arrSortNo(i)
			sqlStr = sqlStr & " where idx=" & arrIdx(i) & "; " & vbCrLf
		next

		'response.Write sqlStr
		rsget.Open sqlStr,dbget,1

		refer = request.ServerVariables("HTTP_REFERER")

	Case "add"
		'�ű� ��ǰ �߰�
		if cdl<>"110" then
			sqlStr = "insert into [db_sitemaster].[dbo].tbl_category_left_bigchance"
			sqlStr = sqlStr + " (cdl, itemid)"
			sqlStr = sqlStr + " select  '" + cdl + "', itemid"
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item"
			sqlStr = sqlStr + " where itemid in (" + itemid + ")"
			sqlStr = sqlStr + " and itemid not in ("
			sqlStr = sqlStr + " select itemid from [db_sitemaster].[dbo].tbl_category_left_bigchance"
			sqlStr = sqlStr + " where cdl='" + cdl + "'"
			sqlStr = sqlStr + " and itemid in (" + itemid + ")"
			sqlStr = sqlStr + ")"
		else
			'����ä���� ��� �ߺз� �߰�
			sqlStr = "insert into [db_sitemaster].[dbo].tbl_category_left_bigchance"
			sqlStr = sqlStr + " (cdl, cdm, itemid)"
			sqlStr = sqlStr + " select  '" + cdl + "', '" + cdm + "', itemid"
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item"
			sqlStr = sqlStr + " where itemid in (" + itemid + ")"
			sqlStr = sqlStr + " and itemid not in ("
			sqlStr = sqlStr + " select itemid from [db_sitemaster].[dbo].tbl_category_left_bigchance"
			sqlStr = sqlStr + " where cdl='" + cdl + "'"
			sqlStr = sqlStr + " and cdm='" + cdm + "'"
			sqlStr = sqlStr + " and itemid in (" + itemid + ")"
			sqlStr = sqlStr + ")"
		end if

		rsget.Open sqlStr,dbget,1

		refer = "category_left_Bigchance.asp?menupos=" & menupos
end Select

%>
<script language="javascript">
<!--
	// ������� ����
	alert("�����߽��ϴ�.");
	self.location = "<%=refer%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
