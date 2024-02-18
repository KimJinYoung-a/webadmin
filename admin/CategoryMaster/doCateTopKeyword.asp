<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doCateTopKeyword.asp
' Discription : ī�װ� žŰ���� ó�� ������
' History : 2008.03.31 ������ ����
'         : 2008.10.27 ��ī�װ� ó�� �߰�(������)
'         : 2009.04.16 ���û�ǰ �߰�(������)
'###############################################

'// ���� ���� �� �Ķ���� ����
dim menupos, mode, selIdx, SortNo, allusing, sqlStr
dim arrIdx, arrSort, lp
dim idx, cdl, cdm, keyword, linkinfo, itemid

menupos		= Request("menupos")
mode		= Request("mode")
allusing	= Request("allusing")
selIdx		= Replace(Request("selIdx")," ","")
SortNo		= Replace(Request("arrSort")," ","")
itemid		= Request("itemid")
idx			= Request("idx")
cdl			= Request("cdl")
if cdl="110" then
	cdm			= Request("cdm")
end if
keyword		= html2db(Request("keyword"))
linkinfo	= html2db(Request("linkinfo"))

if SortNo="" then	SortNo = html2db(Request("SortNo"))

'// ��忡 ���� �б�
Select Case mode
	Case "changeUsing"
		'��뿩�� �ϰ� ����
		if selIdx<>"" then
			sqlStr = "Update [db_sitemaster].[dbo].tbl_category_keyword " &_
					" Set isusing='" & allusing & "'" &_
					" Where idx in (" & selIdx & ")"
			dbget.Execute(sqlStr)
		end if

	Case "changeSort"
		'ǥ�ü��� �ϰ� ����
		if selIdx<>"" then
			arrIdx = split(selIdx,",")
			arrSort = split(SortNo,",")

			for lp=0 to ubound(arrIdx)
				sqlStr = sqlStr & "Update [db_sitemaster].[dbo].tbl_category_keyword " &_
						" Set sortNo=" & arrSort(lp) &_
						" Where idx=" & arrIdx(lp) & ";" & vbCrLf
			next
			dbget.Execute(sqlStr)
		end if

	Case "add"
		'�ű� ���
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_category_keyword " &_
				" (cdl, cdm, keyword, linkinfo, itemid, SortNo) values " &_
				" ('" & cdl & "'" &_
				" ,'" & cdm & "'" &_
				" ,'" & keyword & "'" &_
				" ,'" & linkinfo & "'" &_
				" ,'" & itemid & "'" &_
				" ," & SortNo & ")"
		dbget.Execute(sqlStr)

	Case "modify"
		'���� ����
		sqlStr = "Update [db_sitemaster].[dbo].tbl_category_keyword " &_
				" Set cdl='" & cdl & "'" &_
				" 	,cdm='" & cdm & "'" &_
				" 	,keyword='" & keyword & "'" &_
				" 	,linkinfo='" & linkinfo & "'" &_
				" 	,itemid='" & itemid & "'" &_
				" 	,SortNo=" & SortNo &_
				" Where idx=" & idx
		dbget.Execute(sqlStr)

End Select

%>
<script language="javascript">
<!--
	// ������� ����
	alert("�����߽��ϴ�.");
	self.location = "category_left_topKeyword.asp?menupos=<%=menupos%>&cdl=<%=cdl%>&cdm=<%=cdm%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
