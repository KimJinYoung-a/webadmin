<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doMainTopKeyword.asp
' Discription : ���� žŰ���� ó�� ������
' History : 2008.04.18 ������ ����
'           2012.01.09 ������ : ����Ʈ���� �߰�
'###############################################

'// ���� ���� �� �Ķ���� ����
dim menupos, mode, selIdx, SortNo, allusing, sqlStr, siteDiv
dim arrIdx, arrSort, lp
dim idx, textname, linkinfo
Dim enddate

menupos	= Request("menupos")
mode		= Request("mode")
allusing		= Request("allusing")
selIdx		= Replace(Request("selIdx")," ","")
SortNo		= Replace(Request("arrSort")," ","")
idx			= Request("idx")
textname	= html2db(Request("keyword"))
linkinfo		= html2db(Request("linkinfo"))
siteDiv		= Request("siteDiv")
enddate	= Request("prevDate")


if SortNo="" then	SortNo = html2db(Request("SortNo"))

'// ��忡 ���� �б�
Select Case mode
	Case "changeUsing"
		'��뿩�� �ϰ� ����
		if selIdx<>"" then
			sqlStr = "Update [db_sitemaster].[dbo].tbl_mainTextissue " &_
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
				sqlStr = sqlStr & "Update [db_sitemaster].[dbo].tbl_mainTextissue " &_
						" Set sortNo=" & arrSort(lp) &_
						" Where idx=" & arrIdx(lp) & ";" & vbCrLf
			next
			dbget.Execute(sqlStr)
		end if

	Case "add"
		'�ű� ���
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mainTextissue " &_
				" (textname, linkinfo, enddate,  SortNo ) values " &_
				" ('" & textname & "'" &_
				" ,'" & linkinfo & "'" &_
				" ,'" & enddate & "'" &_
				" ," & SortNo & ")"
		dbget.Execute(sqlStr)

	Case "modify"
		'���� ����
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mainTextissue " &_
				" Set textname='" & textname & "'" &_
				" 	,linkinfo='" & linkinfo & "'" &_
				" 	,SortNo=" & SortNo &_
				" 	,enddate='" & enddate & "'" &_
				" Where idx=" & idx
		dbget.Execute(sqlStr)

End Select

%>
<script language="javascript">
<!--
	// ������� ����
	alert("�����߽��ϴ�.");
	self.location = "index.asp?menupos=<%=menupos%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
