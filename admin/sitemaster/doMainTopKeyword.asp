<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : doMainTopKeyword.asp
' Discription : ���� žŰ���� ó�� ������
' History : 2008.04.18 ������ ����
'           2022.07.01 �ѿ�� ����(isms���������, �ҽ�ǥ��ȭ)
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'// ���� ���� �� �Ķ���� ����
dim menupos, mode, selIdx, SortNo, allusing, sqlStr, siteDiv
dim arrIdx, arrSort, lp
dim idx, keyword, linkinfo

menupos		= Request("menupos")
mode		= Request("mode")
allusing	= Request("allusing")
selIdx		= Replace(Request("selIdx")," ","")
SortNo		= Replace(Request("arrSort")," ","")
idx			= Request("idx")
keyword		= html2db(Request("keyword"))
linkinfo	= html2db(Request("linkinfo"))
siteDiv		= Request("siteDiv")

if SortNo="" then	SortNo = html2db(Request("SortNo"))

'// ��忡 ���� �б�
Select Case mode
	Case "changeUsing"
		'��뿩�� �ϰ� ����
		if selIdx<>"" then
			sqlStr = "Update [db_sitemaster].[dbo].tbl_maintopKeyword " &_
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
				sqlStr = sqlStr & "Update [db_sitemaster].[dbo].tbl_maintopKeyword " &_
						" Set sortNo=" & arrSort(lp) &_
						" Where idx=" & arrIdx(lp) & ";" & vbCrLf
			next
			dbget.Execute(sqlStr)
		end if

	Case "add"
		if keyword <> "" and not(isnull(keyword)) then
			keyword = ReplaceBracket(keyword)
		end If
		if linkinfo <> "" and not(isnull(linkinfo)) then
			linkinfo = ReplaceBracket(linkinfo)
		end If

		'�ű� ���
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_maintopKeyword " &_
				" (siteDiv, keyword, linkinfo, SortNo) values " &_
				" ('" & siteDiv & "'" &_
				" ,'" & keyword & "'" &_
				" ,'" & linkinfo & "'" &_
				" ," & SortNo & ")"
		dbget.Execute(sqlStr)

	Case "modify"
		if keyword <> "" and not(isnull(keyword)) then
			keyword = ReplaceBracket(keyword)
		end If
		if linkinfo <> "" and not(isnull(linkinfo)) then
			linkinfo = ReplaceBracket(linkinfo)
		end If

		'���� ����
		sqlStr = "Update [db_sitemaster].[dbo].tbl_maintopKeyword " &_
				" Set siteDiv='" & siteDiv & "'" &_
				"	,keyword='" & keyword & "'" &_
				" 	,linkinfo='" & linkinfo & "'" &_
				" 	,SortNo=" & SortNo &_
				" Where idx=" & idx
		dbget.Execute(sqlStr)

End Select

%>
<script language="javascript">
<!--
	// ������� ����
	alert("�����߽��ϴ�.");
	self.location = "main_TopKeyword.asp?menupos=<%=menupos%>&siteDiv=<%=siteDiv%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
