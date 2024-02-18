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
dim menupos, mode, selIdx, SortNo, allusing, sqlStr
dim arrIdx, arrSort, lp
dim idx, itemid, comment, userid, cate_large, cate_mid

menupos		= Request("menupos")
mode		= Request("mode")
allusing	= Request("allusing")
selIdx		= Replace(Request("selIdx")," ","")
SortNo		= Replace(Request("arrSort")," ","")
idx			= Request("idx")
itemid		= html2db(Request("itemid"))
comment	= html2db(Request("comment"))
userid	= html2db(Request("userid"))
cate_large	= Format00(3,request("cate_large"))
cate_mid	= Format00(3,Request("cate_mid"))


if SortNo="" then	SortNo = html2db(Request("SortNo"))

'// ��忡 ���� �б�
Select Case mode
	Case "changeUsing"
		'��뿩�� �ϰ� ����
		if selIdx<>"" then
			sqlStr = "Update [db_sitemaster].[dbo].tbl_main_review " &_
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
				sqlStr = sqlStr & "Update [db_sitemaster].[dbo].tbl_main_review" &_
						" Set sortNo=" & arrSort(lp) &_
						" Where idx=" & arrIdx(lp) & ";" & vbCrLf
			next
			dbget.Execute(sqlStr)
		end if

	Case "add"
		'�ű� ���
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_main_review " &_
				" (itemid, comment, SortNo, userid, cate_large, cate_mid) values " &_
				" ('" & itemid & "'" &_
				" ,'" & comment & "'" &_
				" ,'" & SortNo & "'" &_
				" ,'" & userid & "'" &_
				" ,'" & cate_large & "'" &_
				" ,'" & cate_mid & "')"
		dbget.Execute(sqlStr)

	Case "modify"
		'���� ����
		sqlStr = "Update [db_sitemaster].[dbo].tbl_main_review " &_
				" Set itemid='" & itemid & "'" &_
				"	,comment='" & comment & "'" &_
				" 	,SortNo=" & SortNo &_
				" 	,userid='" & userid & "'" &_
				" Where idx=" & idx
		dbget.Execute(sqlStr)

End Select

%>
<script language="javascript">
<!--
	// ������� ����
	alert("�����߽��ϴ�.");
	self.location = "main_review.asp?menupos=<%=menupos%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
