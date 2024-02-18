<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<%
'###############################################
' PageName : doMainTopKeyword.asp
' Discription : ���� žŰ���� ó�� ������
' History : 2009.09.16 �ѿ�� 10x10���� ������ ����
'###############################################

'// ���� ���� �� �Ķ���� ����
dim menupos, mode, selIdx, SortNo, allusing, sqlStr ,arrIdx, arrSort, lp , keyword_gubun
dim idx, keyword, linkinfo
	menupos		= RequestCheckvar(Request("menupos"),10)
	mode		= RequestCheckvar(Request("mode"),16)
	allusing	= RequestCheckvar(Request("allusing"),1)
	keyword_gubun	= RequestCheckvar(Request("keyword_gubun"),10)
	selIdx		= Replace(Request("selIdx")," ","")
	SortNo		= Replace(Request("arrSort")," ","")
	idx			= RequestCheckvar(Request("idx"),10)
	keyword		= html2db(RequestCheckvar(Request("keyword"),32))
	linkinfo	= html2db(Request("linkinfo"))
  	if selIdx <> "" then
		if checkNotValidHTML(selIdx) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
  	if SortNo <> "" then
		if checkNotValidHTML(SortNo) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	if SortNo="" then	SortNo = html2db(Request("SortNo"))

'// ��忡 ���� �б�
Select Case mode
	Case "changeUsing"
		'��뿩�� �ϰ� ����
		if selIdx<>"" then
			sqlStr = "Update [db_academy].[dbo].tbl_maintopKeyword " &_
					" Set isusing='" & allusing & "'" &_
					" Where idx in (" & selIdx & ")"
			dbacademyget.Execute(sqlStr)
		end if

	Case "changeSort"
		'ǥ�ü��� �ϰ� ����
		if selIdx<>"" then
			arrIdx = split(selIdx,",")
			arrSort = split(SortNo,",")

			for lp=0 to ubound(arrIdx)
				sqlStr = sqlStr & "Update [db_academy].[dbo].tbl_maintopKeyword " &_
						" Set sortNo=" & arrSort(lp) &_
						" Where idx=" & arrIdx(lp) & ";" & vbCrLf
			next
			dbacademyget.Execute(sqlStr)
		end if

	Case "add"
		'�ű� ���
		sqlStr = "Insert Into [db_academy].[dbo].tbl_maintopKeyword " &_
				" (keyword, keyword_gubun,linkinfo, SortNo) values " &_
				" ('" & keyword & "'" &_
				" ," & keyword_gubun & "" &_
				" ,'" & linkinfo & "'" &_
				" ," & SortNo & ")"
		
		'response.write sqlStr &"<Br>"
		dbacademyget.Execute(sqlStr)

	Case "modify"
		'���� ����
		sqlStr = "Update [db_academy].[dbo].tbl_maintopKeyword " &_
				" Set keyword='" & keyword & "'" &_
				" 	,keyword_gubun='" & keyword_gubun & "'" &_
				" 	,linkinfo='" & linkinfo & "'" &_				
				" 	,SortNo=" & SortNo &_
				" Where idx=" & idx

		'response.write sqlStr &"<Br>"				
		dbacademyget.Execute(sqlStr)

End Select

%>

<script language="javascript">

	// ������� ����
	alert("�����߽��ϴ�.");
	self.location = "main_TopKeyword.asp?menupos=<%=menupos%>";

</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->