<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doexhibition.asp
' Discription : exhibition ó�� ������
' History : 2016.04.07 ����ȭ ����
'###############################################

'// ���� ���� �� �Ķ���� ����
dim menupos, mode, selIdx, SortNo, allusing, sqlStr
dim arrIdx, arrSort, lp
dim idx, exhibitiontitle, linkinfo
Dim qdate
Dim startdate , enddate , isusing

Dim listidx , itemid , sortnum , topview , linkurl
Dim subidx

menupos			= Request("menupos")
isusing			= Request("isusing")
mode			= Request("mode")
allusing		= Request("allusing")
selIdx			= Replace(Request("selIdx")," ","")
SortNo			= Replace(Request("arrSort")," ","")
idx				= getNumeric(Request("idx"))
exhibitiontitle	= html2db2017(Request("exhibitiontitle"))
linkinfo		= html2db2017(Request("linkinfo"))

qdate			= Request("prevDate")

startdate		= Request("StartDate")& " " &Request("sTm")&":00:00"
enddate			= Request("EndDate")& " " &Request("eTm")&":59:59"

listidx			= getNumeric(Request("listidx"))
subidx			= getNumeric(Request("subidx"))
itemid			= getNumeric(Request("subItemid"))
sortnum			= getNumeric(Request("sortnum"))

topview			= Request("topview") '//�ٷ� ���⿩�� Y N
linkurl			= Trim(Request("linkurl")) '//��ũ

if SortNo="" Then SortNo = html2db2017(Request("SortNo"))

'// ��忡 ���� �б�
Select Case mode
	 Case "subadd"
		'subitem �ű� ���
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_exhibition_item " &_
					" (listidx, itemid , sortnum) values " &_
					" ('" & listidx  & "'" &_
					" ,'" & itemid &"'" &_
					" ,'" & sortnum &"')"
		dbget.Execute(sqlStr)

	 Case "submodify"
		'subitem ����
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_exhibition_item " &_
				" Set itemid='" & itemid & "'" &_
				" 	,sortnum='" & sortnum & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" Where subidx=" & subidx
		dbget.Execute(sqlStr)

		'// ���������� ���� ������ ������Ʈ
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_exhibition_list " + VbCrlf
		sqlStr = sqlStr + " Set lastadminid='" & session("ssBctId") & "'" + VbCrlf
		sqlStr = sqlStr + " ,lastupdate=getdate() " + VbCrlf
		sqlStr = sqlStr + " where idx=" + cstr(listidx)

'		response.write sqlStr
'		response.end
		dbget.Execute(sqlStr)

	Case "add"
		'�ű� ���
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_exhibition_list " &_
					" (exhibitiontitle, startdate , enddate , linkurl , adminid ) values " &_
					" (N'" & exhibitiontitle &"'" &_
					" ,'" & startdate &"'" &_
					" ,'" & enddate &"'" &_
					" ,N'" & linkurl &"'" &_
					" ,'" & session("ssBctId") &"'" &_
					")"
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "modify"
		'���� ����
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_exhibition_list " &_
				" Set exhibitiontitle=N'" & exhibitiontitle & "'" &_
				" 	,startdate='" & startdate & "'" &_
				" 	,enddate='" & enddate & "'" &_
				" 	,lastadminid='" & session("ssBctId") & "'" &_
				" 	,lastupdate=getdate()" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,topview='" & topview & "'" &_
				" 	,linkurl=N'" & linkurl & "'" &_
				" Where idx=" & idx
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "quickadd"
		'�ű� ���
		Dim extratime1 , extratime2 , titledate

		titledate	  = dateadd("d",0,qdate)
		extratime1 = dateadd("d",0,qdate) &" 00:00:00"
		extratime2 = dateadd("d",0,qdate) &" 23:59:59"


		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_exhibition_list " &_
				" (exhibitiontitle, startdate , enddate , adminid) values " &_
				" ('������� --&gt;" & titledate &" ���� ���� ��ȹ�� �Դϴ�'" &_
				" ,'" & extratime1 &"'" &_
				" ,'" & extratime2 &"'" &_
				" ,'" & session("ssBctId") &"')"
		dbget.Execute(sqlStr)

End Select

%>
<% If mode = "subadd"  Or mode = "submodify" then%>
<script>
<!--
	// ������� ����
	alert("�����߽��ϴ�.");
	window.opener.document.location.href = window.opener.document.URL;    // �θ�â ���ΰ�ħ
	 self.close();        // �˾�â �ݱ�
//-->
</script>
<% Else %>
<script language="javascript">
<!--
	// ������� ����
	alert("�����߽��ϴ�.");
	self.location = "index.asp?menupos=<%=menupos%>&prevDate=<%=qdate%>";
//-->
</script>
<% End If %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
