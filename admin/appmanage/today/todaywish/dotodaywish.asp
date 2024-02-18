<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : dotodaywish.asp
' Discription : todaywish ó�� ������
' History : 2013.12.16 ����ȭ ����
'###############################################

'// ���� ���� �� �Ķ���� ����
dim menupos, mode, selIdx, SortNo, allusing, sqlStr
dim arrIdx, arrSort, lp
dim idx, wishtitle, linkinfo
Dim qdate
Dim startdate , enddate , gubun , isusing

Dim listidx , itemid , sortnum 
Dim subidx

menupos	= Request("menupos")
isusing		= Request("isusing")
mode		= Request("mode")
gubun		= Request("gubun")
allusing		= Request("allusing")
selIdx		= Replace(Request("selIdx")," ","")
SortNo		= Replace(Request("arrSort")," ","")
idx			= Request("idx")
wishtitle	= html2db(Request("wishtitle"))
linkinfo		= html2db(Request("linkinfo"))

qdate		=	Request("prevDate")

startdate	= Request("StartDate")& " " &Request("sTm")
enddate	= Request("EndDate")& " " &Request("eTm")

listidx		= Request("listidx")
subidx		= Request("subidx")
itemid		= Request("subItemid")
sortnum	= Request("sortnum")

if SortNo="" then	SortNo = html2db(Request("SortNo"))

'// ��忡 ���� �б�
Select Case mode
	 Case "subadd"
		'subitem �ű� ���
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_todaywish_item " &_
					" (listidx, itemid , sortnum) values " &_
					" ('" & listidx  & "'" &_
					" ,'" & itemid &"'" &_
					" ,'" & sortnum &"')"
		dbget.Execute(sqlStr)

	 Case "submodify"
		'subitem ����
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_todaywish_item " &_
				" Set itemid='" & itemid & "'" &_
				" 	,sortnum='" & sortnum & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" Where subidx=" & subidx
		dbget.Execute(sqlStr)

		'// ���������� ���� ������ ������Ʈ
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_todaywish_list " + VbCrlf
		sqlStr = sqlStr + " Set lastadminid='" & session("ssBctId") & "'" + VbCrlf
		sqlStr = sqlStr + " ,lastupdate=getdate() " + VbCrlf
		sqlStr = sqlStr + " where idx=" + cstr(listidx)

'		response.write sqlStr
'		response.end
		dbget.Execute(sqlStr)

	Case "add"
		'�ű� ���
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_todaywish_list " &_
					" (wishtitle, startdate , enddate , adminid ) values " &_
					" ('" & wishtitle &"'" &_
					" ,'" & startdate &"'" &_
					" ,'" & enddate &"'" &_
					" ,'" & session("ssBctId") &"'" &_
					")"
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "modify"
		'���� ����
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_todaywish_list " &_
				" Set wishtitle='" & wishtitle & "'" &_
				" 	,startdate='" & startdate & "'" &_
				" 	,enddate='" & enddate & "'" &_
				" 	,lastadminid='" & session("ssBctId") & "'" &_
				" 	,lastupdate=getdate()" &_
				" 	,isusing='" & isusing & "'" &_
				" Where idx=" & idx
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "quickadd"
		'�ű� ���
		Dim extratime1 , extratime2 , titledate
			
		titledate	  = dateadd("d",0,qdate)
		extratime1 = dateadd("d",0,qdate) &" 00:00:00"
		extratime2 = dateadd("d",0,qdate) &" 23:59:59"


		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_todaywish_list " &_
				" (wishtitle, startdate , enddate , adminid) values " &_
				" ('������� --&gt;" & titledate &" ���� todaywish �Դϴ�'" &_
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
