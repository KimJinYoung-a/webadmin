<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'// ���� ���� �� �Ķ���� ����
dim menupos, mode, selIdx, SortNo, allusing, sqlStr
dim arrIdx, arrSort, lp
dim idx, title, linkinfo
Dim qdate
Dim startdate , enddate , gubun , isusing, paramisusing

Dim listidx , itemid , sortnum , is1day
Dim subidx
Dim todayban , extraurl
Dim subtitle , saleper '//�ָ�or����Ư����
Dim maxsaleper
Dim pcimage, mobileimage, pclinkurl, mobilelinkurl, price
Dim platform, vType, workertext
Dim frontimage, bannerimage, linkurl



menupos		= Request("menupos")
isusing		= Request("isusing")
mode		= Request("mode")
gubun		= Request("gubun")
allusing	= Request("allusing")
selIdx		= Replace(Request("selIdx")," ","")
SortNo		= Replace(Request("arrSort")," ","")
idx			= Request("idx")
title		= Newhtml2db(Request("title"))
linkinfo	= Request("linkinfo")

qdate		= Request("prevDate")

startdate	= Request("StartDate")& " " &Request("sTm")
enddate		= Request("EndDate")& " " &Request("eTm")

paramisusing= Request("paramisusing")

listidx		= Request("listidx")
subidx		= Request("subidx")
itemid		= Request("Itemid")
sortnum		= Request("sortnum")
saleper		= Request("saleper")

frontimage		= Request("frontimage")
bannerimage		= Request("bannerimage")
linkurl		= Request("linkurl")
price		= Request("price")

platform		= Request("platform")
vType			= Request("type")
workertext		= Request("workertext")

if SortNo="" then	SortNo = Request("SortNo")


If mode = "add" Then '//��¥ üũ
	Dim itemcount
	sqlStr = "select count(*) from [db_sitemaster].[dbo].[tbl_just1day2018_list] where startdate = '"& startdate &"' and isusing = 'Y' AND platform='"&platform&"' " 
	rsget.Open SqlStr, dbget, 1
	if Not rsget.Eof Then
		itemcount = rsget(0)
	end If
	rsget.Close

	If itemcount > 0 Then 
	%>
	<script style="css/text">
		alert("���ϳ�¥ Ȥ�� ���Ͻð� �뿡 ���� �ϴ� �������� �ֽ��ϴ�.");
		history.back();
		//self.location = "index.asp?menupos=<%=menupos%>&prevDate=<%=qdate%>&isusing=<%=paramisusing%>";
	</script>
	<%
	Response.end
	End If
End If 

'// ��忡 ���� �б�
Select Case mode
	 Case "subadd"
		'subitem �ű� ���
		sqlStr = "Insert Into [db_sitemaster].[dbo].[tbl_just1day2018_item] " &_
					" (listidx, title, itemid, frontimage, price, saleper, adminid, isusing, sortnum) values " &_
					" ('" & listidx  & "'" &_
					" ,'" & title &"'" &_
					" ,'" & itemid &"'" &_
					" ,'" & frontimage &"'" &_
					" ,'" & price &"'" &_
					" ,'" & saleper &"'" &_
					" ,'" & session("ssBctId") &"'" &_
					" ,'" & isusing &"'" &_
					" ,'" & sortnum &"')"
		dbget.Execute(sqlStr)

	 Case "submodify"
		'subitem ����
		sqlStr = "Update [db_sitemaster].[dbo].tbl_just1day2018_item " &_
				" Set itemid='" & itemid & "'" &_
				" 	,title='" & title & "'" &_
				" 	,frontimage='" & frontimage & "'" &_
				" 	,price='" & price & "'" &_
				" 	,saleper='" & saleper & "'" &_
				" 	,sortnum='" & sortnum & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,lastadminid='" & session("ssBctId") & "'" &_
				" 	,lastupdate=getdate()" &_
				" Where subidx=" & subidx
		dbget.Execute(sqlStr)

		'// ���������� ���� ������ ������Ʈ
		sqlStr = "Update [db_sitemaster].[dbo].[tbl_just1day2018_list] " + VbCrlf
		sqlStr = sqlStr + " Set lastadminid='" & session("ssBctId") & "'" + VbCrlf
		sqlStr = sqlStr + " ,lastupdate=getdate() " + VbCrlf
		sqlStr = sqlStr + " where idx=" + cstr(listidx)

'		response.write sqlStr
'		response.end
		dbget.Execute(sqlStr)

	Case "add"
		'�ű� ���
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_just1day2018_list " &_
					" (title, startdate , enddate , adminid, isusing , maxsaleper, type, bannerimage, linkurl, workertext, platform) values " &_
					" ('" & title &"'" &_
					" ,'" & startdate &"'" &_
					" ,'" & enddate &"'" &_
					" ,'" & session("ssBctId") &"'" &_
					" ,'" & isusing & "'" &_
					" ,'" & saleper & "'" &_
					" ,'" & vType & "'" &_
					" ,'" & bannerimage & "'" &_
					" ,'" & linkurl & "'" &_
					" ,'" & workertext & "'" &_
					" ,'" & platform & "'" &_
					")"
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "modify"
		'���� ����
		sqlStr = "Update [db_sitemaster].[dbo].tbl_just1day2018_list " &_
				" Set title='" & title & "'" &_
				" 	,startdate='" & startdate & "'" &_
				" 	,enddate='" & enddate & "'" &_
				" 	,lastadminid='" & session("ssBctId") & "'" &_
				" 	,lastupdate=getdate()" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,maxsaleper='" & saleper & "'" &_
				" 	,type='" & vType & "'" &_
				" 	,bannerimage='" & bannerimage & "'" &_
				" 	,linkurl='" & linkurl & "'" &_
				" 	,workertext='" & workertext & "'" &_
				" 	,platform='" & platform & "'" &_
				" Where idx=" & idx
		'response.write sqlStr
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
	self.location = "index.asp?menupos=<%=menupos%>&prevDate=<%=qdate%>&isusing=<%=paramisusing%>";
//-->
</script>
<% End If %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
