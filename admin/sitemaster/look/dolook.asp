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
Dim orderby, copyimageurl, bgcolor, displaytype, itemimage, displaysale

menupos		= Request("menupos")
isusing		= Request("isusing")
mode		= Request("mode")
allusing	= Request("allusing")
idx			= Request("idx")
qdate		= Request("prevDate")
startdate	= Request("StartDate")& " " &Request("sTm")
enddate		= Request("EndDate")& " " &Request("eTm")
paramisusing= Request("paramisusing")
listidx		= Request("listidx")
subidx		= Request("subidx")
itemid		= Request("Itemid")
orderby		= Request("orderby")
copyimageurl = Request("copyimageurl")
bgcolor		= Request("bgcolor")
displaytype	= Request("displaytype")
itemimage	= Request("itemimage")
displaysale	= Request("displaysale")

bgcolor = Replace(bgcolor, "#", "")


if orderby="" then	orderby = Request("orderby")


If mode = "add" Then '//��¥ üũ
	Dim itemcount
	sqlStr = "select count(*) from [db_sitemaster].[dbo].tbl_pc_main_look_list where startdate = '"& startdate &"' and isusing = 'Y'" 
	rsget.Open SqlStr, dbget, 1
	if Not rsget.Eof Then
		itemcount = rsget(0)
	end If
	rsget.Close

	If itemcount > 0 Then 
	%>
	<script style="css/text">
		alert("���ϳ�¥ Ȥ�� ���Ͻð� �뿡 ���� �ϴ� �������� �ֽ��ϴ�.");
		self.location = "index.asp?menupos=<%=menupos%>&prevDate=<%=qdate%>&isusing=<%=paramisusing%>";
	</script>
	<%
	Response.end
	End If
End If 

'// ��忡 ���� �б�
Select Case mode
	 Case "subadd"
		'subitem �ű� ���
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_pc_main_look_item " &_
					" (listidx, itemid , isusing, itemimage, orderby, displaysale, displaytype) values " &_
					" ('" & idx  & "'" &_
					" ,'" & itemid &"'" &_
					" ,'" & isusing &"'" &_
					" ,'" & itemimage &"'" &_
					" ,'" & orderby &"'" &_
					" ,'" & displaysale &"'" &_
					" ,'" & displaytype &"')"
		dbget.Execute(sqlStr)

	 Case "submodify"
		'subitem ����
		sqlStr = "Update [db_sitemaster].[dbo].tbl_pc_main_look_item " &_
				" Set itemid='" & itemid & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,itemimage='" & itemimage & "'" &_
				" 	,orderby='" & orderby & "'" &_
				" 	,displaytype='" & displaytype & "'" &_
				" 	,displaysale='" & displaysale & "'" &_
				" Where idx=" & subidx
		dbget.Execute(sqlStr)

		'// ���������� ���� ������ ������Ʈ
		sqlStr = "Update [db_sitemaster].[dbo].tbl_pc_main_look_list " + VbCrlf
		sqlStr = sqlStr + " Set lastadminid='" & session("ssBctId") & "'" + VbCrlf
		sqlStr = sqlStr + " ,lastupdate=getdate() " + VbCrlf
		sqlStr = sqlStr + " where idx=" + cstr(idx)

'		response.write sqlStr
'		response.end
		dbget.Execute(sqlStr)

	Case "add"
		'�ű� ���
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_pc_main_look_list " &_
					" (copyimageurl, bgcolor, startdate, enddate, adminid, isusing, orderby) values " &_
					" ('" & copyimageurl &"'" &_
					" ,'" & bgcolor &"'" &_
					" ,'" & startdate &"'" &_
					" ,'" & enddate &"'" &_
					" ,'" & session("ssBctId") &"'" &_
					" ,'" & isusing & "'" &_
					" ,'" & orderby & "'" &_
					")"
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "modify"
		'���� ����
		sqlStr = "Update [db_sitemaster].[dbo].tbl_pc_main_look_list " &_
				" Set copyimageurl='" & copyimageurl & "'" &_
				" 	,bgcolor='" & bgcolor & "'" &_
				" 	,startdate='" & startdate & "'" &_
				" 	,enddate='" & enddate & "'" &_
				" 	,lastadminid='" & session("ssBctId") & "'" &_
				" 	,lastupdate=getdate()" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,orderby='" & orderby & "'" &_
				" Where idx=" & idx
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "quickadd"
		'�ű� ���
		Dim extratime1 , extratime2 , titledate
			
		titledate	  = dateadd("d",0,qdate)
		extratime1 = dateadd("d",0,qdate) &" 00:00:00"
		extratime2 = dateadd("d",0,qdate) &" 23:59:59"


		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_pc_main_just1day_list " &_
				" (title, startdate , enddate , adminid) values " &_
				" ('������� --&gt;" & titledate &" ���� mdpick �Դϴ�'" &_
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
	<% if mode = "subadd" then %>
		document.location.href = "/admin/sitemaster/look/pop_LookItemAddInfo.asp?menupos=<%=menupos%>&prevDate=<%=qdate%>&paramisusing=<%=paramisusing%>&usingyn=<%=isusing%>&idx=<%=idx%>";
		opener.document.location.href = "/admin/sitemaster/look/look_insert.asp?menupos=<%=menupos%>&idx=<%=idx%>&prevDate=<%=qdate%>&paramisusing=<%=paramisusing%>";    // �θ�â ���ΰ�ħ
	<% else %>
		opener.document.location.href = "/admin/sitemaster/look/look_insert.asp?menupos=<%=menupos%>&idx=<%=idx%>&prevDate=<%=qdate%>&paramisusing=<%=paramisusing%>";    // �θ�â ���ΰ�ħ
		self.close();
	<% end if %>
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
