<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : dotodaydeal.asp
' Discription : todaydeal ó�� ������
' History : 2014.06.30 ����ȭ ����
'###############################################

'// ���� ���� �� �Ķ���� ����
dim menupos, mode, selIdx, SortNo, allusing, sqlStr
dim arrIdx, arrSort, lp
dim idx, dealtitle, itemurl , itemurlmo
Dim qdate
Dim startdate , enddate , gubun1 , gubun2  , isusing

Dim itemid , sortnum 
Dim itemname

menupos		= Request("menupos")
isusing		= Request("isusing")
mode		= Request("mode")
gubun1		= Request("gubun1")
gubun2		= Request("gubun2")
idx			= Request("idx")

dealtitle	= html2db(Request("dealtitle"))
itemurl		= html2db(Request("itemurl"))
itemurlmo	= html2db(Request("itemurlmo"))

qdate		= Request("prevDate")

startdate	= Request("StartDate")& " " &Request("sTm")
enddate		= Request("EndDate")& " " &Request("eTm")

itemid		= Request("itemid")
itemname	= Request("itemname")
sortnum		= Request("sortnum")

'// ��忡 ���� �б�
Select Case mode
	Case "add"
		'�ű� ���
		sqlStr = "Insert Into db_sitemaster.dbo.tbl_mobile_main_todaydeal " &_
					" (startdate , enddate , adminid , isusing , sortnum , itemid , itemname , dealtitle , gubun1 , gubun2 , itemurl , itemurlmo ) values " &_
					" ('" & startdate &"'" &_
					" ,'" & enddate &"'" &_
					" ,'" & session("ssBctId") &"'" &_
					" ,'" & isusing &"'" &_
					" ,'" & sortnum &"'" &_
					" ,'" & itemid &"'" &_
					" ,'" & itemname &"'" &_
					" ,'" & dealtitle &"'" &_
					" ,'" & gubun1 &"'" &_
					" ,'" & gubun2 &"'" &_
					" ,'" & itemurl &"'" &_
					" ,'" & itemurlmo &"'" &_
					")"
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "modify"
		'���� ����
		sqlStr = "Update db_sitemaster.dbo.tbl_mobile_main_todaydeal " &_
				" Set startdate='" & startdate & "'" &_
				" 	,enddate='" & enddate & "'" &_
				" 	,lastadminid='" & session("ssBctId") & "'" &_
				" 	,lastupdate=getdate()" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,sortnum='" & sortnum & "'" &_
				" 	,itemid='" & itemid & "'" &_
				" 	,itemname='" & itemname & "'" &_
				" 	,dealtitle='" & dealtitle & "'" &_
				" 	,gubun1='" & gubun1 & "'" &_
				" 	,gubun2='" & gubun2 & "'" &_
				" 	,itemurl='" & itemurl & "'" &_
				" 	,itemurlmo='" & itemurlmo & "'" &_
				" Where idx=" & idx
		'response.write sqlStr
		dbget.Execute(sqlStr)

End Select
%>
<script>
<!--
	// ������� ����
	alert("�����߽��ϴ�.");
	self.location = "index.asp?menupos=<%=menupos%>&prevDate=<%=qdate%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
