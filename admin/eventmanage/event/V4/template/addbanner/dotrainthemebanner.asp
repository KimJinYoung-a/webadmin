<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###############################################
' PageName : dotrainthemebanner.asp
' Discription : H�� ������ ��ũ �̹��� ���
' History : 2018.08.14 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	Dim eventid , mode , device , idx, saveafter
	Dim slideimg , linkurl , sorting , isusing , sbgimg '�����̵� �̹���
	Dim topimg , btmimg , topaddimg '���ø� �̹�����
	Dim btmYN , btmcode , btmaddimg , pcadd1 , moadd1 , gubun
	Dim sqlStr, GroupItemCheck

	Dim sIdx, sSortNo, sIsUsing, i , slinkurl '//�����̵�

	idx 		= requestCheckVar(Request.form("idx"),10)
	eventid 	= requestCheckVar(Request.form("eventid"),10)
	mode 		= requestCheckVar(Request.form("mode"),6)
	device 		= requestCheckVar(Request.form("device"),1)
	slideimg 	= requestCheckVar(Request.form("slideimg"),200)
	sbgimg	 	= requestCheckVar(Request.form("bgslideimg"),200)
	linkurl 	= requestCheckVar(Request.form("linkurl"),200)

	topimg 		= requestCheckVar(Request.form("topimg"),200)
	btmimg 		= requestCheckVar(Request.form("btmimg"),200)
	topaddimg 	= requestCheckVar(Request.form("topaddimg"),200)

	btmYN		= requestCheckVar(Request.form("btmYN"),1)
	btmcode		= html2db(Request.form("btmcode"))
	btmaddimg	= requestCheckVar(Request.form("btmaddimg"),200)
	pcadd1		= requestCheckVar(Request.form("pcadd1"),200)
	moadd1		= requestCheckVar(Request.form("moadd1"),200)
	gubun		= requestCheckVar(Request.form("gubun"),200)
	GroupItemCheck 		= requestCheckVar(Request.form("GroupItemCheck"),1)
'Response.write mode & "<br>"
Select Case mode
	 Case "SI"
		'slide�̹��� �ű� ���
		sqlStr = "Insert Into [db_event].[dbo].[tbl_event_manual_group] " &_
					" (evt_code, imgurl,grouptype) values " &_
					" ('" & eventid  & "'" &_
					" ,'" & slideimg &"','B')"
		dbget.Execute(sqlStr)
		saveafter="SI"
	Case "SU"
		'//����Ʈ��������
		for i=1 to request.form("chkIdx").count
			sIdx = request.form("chkIdx")(i)
			sSortNo = request.form("sort"&sIdx)
			sIsUsing = request.form("use"&sIdx)
			slinkurl = request.form("linkurl"&sIdx)
			if sSortNo="" then sSortNo="0"
			if sIsUsing="" then sIsUsing="N"

			sqlStr = sqlStr & " Update [db_event].[dbo].[tbl_event_manual_group] Set "
			sqlStr = sqlStr & " viewidx=" & sSortNo & ""
			sqlStr = sqlStr & " Where idx='" & sIdx & "';" & vbCrLf
		Next

		If sqlStr <> "" then
			dbget.Execute sqlStr
		Else
			Call Alert_return("������ ������ �����ϴ�.")
			dbget.Close: Response.End
		End If 
		saveafter="SU"
	Case "SD" '����
		sIdx = request.form("chkIdx")

		sqlStr = "delete from db_event.dbo.tbl_event_manual_group Where idx='"& sIdx &"'"
		dbget.Execute sqlStr
End Select

	sqlStr = " Update [db_event].[dbo].[tbl_event_md_theme]"
	sqlStr = sqlStr & " Set GroupItemType='B'"
	sqlStr = sqlStr & " ,GroupItemCheck='" + Cstr(GroupItemCheck) + "'"
	sqlStr = sqlStr & " Where evt_code=" & eventid
	dbget.Execute sqlStr
%>
<script language="javascript">
<!--
	// ������� ����
	alert("<%=chkiif(mode="SD","���� �Ϸ�.","����/���� �Ϸ�.")%>");
	self.location = "pop_train_theme_addbanner.asp?eC=<%=eventid%>&smode=<%=mode%>&saveafter=<%=saveafter%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
