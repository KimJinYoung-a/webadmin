<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : mdpick_Process.asp
' Discription : mdpick ó�� ������
' History : 2016.08.02 ���¿�
'###############################################

dim menupos, mode, sqlStr, lp
dim idx, itemid, title, startdate, enddate, isusing, sortno	''img1
idx			= RequestCheckvar(Request("idx"),10)
mode		= RequestCheckvar(Request("mode"),16)
sortno		= RequestCheckvar(Request("sortno"),10)
'img1		= Request("image1")
menupos	= RequestCheckvar(Request("menupos"),10)
enddate	= RequestCheckvar(Request("enddate"),10)
isusing	= RequestCheckvar(Request("isusing"),1)
startdate	= RequestCheckvar(Request("startdate"),10)
itemid		= getNumeric(Request("itemid"))
title		= html2db(Request("mdpicktitle"))
  	if title <> "" then
		if checkNotValidHTML(title) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
'// Ʈ������ ����
dbACADEMYget.beginTrans

'// ��忡 ���� �б�
Select Case mode
	Case "add"
		'// �ű� ���
		sqlStr = "Insert Into [db_academy].[dbo].tbl_mdpick " &_
				" (itemid, title, startdate, enddate, isusing, sortno, adminid) values " &_
				" (" & itemid & "" &_
				" ,'" & title & "' " &_
				" ,'" & startdate & "' " &_
				" ,'" & enddate & "' " &_
				" ,'" & isusing & "' " &_
				" ,'" & sortno & "' " &_
				" ,'" & session("ssBctId") & "')"

		dbACADEMYget.Execute(sqlStr)

	Case "edit"
		'// ���� ����
		sqlStr = "Update [db_academy].[dbo].tbl_mdpick " &_
				" Set title='" & title & "'" &_
				" 	,startdate='" & startdate & "'" &_
				" 	,enddate='" & enddate & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,sortno='" & sortno & "'" &_
				" Where idx='" & idx & "'"
'response.write sqlStr
'response.end
		dbACADEMYget.Execute(sqlStr)

	Case "delete"
'		if idx <> "" then
'			sqlStr = sqlStr & "delete [db_sitemaster].[dbo].tbl_mdpick " &_
'					" Where idx='" & idx & "';" & vbCrLf
'			dbACADEMYget.Execute(sqlStr)
'		end if
End Select


'// Ʈ������ �˻� �� ����
If Err.Number = 0 Then
        dbACADEMYget.CommitTrans
Else
        dbACADEMYget.RollBackTrans
		Alert_return("����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.")
		dbACADEMYget.close()	:	response.End
End If

%>
<script language="javascript">
<!--
	// ������� ����
	alert("�����߽��ϴ�.");
	self.location = "mdpick_list.asp?menupos=<%=menupos%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
