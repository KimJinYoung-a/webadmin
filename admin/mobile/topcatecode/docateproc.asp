<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : domdpick.asp
' Discription : mdpick ó�� ������
' History : 2013.12.16 ����ȭ ����
'###############################################

'// ���� ���� �� �Ķ���� ����
dim menupos, mode, sqlStr , totcnt
dim idx, isusing
Dim gcode ,  dcode

menupos	= Request("menupos")
isusing		= Request("isusing")
mode		= Request("mode")
idx			= getNumeric(Request("idx"))
gcode		= getNumeric(Request("gcode"))
dcode		= getNumeric(Request("dcode"))

'// ��忡 ���� �б�
Select Case mode
	Case "add"

		SqlStr = "select count(*) "
        SqlStr = SqlStr + " from db_sitemaster.[dbo].[tbl_mobile_main_topsubcode] "
        SqlStr = SqlStr + " where dispcode=" + CStr(dcode) 
		rsget.Open SqlStr, dbget, 1
        if Not rsget.Eof then
            totcnt = rsget(0)
        end if
        rsget.close

		If totcnt = 0 then
			'�ű� ���
			sqlStr = "Insert Into db_sitemaster.dbo.tbl_mobile_main_topsubcode " &_
						" (gnbcode, dispcode , adminid , isusing ) values " &_
						" ('" & gcode &"'" &_
						" ,'" & dcode &"'" &_
						" ,'" & session("ssBctId") &"'" &_
						" ,'" & isusing &"'" &_
						")"
			'response.write sqlStr
			dbget.Execute(sqlStr)
		Else
			Response.Write "<script>alert('�̹� ��ϵ� ī�װ� �Դϴ�.'); history.back(-1);</script>"
			dbget.close() : Response.End
		End If 

	Case "modify"
		'���� ����
		sqlStr = "Update db_sitemaster.dbo.tbl_mobile_main_topsubcode " &_
				" Set gnbcode='" & gcode & "'" &_
				" 	,dispcode='" & dcode & "'" &_
				" 	,lastadminid='" & session("ssBctId") & "'" &_
				" 	,lastupdate=getdate()" &_
				" 	,isusing='" & isusing & "'" &_
				" Where idx=" & idx
		dbget.Execute(sqlStr)
End Select

%>
<script>
<!--
	// ������� ����
	alert("�����߽��ϴ�.");
	window.opener.document.location.href = window.opener.document.URL;    // �θ�â ���ΰ�ħ
	 self.close();        // �˾�â �ݱ�
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->