<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###############################################
' PageName : pop_themeslide_proc.asp
' Discription : ����� slide process
' History : 2019-02-11 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	Dim menuidx , mode , device , idx, saveafter
	Dim slideimg , linkurl , sorting , isusing '�����̵� �̹���
	Dim sqlStr, videoLink, hvideotype, videoFullLink, eCode
	Dim sIdx, sSortNo, sIsUsing, i , slinkurl, bgleft, bgright '//�����̵�

	idx	= requestCheckVar(Request.form("idx"),10)
	menuidx	= requestCheckVar(Request.form("menuidx"),10)
	mode = requestCheckVar(Request.form("mode"),6)
	device = requestCheckVar(Request.form("device"),1)
	slideimg = requestCheckVar(Request.form("slideimg"),200)
	eCode	= requestCheckVar(Request.form("evt_code"),10)
'Response.write mode & "<br>"
Select Case mode
	Case "SU"
		'//����Ʈ��������
		for i=1 to request.form("idx").count
			sIdx = request.form("idx")(i)
			sSortNo = request.form("viewidx")(i)
			if sSortNo="" then sSortNo="0"

			sqlStr = sqlStr & "Update db_event.dbo.tbl_event_multi_contents Set " & vbCrLf
			sqlStr = sqlStr & " viewidx=" & sSortNo & "" & vbCrLf
			sqlStr = sqlStr & " Where idx='" & sIdx & "';"
		Next

		If sqlStr <> "" then
			dbget.Execute sqlStr
		Else
			Call Alert_return("������ ������ �����ϴ�.")
			dbget.Close: Response.End
		End If 
        response.write "<script type='text/javascript'>"
        response.write "    window.document.domain = ""10x10.co.kr"";"
        response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=5&viewset='+opener.document.frmEvt.viewset.value);"
        'response.write "    location.replace('" + refer + "');"
        response.write "    self.close();"
        response.write "</script>"
        dbget.close()	:	response.End
	Case "SD" '����
		sIdx = request.form("idx")

		sqlStr = "delete from db_event.dbo.tbl_event_multi_contents Where idx='"& sIdx &"' and device = '"& device &"'"
		dbget.Execute sqlStr
        response.write "<script type='text/javascript'>"
        response.write "    window.document.domain = ""10x10.co.kr"";"
        response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=5&viewset='+opener.document.frmEvt.viewset.value);"
        'response.write "    location.replace('" + refer + "');"
        response.write "    self.close();"
        response.write "</script>"
        dbget.close()	:	response.End
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->