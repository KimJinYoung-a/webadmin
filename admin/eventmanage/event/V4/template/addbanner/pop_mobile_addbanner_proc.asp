<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###############################################
' PageName : pop_mobile_addbanner_proc.asp
' Discription : ����� slide process
' History : 2016-02-16 ����ȭ
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	Dim eventid , mode , idx , gubun
	Dim bimg , btitle , blink , bdate_flag , bst_date , bed_date , isusing '�����̵� �̹���
	Dim sqlStr
	Dim sIdx, sSortNo, sIsUsing, i , sBlink , sGubun , sBtitle , sbst_date , sbed_date , sbimg , sbdate_flag '//�����̵�
	Dim sDt , eDt

	mode 		= requestCheckVar(Request.form("mode"),6)

	idx 		= requestCheckVar(Request.form("idx"),10)
	eventid 	= requestCheckVar(Request.form("eventid"),10)
	gubun		= requestCheckVar(Request.form("gubun"),1)

	bimg 		= requestCheckVar(Request.form("bimg"),200)
	btitle 		= trim(requestCheckVar(Request.form("btitle"),200))
	blink		= Trim(requestCheckVar(Request.form("blink"),200))

	bdate_flag	= requestCheckVar(Request.form("bdate_flag"),1)
	bst_date	= requestCheckVar(Request.form("bst_date"),10)
	bed_date	= requestCheckVar(Request.form("bed_date"),10)

	isusing		= requestCheckVar(Request.form("isusing"),1)
'	sDt			= requestCheckvar(request("sDt"),10) '//�̺�Ʈ ������
'	eDt			= requestCheckvar(request("eDt"),10) '//�̺�Ʈ ������

'//// ������� �̹��� ������ ����
Sub fnevtaddimgcnt()
	Dim imgcnt : imgcnt = 0
	sqlStr = "SELECT count(*) FROM db_event.dbo.tbl_event_mobile_addbanner where evt_code = '"&eventid&"' and isusing = 'Y'" 
	rsget.Open sqlStr,dbget,1
	IF Not rsget.Eof Then
		imgcnt = rsget(0)
	End If
	rsget.close()

	sqlStr = "update db_event.dbo.tbl_event_display set evt_m_addimg_cnt = "& imgcnt &" where evt_code = '"& eventid &"'" 
	dbget.Execute(sqlStr)
End sub

Select Case mode
	 Case "SI"
		'slide�̹��� �ű� ���
		sqlStr = "Insert Into db_event.dbo.tbl_event_mobile_addbanner " &_
					" (evt_code, gubun , bimg , btitle , blink , bdate_flag , bst_date , bed_date , isusing) values " &_
					" ('" & eventid  & "'" &_
					" ,'" & gubun &"'" &_
					" ,'" & bimg &"'" &_
					" ,'" & btitle &"'" &_
					" ,'" & blink &"'" &_
					" ,'" & bdate_flag &"'" &_
					" ,'" & bst_date &"'" &_
					" ,'" & bed_date &"'" &_
					" ,'Y')"
		dbget.Execute(sqlStr)

	    Call fnevtaddimgcnt()

	Case "SU"
		'//����Ʈ��������
		for i=1 to request.form("chkIdx").count
			sIdx = request.form("chkIdx")(i)
			sGubun = request.form("gubun"&sIdx)
			sbimg = request.form("bimg"&sIdx)
			sBtitle = request.form("btitle"&sIdx)
			sIsUsing = request.form("isusing"&sIdx)
			sBlink = request.form("blink"&sIdx)
			sbdate_flag = request.form("bdate_flag"&sIdx)
			sbst_date = request.form("bst_date"&sIdx)
			sbed_date = request.form("bed_date"&sIdx)
			if sIsUsing="" then sIsUsing="N"

			sqlStr = sqlStr & " Update db_event.dbo.tbl_event_mobile_addbanner Set "
			sqlStr = sqlStr & " gubun='" & sGubun & "'"
			sqlStr = sqlStr & " ,bimg='" & sbimg & "'"
			sqlStr = sqlStr & " ,Btitle='" & sBtitle & "'"
			sqlStr = sqlStr & " ,isusing='" & sIsUsing & "'"
			sqlStr = sqlStr & " ,blink='" & sBlink & "'"
			sqlStr = sqlStr & " ,bst_date='" & sbst_date & "'"
			sqlStr = sqlStr & " ,bed_date='" & sbed_date & "'"
			sqlStr = sqlStr & " ,bdate_flag='" & sbdate_flag & "'"
			sqlStr = sqlStr & " Where idx='" & sIdx & "';" & vbCrLf
		Next

		If sqlStr <> "" then
			dbget.Execute sqlStr

		    Call fnevtaddimgcnt()
		Else
			Call Alert_return("������ ������ �����ϴ�.")
			dbget.Close: Response.End
		End If 
	
	Case "SD" '����
		sIdx = request.form("chkIdx")

		sqlStr = "delete from db_event.dbo.tbl_event_mobile_addbanner Where idx='"& sIdx &"'"
		dbget.Execute sqlStr

	    Call fnevtaddimgcnt()
End Select
%>
<script language="javascript">
<!--
	// ������� ����
	alert("<%=chkiif(mode="SD","���� �Ϸ�.","����/���� �Ϸ�.")%>");
	//self.location = "pop_mobile_addbanner.asp?eC=<%=eventid%>&sDt=<%=sDt%>&eDt=<%=eDt%>";
	self.location = "pop_mobile_addbanner.asp?eC=<%=eventid%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
