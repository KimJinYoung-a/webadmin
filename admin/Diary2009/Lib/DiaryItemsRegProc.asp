<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary2009/lib/include_event_code.asp"-->
<%
'// ���̾ ��Ƽ ���� ó�� ������ 2018-08-17 ����ȭ
dim i
dim CateCode , mode
dim idx , tempidx
dim itemcount
dim strSQL
dim idxStrSQL

mode = request.Form("mode")
CateCode = request.Form("cate")
itemcount = request.Form("chkitem").count

if CateCode = "" then CateCode = 0
	
IF mode="I" Then
        dbget.beginTrans
    For i = 1 To itemcount	'���ϰ��� ��ŭ ���ε�
	    strSQL =" INSERT INTO db_diary2010.[dbo].tbl_DiaryMaster " & vbcrlf
        strSQL = strSQL & " (Cate,Itemid,isusing,commentyn,event_code,eventgroup_code,comment_img ,weight, mdpick, limited, storytext , mdpicksort, event_start, event_end) " & vbcrlf
        strSQL = strSQL & " VALUES("  & vbcrlf
        strSQL = strSQL & "'" & CateCode & "' "  & vbcrlf
        strSQL = strSQL & ",'" & request.Form("chkitem")(i) & "' "  & vbcrlf
        strSQL = strSQL & ",'Y' "  & vbcrlf
        strSQL = strSQL & ",'' "  & vbcrlf
        strSQL = strSQL & ",'0' "  & vbcrlf
        strSQL = strSQL & ",'0' "  & vbcrlf
        strSQL = strSQL & ",'' "  & vbcrlf
        strSQL = strSQL & ",'0' "  & vbcrlf
        strSQL = strSQL & ",'x' "  & vbcrlf
        strSQL = strSQL & ",'x' "  & vbcrlf
        strSQL = strSQL & ",''"  & vbcrlf
        strSQL = strSQL & ",'0' "  & vbcrlf
	    strSQL = strSQL & ",null "  & vbcrlf
        strSQL = strSQL & ",null "  & vbcrlf
        strSQL = strSQL & " )"

	    'response.write strSQL&"<br>"
	    dbget.execute(strSQL)

        idxStrSQL = "SELECT SCOPE_IDENTITY()"
        rsget.open idxStrSQL,dbget,2
        IF not rsget.Eof Then
            tempidx = rsget(0)
        End IF
        rsget.close

	    idx = tempidx

        '2019 ���̾ ���� ����
        strSQL = " INSERT INTO [db_diary2010].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " & vbcrlf
        strSQL = strSQL & " VALUES " & vbcrlf
        'strSQL = strSQL & "('" & idx & "','22','1����','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','23','�б⺰','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','24','6����','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','25','1��','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','26','1�� �̻�','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','27','����������','0') ," & vbcrlf
        'strSQL = strSQL & "('" & idx & "','28','����������','0') ," & vbcrlf
        'strSQL = strSQL & "('" & idx & "','29','�ְ�������','0') ," & vbcrlf
        'strSQL = strSQL & "('" & idx & "','30','�Ͻ�����','0') ," & vbcrlf
        'strSQL = strSQL & "('" & idx & "','31','ĳ�ú�','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','32','����','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','33','���','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','34','��Ȧ��','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','35','������','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','36','2019 ��¥��','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','37','�ս���','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','38','��Ŭ��','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','39','���ϸ�','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','40','���̾','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','41','���͵�','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','42','�����','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','43','�ڱ���','0') " & vbcrlf	
        dbget.execute(strSQL)
    Next 
	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)
        Alert_return "���� �Ǿ����ϴ�."
        response.write "<script type='text/javascript'>parent.window.close();</script>"
	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)
		Alert_return "ó���� ������ �߻��߽��ϴ�."
	End If
End IF
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->