<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
	'// ���� ��� ����
	dim mode, sqlStr
	mode = Request("mode")

    '// ��ǰ��ȣ/�ɼǹ�ȣ�� �޴´� //
    dim cardItemId, cardOption
    cardItemId = Request("cardItemId")
    cardOption = Request("cardOption")

	'// Ʈ������ ����
	dbget.beginTrans

	Select Case mode
		Case "add"		'// �ɼ� �ű� ���
			'�ɼǹ�ȣ ����
			sqlStr = "Select Max(cardOption) as cardOption From db_item.dbo.tbl_giftcard_option Where cardItemId=" & cardItemId
			rsget.Open sqlStr,dbget,1
				cardOption = Num2Str(cInt(rsget("cardOption"))+1,4,"0","R")
			rsget.close

			sqlStr = "Insert into db_item.dbo.tbl_giftcard_option " & vbCrLf
			sqlStr = sqlStr & "(cardItemId, cardOption, cardOptionName, cardSellCash, cardSalePrice, cardOrgPrice, optSellYn) " & vbCrLf
			sqlStr = sqlStr & "values " & vbCrLf
			sqlStr = sqlStr & "(" & cardItemId & vbCrLf
			sqlStr = sqlStr & ",'" & cardOption & "'" & vbCrLf
			sqlStr = sqlStr & ",'" & html2db(Request("cardOptionName")) & "'" & vbCrLf
			sqlStr = sqlStr & "," & Request("cardSellCash") & ",0," & Request("cardSellCash") & vbCrLf
			sqlStr = sqlStr & ",'" & Request("optSellYn") & "')" & vbCrLf
			
			dbget.execute(sqlStr)

		Case "modi"	'// �ɼ� ����
			sqlStr = "Update db_item.dbo.tbl_giftcard_option Set " & vbCrLf
			sqlStr = sqlStr & " cardOptionName='" & html2db(Request("cardOptionName")) & "'" & vbCrLf
			sqlStr = sqlStr & " ,cardSellCash=" & Request("cardSellCash") & vbCrLf
			sqlStr = sqlStr & " ,cardOrgPrice=" & Request("cardSellCash") & vbCrLf
			sqlStr = sqlStr & " ,optSellYn='" & Request("optSellYn") & "'" & vbCrLf
			sqlStr = sqlStr & " Where cardItemId=" & cardItemId & vbCrLf
			sqlStr = sqlStr & "		and cardOption='" & cardOption & "'" & vbCrLf

			dbget.execute(sqlStr)

		Case "del"	'// �ɼ� ����
			sqlStr = "Update db_item.dbo.tbl_giftcard_option Set " & vbCrLf
			sqlStr = sqlStr & " optIsUsing='N' " & vbCrLf
			sqlStr = sqlStr & " Where cardItemId=" & cardItemId & vbCrLf
			sqlStr = sqlStr & "		and cardOption='" & cardOption & "'" & vbCrLf

			dbget.execute(sqlStr)

	End Select

	'##### DB ���� ó�� #####
    If Err.Number = 0 Then
    	dbget.CommitTrans				'Ŀ��(����)
    	Response.Write "<script language='javascript'>" & vbCrLf
    	Response.Write "alert('�����͸� �����Ͽ����ϴ�.');" & vbCrLf
    	Response.Write "top.opener.history.go(0);" & vbCrLf
    	Response.Write "top.self.close();" & vbCrLf
    	Response.Write "</script>"
    	
    Else
        dbget.RollBackTrans				'�ѹ�(�����߻���)
        Call Alert_msg("ó���� ������ �߻��߽��ϴ�.")
    End If
%>