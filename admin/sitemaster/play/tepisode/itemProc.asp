<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
'###########################################################
' Description : t-episode ��ǰ��� ������
' Hieditor : 2014-11-20 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim idx, vQuery, i, itemid, mode , number
	idx  = requestCheckVar(request("idx"),10)
	itemid = requestCheckVar(request("itemid"),200)
	mode = requestCheckVar(request("mode"),10)
	
	IF idx = "" THEN
		Response.Write "<script>alert('�߸��� ����Դϴ�.\nNo. ��ȣ�� �־�� �մϴ�.');</script>"
		dbget.close()
		Response.End
	END IF	
	IF IsNumeric(idx) = False THEN
		Response.Write "<script>alert('�߸��� ����Դϴ�.\nNo. ��ȣ�� �־�� �մϴ�.');</script>"
		dbget.close()
		Response.End
	END IF

	If mode = "insert" Then
		For i = LBound(Split(itemid,",")) To UBound(Split(itemid,","))
			vQuery = vQuery & " IF NOT EXISTS(select subidx from db_sitemaster.dbo.tbl_play_photopick_item where subidx = '" & idx & "' and itemid = '" & Trim(Split(itemid,",")(i)) & "') " & vbCrLf
			vQuery = vQuery & " 	BEGIN " & vbCrLf
			vQuery = vQuery & " 		insert into db_sitemaster.dbo.tbl_play_photopick_item (subidx, itemid) values('" & idx & "', '" & Trim(Split(itemid,",")(i)) & "') " & vbCrLf
			vQuery = vQuery & " 	END " & vbCrLf

			dbget.execute vQuery
		Next
	ElseIf mode = "delete" Then
		vQuery = "delete db_sitemaster.dbo.tbl_play_photopick_item where subidx = '" & idx & "' and itemid IN(" & itemid & ")"
		dbget.execute vQuery
	End If
%>

<script type="text/javascript">
opener.document.location.reload();
document.location.href = "pop_itemReg.asp?idx=<%=idx%>";
</script>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->	