<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim sqlStr, itemid, cDisp, cmdparam, chgItemname
Dim arrItemid : arrItemid = request("cksel")
Dim isdisplay

itemid 		= request("itemid")
chgItemname	= Trim(requestCheckVar(request("chgItemname"),64))
cmdparam	= request("cmdparam")
isdisplay	= Trim(request("isdisplay"))

If cmdparam = "chgname" Then
	sqlStr = ""
	sqlStr = sqlStr & " UPDATE db_outmall.dbo.tbl_between_cate_item SET "
	sqlStr = sqlStr & " chgItemname = '"&html2db(chgItemname)&"' "
	sqlStr = sqlStr & " WHERE itemid = '"& itemid &"' " 
	dbCTget.execute sqlStr
	response.write "<script language='javascript'>alert('����Ǿ����ϴ�');opener.location.reload();window.close();</script>"
ElseIf cmdparam = "EditDisplay" Then
    arrItemid = Trim(arrItemid)
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('���õ� ��ǰ�� �����ϴ�.\nȮ�� �� �ٽ� �õ����ּ���.');history.back(-1);</script>"
		dbCTget.Close: Response.End
	End If

	sqlStr = ""
	sqlStr = sqlStr & " UPDATE db_outmall.dbo.tbl_between_cate_item SET "
	sqlStr = sqlStr & " isdisplay = '"&isdisplay&"' "
	sqlStr = sqlStr & " WHERE itemid in ("& arrItemid &") " 
	dbCTget.execute sqlStr
	response.write "<script language='javascript'>alert('����Ǿ����ϴ�');parent.location.reload()</script>"
End If

%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->