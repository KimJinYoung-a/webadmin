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
	response.write "<script language='javascript'>alert('변경되었습니다');opener.location.reload();window.close();</script>"
ElseIf cmdparam = "EditDisplay" Then
    arrItemid = Trim(arrItemid)
	If arrItemid = "" Then
		Response.Write "<script language=javascript>alert('선택된 상품이 없습니다.\n확인 후 다시 시도해주세요.');history.back(-1);</script>"
		dbCTget.Close: Response.End
	End If

	sqlStr = ""
	sqlStr = sqlStr & " UPDATE db_outmall.dbo.tbl_between_cate_item SET "
	sqlStr = sqlStr & " isdisplay = '"&isdisplay&"' "
	sqlStr = sqlStr & " WHERE itemid in ("& arrItemid &") " 
	dbCTget.execute sqlStr
	response.write "<script language='javascript'>alert('변경되었습니다');parent.location.reload()</script>"
End If

%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->