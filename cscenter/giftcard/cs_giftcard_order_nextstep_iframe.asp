<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

dim requserid, reqhp

requserid    = request("requserid")
reqhp        = request("reqhp")

dim sqlStr

if (requserid = "") and (reqhp = "") then
	dbget.close()	:	response.End
end if



sqlStr = " select top 1 "
sqlStr = sqlStr + " 	userid, usercell as reqhp "
sqlStr = sqlStr + " from "
sqlStr = sqlStr + " 	db_user.dbo.tbl_user_n "
sqlStr = sqlStr + " where "

if (requserid <> "") then
    sqlStr = sqlStr + " 	userid = '" + CStr(requserid) + "' "
else
	sqlStr = sqlStr + " 	usercell = '" + CStr(reqhp) + "' "
end if
rsget.Open sqlStr,dbget,1
if Not rsget.Eof then

	if (requserid <> "") then
%>
		<script>
		if (parent.frm.reqhp.value == "<%= rsget("reqhp") %>") {
			alert("�ǹٸ� ���̵��Դϴ�.");
		} else {
			alert("�˻��� �ڵ��� ��ȣ�� ��ġ���� �ʽ��ϴ�.\n\n�˻��� �ڵ�����ȣ : <%= rsget("reqhp") %>");
		}
		</script>
<%
	else
%>
		<script>
		alert("���̵� �˻��Ǿ����ϴ�.\n\n�˻��� ���̵� : <%= rsget("userid") %>");
		parent.frm.requserid.value = "<%= rsget("userid") %>";
		</script>
<%
	end if
else
	response.write "<script>alert('�˻������ �����ϴ�.')</script>"
end if
rsget.close

response.End

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
