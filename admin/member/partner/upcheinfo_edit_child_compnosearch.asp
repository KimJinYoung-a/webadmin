<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ü�������/����
' History : 2015.05.27 ���ر� ����
'			2021.12.06 �ѿ�� ����(���Ѽ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/admin/member/partner/partnerCls.asp"-->

<%
	Dim ogroup,i, vTIdx, vAction, socno, vMessage, vQuery
	Dim groupid, vSearchGroupID, vGubun, vCompNOchgOX
	
	vTIdx 			= request("tidx")
	groupid 		= request("groupid")
	vGubun 			= Request("gb")
	vAction			= Request("action")
	socno 			= request("socno")
	
	
	If vAction = "search" Then
		'// ����ڹ�ȣ �Ǵ� �ֹι�ȣ
		If Len(socno) <> 12 and Len(socno) <> 14 Then
			Response.Write "<script>alert('����ڵ�Ϲ�ȣ ������ �߸��Ǿ����ϴ�.\n�ٽ� Ȯ���Ͻð� �˻����ּ���.1');history.back();</script>"
			dbget.close()
			Response.End
		Else
			If Not (12 - Len(Replace(socno,"-","")) = 2) and Not (14 - Len(Replace(socno,"-","")) = 1) Then
				Response.Write "<script>alert('����ڵ�Ϲ�ȣ ������ �߸��Ǿ����ϴ�.\n�ٽ� Ȯ���Ͻð� �˻����ּ���.');history.back();</script>"
				dbget.close()
				Response.End
			End IF
		End IF
		
		set ogroup = new CPartnerGroup
		
		ogroup.FPageSize = 20
		ogroup.FCurrPage = 1
		ogroup.FRectsocno = socno
		
		ogroup.GetGroupInfoList
		
		if (ogroup.FResultCount > 0) then
			vMessage = ogroup.FItemList(0).Fcompany_name & "(" & socno & ")<br>�̹� ��ϵ� ��ü�Դϴ�."
			vSearchGroupID = ogroup.FItemList(0).Fgroupid
		else
			If vGubun = "newcomp" Then
				vQuery = "SELECT TOP 1 (SELECT username FROM [db_partner].[dbo].[tbl_user_tenbyten] WHERE userid = A.reguserid) FROM [db_partner].[dbo].[tbl_partner_temp_info] AS A WHERE company_no = '" & socno & "' AND status NOT IN ('0','3') "
				rsget.Open vQuery,dbget
				IF Not rsget.EOF THEN
					vMessage = "" & rsget(0) & " ���� �� ����ڹ�ȣ��<br>��û�� ������� �ֽ��ϴ�.<br>�� ���� �Ϸ�� �� ��û�� �� �ֽ��ϴ�."
					vSearchGroupID = "x"
				Else
					vMessage = "��ϰ����� ����ڹ�ȣ�Դϴ�."
				END IF
				rsget.close()
			Else
				vMessage = "��ϰ����� ����ڹ�ȣ�Դϴ�."
			End IF
		end if
		
		set ogroup = nothing
	End If
%>

<script language="javascript">
function goThisGroupcode(gcode)
{
	if(gcode == "")
	{
		document.location.href = "/admin/member/partner/upcheinfo_edit_child2.asp?socno=<%=socno%>&groupid_old=<%=groupid%>&gb=<%=vGubun%>&tidx=<%=vTIdx%>";
	}
	else
	{
		document.location.href = "/admin/member/partner/upcheinfo_edit_child2.asp?groupid="+gcode+"&groupid_old=<%=groupid%>&gb=<%=vGubun%>&tidx=<%=vTIdx%>";
	}
}
function goNewCompReg(){
	document.location.href = "/admin/member/partner/upcheinfo_new.asp?socno=<%=socno%>&gb=newcompreg";
}
</script>

<form name="frm1" method="post">
<input type="hidden" name="action" value="search">
<input type="hidden" name="tidx" value="<%=vTIdx%>">
<input type="hidden" name="gb" value="<%=vGubun%>">
<input type="hidden" name="groupid" value="<%=groupid%>">
<br><br><br><br><br><br><br><br>
<table width="250" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td style="padding-left:15px;">* ����ڹ�ȣ �Է�(<font color="blue">- �� �� �־��ּ���.</font>)</td>
</tr>
<tr height="50" valign="middle" bgcolor="FFFFFF">
	<td align="center" style="padding:5px 0 5px 0;">
		<input type="text" name="socno" value="<%=socno%>" maxlength="14" size="20">&nbsp;<input type="submit" class="button" value="�˻�"><%'' 12->14 %>
		<% If vMessage <> "" Then %>
		<br>&nbsp;<br><br>
		<%=vMessage%>&nbsp;
			<% If vGubun = "newcomp" Then %>
				<% If vSearchGroupID = "" Then %>
					<input type="button" class="button" value="����" onClick="goNewCompReg('<%=socno%>');">
				<% End If %>

			<% Else %>
				<% if C_MngPart or C_ADMIN_AUTH then %>
					<Br><Br>[�����ڸ��]
					<br>�׷��ڵ�:<input type="text" name="SearchGroupID" value="<%=vSearchGroupID%>" size=10 maxlength=10>
					<input type="button" class="button" value="����" onClick="goThisGroupcode(frm1.SearchGroupID.value);">
				<% else %>
					<input type="button" class="button" value="����" onClick="goThisGroupcode('<%=vSearchGroupID%>');">
				<% end if %>
			<% End If %>
		<% End IF %>
	</td>
</tr>
</table>
</form>

<script language="javascript">
frm1.socno.focus();
</script>

<%
function checkidexist(userid)
        dim sql

        sql = "select top 1 userid from [db_user].[dbo].tbl_logindata where userid = '" + userid + "'"
        rsget.Open sql,dbget,1

        checkidexist = (not rsget.EOF)

        rsget.close

        sql = "select userid from [db_user].[dbo].tbl_deluser where userid = '" + userid + "'"
		rsget.Open sql, dbget, 1
		checkidexist = checkidexist or (Not rsget.Eof)
		rsget.Close
end function

function checksocnoexist(socno)
        dim sql

        sql = "select top 1 userid from [db_user].[dbo].tbl_user_c where socno = '" + socno + "'"
        rsget.Open sql,dbget,1

        checksocnoexist = (not rsget.EOF)

        rsget.close
end function


function checkspecialpass(target)
        dim buf, result, index

        index = 1
        do until index > len(target)
                buf = mid(target, index, cint(1))
                if (buf="'") or (buf="`") then
                        checkspecialpass = true
                        exit function
                else
                        result = false
                end if
                index = index + 1
        loop
        checkspecialpass = false
end function

function checkspecialchar(target)
        dim buf, result, index

        index = 1
        do until index > len(target)
                buf = mid(target, index, cint(1))
                if (lcase(buf) >= "a" and lcase(buf) <= "z") then
                        result = false
                elseif (buf >= "0" and buf <= "9") then
                        result = false
                else
                        checkspecialchar = true
                        exit function
                end if
                index = index + 1
        loop
        checkspecialchar = false
end function
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->