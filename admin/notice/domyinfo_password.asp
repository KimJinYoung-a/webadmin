<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��Ʈ��� �������� 
' History : 2007.07.30 �ѿ�� ����
'           2008.12.15 ������ ����(�н����� ��ȭ ��å ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->


<%

function IsSpecialCharExist(s)
    dim buf, result, index

    index = 1
    do until index > len(s)
            buf = mid(s, index, cint(1))
            if (lcase(buf) >= "a" and lcase(buf) <= "z") then
                    result = false
            elseif (buf >= "0" and buf <= "9") then
                    result = false
            else
                    IsSpecialCharExist = true
                    exit function
            end if
            index = index + 1
    loop

    IsSpecialCharExist = false
end function

dim refer
refer = request.ServerVariables("HTTP_REFERER")	

dim userid,txName,txintro,txpass1,txpass2,txpass3
	userid = session("ssBctId")
	txName = html2db(request("txName"))
	txintro = html2db(request("txintro"))
	txpass1 = html2db(request("txpass1"))
	txpass2 = html2db(request("txpass2"))
	txpass3 = html2db(request("txpass3"))
	
	dim sqlseach , virid
	sqlseach = " select id , password from [db_partner].[dbo].tbl_partner where id = '"& userid &"'"
	rsget.open sqlseach,dbget,1
		virid = rsget("password")
	
		if txpass1 = virid then		'������й�ȣ�� �ٲܺ�й�ȣ ��

			'//�н����� ��å �˻�
			if chkPasswordComplex(userid,txpass2)<>"" then
				response.write "<script language='javascript'>" &vbCrLf &_
								"	alert('" & chkPasswordComplex(userid,txpass2) & "\n��й�ȣ�� Ȯ���� �ٽ� �õ����ּ���.');" &vbCrLf &_
								" 	location.replace('" & refer & "');" &vbCrLf &_
								"</script>"
				dbget.close()	:	response.End
			end if

			'//�н����� ����
			dim sqlpassword
			sqlpassword = "update [db_partner].[dbo].tbl_partner"			& VbCrlf
			sqlpassword = sqlpassword + " set lastInfoChgDT=getdate(), password='" & txpass2 & "'"	& VbCrlf
			sqlpassword = sqlpassword + " where id='" + CStr(userid) + "'"	
			dbget.execute sqlpassword
	rsget.close
	%>

		<script language="javascript">
		alert('��й�ȣ�� ���� �Ǿ����ϴ�.');
		location.replace('<%= refer %>');
		</script>

		<% else %>

		<script language="javascript">
		alert("������й�ȣ�� Ʋ���ϴ�. ��й�ȣ �нǽ� ���ǻ��� : �ý�����");
		location.replace('<%= refer %>');
		</script>

<% end if %>