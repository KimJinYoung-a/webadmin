<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ������ �ֹ��� �ۼ�
' History : 2012.09.13 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->

<%
dim menupos , mode , shopid , cwflag ,shopdiv
	menupos = requestCheckVar(request("menupos"),10)
	mode = requestCheckVar(request("mode"),32)
	shopid = requestCheckVar(request("shopid"),32)
	cwflag = requestCheckVar(request("cwflag"),32)

if mode="chcwflag" then
	if shopid = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('���� ���̵� �����ϴ�.');"
		response.write "</script>"
		response.end	:	dbget.close()
	end if
	if cwflag = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��� �����ڰ� �����ϴ�.');"
		response.write "</script>"
		response.end	:	dbget.close()
	end if

	shopdiv = getoffshopdiv(shopid)
	
	'//�ش������ �����Ź ����� �ְ�, ���屸���� ��������ϰ�� ������ �����Ź���� �⺻ ����
	if getcwflag(shopid,"B013") = "1" and shopdiv = "13" then
		response.write "<script type='text/javascript'>"
		response.write "	var divcwflag = parent.document.getElementById('divcwflag');"
		response.write "	divcwflag.style.display = '';"

		response.write "	var cwflag = parent.document.getElementsByName('cwflag');"
		response.write "	cwflag[1].checked = true;"
	
		response.write "</script>"
		response.end	:	dbget.close()

	'//�ش������ �����Ź ����� �ִ°��, ����� �����Ҽ� �ְ�
	elseif getcwflag(shopid,"B013") = "1" then
		response.write "<script type='text/javascript'>"
		response.write "	var divcwflag = parent.document.getElementById('divcwflag');"
		response.write "	divcwflag.style.display = '';"
		response.write "</script>"
		response.end	:	dbget.close()		
	else
		response.write "<script type='text/javascript'>"
		response.write "	var divcwflag = parent.document.getElementById('divcwflag');"
		response.write "	divcwflag.style.display = 'none';"
		response.write "	var cwflag = parent.document.getElementsByName('cwflag');"
		response.write "	cwflag[0].checked = true;"		
		response.write "</script>"
		response.end	:	dbget.close()
	end if
	
else
	response.write "<script type='text/javascript'>"
	response.write "	alert('�����ڰ� �����ϴ�.');"
	response.write "</script>"
	response.end	:	dbget.close()
end if

%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->