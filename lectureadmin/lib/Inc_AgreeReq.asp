<%
if (session("isAgreeReq")="Y") then
    ''if (application("Svr_Info")	= "Dev") then  ''���� ���� �� �ּ�ó��.
    response.write "<html><body><br><br>"
    response.write "<p align='center'><font size=2>���� ���/��� ������ �̿밡���մϴ�.</font></p>"
    response.write "<p align='center'><input type='button' value='���/��� ���� �޴��� �̵�' onClick=""location.href='/lectureadmin/contract/ctrListBrand.asp?menupos=1816'""></p>"
    response.write "</body></html>"
    response.end
    ''end if
end if
%>