<%
if (session("isAgreeReq")="Y") then
    ''if (application("Svr_Info")	= "Dev") then  ''서비스 오픈 후 주석처리.
    response.write "<html><body><br><br>"
    response.write "<p align='center'><font size=2>먼저 약관/계약 동의후 이용가능합니다.</font></p>"
    response.write "<p align='center'><input type='button' value='약관/계약 동의 메뉴로 이동' onClick=""location.href='/lectureadmin/contract/ctrListBrand.asp?menupos=1816'""></p>"
    response.write "</body></html>"
    response.end
    ''end if
end if
%>