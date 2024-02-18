<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
'####################################################
' Description :  내부 코드 바코드로 변환
' History : 2012.07.05 한용민 생성
'####################################################

dim barcode ,image , barcodetype ,height ,barwidth
	barcode = requestCheckVar(request("barcode"),30)
	image = requestCheckVar(request("image"),1)
	barcodetype = requestCheckVar(request("barcodetype"),2)
	height = requestCheckVar(request("height"),3)
	barwidth = requestCheckVar(request("barwidth"),3)
					
if barcode = "" then
	response.write "바코드가 지정되지 않았습니다."
	response.end
end if

if image = "" then image = 3
if barcodetype = "" then barcodetype = 23
if height = "" then height = 40
if barwidth = "" then barwidth = 1
%>
<Br><Br><Br><Br><Br><Br><Br><Br>
<div id="barocde">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<img src="http://company.10x10.co.kr/barcode/barcode.asp?image=<%=image%>&type=<%=barcodetype%>&data=<%=trim(barcode)%>&height=<%=height%>&barwidth=<%=barwidth%>">
</div>
