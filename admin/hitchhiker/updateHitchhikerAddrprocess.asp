<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	PageName 	: /admin/hitchhiker/downHitchhiker.asp
'	Description : 히치하이커
'	History		: 2006.11.30 정윤정 생성
'				  2016.07.07 한용민 수정 SSL 적용
'#############################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
Dim strSql, iHVol, sUID, zipcode, addr2, userphone, usercell, recevieName, addr1
	iHVol= request("iHV")
	sUID = request("sUID")
	recevieName = request("recevieName")
	zipcode = request("zipcode")
	addr1 = requestcheckvar(request("addr1"),128)
	addr2 = request("addr2")
	userphone = request("userphone1")&"-"&request("userphone2")&"-"&request("userphone3")
	usercell = request("usercell1")&"-"&request("usercell2")&"-"&request("usercell3")

strSql = " UPDATE db_user.dbo.tbl_user_hitchhiker "&_
 		" SET "&_
 		" recevieName='" + recevieName + "'"&_	
 		" ,zipcode='" + zipcode + "'"&_	 
 		" ,zipaddr='" + addr1 + "'"&_  
		" ,useraddr='" + addr2 + "'"&_  
 		" ,userphone='" + userphone + "'"&_  
 		" ,usercell='" + usercell + "'"  &_	 
 		" where userid='" + sUID + "' and HVol='"+iHVol+"'" 

'response.write strSql & "<Br>"
dbget.execute strSql 
%>

<script type="text/javascript">
	alert("수정되었습니다");
	self.close();
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->