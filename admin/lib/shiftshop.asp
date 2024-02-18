<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 오프샵 직영매장 직원 권한 재설정
' Hieditor : 2011.01.27 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim userid ,shiftid ,IsValidShiftID ,sqlStr
	userid  = session("ssBctId")
	shiftid = request("shiftid")

dim ref
	ref = request.ServerVariables("HTTP_REFERER")

''로그인
sqlStr = "select top 30" + VbCrlf
sqlStr = sqlStr & " ut.userid , ps.shopid" + VbCrlf
sqlStr = sqlStr + " from db_partner.dbo.tbl_user_tenbyten ut" + vbcrlf 
sqlStr = sqlStr + " join db_partner.dbo.tbl_partner_shopuser ps" + vbcrlf 
sqlStr = sqlStr + " 	on ps.empno = ut.empno" + vbcrlf
sqlStr = sqlStr & " where ut.userid = '"&userid&"'" & vbcrlf
sqlStr = sqlStr & " and ut.isusing=1" & vbcrlf

' 퇴사예정자 처리	' 2018.10.16 한용민
sqlStr = sqlStr & " and (ut.statediv ='Y' or (ut.statediv ='N' and datediff(dd,ut.retireday,getdate())<=0))" & vbcrlf
sqlStr = sqlStr & " and ps.shopid = '"&shiftid&"'" & vbcrlf

'response.write sqlStr & "<br>"
rsget.Open sqlStr,dbget,1

if not rsget.EOF  then        
    session("ssBctBigo") = rsget("shopid")
	response.Cookies("partner").domain = "10x10.co.kr"           
    IsValidShiftID = true
else
	IsValidShiftID = false
end if
rsget.close
%>

<% if Not IsValidShiftID then %>
	<script language='javascript'>
		alert('매장이 지정되어 있지 않습니다.관리자 문의하세요');
	</script>
<% end if %>

<script language='javascript'>
	location.replace('<%= ref %>')
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->