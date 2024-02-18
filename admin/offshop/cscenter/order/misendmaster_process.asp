<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 고객센터
' Hieditor : 2011.03.10 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->
<%
dim mode ,detailIDx ,orderno, didx, finishstr, state ,sqlStr,i,masteridx
	detailIDx = requestCheckVar(request("detailIDx"),10)
	mode = requestCheckVar(request("mode"),32)
	orderno = requestCheckVar(request("orderno"),16)
	masteridx = requestCheckVar(request("masteridx"),10)
	didx        = request("didx")
	finishstr   = request("finishstr")
	state       = request("state")
	
	didx = didx + ",,"
	didx = Split(didx, ",")
	
	finishstr = finishstr + ",,"
	finishstr = Split(finishstr, ",")
	
	state = state + ",,"
	state = Split(state, ",")

dim refer 
	refer = request.ServerVariables("HTTP_REFERER")

if (mode="SendCallChange") then
    sqlStr = "update [db_shop].dbo.tbl_shopbeasong_mibeasong_list set" &VbCRLF
	sqlStr = sqlStr + " isSendCall = 'Y' " &VbCRLF
	sqlStr = sqlStr + " ,state=4" &VbCRLF
	sqlStr = sqlStr + " where detailidx=" + CStr(detailIDx) &VbCRLF
    
    'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
else
    for i = 0 to UBound(didx)
    	if (trim(didx(i)) <> "") then
    		sqlStr = "update [db_shop].dbo.tbl_shopbeasong_mibeasong_list set" &VbCRLF
    		sqlStr = sqlStr + " finishstr = '" + requestCheckVar(trim(finishstr(i)),64) + "'"&VbCRLF
    		sqlStr = sqlStr + " ,state = '" + requestCheckVar(trim(state(i)),1) + "' " &VbCRLF
    		sqlStr = sqlStr + " where detailidx=" + CStr(requestCheckVar(didx(i),10)) &VbCRLF
			
			'response.write sqlStr &"<Br>"
    		dbget.Execute sqlStr
    	end if
    next
end if
%>

<script type='text/javascript'>
	alert('저장 되었습니다.');
	location.replace('misendmaster_main.asp?masteridx=<%= masteridx %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->