<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : �����ν� ���°���
' Hieditor : 2011.03.22 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim placeid , placeiname,imagetype ,sqlStr ,validpart , tmpplaceid ,mode ,isusing
	placeid   = requestCheckVar(request("placeid"),10)
	placeiname   = requestCheckVar(request("placeiname"),32)
	validpart  = requestCheckVar(request("validpart"),10)
	mode = requestCheckVar(request("mode"),32)
	isusing = requestCheckVar(request("isusing"),1)
		
	tmpplaceid = false
	
dim referer
	referer = request.ServerVariables("HTTP_REFERER")

'//����
if mode = "EDIT" then
    sqlStr = " update db_partner.dbo.tbl_user_inouttime_place set" + VbCrlf
    sqlStr = sqlStr + " placeiname='" + html2db(placeiname) + "'" + VbCrlf
    sqlStr = sqlStr + " ,validpart='" + validpart + "'" + VbCrlf
    sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VbCrlf
    sqlStr = sqlStr + " where placeid=" + CStr(placeid) + VbCrlf
    
    'response.write sqlStr
    dbget.Execute sqlStr

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	'response.write "	opener.location.reload();"
	response.write "	location.replace('/common/member/fingerprints/fingerprints_poscode.asp?placeid="&placeid&"&mode=EDIT');"
	response.write "</script>"

'//�űԵ��
elseif mode = "ADD" then

	sqlStr = "select placeid "
	sqlStr = sqlStr + " from db_partner.dbo.tbl_user_inouttime_place" + VbCrlf
	sqlStr = sqlStr + " where placeid =" + placeid
	
	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1
		if not rsget.eof then
	    	tmpplaceid = true
	    end if
	rsget.Close
	
	if tmpplaceid then
		response.write "<script type='text/javascript'>"
		response.write "	alert('�����Ǿ��ų� ,�̹� ��ϵǾ� �ִ� ��ȣ �Դϴ�');"
		response.write "	location.replace('/common/member/fingerprints/fingerprints_poscode.asp?placeid="&placeid&"&placeiname="&placeiname&"&validpart="&validpart&"&isusing="&isusing&"&mode=ADD');"
		response.write "</script>"		
	    dbget.close()	:	response.end
	end if

    sqlStr = " insert into db_partner.dbo.tbl_user_inouttime_place" + VbCrlf
    sqlStr = sqlStr + " (placeid,placeiname,validpart,isusing)"+ VbCrlf
    sqlStr = sqlStr + " values("
    sqlStr = sqlStr + " " + CStr(placeid) + VbCrlf
    sqlStr = sqlStr + " ,'" + html2db(placeiname) + "'" + VbCrlf
    sqlStr = sqlStr + " ," + validpart + "" + VbCrlf
    sqlStr = sqlStr + " ,'" + isusing + "'" + VbCrlf    
    sqlStr = sqlStr + " )" + VbCrlf
    
    'response.write sqlStr
    dbget.Execute sqlStr

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	'response.write "	opener.location.reload();"
	response.write "	location.replace('/common/member/fingerprints/fingerprints_poscode.asp?placeid="&placeid&"&mode=EDIT');"
	response.write "</script>"	
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
