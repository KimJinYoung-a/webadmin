<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��������ֹ�
' History : 2012.05.21 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/lib/classes/order/ordergiftcls.asp"-->

<%
dim orderserial , mode ,ojumun ,isFinishValid , sqlstr, userid
	orderserial = requestCheckVar(request("orderserial"),11)
	mode = requestCheckVar(request("mode"),32)

isFinishValid= false

'/��������ֹ��Ϸ�ó��	
if mode = "siteReceivefinsh" then
	if orderserial = "" then
		response.write "<script language='javascript'>"
		response.write "	alert('�ֹ���ȣ�� �����ϴ�');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	
	'/�ֹ���ȸ
	set ojumun = new CJumunMaster
		ojumun.FRectOrderSerial = orderserial
		ojumun.SearchJumunList

	if ojumun.ftotalcount < 1 then
		response.write "<script language='javascript'>"
		response.write "	alert('�ֹ����� �����ϴ�');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	
	'/���°��� �����Ϸ�� , �Ϸ�ó�� �����̰�
	isFinishValid = (ojumun.FMasterItemList(0).FIpkumdiv>3) and (ojumun.FMasterItemList(0).FIpkumdiv<8)
	
	'/�ּ��ϰ��
	isFinishValid = isFinishValid and (ojumun.FMasterItemList(0).FCancelyn="N")
	
    '/cs �޸� �����ϱ����� 
    userid = ojumun.FMasterItemList(0).FUserID
    
	if (Not isFinishValid) then
		if (ojumun.FMasterItemList(0).FIpkumdiv>7) then
			response.write "<script language='javascript'>"
			response.write "	alert('�̹� ó�� �Ϸ�� �ֹ� �Դϴ�.');"
			response.write "</script>"
			dbget.close()	:	response.End
		elseif (ojumun.FMasterItemList(0).FIpkumdiv<4) then
			response.write "<script language='javascript'>"
			response.write "	alert('�������� �ֹ��� �Դϴ�.');"
			response.write "</script>"
			dbget.close()	:	response.End
		elseif (ojumun.FMasterItemList(0).FCancelyn<>"N") then
			response.write "<script language='javascript'>"
			response.write "	alert('��ҵ� �ֹ� �Դϴ�.');"
			response.write "</script>"
			dbget.close()	:	response.End			
		elseif (ojumun.FMasterItemList(0).Fjumundiv<>"7") then
			response.write "<script language='javascript'>"
			response.write "	alert('������� �ֹ����� �ƴմϴ�.');"
			response.write "</script>"
			dbget.close()	:	response.End
		end if
	end if
	
    
    
	set ojumun = nothing

	sqlStr = "update D" + vbcrlf
	sqlStr = sqlStr + " set currstate='7'" + vbcrlf
	sqlStr = sqlStr + " ,songjangno='��Ÿ���'" + vbcrlf
	sqlStr = sqlStr + " ,songjangdiv='99'" + vbcrlf
	sqlStr = sqlStr + " ,beasongdate=IsNULL(beasongdate,getdate())" + vbcrlf
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail D" + vbcrlf
	sqlStr = sqlStr + " Join [db_order].[dbo].tbl_order_master m" + vbcrlf
    sqlStr = sqlStr + " 	on m.orderserial=d.orderserial" + vbcrlf
	sqlStr = sqlStr + " where d.orderserial = '"&orderserial&"'" + vbcrlf
	sqlStr = sqlStr + " and d.cancelyn<>'Y'" + vbcrlf
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and m.jumundiv='7'"
	
	'response.write sqlstr & "<Br>"
	dbget.execute sqlstr

    sqlStr = "update [db_order].[dbo].tbl_order_master" + vbcrlf
    sqlStr = sqlStr + " set ipkumdiv='8'" + vbcrlf
    sqlStr = sqlStr + " , beadaldate=getdate()" + vbcrlf
	sqlstr = sqlstr & " ,songjangdiv = '99'" + vbcrlf
	sqlstr = sqlstr & " ,deliverno = '��Ÿ���'" + vbcrlf
    sqlStr = sqlStr + " where orderserial in (" + vbcrlf
    sqlStr = sqlStr + "     select m.orderserial" + vbcrlf
    sqlStr = sqlStr + "     from [db_order].[dbo].tbl_order_master m" + vbcrlf
    sqlStr = sqlStr + "     left join [db_order].[dbo].tbl_order_detail d" + vbcrlf
    sqlStr = sqlStr + "         on m.orderserial=d.orderserial" + vbcrlf
    sqlStr = sqlStr + "     where m.orderserial = '"&orderserial&"'" + vbcrlf
    sqlStr = sqlStr + "     and m.cancelyn='N'" + vbcrlf
    sqlStr = sqlStr + "     and m.jumundiv<>9" + vbcrlf
    sqlStr = sqlStr + "     and d.itemid<>0" + vbcrlf
    sqlStr = sqlStr + "     and d.cancelyn<>'Y'" + vbcrlf
	sqlStr = sqlStr + " 	and m.jumundiv='7'"    
    sqlStr = sqlStr + "     group by m.orderserial" + vbcrlf
    sqlStr = sqlStr + "     having sum(case when IsNull(d.currstate,'0')<>'7' then 1 else 0 end )=0"
    sqlStr = sqlStr + " ) "
	
	'response.write sqlstr & "<Br>"
	dbget.execute sqlstr

    ''CS �޸� ���� // ������ �߰�
    sqlStr = " insert into [db_cs].[dbo].tbl_cs_memo(orderserial, divcd, userid, mmgubun, qadiv, phoneNumber, writeuser, finishuser, contents_jupsu, finishyn,finishdate,regdate) "
    sqlStr = sqlStr + " values('" + CStr(orderserial) + "','1','" + CStr(userid) + "','0','20','','" + session("ssBctId") + "','" + session("ssBctId") + "','������� �Ϸ�����','Y',getdate(),getdate()) "
    dbget.Execute sqlStr
        
	response.write "<script language='javascript'>"
	''response.write "	parent.opener.location.reload();"	
	response.write "	alert('���Ϸ� ó�� �Ǿ����ϴ�.');"
	response.write "	parent.location.href='/admin/apps/siteReceive/popSiteReceive.asp?orderserial="&orderserial&"&aplot=Y';"
	''response.write "	parent.plotReceipt();"
	response.write "</script>"
	dbget.close()	:	response.End

else
	response.write "<script language='javascript'>"
	response.write "	alert('�����ڰ� �����ϴ�');"
	response.write "</script>"
	dbget.close()	:	response.End
end if	
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->