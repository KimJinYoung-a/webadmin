<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim itemgubun, itemid, itemoption

itemgubun   = request("itemgubun")
itemid      = request("itemid")
itemoption  = request("itemoption")


response.write itemgubun & "<br>"
response.write itemid & "<br>"
response.write itemoption & "<br>"



''���� ������ �������� Check.
dim sqlStr
dim ErrStr
ErrStr = ""


''�ֱ� �Ǹų��� : �Ǹų����� ����ΰ�쵵.. Code�� �ٲپ����?
sqlStr = "select top 1 * from "
sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
sqlStr = sqlStr + " where d.itemid=" + CStr(itemid)
sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"

rsget.Open sqlStr, dbget, 1
if Not rsget.Eof then
	ErrStr = "�����Ϸ��� �ɼ����� �Ǹŵ� ����(6�����̳�)�� �ֽ��ϴ�. �����Ͻ� �� �����ϴ�."
end if
rsget.close

''6���� ���� �Ǹų��� : �Ǹų����� ����ΰ�쵵.. Code�� �ٲپ����?
if ErrStr="" then
	sqlStr = "select top 1 * from "
	sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d"
	sqlStr = sqlStr + " where d.itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"

	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		ErrStr = "�����Ϸ��� �ɼ����� �Ǹŵ� ����(6��������)�� �ֽ��ϴ�. �����Ͻ� �� �����ϴ�."
	end if
	rsget.close
end if

''�������
if ErrStr="" then
	sqlStr = "select top 1 * from [db_storage].[dbo].tbl_acount_storage_detail d,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.code=d.mastercode"
	sqlStr = sqlStr + " and m.deldt is NULL"
	sqlStr = sqlStr + " and d.iitemgubun='10'"
	sqlStr = sqlStr + " and d.itemid=" + CStr(itemid)
	sqlStr = sqlStr + " and d.itemoption='" + itemoption + "'"
    sqlStr = sqlStr + " and d.deldt is NULL"
    
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		ErrStr = "�����Ϸ��� �ɼ����� ����� ������ �ֽ��ϴ�. �����Ͻ� �� �����ϴ�."
	end if
	rsget.close
end if
	        	

'' �¶��� ���곻��
if ErrStr="" then
    sqlStr = "select top 1 * from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
    sqlStr = sqlStr + " where itemid=" & itemid
    sqlStr = sqlStr + " and itemoption='" & itemoption & "'"
    
    rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		ErrStr = "�����Ϸ��� �ɼ����� �¶��� ���� ������ �ֽ��ϴ�. �����Ͻ� �� �����ϴ�."
	end if
	rsget.close
end if

'' �������� ���곻��
if ErrStr="" then
    sqlStr = "select top 1 * from [db_jungsan].[dbo].tbl_off_jungsan_detail"
    sqlStr = sqlStr + " where itemgubun='10'"
    sqlStr = sqlStr + " and itemid=" & itemid
    sqlStr = sqlStr + " and itemoption='" & itemoption & "'"
    
    rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		ErrStr = "�����Ϸ��� �ɼ����� OFF ���� ������ �ֽ��ϴ�. �����Ͻ� �� �����ϴ�."
	end if
	rsget.close
end if



if (ErrStr<>"") then
    response.write "<script>alert('������ �� �����ϴ�.\n\n" & ErrStr & "');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if

'response.write "����process"
'dbget.close()	:	response.End


sqlStr = " delete from [db_summary].[dbo].tbl_daily_logisstock_summary" + VbCrlf 
sqlStr = sqlStr + " where itemgubun='" + CStr(itemgubun) + "'" + VbCrlf 
sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf 
sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"

dbget.Execute sqlStr


sqlStr = " delete from [db_summary].[dbo].tbl_erritem_daily_summary"
sqlStr = sqlStr + " where itemgubun='" + CStr(itemgubun) + "'" + VbCrlf 
sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf 
sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"
dbget.Execute sqlStr



sqlStr = " delete from [db_summary].[dbo].tbl_monthly_logisstock_summary"
sqlStr = sqlStr + " where itemgubun='" + CStr(itemgubun) + "'" + VbCrlf 
sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf 
sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"
dbget.Execute sqlStr



sqlStr = " delete from [db_summary].[dbo].tbl_LAST_monthly_logisstock"
sqlStr = sqlStr + " where itemgubun='" + CStr(itemgubun) + "'" + VbCrlf 
sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf 
sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"
dbget.Execute sqlStr


sqlStr = " delete from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary"
sqlStr = sqlStr + " where itemgubun='" + CStr(itemgubun) + "'" + VbCrlf 
sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf 
sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"
dbget.Execute sqlStr


sqlStr = " delete from [db_summary].[dbo].tbl_current_logisstock_summary"
sqlStr = sqlStr + " where itemgubun='" + CStr(itemgubun) + "'" + VbCrlf 
sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf 
sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"
dbget.Execute sqlStr


response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
response.write "<script>window.close();</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->