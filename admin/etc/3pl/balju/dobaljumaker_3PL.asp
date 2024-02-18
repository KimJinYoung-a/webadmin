<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 80
%>
<%
'###########################################################
' Description : �������
' History : �̻� ����
'			2018.03.28 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
function FormatStr(n,orgData)
	dim tmp
	if (n-Len(CStr(orgData))) < 0 then
		FormatStr = CStr(orgData)
		Exit Function
	end if

	tmp = String(n-Len(CStr(orgData)), "0") & CStr(orgData)
	FormatStr = tmp
end Function

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode
dim orderserial
dim differencekey, workgroup, songjangdiv
dim tplcompanyid

mode        = request("mode")
orderserial = request("orderserial")
workgroup   = request("workgroup")
tplcompanyid   = request("extSiteName")
songjangdiv = request("songjangdiv")

dummiserial = orderserial
orderserial = split(orderserial,"|")

dim sqlStr,i
dim iid
dim dummiserial
dim obaljudate
dim errcode

if mode="arr" then
	''��ȿ��üũ.
	dummiserial = Mid(dummiserial,2,Len(dummiserial))
	dummiserial = replace(dummiserial,"|","','")
	sqlStr = " select top 1 orderserial from [db_threepl].[dbo].[tbl_tpl_balju_detail]"
	sqlStr = sqlStr + " where orderserial in ('" + dummiserial + "')"

	rsget_TPL.Open sqlStr,dbget_TPL,1
	if Not rsget_TPL.Eof then
		response.write "<script language='javascript'>"
		response.write "alert('" + rsget_TPL("orderserial") + " : �̹� ������õ� �ֹ��Դϴ�. \n\n�̹� ��������� �ֹ����� �ߺ� ������� �� �� �����ϴ�.');"
		response.write "location.replace('" + CStr(refer) + "');"
		response.write "</script>"
		dbget_TPL.close()	:	response.End
	end if
	rsget_TPL.Close

	sqlStr = " select (count(idx) + 1) as differencekey"
	sqlStr = sqlStr + " from [db_threepl].[dbo].[tbl_tpl_balju_master]"
	sqlStr = sqlStr + " where convert(varchar(10),baljudate,21)=convert(varchar(10),getdate(),21)"

	rsget_TPL.Open sqlStr,dbget_TPL,1
		differencekey = rsget_TPL("differencekey")
	rsget_TPL.close


On Error Resume Next
dbget_TPL.beginTrans

If Err.Number = 0 Then
        errcode = "001"


    '' �ù�纰 ������� system���� ����

	'#######������� ������###############
	sqlStr = "insert into [db_threepl].[dbo].[tbl_tpl_balju_master](baljudate,differencekey,workgroup, songjangdiv, tplcompanyid)"
	sqlStr = sqlStr + " values(getdate()," + CStr(differencekey) + ",'" + workgroup + "','" + songjangdiv + "', '" + CStr(tplcompanyid) + "')"
	rsget_TPL.Open sqlStr,dbget_TPL,1

	sqlStr = "select top 1 idx, convert(varchar(19),baljudate,21) as baljudate from [db_threepl].[dbo].[tbl_tpl_balju_master] order by idx desc"
	rsget_TPL.Open sqlStr,dbget_TPL,1
	    iid = rsget_TPL("idx")
	    obaljudate = rsget_TPL("baljudate")
	rsget_TPL.Close

	'#######������� ������###############
	for i=0 to Ubound(orderserial)
		if orderserial(i)<>"" then
			sqlStr = "insert into [db_threepl].[dbo].[tbl_tpl_balju_detail](baljuid,orderserial)"
			sqlStr = sqlStr + " values(" + CStr(iid) + ","
			sqlStr = sqlStr + " '" + orderserial(i) + "'"
			sqlStr = sqlStr + " )"
			rsget_TPL.Open sqlStr,dbget_TPL,1
		end if
	next

    ''** [db_threepl].[dbo].[tbl_tpl_balju_detail].baljusongjangno is NULL �ΰ�� ��ü������� �ν� (Logics ���� �ý���)
    ''�ٹ����� ����� ��� �����ȣ�� not null ������ �Է�..

    sqlStr = "update [db_threepl].[dbo].[tbl_tpl_balju_detail]" + VbCrlf
	sqlStr = sqlStr + " set baljusongjangno=''"
	sqlStr = sqlStr + " where baljuid=" + CStr(iid)
	sqlStr = sqlStr + " and baljusongjangno is NULL "
	rsget_TPL.Open sqlStr,dbget_TPL,1
end if


If Err.Number = 0 Then
        errcode = "002"

	''' �ֹ����� ����������� �Է�
	sqlStr = "update [db_threepl].[dbo].[tbl_tpl_orderMaster]"
	sqlStr = sqlStr + " set baljudate='" + CStr(obaljudate) + "'"
	sqlStr = sqlStr + " where ipkumdiv > 4 "
	sqlStr = sqlStr + " and baljudate is NULL"
	sqlStr = sqlStr + " and orderserial in "
	sqlStr = sqlStr + " (select orderserial from [db_threepl].[dbo].[tbl_tpl_balju_detail]"
	sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
	sqlStr = sqlStr + " )"
	rsget_TPL.Open sqlStr,dbget_TPL,1


	'#######�ֹ� �뺸 ���� ############### (��������� ����)
	sqlStr = "update [db_threepl].[dbo].[tbl_tpl_orderMaster]"
	sqlStr = sqlStr + " set ipkumdiv='5'"
	sqlStr = sqlStr + " ,baljudate='" + CStr(obaljudate) + "'"
	sqlStr = sqlStr + " where ipkumdiv=4"
	sqlStr = sqlStr + " and orderserial in "
	sqlStr = sqlStr + " (select orderserial from [db_threepl].[dbo].[tbl_tpl_balju_detail]"
	sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
	sqlStr = sqlStr + " )"
	rsget_TPL.Open sqlStr,dbget_TPL,1
end if


If Err.Number = 0 Then
        errcode = "004"
	'###### Order Detail �ٹ����� ��� ��������� ���� ############
	sqlStr = "update [db_threepl].[dbo].[tbl_tpl_orderDetail]"
	sqlStr = sqlStr + " set currstate='2'"
	sqlStr = sqlStr + " ,songjangdiv=" + CStr(songjangdiv)
	sqlStr = sqlStr + " where orderserial in "
	sqlStr = sqlStr + " (select orderserial from [db_threepl].[dbo].[tbl_tpl_balju_detail]"
	sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
	sqlStr = sqlStr + " )"
	sqlStr = sqlStr + " and [db_threepl].[dbo].[tbl_tpl_orderDetail].itemid<>0"
	rsget_TPL.Open sqlStr,dbget_TPL,1
end if


If Err.Number = 0 Then
        errcode = "013"
    '' ��� ������Ʈ
    ''sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_balju " & iid
    ''dbget_TPL.execute sqlStr
end if


If Err.Number = 0 Then
    dbget_TPL.CommitTrans
Else
    dbget_TPL.RollBackTrans
    response.write "<script>alert('����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n������ ���� ��� (�����ڵ� : " + CStr(errcode) + ")');</script>"
    response.write "<script>history.back()</script>"
    dbget_TPL.close()	:	dbget.close()	:	response.End
End If
on error Goto 0

end if

%>

<script language="javascript">
alert('������ü��� ���� �Ǿ����ϴ�.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_TPLclose.asp" -->