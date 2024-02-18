<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim designer
dim isusing, isextusing, streetusing, extstreetusing, specialbrand, partnerusing, isoffusing

designer		= requestCheckvar(request.form("designer"),40)
isusing			= requestCheckvar(request.form("isusing"),10)
isextusing		= requestCheckvar(request.form("isextusing"),10)
streetusing		= requestCheckvar(request.form("streetusing"),10)
extstreetusing	= requestCheckvar(request.form("extstreetusing"),10)
specialbrand	= requestCheckvar(request.form("specialbrand"),10)
partnerusing	= requestCheckvar(request.form("partnerusing"),10)
isoffusing      = requestCheckvar(request.form("isoffusing"),10)

dim sqlStr

dim mode, adminid
dim totalitemcount_m, totalitemcount_w
dim defaultdeliveryType, defaultFreeBeasongLimit, defaultDeliverPay, orgdefaultdeliveryType
adminid = session("ssBctID")
mode						= requestCheckvar(request.form("mode"),40)
defaultdeliveryType			= requestCheckvar(request.form("defaultdeliveryType"),30)
defaultFreeBeasongLimit		= requestCheckvar(request.form("defaultFreeBeasongLimit"),30)
defaultDeliverPay			= requestCheckvar(request.form("defaultDeliverPay"),30)
orgdefaultdeliveryType			= requestCheckvar(request.form("orgdefaultdeliveryType"),30)

dim pisusing

if (mode = "policy") then
	'defaultdeliveryType
	'defaultFreeBeasongLimit
	'defaultDeliverPay

	sqlStr = " select "
	sqlStr = sqlStr + " sum(case when mwdiv='M' then 1 else 0 end) as totalitemcount_m, "
	sqlStr = sqlStr + " sum(case when mwdiv='W' then 1 else 0 end) as totalitemcount_w "
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item"
	sqlStr = sqlStr + " where makerid='" + designer + "'"
	sqlStr = sqlStr + "		and isusing = 'Y' " '2014-12-11 ������ �߰� (������ ��ǰ�� ����)
	rsget.Open sqlStr,dbget,1

	totalitemcount_m = rsget("totalitemcount_m")
	totalitemcount_w = rsget("totalitemcount_w")

	if IsNULL(totalitemcount_m) then totalitemcount_m = 0 end if
	if IsNULL(totalitemcount_w) then totalitemcount_w = 0 end if

	rsget.Close

    ' 2019.02.21 �ѿ�� ����(�̹��� �̻�� ����)
	if ((totalitemcount_m <> 0) or (totalitemcount_w <> 0)) and (defaultdeliveryType<>"") and (orgdefaultdeliveryType="") then
		response.write "<script type='text/javascript'>alert('���� �Ǵ� ��Ź ��ǰ�� �ִ°�� �⺻��å�� �����մϴ�.');</script>"
		response.write "<script type='text/javascript'>location.replace('" & refer & "');</script>"
		response.end
	else
		'// ���� ��ۺ� ��å ����
		sqlStr = " insert into [db_user].[dbo].[tbl_user_c_defaultdelivery_log](userid, defaultFreeBeasongLimit, defaultDeliverPay, defaultDeliveryType) "
		sqlStr = sqlStr + " select top 1 userid, defaultFreeBeasongLimit, defaultDeliverPay, defaultDeliveryType "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_user].[dbo].[tbl_user_c] "
		sqlStr = sqlStr + " where userid = '" & designer & "' "
		dbget.execute sqlStr

		if (defaultdeliveryType <> "9") then
			defaultFreeBeasongLimit = 0
			''defaultDeliverPay = 0   ''�⺻ ��ۺ� ��밡�� (������ ����)
		end if

		sqlStr = "update [db_user].[dbo].tbl_user_c" + VbCrlf

		if (defaultdeliveryType = "") then
			sqlStr = sqlStr + " set defaultdeliveryType=null " + VbCrlf
		else
			sqlStr = sqlStr + " set defaultdeliveryType='" + CStr(defaultdeliveryType)  + "' " + VbCrlf
		end if

		sqlStr = sqlStr + " ,defaultFreeBeasongLimit=" + CStr(defaultFreeBeasongLimit)  + " " + VbCrlf
		sqlStr = sqlStr + " ,defaultDeliverPay=" + CStr(defaultDeliverPay)  + " " + VbCrlf
		sqlStr = sqlStr + " where userid='" + designer + "'" + VbCrlf

		dbget.execute sqlStr
	end if

elseif (mode="using") then
    sqlStr = "select isusing from [db_partner].[dbo].tbl_partner" + VbCrlf
    sqlStr = sqlStr + " where id='" + designer + "'" + VbCrlf
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        pisusing = rsget("isusing")
    end if
    rsget.close

    if (pisusing<>partnerusing) then
    	sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
    	sqlStr = sqlStr + " set lastInfoChgDT=getdate(), isusing='" + partnerusing + "'" + VbCrlf
    	if (partnerusing="N") then
    	    sqlStr = sqlStr + " ,lastExpireDT=isNULL(lastExpireDT,getdate())"
        end if
    	sqlStr = sqlStr + " where id='" + designer + "'" + VbCrlf

    	dbget.execute sqlStr

    	sqlStr = "Insert into db_log.dbo.tbl_partner_login_log" + VbCrlf
		sqlStr = sqlStr&" (userid,refip,regdate,loginSuccess,UsbTokenSn)" + VbCrlf
		sqlStr = sqlStr&" values('"&designer&"','"&request.ServerVariables("REMOTE_ADDR")&"',getdate(),'"&CHKIIF(partnerusing="N","X","A")&"',convert(varchar(24),'"& adminid &"'))" + VbCrlf

		dbget.execute sqlStr

		' ����, �귣�� ����α�
		fnChkauthlog "", designer, "11", "SCM �귣�� ���Ѻ���:��뿩�� " & pisusing & "->" & partnerusing, adminid
    end if

	sqlStr = "update [db_user].[dbo].tbl_user_c" + VbCrlf
	sqlStr = sqlStr + " set isusing='" + (isusing)  + "'" + VbCrlf
	sqlStr = sqlStr + " ,isextusing='" + (isextusing)  + "'" + VbCrlf
	sqlStr = sqlStr + " ,isoffusing='" + (isoffusing)  + "'" + VbCrlf
	sqlStr = sqlStr + " ,streetusing='" + (streetusing)  + "'" + VbCrlf
	sqlStr = sqlStr + " ,extstreetusing='" + (extstreetusing)  + "'" + VbCrlf
	sqlStr = sqlStr + " ,specialbrand='" + (specialbrand)  + "'" + VbCrlf
	sqlStr = sqlStr + " where userid='" + designer + "'" + VbCrlf

	dbget.execute sqlStr
else
    rw "["&mode&"] ������"
end if

%>


<script type='text/javascript'>alert('����Ǿ����ϴ�.');</script>
<script type='text/javascript'>location.replace('<%= refer %>');</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
