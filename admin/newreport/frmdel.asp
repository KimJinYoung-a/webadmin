<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<%
'��������
dim idx, fequip_code, fequip_gubun, fequip_name ,fmodel_name, fmanufacture_company ,fbuy_company_code ,fbuy_company_name ,fbuy_date
dim fbuy_cost, fbuy_vat ,fbuy_sum, fequip_no, fdurability_month, fdetail_quality1 , fdetail_quality2, fdetail_qualityetc
dim fdetail_ip, fetc_str, fusinguserid, fpart_code, fregdate, flastupdate, freguserid, fmodiuserid , fuser_id , fdate

idx = request("idx")			'�۹�ȣ
fuser_id = request("ssBctId")		'���� ���� �α������ξƵ�
fdate = now()

dim sql 			'�������˻��� ��������

<!-- ������ ��� idx ������ ���� ���� �´�.����-->
sql = "select * from [db_partner].[dbo].tbl_equipment_list where idx=" +idx
rsget.open sql,dbget,1
	fequip_code = rsget("equip_code")			'����ڵ�
	fequip_gubun = rsget("equip_gubun")			'��񱸺�
	fequip_name = rsget("equip_name")			'����̸�
	fmanufacture_company = rsget("manufacture_company")
	fbuy_company_code = rsget("buy_company_code")
	fbuy_company_name  = rsget("buy_company_name")
	fbuy_date = rsget("buy_date")				'�����ѳ�¥
	fbuy_cost = rsget("buy_cost")				'���Ժ��
	fbuy_vat = rsget("buy_vat")
	fbuy_sum = rsget("buy_sum")				'�����Ѱ������� �ѱݾ�
	fequip_no = rsget("equip_no")
	fdurability_month = rsget("durability_month")		'���� 36����
	fdetail_quality1 = rsget("detail_quality1")
	fdetail_quality2 = rsget("detail_quality2")
	fdetail_qualityetc = rsget("detail_qualityetc")
	fdetail_ip = rsget("detail_ip")				'���ip
	fetc_str = rsget("etc_str")
	fusinguserid = rsget("usinguserid")			'�����id
	fpart_code = rsget("part_code")
	fregdate = rsget("regdate")				'������
	flastupdate = rsget("lastupdate")			'�����������ѳ�¥
	freguserid = rsget("reguserid")				'������ѻ��id
	fmodiuserid = rsget("modiuserid")			'��񸦸�����������id
rsget.close
<!-- ������ ��� idx ������ ���� ���� �´�.��-->

<!-- ������ ��� �α� ���̺� ���� ����-->
dim sql1		'��������
sql1 = "INSERT INTO [db_partner].[dbo].tbl_equipment_log(equip_code,equip_gubun,equip_name,model_name,manufacture_company,buy_company_code,buy_company_name,buy_date,buy_cost,buy_vat,buy_sum,equip_no,durability_month,detail_quality1,detail_quality2,detail_qualityetc,detail_ip,etc_str,usinguserid,part_code,regdate,lastupdate,reguserid,modiuserid,del_id,del_date) VALUES" 
sql1 = sql1 & "('" & fequip_code & "'"
sql1 = sql1 & ",'" & fequip_gubun & "'"
sql1 = sql1 & ",'" & fequip_name & "'"
sql1 = sql1 & ",'" & fmodel_name & "'"
sql1 = sql1 & ",'" & fmanufacture_company & "'"
sql1 = sql1 & ",'" & fbuy_company_code & "'"
sql1 = sql1 & ",'" & fbuy_company_name & "'"
sql1 = sql1 & ",'" & fbuy_date & "'"
sql1 = sql1 & "," & fbuy_cost & ""
sql1 = sql1 & "," & fbuy_vat & ""
sql1 = sql1 & "," & fbuy_sum & ""
sql1 = sql1 & ",'" & fequip_no & "'"
sql1 = sql1 & ",'" & fdurability_month & "'"
sql1 = sql1 & ",'" & fdetail_quality1 & "'"
sql1 = sql1 & ",'" & fdetail_quality2 & "'"
sql1 = sql1 & ",'" & fdetail_qualityetc & "'"
sql1 = sql1 & ",'" & fdetail_ip & "'"
sql1 = sql1 & ",'" & fetc_str & "'"
sql1 = sql1 & ",'" & fusinguserid & "'"
sql1 = sql1 & ",'" & fpart_code & "'"
sql1 = sql1 & ",'" & fregdate & "'"
sql1 = sql1 & ",'" & flastupdate & "'"
sql1 = sql1 & ",'" & freguserid & "'"
sql1 = sql1 & ",'" & fmodiuserid & "'"
sql1 = sql1 & ",'" & fuser_id & "'"
sql1 = sql1 & ",'" & fdate & "')"
dbget.execute sql1
<!-- ������ ��� �α� ���̺� ���� ��-->

<!-- ������̺��� ��� ���� ����-->
dim sqlStr		'��������
sqlStr = " delete from [db_partner].[dbo].tbl_equipment_list" + VBCrlf
sqlStr = sqlStr + " where idx=" & idx
rsget.Open sqlStr,dbget,1
<!-- ������̺��� ��� ���� ��-->

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>

<script language="javascript">
alert('ó�� �Ǿ����ϴ�.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->





