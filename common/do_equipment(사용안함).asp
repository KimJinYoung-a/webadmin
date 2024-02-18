<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'#######################################################
'	2009년 01월 19일 한용민 수정
'#######################################################
%>
<%
function makeEquipCode(byval idx,byval equip_gubun,byval part_code,byval buy_date)
	if buy_date="" then buy_date="00000000"

	makeEquipCode = equip_gubun + "-" + replace(Left(CStr(buy_date),10),"-","") + "-" + format00(6,idx)
end function

dim idx
dim equip_code
dim equip_gubun
dim part_code
dim equip_name
dim model_name
dim manufacture_company
dim buy_company_code
dim buy_company_name
dim buy_date
dim buy_cost
dim buy_vat
dim buy_sum
dim equip_no
dim durability_month
dim detail_quality1
dim detail_quality2
dim detail_qualityetc
dim detail_ip
dim etc_str
dim usinguserid

idx       				=	request("idx")
equip_code         		=	request("equip_code")
equip_gubun        		=	request("equip_gubun")
part_code				=	request("part_code")
equip_name         		=	html2db(request("equip_name"))
model_name         		=	html2db(request("model_name"))
manufacture_company		=	html2db(request("manufacture_company"))
buy_company_code   		=	request("buy_company_code")
buy_company_name   		=	html2db(request("buy_company_name"))
buy_date           		=	request("buy_date")
buy_sum            		=	request("buy_sum")
equip_no           		=	request("equip_no")
durability_month   		=	request("durability_month")
detail_quality1    		=	html2db(request("detail_quality1"))
detail_quality2    		=	html2db(request("detail_quality2"))
detail_qualityetc  		=	html2db(request("detail_qualityetc"))
detail_ip          		=	request("detail_ip")
etc_str            		=	html2db(request("etc_str"))
usinguserid        		=	request("usinguserid")

if not IsNumeric(buy_sum) then buy_sum=0
buy_cost = Clng(buy_sum*10/11)
buy_vat  = buy_sum-buy_cost

dim mode

if idx<>"" then
	mode="edit"
else
	mode="add"
end if


dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr, i

if (mode="edit") then

	equip_code = makeEquipCode(idx, equip_gubun, part_code, buy_date)
	''수정
	sqlStr = " update [db_partner].[dbo].tbl_equipment_list" + VbCrlf
	sqlStr = sqlStr + " set equip_code='" + equip_code + "'" + VbCrlf
	sqlStr = sqlStr + " ,equip_gubun='" + equip_gubun + "'" + VbCrlf
	sqlStr = sqlStr + " ,part_code='" + part_code + "'" + VbCrlf
	sqlStr = sqlStr + " ,equip_name='" + equip_name + "'" + VbCrlf
	sqlStr = sqlStr + " ,model_name='" + model_name + "'" + VbCrlf
	sqlStr = sqlStr + " ,manufacture_company='" + manufacture_company + "'" + VbCrlf
	sqlStr = sqlStr + " ,buy_company_code='" + buy_company_code + "'" + VbCrlf
	sqlStr = sqlStr + " ,buy_company_name='" + buy_company_name + "'" + VbCrlf
	if buy_date<>"" then
		sqlStr = sqlStr + " ,buy_date='" + buy_date + "'" + VbCrlf
	end if
	sqlStr = sqlStr + " ,buy_cost=" + CStr(buy_cost) + "" + VbCrlf
	sqlStr = sqlStr + " ,buy_vat=" + CStr(buy_vat) + "" + VbCrlf
	sqlStr = sqlStr + " ,buy_sum=" + CStr(buy_sum) + "" + VbCrlf
	sqlStr = sqlStr + " ,equip_no=1" + VbCrlf
	sqlStr = sqlStr + " ,durability_month=36" + VbCrlf
	sqlStr = sqlStr + " ,detail_quality1='" + detail_quality1 + "'" + VbCrlf
	sqlStr = sqlStr + " ,detail_quality2='" + detail_quality2 + "'" + VbCrlf
	sqlStr = sqlStr + " ,detail_qualityetc='" + detail_qualityetc + "'" + VbCrlf
	sqlStr = sqlStr + " ,detail_ip='" + detail_ip + "'" + VbCrlf
	sqlStr = sqlStr + " ,etc_str='" + etc_str + "'" + VbCrlf
	sqlStr = sqlStr + " ,usinguserid='" + usinguserid + "'" + VbCrlf
	sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
	sqlStr = sqlStr + " ,modiuserid='" + session("ssBctId") + "'" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)
'response.write sqlStr
	dbget.execute sqlStr
else
	sqlStr = " select * from [db_partner].[dbo].tbl_equipment_list where 1=0" + VbCrlf
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("equip_gubun") = equip_gubun
	rsget("part_code") = part_code
	rsget("equip_name") = equip_name
	rsget("model_name") = model_name
	rsget("manufacture_company") = manufacture_company
	rsget("buy_company_code") = buy_company_code
	rsget("buy_company_name") = buy_company_name

	if buy_date<>"" then
		rsget("buy_date") = buy_date
	end if

	if buy_cost<>"" then
		rsget("buy_cost") = buy_cost
		rsget("buy_vat") = buy_cost
		rsget("buy_sum") = buy_sum
	end if

	rsget("equip_no") = 1
	rsget("durability_month") = 36

	rsget("detail_quality1") = detail_quality1
	rsget("detail_quality2") = detail_quality2
	rsget("detail_qualityetc") = detail_qualityetc
	rsget("detail_ip") = detail_ip
	rsget("etc_str") = etc_str
	rsget("usinguserid") = usinguserid
	rsget("reguserid") = session("ssBctId")


	rsget.update
		idx = rsget("idx")
	rsget.close

	equip_code = makeEquipCode(idx, equip_gubun, part_code, buy_date)

	sqlStr = "update [db_partner].[dbo].tbl_equipment_list"
	sqlStr = sqlStr + " set equip_code='" + equip_code + "'"
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	dbget.execute sqlStr

end if
%>
<script language='javascript'>
location.replace('/common/pop_equipmentreg.asp?idx=<%= idx %>');
opener.location.reload();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->