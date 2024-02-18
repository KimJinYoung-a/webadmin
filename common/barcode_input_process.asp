<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description : ���ڵ� ã��
' History : 2017.04.10 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<%
Const DELETE_KEY = "000000000100"
dim IsDeleteBarCode

dim itemgubun, itemid, itemoption, publicbarcode

itemgubun 	= trim(requestCheckVar(request("itemgubun"),2))
itemid		= trim(requestCheckVar(request("itemid"),10))
itemoption	= trim(requestCheckVar(request("itemoption"),4))
publicbarcode	= trim(requestCheckVar(request("publicbarcode"),20))


IsDeleteBarCode = (DELETE_KEY=publicbarcode)


dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr, barcodeAlreadyExixts
dim existsitemgubun, existsitemid, existsitemoption, existsitemname, existeitemoptionname

dim stockitemexists
''��ϵǾ��ִ� ���ڵ����� üũ
barcodeAlreadyExixts = false

sqlStr = " select top 1 * from [db_item].[dbo].tbl_item_option_stock" + VbCrlf
sqlStr = sqlStr + " where barcode='" + publicbarcode + "'" + VbCrlf
sqlStr = sqlStr + " and not ("
sqlStr = sqlStr + " 	itemgubun='" + itemgubun + "'" + VbCrlf
sqlStr = sqlStr + " 	and itemid=" + CStr(itemid) + "" + VbCrlf
sqlStr = sqlStr + " 	and itemoption='" + CStr(itemoption) + "'" + VbCrlf
sqlStr = sqlStr + " ) "
rsget.Open sqlStr,dbget,1
if not rsget.Eof then
	barcodeAlreadyExixts = true
end if
rsget.close



if barcodeAlreadyExixts=false then
	sqlStr = " select top 1 shopitemid from [db_shop].[dbo].tbl_shop_item" + VbCrlf
	sqlStr = sqlStr + " where extbarcode='" + publicbarcode + "'" + VbCrlf
	sqlStr = sqlStr + " and not (" + VbCrlf
	sqlStr = sqlStr + " itemgubun='" + itemgubun + "'" + VbCrlf
	sqlStr = sqlStr + " and shopitemid=" + CStr(itemid) + "" + VbCrlf
	sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'" + VbCrlf
	sqlStr = sqlStr + " )" + VbCrlf
	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		barcodeAlreadyExixts = true
	end if
	rsget.close
end if


if (IsDeleteBarCode) then
    publicbarcode = ""
end if
response.write barcodeAlreadyExixts
if barcodeAlreadyExixts=false then
	''������ڵ� �Է�.
	''1.�������λ�ǰ - �������� �����迡 ���� �ִ� ��ǰ�� ������Ʈ.
	sqlStr = " update [db_shop].[dbo].tbl_shop_item" + VbCrlf
	sqlStr = sqlStr + " set extbarcode='" + publicbarcode + "'," + VbCrlf
	sqlStr = sqlStr + " updt=getdate()" + VbCrlf
	''sqlStr = sqlStr + " ,franupdt=getdate()," + VbCrlf
	''sqlStr = sqlStr + " cmsupdt=getdate()" + VbCrlf
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'" + VbCrlf
	sqlStr = sqlStr + " and shopitemid=" + CStr(itemid) + "" + VbCrlf
	sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'" + VbCrlf

	rsget.Open sqlStr,dbget,1


	sqlStr = " select top 1 * from [db_item].[dbo].tbl_item_option_stock" + VbCrlf
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'" + VbCrlf
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
	sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'" + VbCrlf
	rsget.Open sqlStr,dbget,1
	stockitemexists = (not rsget.Eof)
	rsget.close

	if (stockitemexists) then
		sqlStr = " update [db_item].[dbo].tbl_item_option_stock" + VbCrlf
		sqlStr = sqlStr + " set barcode='" + publicbarcode + "'" + VbCrlf
		sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'" + VbCrlf
		sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
		sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'" + VbCrlf

		rsget.Open sqlStr,dbget,1
	else
	    if Not (IsDeleteBarCode) then
    		sqlStr = " insert into [db_item].[dbo].tbl_item_option_stock" + VbCrlf
    		sqlStr = sqlStr + " (itemgubun,itemid,itemoption,barcode)" + VbCrlf
    		sqlStr = sqlStr + " values("
    		sqlStr = sqlStr + " '" + itemgubun + "'," + VbCrlf
    		sqlStr = sqlStr + " " + CStr(itemid) + "," + VbCrlf
    		sqlStr = sqlStr + " '" + itemoption + "'," + VbCrlf
    		sqlStr = sqlStr + " '" + publicbarcode + "'" + VbCrlf
    		sqlStr = sqlStr + " )" + VbCrlf
    		rsget.Open sqlStr,dbget,1
    	end if
	end if
end if

if barcodeAlreadyExixts=true then
	sqlStr = " select top 1 s.itemgubun,s.itemid,s.itemoption, i.itemname, IsNULL(v.opt2name,'') as itemoptionname "
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option_stock s"
	sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on s.itemgubun='10' and s.itemid=i.itemid"
	sqlStr = sqlStr + " left join [db_item].[dbo].vw_itemoption v on s.itemgubun='10' and s.itemid=v.itemid and s.itemoption=v.itemoption"
	sqlStr = sqlStr + " where barcode='" + publicbarcode + "'"

	rsget.Open sqlStr,dbget,1
		if not rsget.Eof then
			existsitemgubun = rsget("itemgubun")
			existsitemid	= rsget("itemid")
			existsitemoption = rsget("itemoption")
			existsitemname	= replace(db2html(rsget("itemname")),"'","")
			existeitemoptionname = replace(db2html(rsget("itemoptionname")),"'","")
		end if
	rsget.close
end if
%>
<% if (IsDeleteBarCode) then %>
    <script type='text/javascript'>
	    alert('���� �Ǿ����ϴ�.');
	    history.back();
    </script>
<% else %>
    <% if (barcodeAlreadyExixts) then %>
	    <script type='text/javascript'>
		    alert('�̹� ������� ���ڵ�(<%= publicbarcode %>) �Դϴ�. ��ǰ��ȣ : (<%= existsitemid %>) <%= existsitemname %>,<%= existeitemoptionname %>');
		    history.back();
	    </script>
    <% else %>
	    <script type='text/javascript'>
		    alert('��� �Ǿ����ϴ�.');
		    history.back();
	    </script>
    <% end if %>
<% end if %>

<!-- #include virtual="/lib/db/dbclose.asp" -->
