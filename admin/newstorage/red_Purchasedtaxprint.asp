<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���ϰ�(��36524) ���ݰ�꼭 ����
' History : 2022.08.10 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/PurchasedProductCls.asp"-->
<%
dim tax_no, groupcode, sqlstr, TaxResultCount, biz_no
    groupcode = request("groupcode")
    tax_no = request("tax_no")

sqlstr = "select top 10 *"
sqlstr = sqlstr & " from db_jungsan.dbo.tbl_pp_product_sheet_tax_history_master t with (nolock)"
sqlstr = sqlstr & " where t.tax_no='" + tax_no + "'"
sqlstr = sqlstr & " and t.resultmsg='OK'"
sqlstr = sqlstr & " and t.deleteyn='N'"
sqlstr = sqlstr & " and t.groupcode='"& groupcode &"'"

'response.write sqlstr & "<Br>"
rsget.CursorLocation = adUseClient
rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
TaxResultCount = rsget.RecordCount
if not rsget.EOF  then
	biz_no = rsget("biz_no")
end if
rsget.close

if TaxResultCount<>1 then
	response.write "<script type='text/javascript'>alert('�ùٸ� ��꼭��ȣ�� �ƴҼ� �ֽ��ϴ�. ������ ���ǿ��" + CStr(TaxResultCount) + "');</script>"
else
    if Left(tax_no,2)="TX" then
        if (application("Svr_Info") = "Dev") then
    	    ''TEST URL
    	    response.write "<script type='text/javascript'>location.replace('http://www.bill36524.com/popupBillTax.jsp?NO_TAX="+ tax_no +"&NO_BIZ_NO="+ biz_no + "');</script>"
    	else
    	    ''REAL URL
    	    response.write "<script type='text/javascript'>location.replace('http://www.bill36524.com/popupBillTax.jsp?NO_TAX="+ tax_no +"&NO_BIZ_NO="+ biz_no + "');</script>"
        end if
    else
     ''response.write "<script type='text/javascript'>location.replace('http://www.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no=" + tax_no + "&cur_biz_no=" + biz_no + "');</script>"
	    response.write "<script type='text/javascript'>location.replace('http://www.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no=" + tax_no + "&cur_biz_no=" + biz_no + "&s_biz_no=" + biz_no + "&b_biz_no=2118700620');</script>"

    end if
end if
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->