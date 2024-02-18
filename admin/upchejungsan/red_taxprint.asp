<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim designer, tax_no
designer = request("makerid")
tax_no = request("tax_no")

dim sqlstr, TaxResultCount
dim biz_no
sqlstr = "select top 10 * from  [db_jungsan].[dbo].tbl_tax_history_master"
sqlstr = sqlstr + " where tax_no='" + tax_no + "'"
sqlstr = sqlstr + " and makerid='" + designer + "'"
sqlstr = sqlstr + " and resultmsg='OK'"
sqlstr = sqlstr + " and deleteyn='N'"

'response.write sqlstr
rsget.Open sqlStr,dbget,1
TaxResultCount = rsget.RecordCount
if  not rsget.EOF  then
	biz_no = rsget("biz_no")
end if
rsget.close

if (TaxResultCount<>1) and ((Left(tax_no,2)<>"TX") and  (Left(tax_no,2)<>"FX")) then
	response.write "<script>alert('올바른 계산서번호가 아닐수 있습니다. 관리자 문의요망" + CStr(TaxResultCount) + "');</script>"
else
	''response.write "<script>location.replace('http://www.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no=" + tax_no + "&cur_biz_no=" + biz_no + "');</script>"
	'' cur_biz_no 에 따라 공급자/공급받는자 세금계산서가 나옴.
	if (Left(tax_no,2)="TX") or (Left(tax_no,2)="FX") then
        if (application("Svr_Info") = "Dev") then
    	    ''TEST URL
    	    response.write "<script>location.replace('http://www.bill36524.com/popupBillTax.jsp?NO_TAX="+ tax_no +"&NO_BIZ_NO="+ biz_no + "');</script>"
    	else
    	    ''REAL URL
    	    response.write "<script>location.replace('http://www.bill36524.com/popupBillTax.jsp?NO_TAX="+ tax_no +"&NO_BIZ_NO="+ biz_no + "');</script>"
        end if
    else
	    response.write "<script>location.replace('http://www.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no=" + tax_no + "&cur_biz_no=" + "2118700620" + "&s_biz_no=" + biz_no + "&b_biz_no=2118700620');</script>"
	end if
end if
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->