<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim designer, tax_no, groupid
designer = session("ssBctID")
tax_no = request("tax_no")
groupid = getPartnerId2GroupID(designer)

dim sqlstr, TaxResultCount
dim biz_no
sqlstr = "select top 10 * from  [db_jungsan].[dbo].tbl_tax_history_master t"
sqlstr = sqlstr + "     Join db_partner.dbo.tbl_partner p "
sqlstr = sqlstr + "     on p.id=t.makerid and p.groupid='"&groupid&"'"
sqlstr = sqlstr + " where t.tax_no='" + tax_no + "'"
sqlstr = sqlstr + " and t.resultmsg='OK'"
sqlstr = sqlstr + " and t.deleteyn='N'"

'response.write sqlstr
rsget.Open sqlStr,dbget,1
TaxResultCount = rsget.RecordCount
if  not rsget.EOF  then
	biz_no = rsget("biz_no")
end if
rsget.close

if TaxResultCount<>1 then
	response.write "<script>alert('올바른 계산서번호가 아닐수 있습니다. 관리자 문의요망" + CStr(TaxResultCount) + "');</script>"
else
    if Left(tax_no,2)="TX" then
        if (application("Svr_Info") = "Dev") then
    	    ''TEST URL
    	    response.write "<script>location.replace('http://www.bill36524.com/popupBillTax.jsp?NO_TAX="+ tax_no +"&NO_BIZ_NO="+ biz_no + "');</script>"
    	else
    	    ''REAL URL
    	    response.write "<script>location.replace('http://www.bill36524.com/popupBillTax.jsp?NO_TAX="+ tax_no +"&NO_BIZ_NO="+ biz_no + "');</script>"
        end if
    else
     ''response.write "<script>location.replace('http://www.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no=" + tax_no + "&cur_biz_no=" + biz_no + "');</script>"
	    response.write "<script>location.replace('http://www.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no=" + tax_no + "&cur_biz_no=" + biz_no + "&s_biz_no=" + biz_no + "&b_biz_no=2118700620');</script>"

    end if
end if
%>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->