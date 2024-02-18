<%@ language=vbscript %>
<% option explicit %>

<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 세금계산서 발행후 저장
' History : 서동석 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/jungsan/electaxcls.asp" -->
<%
dim IsBill36524 : IsBill36524 = (request("billsite")="B")

dim oelectaxitem, oelectax



''기발행 세금계산서인지 체크
dim sqlstr,IsExistsTax
sqlstr = "select count(*) as cnt from  [db_jungsan].[dbo].tbl_tax_history_master"
sqlstr = sqlstr + " where jungsanid=" + CStr(request("jungsanid"))
sqlstr = sqlstr + " and jungsangubun='" + CStr(request("jungsangubun")) + "'"
sqlstr = sqlstr + " and makerid='" + CStr(request("makerid")) + "'"
sqlstr = sqlstr + " and resultmsg='OK'" '''기존
sqlstr = sqlstr + " and deleteyn='N'"
sqlstr = sqlstr + " and jgubun='" + CStr(request("jgubun")) + "'"
rsget.Open sqlStr,dbget,1
	IsExistsTax = rsget("cnt")>0
rsget.close

if IsExistsTax then
    if (request("autotype")="V2") then
        response.write "<script>"&vbCRLF
        response.write "parent.opener.addResultLog('"&request("jungsanid")&"','기발행계산서');"&vbCRLF
        response.write "parent.opener.fnNextEvalProc();"&vbCRLF
        response.write "</script>"
    else
	    response.write "<script>alert('이미 발행된 세금계산서 또는 통신중 오류건 입니다.');</script>"
	    response.write "<script>history.back();<script>"
    end if 
	dbget.close()	:	response.End
end if


set oelectaxitem = new CElecTaxRegItem
oelectaxitem.Fjungsanid  = request("jungsanid")
oelectaxitem.Fjungsangubun = request("jungsangubun")
oelectaxitem.Fmakerid     = request("makerid")
oelectaxitem.Fjungsanname= request("jungsanname")

oelectaxitem.Fbiz_no = request("biz_no")
oelectaxitem.Fcorp_nm = request("corp_nm")
oelectaxitem.Fceo_nm = request("ceo_nm")
oelectaxitem.Fbiz_status = request("biz_status")
oelectaxitem.Fbiz_type = request("biz_type")
oelectaxitem.Faddr = request("addr")
oelectaxitem.Fdam_nm = request("dam_nm")
oelectaxitem.Femail = request("email")
oelectaxitem.Fhp_no1 = request("hp_no1")
oelectaxitem.Fhp_no2 = request("hp_no2")
oelectaxitem.Fhp_no3 = request("hp_no3")

oelectaxitem.Fwrite_date = request("write_date")
oelectaxitem.Fsb_type = request("sb_type")
oelectaxitem.Ftax_type   = request("tax_type")
oelectaxitem.Fbill_type  = request("bill_type")
oelectaxitem.Fpc_gbn     = request("pc_gbn")
oelectaxitem.Fvol_no     = ""
oelectaxitem.Fissue_no   = ""
oelectaxitem.Fserial_no  = ""
oelectaxitem.Fremark     = ""

oelectaxitem.Fitem_count  = request("item_count")
oelectaxitem.Fitem_nm     = request("item_nm")
oelectaxitem.Fitem_std    = ""
oelectaxitem.Fitem_qty    = request("item_qty")
oelectaxitem.Fitem_price  = request("item_price")
oelectaxitem.Fapprove_type= "01"                     ' 01 공급받는자승인  11 공급자가 승인
oelectaxitem.Fitem_amt    = request("item_amt")
oelectaxitem.Fitem_vat	  = request("item_vat")
oelectaxitem.Fitem_remark = request("item_remark")

oelectaxitem.Fcur_u_user_no = request("cur_u_user_no")
oelectaxitem.Fcur_dam_nm = request("cur_dam_nm")
oelectaxitem.Fcur_email  = request("cur_email")
oelectaxitem.Fcur_hp_no1 = request("cur_hp_no1")
oelectaxitem.Fcur_hp_no2 = request("cur_hp_no2")
oelectaxitem.Fcur_hp_no3 = request("cur_hp_no3")

''외상미수 추가
oelectaxitem.Fcredit_amt = request("credit_amt")

oelectaxitem.FRectBillSite = request("billsite")
oelectaxitem.FJGubun = request("jgubun")

set oelectax = new CElecTaxReg
set oelectax.FRectOneRegitem = oelectaxitem

on Error resume Next
oelectax.SavePreData
If Err then
	response.write "DB작업중 오류 - " + Err.Description + " 관리자 문의 요망"
	dbget.close()	:	response.End
end if
on Error goto 0

dim psavedIdx : psavedIdx=0
if (IsBill36524) then
   ''bill36524 flexApi로 발행 하므로 SKIP
   ''oelectax.ExecReverseTaxBill36524
    psavedIdx = oelectax.FRectOneRegitem.Fidx
else
    oelectax.ExecDTIXmlDom
end if

dim IsSuccess, itax_no, ierrmsg, ibizno
IsSuccess = (oelectax.Fresultmsg="OK")
itax_no = oelectax.Ftax_no
ierrmsg = oelectax.Fresultmsg
ibizno  = oelectaxitem.Fbiz_no
'if (Not IsSuccess) then
'	response.write "통신중 오류 관리자 문의 요망 : ERR No(" + oelectax.Ftax_no + ") " + oelectax.Fresultmsg
'	dbget.close()	:	response.End
'end if


set oelectaxitem = nothing
set oelectax = nothing


%>
 <% if (IsBill36524) then %>
        <script language='javascript'>
            parent.billTaxEvalFlexApi('<%= psavedIdx %>');

        </script>
<% else %>
    <% if IsSuccess then %>
            <script language='javascript'>
            alert('세금계산서가 발행 되었습니다.');
            opener.location.reload();
            window.close();
            </script>
    <% else %>
        <%
        response.redirect "taxregresult.asp?itax_no=" + CStr(itax_no) + "&ierrmsg=" + (ierrmsg)
        %>
    <% end if %>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->