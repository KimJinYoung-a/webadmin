<%@ language=vbscript %>
<% option explicit %>
<%Response.Addheader "P3P","policyref='http://www.10x10.co.kr/w3c/p3p.xml', CP='NOI DSP LAW NID PSA ADM OUR IND NAV COM'"%>
<%
'####################################################
' Description : 세금계산서
' History : 서동석 생성
'			2017.04.12 한용민 수정(보안관련처리)
'####################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/electaxcls.asp" -->
<%
dim oelectaxitem, oelectax

''기발행 세금계산서인지 체크
dim sqlstr,IsExistsTax
sqlstr = "select count(*) as cnt from  [db_jungsan].[dbo].tbl_tax_history_master"
sqlstr = sqlstr + " where jungsanid=" + CStr(requestCheckVar(request("jungsanid"),10))
sqlstr = sqlstr + " and jungsangubun='" + CStr(requestCheckVar(request("jungsangubun"),16)) + "'"
sqlstr = sqlstr + " and makerid='" + CStr(requestCheckVar(request("makerid"),32)) + "'"
sqlstr = sqlstr + " and resultmsg='OK'"
sqlstr = sqlstr + " and deleteyn='N'"

rsget.Open sqlStr,dbget,1
	IsExistsTax = rsget("cnt")>0
rsget.close

if IsExistsTax then
	response.write "<script type='text/javascript'>alert('이미 발행된 세금계산서 입니다.\n 재발행 하시려면 취소후 다시 발행하셔야 합니다. \n- 관리자 문의 요망');</script>"
	response.write "<script type='text/javascript'>history.back();<script type='text/javascript'>"
	dbget.close()	:	response.End
end if


set oelectaxitem = new CElecTaxRegItem
oelectaxitem.Fjungsanid  = requestCheckVar(request("jungsanid"),10)
oelectaxitem.Fjungsangubun = requestCheckVar(request("jungsangubun"),16)
oelectaxitem.Fmakerid     = requestCheckVar(request("makerid"),32)
oelectaxitem.Fjungsanname= requestCheckVar(request("jungsanname"),128)

oelectaxitem.Fbiz_no = requestCheckVar(request("biz_no"),16)
oelectaxitem.Fcorp_nm = requestCheckVar(request("corp_nm"),128)
oelectaxitem.Fceo_nm = requestCheckVar(request("ceo_nm"),64)
oelectaxitem.Fbiz_status = requestCheckVar(request("biz_status"),64)
oelectaxitem.Fbiz_type = requestCheckVar(request("biz_type"),64)
oelectaxitem.Faddr = requestCheckVar(request("addr"),255)
oelectaxitem.Fdam_nm = requestCheckVar(request("dam_nm"),64)
oelectaxitem.Femail = requestCheckVar(request("email"),128)
oelectaxitem.Fhp_no1 = requestCheckVar(request("hp_no1"),4)
oelectaxitem.Fhp_no2 = requestCheckVar(request("hp_no2"),4)
oelectaxitem.Fhp_no3 = requestCheckVar(request("hp_no3"),4)

oelectaxitem.Fwrite_date = requestCheckVar(request("write_date"),30)
oelectaxitem.Fsb_type    = requestCheckVar(request("sb_type"),2)
oelectaxitem.Ftax_type   = requestCheckVar(request("tax_type"),2)
oelectaxitem.Fbill_type  = requestCheckVar(request("bill_type"),2)
oelectaxitem.Fpc_gbn     = requestCheckVar(request("pc_gbn"),1)
oelectaxitem.Fvol_no     = ""
oelectaxitem.Fissue_no   = ""
oelectaxitem.Fserial_no  = ""
oelectaxitem.Fremark     = ""

oelectaxitem.Fitem_count  = requestCheckVar(request("item_count"),10)
oelectaxitem.Fitem_nm     = requestCheckVar(request("item_nm"),128)
oelectaxitem.Fitem_std    = ""
oelectaxitem.Fitem_qty    = requestCheckVar(request("item_qty"),10)
oelectaxitem.Fitem_price  = requestCheckVar(request("item_price"),30)
oelectaxitem.Fapprove_type= "11" ''""                     ' 01 공급받는자승인  11 공급자가 승인
oelectaxitem.Ffinal_status= "01"
oelectaxitem.Fitem_amt    = requestCheckVar(request("item_amt"),30)
oelectaxitem.Fitem_vat	  = requestCheckVar(request("item_vat"),30)
oelectaxitem.Fitem_remark = requestCheckVar(request("item_remark"),64)

oelectaxitem.Fcur_u_user_no = requestCheckVar(request("cur_u_user_no"),32)
oelectaxitem.Fcur_dam_nm = requestCheckVar(request("cur_dam_nm"),64)
oelectaxitem.Fcur_email  = requestCheckVar(request("cur_email"),128)
oelectaxitem.Fcur_hp_no1 = requestCheckVar(request("cur_hp_no1"),4)
oelectaxitem.Fcur_hp_no2 = requestCheckVar(request("cur_hp_no2"),4)
oelectaxitem.Fcur_hp_no3 = requestCheckVar(request("cur_hp_no3"),4)


''외상미수 추가
oelectaxitem.Fcredit_amt = requestCheckVar(request("credit_amt"),30)

set oelectax = new CElecTaxReg
set oelectax.FRectOneRegitem = oelectaxitem

on Error resume Next
oelectax.SavePreData
If Err then
	response.write "DB작업중 오류 - " + Err.Description + " 관리자 문의 요망"
	dbget.close()	:	response.End
end if
on Error goto 0

oelectax.ExecDTIXmlDom

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

<% if IsSuccess then %>
<script type='text/javascript'>
function PopTaxPrint(itax_no,ibizno){
	var popwinsub = window.open("http://www.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no=" + itax_no + "&cur_biz_no=" + ibizno,"taxview","width=670,height=620,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popwinsub.focus();
}

//var popwinsub = window.open("http://www.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no=<%= itax_no %>&cur_biz_no=<%= ibizno %>","taxview","width=670,height=620,status=no, scrollbars=auto, menubar=no, resizable=yes");
//popwinsub.focus();
alert('세금계산서가 발행 되었습니다.');

//PopTaxPrint('<%= itax_no %>','<%= ibizno %>');
opener.location.reload();
window.close();
</script>
<% else %>
<%
response.redirect "taxregresult.asp?itax_no=" + CStr(itax_no) + "&ierrmsg=" + (ierrmsg)
%>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->