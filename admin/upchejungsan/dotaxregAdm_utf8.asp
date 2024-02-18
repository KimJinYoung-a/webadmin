<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 세금계산서 발행후 저장
' History : 서동석 생성
'           2022.11.08 한용민 수정(빌36524 플래시연동발행 에서 위하고 api 연동으로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/classes/jungsan/electaxcls.asp" -->
<%
dim IsBill36524, oelectaxitem, oelectax, sqlstr,IsExistsTax, billsite, credit_amt, IsWEHAGO
dim jungsanid, jungsangubun, makerid, jungsanname, biz_no, corp_nm, ceo_nm, biz_status, biz_type, addr, dam_nm, email
dim hp_no1, hp_no2, hp_no3, write_date, sb_type, tax_type, bill_type, pc_gbn, item_count, item_nm, item_qty
dim item_price, item_amt, item_vat, item_remark, cur_u_user_no, cur_dam_nm, cur_email, cur_hp_no1, cur_hp_no2, cur_hp_no3
dim jgubun, IsDateExistsTax
    jungsanid = requestCheckvar(getNumeric(request("jungsanid")),10)
    jungsangubun = requestCheckvar(request("jungsangubun"),16)
    makerid = requestCheckvar(request("makerid"),32)
    jungsanname = requestCheckvar(request("jungsanname"),128)
    biz_no = requestCheckvar(request("biz_no"),16)
    corp_nm = requestCheckvar(request("corp_nm"),128)
    ceo_nm = requestCheckvar(request("ceo_nm"),64)
    biz_status = requestCheckvar(request("biz_status"),64)
    biz_type = requestCheckvar(request("biz_type"),64)
    addr = requestCheckvar(request("addr"),255)
    dam_nm = requestCheckvar(request("dam_nm"),64)
    email = requestCheckvar(request("email"),128)
    hp_no1 = requestCheckvar(request("hp_no1"),4)
    hp_no2 = requestCheckvar(request("hp_no2"),4)
    hp_no3 = requestCheckvar(request("hp_no3"),4)
    write_date = requestCheckvar(request("write_date"),32)
    sb_type = requestCheckvar(request("sb_type"),2)
    tax_type = requestCheckvar(request("tax_type"),2)
    bill_type = requestCheckvar(request("bill_type"),2)
    pc_gbn = requestCheckvar(request("pc_gbn"),1)
    item_count = requestCheckvar(request("item_count"),10)
    item_nm = requestCheckvar(request("item_nm"),128)
    item_qty = requestCheckvar(getNumeric(request("item_qty")),10)
    item_price = requestCheckvar(request("item_price"),32)
    item_amt = requestCheckvar(request("item_amt"),32)
    item_vat = requestCheckvar(request("item_vat"),32)
    item_remark = requestCheckvar(request("item_remark"),64)
    cur_u_user_no = requestCheckvar(request("cur_u_user_no"),32)
    cur_dam_nm = requestCheckvar(request("cur_dam_nm"),64)
    cur_email = requestCheckvar(request("cur_email"),128)
    cur_hp_no1 = requestCheckvar(request("cur_hp_no1"),4)
    cur_hp_no2 = requestCheckvar(request("cur_hp_no2"),4)
    cur_hp_no3 = requestCheckvar(request("cur_hp_no3"),4)
    credit_amt = requestCheckvar(request("credit_amt"),64)
    billsite = requestCheckvar(request("billsite"),2)
    jgubun = requestCheckvar(request("jgubun"),2)
    IsBill36524 = (billsite="B")
    IsWEHAGO = (billsite="WE")

if write_date="" or isnull(write_date) then write_date=NULL
IsDateExistsTax=false

''기발행 세금계산서인지 체크
sqlstr = "select count(*) as cnt from  [db_jungsan].[dbo].tbl_tax_history_master"
sqlstr = sqlstr + " where jungsanid=" + CStr(request("jungsanid"))
sqlstr = sqlstr + " and jungsangubun='" + CStr(request("jungsangubun")) + "'"
sqlstr = sqlstr + " and makerid='" + CStr(request("makerid")) + "'"
sqlstr = sqlstr + " and resultmsg='OK'" '''기존
sqlstr = sqlstr + " and deleteyn='N'"
sqlstr = sqlstr + " and jgubun='" + CStr(request("jgubun")) + "'"

'response.write sqlstr & "<Br>"
rsget.CursorLocation = adUseClient
rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
	IsExistsTax = rsget("cnt")>0
rsget.close

if IsExistsTax then
    if IsWEHAGO then
        if (request("autotype")="V2") then
            response.write "<script type='text/javascript'>"&vbCRLF
            response.write "parent.opener.addResultLog('"&request("jungsanid")&"','기발행계산서');"&vbCRLF
            session.codePage = 949
            response.write "parent.opener.fnNextEvalProc();"&vbCRLF
            response.write "</script>"
            dbget.close()	:	response.End
        else
			response.write "<script type='text/javascript'>"
			response.write "    alert('이미 발행된 세금계산서 또는 통신중 오류건 입니다.');"
			session.codePage = 949
            response.write "    parent.window.close();"
			response.write "</script>"
			dbget.close()	:	response.End
        end if 
    elseif (IsBill36524) then
        if (request("autotype")="V2") then
            response.write "<script type='text/javascript'>"&vbCRLF
            response.write "parent.opener.addResultLog('"&request("jungsanid")&"','기발행계산서');"&vbCRLF
            session.codePage = 949
            response.write "parent.opener.fnNextEvalProc();"&vbCRLF
            response.write "</script>"
            dbget.close()	:	response.End
        else
			response.write "<script type='text/javascript'>"
			response.write "    alert('이미 발행된 세금계산서 또는 통신중 오류건 입니다.');"
			session.codePage = 949
			response.write "    history.back();"
			response.write "</script>"
			dbget.close()	:	response.End
        end if
    end if
end if

' 스크립트에서 연속클릭으로 두번 중복발행 되는 케이스가 있어서 발행시간 체크
sqlstr = "select count(*) as cnt from [db_jungsan].[dbo].tbl_tax_history_master"
sqlstr = sqlstr & " where jungsanid=" + CStr(request("jungsanid"))
sqlstr = sqlstr & " and jungsangubun='" + CStr(request("jungsangubun")) + "'"
sqlstr = sqlstr & " and makerid='" + CStr(request("makerid")) + "'"
sqlstr = sqlstr & " and deleteyn='N'"
sqlstr = sqlstr & " and jgubun='" + CStr(request("jgubun")) + "'"
sqlstr = sqlstr & " and datediff(s,regdate,getdate())<=3"

'response.write sqlstr & "<Br>"
rsget.CursorLocation = adUseClient
rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
	IsDateExistsTax = rsget("cnt")>0
rsget.close

if IsDateExistsTax then
    if IsWEHAGO then
        if (request("autotype")="V2") then
            response.write "<script type='text/javascript'>"&vbCRLF
            response.write "parent.opener.addResultLog('"&request("jungsanid")&"','3초이내동일계산서발행시도');"&vbCRLF
            session.codePage = 949
            response.write "parent.opener.fnNextEvalProc();"&vbCRLF
            response.write "</script>"
            dbget.close()	:	response.End
        else
			response.write "<script type='text/javascript'>"
			response.write "    alert('3초이내에 동일한 계산서를 발행시도 하셨습니다.\n3초 뒤에 다시 시도해주세요.');"
			session.codePage = 949
            response.write "    parent.window.close();"
			response.write "</script>"
			dbget.close()	:	response.End
        end if 
    elseif (IsBill36524) then
        if (request("autotype")="V2") then
            response.write "<script type='text/javascript'>"&vbCRLF
            response.write "parent.opener.addResultLog('"&request("jungsanid")&"','3초이내동일계산서발행시도');"&vbCRLF
            session.codePage = 949
            response.write "parent.opener.fnNextEvalProc();"&vbCRLF
            response.write "</script>"
            dbget.close()	:	response.End
        else
			response.write "<script type='text/javascript'>"
			response.write "    alert('이미 발행된 세금계산서 또는 통신중 오류건 입니다.\n3초 뒤에 다시 시도해주세요.');"
			session.codePage = 949
			response.write "    history.back();"
			response.write "</script>"
			dbget.close()	:	response.End
        end if
    end if
end if

set oelectaxitem = new CElecTaxRegItem
oelectaxitem.Fjungsanid  = jungsanid
oelectaxitem.Fjungsangubun = jungsangubun
oelectaxitem.Fmakerid     = makerid
oelectaxitem.Fjungsanname= jungsanname
oelectaxitem.Fbiz_no = biz_no
oelectaxitem.Fcorp_nm = corp_nm
oelectaxitem.Fceo_nm = ceo_nm
oelectaxitem.Fbiz_status = biz_status
oelectaxitem.Fbiz_type = biz_type
oelectaxitem.Faddr = addr
oelectaxitem.Fdam_nm = dam_nm
oelectaxitem.Femail = email
oelectaxitem.Fhp_no1 = hp_no1
oelectaxitem.Fhp_no2 = hp_no2
oelectaxitem.Fhp_no3 = hp_no3
oelectaxitem.Fwrite_date = write_date
oelectaxitem.Fsb_type = sb_type
oelectaxitem.Ftax_type   = tax_type
oelectaxitem.Fbill_type  = bill_type
oelectaxitem.Fpc_gbn     = pc_gbn
oelectaxitem.Fvol_no     = ""
oelectaxitem.Fissue_no   = ""
oelectaxitem.Fserial_no  = ""
oelectaxitem.Fremark     = ""
oelectaxitem.Fitem_count  = item_count
oelectaxitem.Fitem_nm     = item_nm
oelectaxitem.Fitem_std    = ""
oelectaxitem.Fitem_qty    = item_qty
oelectaxitem.Fitem_price  = item_price
oelectaxitem.Fapprove_type= "01"                     ' 01 공급받는자승인  11 공급자가 승인
oelectaxitem.Fitem_amt    = item_amt
oelectaxitem.Fitem_vat	  = item_vat
oelectaxitem.Fitem_remark = item_remark
oelectaxitem.Fcur_u_user_no = cur_u_user_no
oelectaxitem.Fcur_dam_nm = cur_dam_nm
oelectaxitem.Fcur_email  = cur_email
oelectaxitem.Fcur_hp_no1 = cur_hp_no1
oelectaxitem.Fcur_hp_no2 = cur_hp_no2
oelectaxitem.Fcur_hp_no3 = cur_hp_no3
''외상미수 추가
oelectaxitem.Fcredit_amt = credit_amt
oelectaxitem.FRectBillSite = billsite
oelectaxitem.FJGubun = jgubun

set oelectax = new CElecTaxReg
set oelectax.FRectOneRegitem = oelectaxitem

'on Error resume Next
oelectax.SavePreData
If Err then
	response.write "DB작업중 오류 - " + Err.Description + " 관리자 문의 요망"
    session.codePage = 949
	dbget.close()	:	response.End
end if
'on Error goto 0

dim psavedIdx : psavedIdx=0
if (IsBill36524 or IsWEHAGO) then
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
<%
if IsWEHAGO then
    session.codePage = 949
%>
    <script type="text/javascript">
        // 세금계산서 발행
        parent.getInvoiceSendTax('<%= psavedIdx %>');
    </script>
<%
elseif (IsBill36524) then
    session.codePage = 949
%>
        <script type="text/javascript">
            parent.FxSendTaxAccount('<%= psavedIdx %>');
        </script>
<% else %>
    <% 
    if IsSuccess then 
        session.codePage = 949
    %>
            <script type="text/javascript">
            alert('세금계산서가 발행 되었습니다.');
            opener.location.reload();
            window.close();
            </script>
    <%
    else
        session.codePage = 949
    %>
        <%
        response.redirect "taxregresult.asp?itax_no=" + CStr(itax_no) + "&ierrmsg=" + (ierrmsg)
        %>
    <% end if %>
<% end if %>
<%
session.codePage = 949
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->