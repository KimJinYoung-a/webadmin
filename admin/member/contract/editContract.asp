<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ü ��� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
 <!-- #include virtual="/lib/ecContractApi_function.asp"-->
<%
Dim ctrKey : ctrKey=requestCheckvar(request("ctrKey"),10)

dim oneContract
dim acctoken, reftoken

set oneContract = new CPartnerContract
oneContract.FRectCtrKey = ctrKey

if ctrKey<>"" then
    oneContract.GetOneContractMaster
end if
	
if onecontract.FResultCount<1 then
    response.write "<script>alert('������ ���ų� ��ȿ�� ����ȣ�� �ƴմϴ�.');</script>"
    dbget.close()	:	response.End
end if

dim oContractDetail
set oContractDetail = new CPartnerContract
oContractDetail.FRectCtrKey = ctrKey
oContractDetail.GetContractDetailList

dim oContractSub
set oContractSub = new CPartnerContract
oContractSub.FRectCtrKey = ctrKey
oContractSub.GetContractSubList

Dim isEcContract : isEcContract = (oneContract.FOneItem.FecCtrSeq <> "" and not  isNull(oneContract.FOneItem.FecCtrSeq ) and oneContract.FOneItem.FecCtrSeq <> "0")
Dim isContractEditVaild : isContractEditVaild=(oneContract.FOneItem.FCtrState=0 and not isEcContract)
Dim isContractDelValid  : isContractDelValid=(oneContract.FOneItem.FCtrState<=1)   ''�߼� �� ������ ��� üũ
Dim isContractFinValid  : isContractFinValid=(oneContract.FOneItem.FCtrState=3)
  
dim i

 dim  con_status, con_info,ecCtrState,con_error
 dim APIpath,strParam,objXML,iRbody,jsResult,strErrMsg
 
if isEcContract then
	
	oneContract.fnGetContractToken
	acctoken = oneContract.Facctoken 	
	reftoken = oneContract.Freftoken 
 

	ecCtrState =  fnViewEcCont(oneContract.FOneItem.FecCtrSeq,oneContract.FOneItem.FBcompany_no,oneContract.FOneItem.FecBUser,acctoken)
	
 
		if Fchkerror = "invalid_token" then
			call sbGetRefToken(reftoken)
 			acctoken = Faccess_token		
		 	ecCtrState =   fnViewEcCont(oneContract.FOneItem.FecCtrSeq,oneContract.FOneItem.FBcompany_no,oneContract.FOneItem.FecBUser,acctoken)
		end if
		 
		if ecCtrState="" and strErrMsg<>"" then
            response.write strErrMsg
            response.end
		end if
  
			 
		if ecCtrState<>"" and ecCtrState <> GetContractEcStateName(oneContract.FOneItem.FCtrState) then
			dim sqlStr
			sqlStr = "update db_partner.dbo.tbl_partner_ctr_master set ctrstate = "&GetContractEcState(ecCtrState)&" , lastupdate = getdate() "
			 sqlstr = sqlstr & " where ctrKey="&oneContract.FOneItem.FCtrKey&VbCRLF
			 dbget.Execute  sqlstr, 1
	  end if 
end if
%>				 
<script type="text/javascript" src="contract.js?v=1.00"></script>
<script language='javascript'>
function edtContract(){
    var frm=document.frmCtrEdt;

    if (confirm('��༭�� �����Ͻðڽ��ϱ�?')){
        frm.mode.value="edt";
        frm.submit();
    }
}

function delContract(){
    var frm=document.frmCtrEdt;

<%if not isEcContract then%>
	    if (confirm('��༭�� �����Ͻðڽ��ϱ�?')){
	        frm.mode.value="del";
	        frm.submit();
	    }
  <%else%>
  	 if (confirm('��༭�� �����Ͻðڽ��ϱ�? ���ڰ���� U+����Ʈ���� ���� �����մϴ�.')){
  			jsEcSubmit();
  }
  <%end if%>
}

function admindelContract(){
    var frm=document.frmCtrEdt;

    if (confirm('��༭�� �����Ͻðڽ��ϱ�?')){
        frm.mode.value="del";
        frm.submit();
    }
}

function delContractOpened(){
    var frm=document.frmCtrEdt;
<%if not isEcContract then%>
    if (confirm('�̹� ���µ� ��༭ �Դϴ�. �����Ͻ÷��� ������ ���ۼ� �ϼž� �մϴ�. �����Ͻðڽ��ϱ�?')){
        frm.mode.value="del";
        frm.submit();
    }
    <%else%>
  	 if (confirm('��༭�� �����Ͻðڽ��ϱ�? ���ڰ���� U+����Ʈ���� ���� �����մϴ�.')){
  			jsEcSubmit();
  }
  <%end if%>
}

function finContract(){
    var frm=document.frmCtrEdt;
<%if not isEcContract then%>
    if (confirm('��༭�� ���� �Ϸ� ó���Ͻðڽ��ϱ�?')){
        frm.mode.value="fin";
        frm.submit();
    }
     <%else%>
  	 if (confirm('��༭�� �Ϸ�ó���Ͻðڽ��ϱ�? ���ڰ���� U+����Ʈ���� �Ϸ�ó�� �����մϴ�.')){
  			jsEcSubmit();
  }
  <%end if%>
}

function jsEcSubmit(){
	document.frmecView.target="_blank";
	document.frmecView.submit();
}
</script>
<form name="frmecView" method="post" action="<%=FecURL%>/w20/contractView.do" style="margin:0px;" > 
<input type="hidden" name="remote_id" value="<%=FecId%>" />  <!-- �ۼ��� LOGIN ID -->
<input type="hidden" name="cont_seq" value="<%=onecontract.FOneItem.FecCtrSeq%>" />  <!-- ��༭ ��ȣ -->
<input type="hidden" name="corp_id" value="<%=onecontract.FOneItem.FAcompany_no%>" /> <!-- ����� ȭ���Ϸ��� ����ڹ�ȣ -->
</form> 
<form name="frmCtrEdt" method="post" action="ctrReg_Process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="edt">
<input type="hidden" name="ctrKey" value="<%=onecontract.FOneItem.FctrKey%>">
<input type="hidden" name="groupid" value="<%=onecontract.FOneItem.Fgroupid%>">
<table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA">
    <tr bgcolor="#FFFFFF">
        <td colspan="3"><b>* <%=onecontract.FOneItem.FcontractName%></b></td>
        <td align="right">
            <%if  isEcContract then %>
                <img src="/images/documents_icon.png" style="cursor:pointer;" onClick="jsEcSubmit();">
                &nbsp;
            <%end if%>
            <% If onecontract.FOneItem.FsignType = "D" Then %>
                <img src="/images/browser_icon.png" style="cursor:pointer" onClick="dnWebAdmDocu('<%=onecontract.FOneItem.FctrKey %>');">
            <% Else %>
                <img src="/images/browser_icon.png" style="cursor:pointer" onClick="dnWebAdm('<%=onecontract.FOneItem.FctrKey %>');">
            <% End If %>
            &nbsp;
            <img src="/images/pdf_icon.png" style="cursor:pointer" onClick="dnPdfAdm('<%=onecontract.FOneItem.getPdfDownLinkUrlAdm %>');">
        </td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#FFDDDD" align="center" >�׷��ڵ�</td>
        <td colspan="3"><%=onecontract.FOneItem.Fgroupid%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#FFDDDD" align="center" width="15%">�����</td>
        <td width="35%"><%=onecontract.FOneItem.Fregdate %></td>
        <td bgcolor="#FFDDDD" align="center" width="15%">�߼���</td>
        <td width="35%"><%=onecontract.FOneItem.Fsenddate%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#FFDDDD" align="center" width="15%">��üȮ����</td>
        <td ><%=onecontract.FOneItem.Fconfirmdate%></td>
        <td bgcolor="#FFDDDD" align="center" width="15%">�Ϸ���</td>
        <td ><%=onecontract.FOneItem.Ffinishdate%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#FFDDDD" align="center" width="20%">��༭����</td>
        <td colspan="3"><%=onecontract.FOneItem.GetContractStateName%></td>
    </tr>
</table>

<%if isEcContract then%>
    <div id="ecDiv"> 
        <iframe id="ifrec" name="ifrec" src="about:blank" frameborder="0" width="1024" height="1200"></iframe>
        <script >
            document.frmecView.target = "ifrec" ;
                document.frmecView.submit();
        </script>
    </div>
<%else%>
    <div>
        <table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA">
            <tr bgcolor="#FFFFFF" >
                <td bgcolor="#DDDDFF" width="20%" align="center" colspan="2">��༭Ÿ��</td>
                <td colspan="3">
                    <%if isEcContract then%>
                        ����(<%=onecontract.FOneItem.FecCtrSeq%>)
                    <%else%>
                        <% If onecontract.FOneItem.FsignType = "D" Then %>
                            DocuSign
                        <% Else %>
                            ����
                        <% End If %>
                    <%end if%>
                </td>
            </tr>
            <tr bgcolor="#FFFFFF" >
                <td bgcolor="#DDDDFF" width="20%" align="center" colspan="2">�������</td>
                <td colspan="3"><%=onecontract.FOneItem.FRegUserName%> (<%=onecontract.FOneItem.FRegUserID%>)</td>
            </tr>
        
            <tr bgcolor="#FFFFFF">
                <td bgcolor="#DDDDFF" rowspan="2" align="center" colspan="2">�ٹ�����</td>
                <td ><input type="text" class="text" name="$$A_UPCHENAME$$" value="<%=oContractDetail.getValueByKey("$$A_UPCHENAME$$")%>"></td>
                <td ><input type="text" class="text" name="$$A_COMPANY_NO$$" value="<%=oContractDetail.getValueByKey("$$A_COMPANY_NO$$")%>"></td>
                <td ><input type="text" class="text" name="$$A_CEONAME$$" value="<%=oContractDetail.getValueByKey("$$A_CEONAME$$")%>"></td>
            </tr>
            <tr bgcolor="#FFFFFF">
                <td colspan="3"><input type="text" class="text" name="$$A_COMPANY_ADDR$$" value="<%=oContractDetail.getValueByKey("$$A_COMPANY_ADDR$$")%>" size="40"></td>
            </tr>
            <tr bgcolor="#FFFFFF">
                <td bgcolor="#DDDDFF" rowspan="2" align="center" colspan="2">���޻�</td>
                <td ><input type="text" class="text" name="$$B_UPCHENAME$$" value="<%=oContractDetail.getValueByKey("$$B_UPCHENAME$$")%>"></td>
                <td ><input type="text" class="text" name="$$B_COMPANY_NO$$" value="<%=oContractDetail.getValueByKey("$$B_COMPANY_NO$$")%>"></td>
                <td ><input type="text" class="text" name="$$B_CEONAME$$" value="<%=oContractDetail.getValueByKey("$$B_CEONAME$$")%>"></td>
            </tr>
            <tr bgcolor="#FFFFFF">
                <td colspan="3"><input type="text" class="text" name="$$B_COMPANY_ADDR$$" value="<%=oContractDetail.getValueByKey("$$B_COMPANY_ADDR$$")%>" size="40"></td>
            </tr>
            <tr bgcolor="#FFFFFF">
                <td bgcolor="#DDDDFF" width="20%" align="center" colspan="2">�����</td>
                <td width="30%"><input type="text" class="text" name="$$CONTRACT_DATE$$" value="<%=oContractDetail.getValueByKey("$$CONTRACT_DATE$$")%>"></td>
                <td bgcolor="#DDDDFF" width="20%" align="center">���������</td>
                <td width="30%"><input type="text" class="text" name="$$ENDDATE$$" value="<%=oContractDetail.getValueByKey("$$ENDDATE$$")%>"></td>
            </tr>
            <tr bgcolor="#FFFFFF">
                <% if (onecontract.FOneItem.IsDefaultContract) then %>
                    <td bgcolor="#DDDDFF" width="20%" align="center" colspan="2">���������</td>
                    <td width="30%" colspan="3"><input type="text" class="text" name="$$DEFAULT_JUNGSANDATE$$" value="<%=oContractDetail.getValueByKey("$$DEFAULT_JUNGSANDATE$$")%>" size="30"></td>
                <% else %>
                    <td colspan="5"></td>
                <% end if %>
            </tr>
        </table>
        <p>
    </div>
<%end if%>
<% if (Not onecontract.FOneItem.IsDefaultContract) then %>
    <% if oContractSub.FResultCount>0 then %>
        <table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA">
            <tr bgcolor="#FFFFFF">
                <td colspan="6"> - ������</td>
            </tr>
            <tr bgcolor="#DDDDFF" align="center">
                <td>�귣��ID</td>
                <td>�Ǹ�ä��</td>
                <td>�������</td>
                <td>�⺻������</td>
                <td>���</td>
                <td></td>
            </tr>
            <% for i=0 to oContractSub.FResultCount-1 %>
                <input type="hidden" name="ctrSubKey" value="<%=oContractSub.FItemList(i).FctrSubKey %>">
                <input type="hidden" name="addsellplace" value="<%=oContractSub.FItemList(i).Fsellplace %>">
                <input type="hidden" name="addmwdiv" value="<%=oContractSub.FItemList(i).Fcontractmwdiv %>">
                <tr bgcolor="#FFFFFF"  align="center">
                    <td><%=oneContract.FOneItem.FMakerid %></td>
                    <td><%=oContractSub.FItemList(i).getSellplaceName %></td>
                    <td><%=fnMaeipdivName(oContractSub.FItemList(i).Fcontractmwdiv) %></td>
                    <td>
                        <input type="text" name="addmargin" value="<%=oContractSub.FItemList(i).Fcontractmargin %>" size="5" maxlength="5"> %
                        <% if (oContractSub.FItemList(i).Fcontractmwdiv="M" or oContractSub.FItemList(i).Fcontractmwdiv="B031") then %>
                            (������:<%=100-oContractSub.FItemList(i).Fcontractmargin%>)
                        <% end if %>
                    </td>
                    <td>
                        <% if (FALSE) and (oContractSub.FItemList(i).Fcontractmwdiv="U") then %> <!-- ��ۺ���� ���ǥ�þ���-->
                            <% if Not isNULL(oContractSub.FItemList(i).FdefaultdeliveryType) then %>
                                <select class="select" name="addON_dlvtype">
                                    <option value="">�⺻��å
                                    <option value="9" <%=CHKIIF(oContractSub.FItemList(i).FdefaultdeliveryType="9","selected","") %> >��ü���ǹ��
                                    <option value="7" <%=CHKIIF(oContractSub.FItemList(i).FdefaultdeliveryType="7","selected","") %> >��ü���ҹ��
                                </select>
                                <br><input type="text" class="text" name="addON_dlvlimit" value="<%=oContractSub.FItemList(i).FdefaultFreebeasongLimit%>" size="7" style="text-align:right">�̸�
                                <br><input type="text" class="text" name="addON_dlvpay" value="<%=oContractSub.FItemList(i).Fdefaultdeliverpay%>" size="5" maxlength="5" style="text-align:right">��
                            <% end if %>
                        <% end if %>
                    </td>
                    <td>
                    </td>
                </tr>
            <% next %>
        </table>
        <p>
    <% end if %>
<% end if %>
 
<table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA">
<% if not isEcContract then %>
    <tr bgcolor="#FFFFFF" >
        <td height="30" align="center">
        <% if (isContractEditVaild) then %>
        <input type="button" value="�� ��" class="button" onClick="edtContract()">
        <% end if %>

        <% if (isContractDelValid) then %>
        &nbsp;
            <% if  oneContract.FOneItem.FCtrState=0 then %>
            <input type="button" value="�� ��" class="button" onClick="delContract()">
            <% else %>
            <input type="button" value="��༭ ���� �� ����" class="button" onClick="delContractOpened()">
            <% end if %>
        <% end if %>

        <% if (isContractFinValid) then %>
        &nbsp;
        <input type="button" value="�Ϸ� ó��" class="button" onClick="finContract()">
        <% end if %>
        </td>
    </tr>
    <tr bgcolor="#FFFFFF" >
        <td height="30" align="center">
            <input type="button" value="�� ��" class="button" onClick="delContract()">
            <input type="button" value="�Ϸ� ó��" class="button" onClick="finContract()">
        </td>
    </tr>
<%end if%>

    <tr bgcolor="#FFFFFF" >
        <td height="30" align="center">
            <input type="button" value="����(��������)" class="button" onClick="admindelContract()">
            <% if C_ADMIN_AUTH then %>            
                ������ ���� :            
                <input type="button" value="�Ϸ� ó��" class="button" onClick="finContract()">
            <% end if %>                
        </td>
    </tr>

</table>
</form>
<%
set oneContract = Nothing
set oContractDetail  = Nothing
set oContractSub = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->