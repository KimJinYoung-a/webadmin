<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������û�� ����Ʈ
' History : 2011.10.13 ������  ����
'			2021.02.24 �ѿ�� ����(�˻������߰�)
'' ToDo ��Ÿ��(�޿�) ���ۺҰ�(DBŸ�Լ��� or ) // ȯ�� ����..
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/payreqListCls.asp"-->
<!-- #include virtual="/lib/classes/approval/payRequestCls.asp"-->
<!-- #include virtual="/lib/classes/approval/commCls.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<%
Dim isUseSerp : isUseSerp = true

Dim clsPay
Dim ipayrequeststate ,sadminId
Dim ipayrequestidx
Dim iCurrpage,ipagesize,iTotCnt,iTotalPage
Dim arrList,intLoop
Dim ipayRequestType,spayRequestTitle
Dim searchType, searchsdate, searchedate, blnTakeDoc, sUserName
Dim arrAccount, intA ,arrAccConts, intAC
Dim iarap_cd,sarap_nm, research, iOutBank, selBiz, payReqType, PayReqPRice, notIncEtc,sCustNm
Dim DocSendErp, payType, payrequesttitle
	payrequesttitle = requestCheckvar(Request("payrequesttitle"),200)
	iPageSize = 30
	iCurrPage = requestCheckvar(getNumeric(trim(Request("iCP"))),10)
	if iCurrPage="" then iCurrPage=1

	sadminId =  session("ssBctId")
	ipayrequestIdx	= requestCheckvar(Request("ipridx"),10)

	searchType = requestCheckvar(Request("selST"),1)
	searchsdate= requestCheckvar(Request("selSD"),10)
	searchedate= requestCheckvar(Request("selED"),10)
	iarap_cd		= requestCheckvar(Request("iaidx"),13)
	sarap_nm		= requestCheckvar(Request("selarap"),50)
    ipayrequeststate= requestCheckvar(Request("selPRS"),4)
	blnTakeDoc= requestCheckvar(Request("selTD"),1)
	sUserName= requestCheckvar(Request("sUnm"),30)
	research = requestCheckvar(Request("research"),10)
	iOutBank = requestCheckvar(Request("selOB"),30)
	selBiz   = requestCheckvar(Request("selBiz"),30)
	payReqType = requestCheckvar(Request("payReqType"),30)
	PayReqPRice = requestCheckvar(getNumeric(trim(Request("PayReqPRice"))),10)
	notIncEtc   = requestCheckvar(Request("notIncEtc"),10)
	sCustNm			= requestCheckvar(Request("sCnm"),50)
	DocSendErp   	= requestCheckvar(Request("DocSendErp"),10)
	payType         = requestCheckvar(Request("payType"),10)
if (research="") and (ipayrequeststate="") then ipayrequeststate="255"
if (research="") and (notIncEtc="") then notIncEtc="ex"

'���� �⺻ �� ���� ��������
set clsPay = new CPayReqList
	clsPay.Fpayrequesttitle = payrequesttitle
	clsPay.FpayRequestIdx =ipayrequestidx
	clsPay.FpayRequestType	        = ipayRequestType
	clsPay.FSearchType		        = searchType
	clsPay.FSDate					= searchsdate
	clsPay.FEDate					= searchedate
 	clsPay.Farap_cd					= iarap_cd
	clsPay.Fpayrequeststate         = ipayrequeststate
	clsPay.FisTakeDoc				= blnTakeDoc
	clsPay.FUsername				= sUserName
	clsPay.FOutBank				    = iOutBank
	clsPay.FBizSection_CD           = selBiz
	clsPay.FpayRequestType          = payReqType
	clsPay.Fpayrequestprice         = PayReqPRice
	clsPay.FnotIncEtc               = notIncEtc
	clsPay.FCustNm                  = sCustNm
	clsPay.FDocSendErp              = DocSendErp
	clsPay.FpayType                 = payType
	clsPay.FCurrpage 				= iCurrpage
	clsPay.FPagesize				= ipagesize
	arrList = clsPay.fnGetPayReqAllList
	iTotCnt = clsPay.FTotCnt

set clsPay = nothing


	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��

Dim TotSum : TotSum=0
%>




 <script language="javascript" src="/admin/approval/eapp/eapp.js"></script>
<script language="javascript">
<!--
	function jsNewReg(){
		var winR = window.open("regPayRequest.asp","popR","width=880, height=600, resizable=yes, scrollbars=yes");
		winR.focus();
	}

	function jsMod(ipridx){
		var winR = window.open("regPayRequest.asp?ipridx="+ipridx,"popR","width=880, height=600, resizable=yes, scrollbars=yes");
		winR.focus();
	}

	function jsSearch(){
	 document.frm.submit();
	}

	// ������ �̵�
function jsGoPage(iCP)
	{
		document.frm.iCP.value=iCP;
		document.frm.submit();
	}

	//�����׸� �ҷ�����
 	function jsGetARAP(){
 			var winARAP = window.open("/admin/linkedERP/arap/popGetARAP.asp","popARAP","width=600,height=600,resizable=yes, scrollbars=yes");
 			winARAP.focus();
 	}

 	function jsReSetARAP(){
 			document.frm.iaidx.value = 0;
 			document.frm.selarap.value = "";
 	}

 	//���� �����׸� ��������
 	function jsSetARAP(dAC, sANM,sACC,sACCNM){
 		document.frm.iaidx.value = dAC;
 		document.frm.selarap.value = sANM;
 	}


function CkeckAll(comp){
    var frm = comp.form;
    var bool =comp.checked;
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    if (e.disabled) continue;
			e.checked=bool;
			AnCheckClick(e)
		}
	}
}

function checkThis(comp){
    AnCheckClick(comp)
}

function jsLinkERP(frm){
    var ischecked =false;

    for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    ischecked = e.checked;
			if (ischecked) break;
		}
	}

	if (!ischecked){
	    alert('���� ������ �����ϴ�.');
	    return;
	}

	if (confirm('���� ������ ERP�� �����Ͻðڽ��ϱ�?')){
	    frm.action = "erpLink_Process.asp";
	    frm.LTp.value="A";
	    frm.submit();
	}
}

function jsLink_SERP_unlock(frm){
    for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    e.disabled=false;
		}
	}
}

function jsLink_SERP(frm){
    var ischecked =false;

   
    for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    ischecked = e.checked;
			if (ischecked) break;
		}
	}


	if (!ischecked){
	    alert('���� ������ �����ϴ�.');
	    return;
	}

	if (confirm('���� ������ S-ERP�� �����Ͻðڽ��ϱ�?')){
	    frm.action = "S_erpLink_Process.asp";
	    frm.LTp.value="A";
	    frm.submit();
	}
}

function jsReceiveERP(frm){
    if (confirm('���� ����� ���� �Ͻðڽ��ϱ�?')){
	    frm.LTp.value="R";
	    frm.action = "erpLink_Process.asp";
	    frm.submit();
	}
}

function jsReceive_SERP(frm){
//alert('�۾���');
//return;
    if (confirm('���� ����� ���� �Ͻðڽ��ϱ�?')){
	    frm.LTp.value="R";
	    frm.action = "S_erpLink_Process.asp";
	    frm.submit();
	}
}


//����Ȯ��
function popConfirmPayrequest(iridx,pidx){
    var iURI = '/admin/approval/eapp/confirmpayrequest.asp?iridx='+iridx+'&ipridx='+pidx+'&ias=1'; //ias Ȯ��..
    var popwin = window.open(iURI,'popConfirmPayrequest','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}
//����ݷ�
function popReturnPayrequest(iridx,iprIdx){
	if(confirm("���縦 �ݷ��Ͻðڽ��ϱ�?")){
		document.frmR.iridx.value = iridx ;
		document.frmR.iprIdx.value = iprIdx ;
	 	document.frmR.submit();  
	}
}
function popModPayDoc(iridx,pidx){
	 var iURI = '/admin/approval/eapp/modeappPayDoc.asp?iridx='+iridx+'&ipridx='+pidx ; //ias Ȯ��..
    var popwin = window.open(iURI,'popConfirmPayrequest','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}
//-->
</script>
<form name="frmR" method="post" action="/admin/approval/eapp/procpayrequest.asp">
	<input type="hidden" name="hidM" value="SR"> 
	<input type="hidden" name="iridx" value=""> 
	<input type="hidden" name="iprIdx" value=""> 
</form>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="/admin/approval/payreqList/index.asp">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="iCP" value="">
			<input type="hidden" name="iPS" value="">
			<input type="hidden" name="research" value="on">
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="5" width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
				<td align="left">
						<select name="selST">
							<option value="1" <%IF searchType="1" THEN%>selected<%END IF%>>������û��</option>
							<option value="2" <%IF searchType="2" THEN%>selected<%END IF%>>����������</option>
							<!--<option value="3" <%IF searchType="3" THEN%>selected<%END IF%>>����(�Ա�)��</option>-->
						</select>
						<input type="text" name="selSD" size="10" value="<%=searchSDate%>"><img src="/images/calicon.gif" align="absmiddle" border="0" onClick="jsPopCal('selSD');"  style="cursor:hand;">
						~
						<input type="text" name="selED" size="10" value="<%=searchEDate%>"><img src="/images/calicon.gif" align="absmiddle" border="0" onClick="jsPopCal('selED');"  style="cursor:hand;">
						&nbsp;&nbsp;
					�ۼ���:
					<input type="text" name="sUnm" size="8" value="<%=sUserName%>">
					&nbsp;&nbsp;
					�ŷ�ó:
						<input type="text" name="sCnm" size="20" value="<%=sCustNm%>">
						&nbsp;&nbsp;
					����idx:
						<input type="text" name="iPRidx" size="20" value="<%=ipayrequestidx%>">
				</td>
				<td rowspan="5" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSearch();">
				</td>
			</tr>
			<tr  bgcolor="#FFFFFF" >
				<td>
					�����׸� :
					<input type="hidden" name="iaidx" value="<%=iarap_cd%>" >
					<input type="text" name="selarap" value="<%=sarap_nm%>" size="13" onClick="jsGetARAP();" readonly>
					&nbsp;&nbsp;
					��������:
					<select name="selPRS">
						<option value="">----</option>
						 <%sbOptPayRequestState ipayrequeststate%>
					</select>
					&nbsp;&nbsp;
					��ݿ�������:
					<select name="selOB">
					<option value="">--����--</option>
					<%
					Dim clsComm

					set clsComm = new CcommCode
					clsComm.Fparentkey = 2 '����� ��������
					clsComm.Fcomm_cd = iOutBank
					clsComm.sbOptCommCD
					set clsComm = nothing
					%>
					</select>
					&nbsp;&nbsp;
					����ι�:
					<%
					Dim clsBS, arrBizList
					Set clsBS = new CBizSection
                    	clsBS.FUSE_YN = "Y"
                    	clsBS.FOnlySub = "Y"
                    	arrBizList = clsBS.fnGetBizSectionList
                    Set clsBS = nothing
                    %>
                    <select name="selBiz">
                    <option value="">--����--</option>
                    <% For intLoop = 0 To UBound(arrBizList,2)	%>
                		<option value="<%=arrBizList(0,intLoop)%>" <%IF Cstr(selBiz) = Cstr(arrBizList(0,intLoop)) THEN%> selected <%END IF%>><%=arrBizList(1,intLoop)%></option>
                	<% Next %>
                    </select>

					&nbsp;&nbsp;

				</td>
			</tr>
			<tr  bgcolor="#FFFFFF" >
			    <td >
			        erp�������ۿ���:
			        <select name="DocSendErp">
						<option value="">----</option>
						<option value="Y" <%IF DocSendErp = "Y" THEN%>selected<%END IF%>>���� ���ۿϷ�</option>
						<option value="N" <%IF DocSendErp = "N" THEN%>selected<%END IF%>>���� ������</option>
					</select>
					&nbsp;&nbsp;
			        �������⿩��:
					<select name="selTD">
						<option value="">----</option>
						<option value="1" <%IF blnTakeDoc = "1" THEN%>selected<%END IF%>>Y</option>
						<option value="0" <%IF blnTakeDoc ="0" THEN%>selected<%END IF%>>N</option>
					</select>
					&nbsp;&nbsp;
					�ڱݿ뵵
					<input type="text" name="payrequesttitle" size="20" value="<%= payrequesttitle %>">
			    </td>
			</tr>
			<tr  bgcolor="#FFFFFF" >
				<td>
			        ������ûType :
					<input type="radio" name="payReqType" value="" <%= CHKIIF(payReqType="","checked","") %> >��ü
					<input type="radio" name="payReqType" value="1" <%= CHKIIF(payReqType="1","checked","") %>  >�Ϲ�
					<input type="radio" name="payReqType" value="9" <%= CHKIIF(payReqType="9","checked","") %> >��ǰ�����������
					<input type="radio" name="payReqType" value="8" <%= CHKIIF(payReqType="8","checked","") %> >������ȯ��
					&nbsp;&nbsp;
					������û�ݾ�:
					<input type="text" name="PayReqPRice" size="10" value="<%= PayReqPRice %>">
					&nbsp;&nbsp;
					<input type="checkbox" name="notIncEtc" value="ex" <%= CHKIIF(notIncEtc="ex","checked","") %> >��Ÿ�ε� �˻�����
				</td>
			</tr>
			<tr  bgcolor="#FFFFFF" >
				<td>
				    ������� :
				    <input type="radio" name="payType" value="" <%= CHKIIF(payType="","checked","") %> >��ü
				    <input type="radio" name="payType" value="2" <%= CHKIIF(payType="2","checked","") %> >������ü
				    <input type="radio" name="payType" value="255" <%= CHKIIF(payType="255","checked","") %> >ETC(��ȭ,�ڵ���ü,����������,Check��ü ��)
				</td>
			</tr>

		</table>
	</td>
</tr>
<tr>
	<td>
	    <table width="100%" cellspacing="0" cellpadding="0">
	    <tr>
	        <td align="left"><input type="button" class="button" value="�űԵ��" onClick="jsNewReg();"></td>
	        
	        <% if (isUseSerp) then %>
	            <td align="right" width="200"><input type="button" class="button" value="sERP ����" onClick="jsLink_SERP(frmAct);"></td>
	            <td align="right" width="170"><input type="button" class="button" value="sERP ������� ����" onClick="jsReceive_SERP(frmAct);"></td>
	        <% else %>
	        <td align="right" width="200"><input type="button" class="button" value="ERP ����" onClick="jsLinkERP(frmAct);"></td>
	        <td align="right" width="170"><input type="button" class="button" value="ERP ������� ����" onClick="jsReceiveERP(frmAct);"></td>
    	    <% if C_ADMIN_AUTH or C_MngPart then %>
    	            
    	        <td align="right" width="400">
    	        <font color=red>sERP[</font>
    	        
    	        <input type="button" value="unlock" onClick="jsLink_SERP_unlock(frmAct)">
    	        <input type="button" class="button" value="sERP ����" onClick="jsLink_SERP(frmAct);">
    	        
    	        <input type="button" class="button" value="sERP ������� ����" onClick="jsReceive_SERP(frmAct);">
    	        
    	        <font color=red>]</td>
    	            
    	    <% end if %>
    	    <% end if %>
	    </tr>
	    </table>
	    </form>
	</td>
</tr>

<tr>
	<td>
		<!-- ��� �� ���� -->
		<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a"   border="0">
		<Form name="frmAct" method="post" action="erpLink_Process.asp">
		<input type="hidden" name="LTp" value="A">
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
				    <td width="20"><input type="checkbox" name="chkAll" onClick="CkeckAll(this)"></td>
					<td>����<br>Idx</td>
					<td width="120">�ڱݿ뵵</td>
					<td>�����׸�</td>
					<td>�μ�����</td>
					<td>�ŷ�ó</td>
					<!--<td>�������</td>//-->
					<td>������û�ݾ�</td>
					<td>������û��</td>
					<td>����������</td>
					<!--<td>����(�Ա�)��</td>-->
					<td>�ۼ���</td>
					<td>�ۼ���</td>
					<td>��������</td>
					<td>ERP<br>��������</td>
					<td>����<br>����</td>
					<td>����<br>����</td>
					<td>����<br>Ȯ��</td>
					<td>����<br>�ݷ�</td>
				</tr>
				<%IF isArray(arrList) THEN
					For intLoop = 0 To UBound(arrList,2)
					TotSum = TotSum + arrList(6,intLoop)
				%>
				<tr bgcolor="#FFFFFF" align="center">
				    <td <%= CHKIIF(arrList(32,intLoop)="2" or (arrList(32,intLoop)="0") or ISNULL(arrList(32,intLoop)),"","bgcolor='#F3F399'") %>><input type="checkbox" name="chk" value="<%=arrList(0,intLoop)%>" onClick="checkThis(this)" <%= CHKIIF((arrList(16,intLoop)="7") AND (arrList(32,intLoop)=2),"","disabled") %> ></td>
					<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>);"><%=arrList(0,intLoop)%></a></td>
					<td width="120"><%=arrList(3,intLoop)%></td>
					<td><%=arrList(19,intLoop)%><br>[<%=arrList(27,intLoop)%>]&nbsp;<font color=gray><%=arrList(28,intLoop)%></font></td>
					<td><%=arrList(25,intLoop)%></td>
					<td><%=arrList(29,intLoop)%></td>
					<!--<td><%=arrList(24,intLoop)%></td>//-->
					<td><%=formatnumber(arrList(6,intLoop),0)%></td>
					<td>
					    <% if isNULL(arrList(5,intLoop)) then %>
                        <font color="red">NULL</font>
					    <% else %>
					    <% if (arrList(16,intLoop)<9) and Left(CStr(arrList(5,intLoop)),10)< Left(CStr(now()),10) then %>
					    <font color=red><%=arrList(5,intLoop)%></font>
					    <% elseif  (arrList(16,intLoop)<9) and Left(CStr(arrList(5,intLoop)),10)=Left(CStr(now()),10) then %>
					    <b><%=arrList(5,intLoop)%></b>
					    <% else %>
					    <%=arrList(5,intLoop)%>
					    <% end if %>
					    <% end if %>
					</td>
					<td><%=arrList(10,intLoop)%></td>
					<!--<td><%=arrList(12,intLoop)%></td>-->
					<td><%=arrList(21,intLoop)%></td>
					<td><%=Replace(arrList(18,intLoop)," ��","<br>��")%></td>
					<td><%=fnGetPayRequestState(arrList(16,intLoop))%></td>
					<td>
					    <% if Not IsNULL(arrList(23,intLoop)) then %>
					    [<%=arrList(22,intLoop)%>]<%=arrList(23,intLoop)%>
					    <% end if %>

					    <% if Not IsNULL(arrList(30,intLoop)) then %>
					    <br>
					    [<%=arrList(30,intLoop)%>]<%=arrList(31,intLoop)%>
					    <% end if %>
					</td>
					<td><%=arrList(26,intLoop)%><br>
							<img src="/images/icon_arrow_link.gif" onClick="popModPayDoc(<%=arrList(2,intLoop)%>,<%=arrList(0,intLoop)%>);" style="cursor:pointer">
						</td>
					<td><%IF arrList(14,intLoop) THEN%>Y<%ELSE%>N<%END IF%></td>
					<td>
					<% if arrList(16,intLoop)=1 Then %>
					<img src="/images/icon_arrow_link.gif" onClick="popConfirmPayrequest(<%=arrList(2,intLoop)%>,<%=arrList(0,intLoop)%>);" style="cursor:pointer">
					<% end if %>
					</td>
					<td>
					<% if arrList(16,intLoop)=1 or arrList(16,intLoop)=7 Then %>
					<img src="/images/icon_arrow_link.gif" onClick="popReturnPayrequest(<%=arrList(2,intLoop)%>,<%=arrList(0,intLoop)%>);" style="cursor:pointer">
					<% end if %>
					</td>
				</tr>
				<%
					Next
				%>
				<tr>
				    <td></td>
				    <td colspan="5"></td>
				    <td><%=formatnumber(TotSum,0) %></td>
				    <td colspan="10"></td>
				</tr>
				<%
					ELSE
				%>
				<tr bgcolor="#FFFFFF">
					<td colspan="17" align="center">��ϵ� ������ �����ϴ�.</td>
				</tr>
				<%END IF%>
				</table>
			</td>
		</tr>
        </form>
<!-- ������ ���� -->
		<%
		Dim iStartPage,iEndPage,iX,iPerCnt
		iPerCnt = 10

		iStartPage = (Int((iCurrPage-1)/iPerCnt)*iPerCnt) + 1

		If (iCurrPage mod iPerCnt) = 0 Then
			iEndPage = iCurrPage
		Else
			iEndPage = iStartPage + (iPerCnt-1)
		End If
		%>
			<tr height="25" >
				<td colspan="17" align="center">
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
					    <tr valign="bottom" height="25">
					        <td valign="bottom" align="center">
					         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
							<% else %>[pre]<% end if %>
					        <%
								for ix = iStartPage  to iEndPage
									if (ix > iTotalPage) then Exit for
									if Cint(ix) = Cint(iCurrPage) then
							%>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
							<%		else %>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
							<%
									end if
								next
							%>
					    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
							<% else %>[next]<% end if %>
					        </td>
					    </tr>
					</table>
				</td>
			</tr>
			</table>
	</td>
</tr>
</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
