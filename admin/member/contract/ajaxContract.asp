<%@ language=vbscript %>
<%
option Explicit
'Response.Buffer = True
Response.CharSet = "euc-kr"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
dim mode : mode = requestCheckvar(request("mode"),32)
dim groupid : groupid = requestCheckvar(request("groupid"),32)
dim vBody
dim oDftCTRPTypeDetail
dim ecAUser

 

if (mode="addDft") then

    set oDftCTRPTypeDetail = new CPartnerContract
    oDftCTRPTypeDetail.FRectContractType = DEFAULT_CONTRACTTYPE
    oDftCTRPTypeDetail.FRectGroupID = groupid
    oDftCTRPTypeDetail.getContractDetailProtoTypeWithGroupInfo

    '' �⺻��༭
    vBody = "<input type='hidden' name='addftkey' value='1'>"
    vBody = vBody&"<input type='hidden' name='contractType' value='"&DEFAULT_CONTRACTTYPE&"'>"
    vBody = vBody&"<table width='100%' border='0' cellspacing='1' cellpadding='4' class='a' bgcolor='#BABABA' >"
    vBody = vBody&"<tr bgcolor='#FFFFFF' >"
    vBody = vBody&"	<td bgcolor='#DDDDFF' width='20%' align='center' colspan='2'>��༭����</td>"
    vBody = vBody&"	<td colspan='3'><input type='checkbox' name='chkCT11' value='1' checked>�ŷ��⺻��༭ <input type='checkbox' name='chkCT12' value='1'>�����԰�༭ <input type='checkbox' name='chkCT13' value='1' onclick='saAreaCheck(this);'>Ư���༭</td>"
    vBody = vBody&"</tr>"
    vBody = vBody&"<tr bgcolor='#FFFFFF' >"
    vBody = vBody&"	<td bgcolor='#DDDDFF' width='20%' align='center' colspan='2'>�������</td>"
    vBody = vBody&"	<td colspan='3'>"&session("ssBctCname")&"("&session("ssBctID")&")</td>"
    vBody = vBody&"</tr>" 
    vBody = vBody&"<tr bgcolor='#FFFFFF'>"
    vBody = vBody&"	<td bgcolor='#DDDDFF' rowspan='2' align='center' colspan='2'>�ٹ�����</td>"
    vBody = vBody&"	<td ><input type='text' class='text' name='$$A_UPCHENAME$$' value='"&oDftCTRPTypeDetail.getDefaultValueByKey("$$A_UPCHENAME$$")&"'></td>"
    vBody = vBody&"	<td ><input type='text' class='text' name='$$A_COMPANY_NO$$' value='"&oDftCTRPTypeDetail.getDefaultValueByKey("$$A_COMPANY_NO$$")&"'></td>"
    vBody = vBody&"	<td ><input type='text' class='text' name='$$A_CEONAME$$' value='"&oDftCTRPTypeDetail.getDefaultValueByKey("$$A_CEONAME$$")&"'></td>"
    vBody = vBody&"</tr>"
    vBody = vBody&"<tr bgcolor='#FFFFFF'>"
    vBody = vBody&"	<td colspan='3'><input type='text' class='text' name='$$A_COMPANY_ADDR$$' value='"&oDftCTRPTypeDetail.getDefaultValueByKey("$$A_COMPANY_ADDR$$")&"' size='40'></td>"
    vBody = vBody&"</tr>"
    vBody = vBody&"<tr bgcolor='#FFFFFF'>"
    vBody = vBody&"<td bgcolor='#DDDDFF' rowspan='2' align='center' colspan='2'>���»�</td>"
    vBody = vBody&"<td ><input type='text' class='text' name='$$B_UPCHENAME$$' value='"&replace(oDftCTRPTypeDetail.getDefaultValueByKey("$$B_UPCHENAME$$"),"'","")&"'></td>"
    vBody = vBody&"<td ><input type='text' class='text' name='$$B_COMPANY_NO$$' value='"&oDftCTRPTypeDetail.getDefaultValueByKey("$$B_COMPANY_NO$$")&"'></td>"
    vBody = vBody&"<td ><input type='text' class='text' name='$$B_CEONAME$$' value='"&oDftCTRPTypeDetail.getDefaultValueByKey("$$B_CEONAME$$")&"'></td>"
    vBody = vBody&"</tr>"
    vBody = vBody&"<tr bgcolor='#FFFFFF'>"
    vBody = vBody&"<td colspan='3'><input type='text' class='text' name='$$B_COMPANY_ADDR$$' value='"&oDftCTRPTypeDetail.getDefaultValueByKey("$$B_COMPANY_ADDR$$")&"' size='40'></td>"
    vBody = vBody&"</tr>"
    vBody = vBody&"<tr bgcolor='#FFFFFF'>"
    vBody = vBody&"<td bgcolor='#DDDDFF' width='20%' align='center' colspan='2'>�����</td>"
    vBody = vBody&"<td width='30%'><input type='text' class='text' name='$$CONTRACT_DATE$$' value='"&oDftCTRPTypeDetail.getDefaultValueByKey("$$CONTRACT_DATE$$")&"'></td>"
    vBody = vBody&"<td bgcolor='#DDDDFF' width='20%' align='center' >���������</td>"
    vBody = vBody&"<td width='30%'><input type='text' class='text' name='$$ENDDATE$$' value='"&oDftCTRPTypeDetail.getDefaultValueByKey("$$ENDDATE$$")&"'></td>"
    vBody = vBody&"<tr bgcolor='#FFFFFF'>" 
    vBody = vBody&"<td bgcolor='#DDDDFF' width='20%' align='center' colspan='2'>���������</td>"
    vBody = vBody&"<td width='30%' colspan='3'><input type='text' class='text' name='$$DEFAULT_JUNGSANDATE$$' value='"&oDftCTRPTypeDetail.getDefaultValueByKey("$$DEFAULT_JUNGSANDATE$$")&"' size='30'></td>"
    vBody = vBody&"</tr>"
    vBody = vBody&"<tr bgcolor='#FFFFFF' id='specialAppointmentArea' style='display:none;'>"
    vBody = vBody&"<td bgcolor='#DDDDFF' width='20%' align='center' colspan='2'>Ư�೻��</td>"
    vBody = vBody&"<td width='30%' colspan='3'><textarea rows='5' cols='100' name='$$CONTENTS_CONTS$$' id='specialAppointmentContents'></textarea></td>"    
    vBody = vBody&"</tr>"
    vBody = vBody&"</table>"

    SET oDftCTRPTypeDetail = Nothing
else
    vBody = "���ǵ��� �ʾҽ��ϴ�. ["&mode&"]"
end if
Response.Write vBody
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->