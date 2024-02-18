<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ü ��� ����
' History : ������ ����
'			2017.12.08 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<%
dim groupid : groupid= requestCheckvar(request("groupid"),10)
dim signtype : signtype = requestCheckvar(request("signtype"),1)
Dim isEcContract : isEcContract = (signtype ="2")
 

dim ocontract
set ocontract = new CPartnerContract
	ocontract.FPageSize=50
	ocontract.FCurrPage = 1
	ocontract.FRectContractState = 0
	ocontract.FRectGroupID = groupid
	ocontract.GetNewContractList

if (ocontract.FResultCount<1) then
    response.write "������ ��༭�� �����ϴ�."
    dbget.Close() : response.end
end if

dim oMdInfoList
set oMdInfoList = new CPartnerContract
oMdInfoList.FRectGroupID = groupid
oMdInfoList.FRectContractState = 0
oMdInfoList.FRectMdId = session("ssBctID")
oMdInfoList.getContractEmailMdList(TRUE)   ''true is TEST

Dim i

dim iMailContents
if signtype ="2" then
	iMailContents = makeEcCtrMailContents(ocontract,oMdInfoList,TRUE,manageUrl)
else
iMailContents = makeCtrMailContents(ocontract,oMdInfoList,TRUE)
end if
%>

<%= iMailContents %>

<% if FALSE then %>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
	<title>��༭ �̸���</title>
</head>
<body>
<table cellspacing="0" cellpadding="0" style="border:0; width:800px; padding:0;">
<tbody>
<tr>
	<td><img width="600" height="60" src="http://fiximage.10x10.co.kr/web2008/mail/mail_header.gif" /></td>
</tr>
<tr>
	<td style="border:5px solid #eee; padding:30px; background:#fff;">
		<table cellspacing="0" cellpadding="0" style="width:100%; padding:0; margin:0">
		<tbody>
		<tr>
			<td style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; line-height:1.6; padding:0; margin:0"><strong>�ȳ��ϼ���. �ٹ����� �Դϴ�.</strong><br />
				 <%if isEcContract  then%>
				�ű� ��༭�� ���� �Ǿ����ϴ�.
				���»� ����(http://scm.10x10.co.kr)�� �α��� �� ��ü ������ �޴����� ���� ������ �������ּ���
				<%else%>
				�ű� ��༭�� �߼� �Ǿ����ϴ�.<br />
				�Ʒ� ��༭�� �ٿ�ε� ������ �� ���/���� �Ͻþ� ����ڿ��� �������� �߼��� �ֽñ�ٶ��ϴ�.<br />
				(�Ʒ� ������ ���޻����(scm.10x10.co.kr) �α����� ��ü������ �޴����� Ȯ�� �����մϴ�.)
				<%end if%>
			</td>
		</tr>
		<tr>
			<td style="padding:10px 0; margin:0;">
				<table cellspacing="0" cellpadding="0" style="width:100%; border-collapse:collapse; empty-cells:show; padding:0; margin:0;">
				<thead>
				<tr>
					<th style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; background:#eee; padding:5px; margin:0; border:1px solid #ccc;">��༭ ��</th>
					<th style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; background:#eee; padding:5px; margin:0; border:1px solid #ccc;">��༭��ȣ</th>
					<th style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; background:#eee; padding:5px; margin:0; border:1px solid #ccc;">�귣��ID</th>
					<th style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; background:#eee; padding:5px; margin:0; border:1px solid #ccc;">�Ǹ�ó</th>
					<th style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; background:#eee; padding:5px; margin:0; border:1px solid #ccc;">�����</th>
					<th style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; background:#eee; padding:5px; margin:0; border:1px solid #ccc;">�ٿ�ε�</th>
				</tr>
				</thead>
				<tbody>
				<% for i=0 to ocontract.FResultCount - 1 %>
				<tr>
					<td style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"><%= ocontract.FITemList(i).FContractName %></td>
					<td style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"><%= ocontract.FITemList(i).FctrNo %></td>
					<td style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"><%= ocontract.FITemList(i).FMakerid %></td>
					<td style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"><%= ocontract.FITemList(i).getMajorSellplaceName %></td>
					<td style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"><%= ocontract.FITemList(i).FcontractDate %></td>
					<td style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"><a target="_blank" href="<%= ocontract.FITemList(i).getPdfDownLinkUrlAdm %>"><img src="http://scm.10x10.co.kr/images/pdficon.gif" style="border:0;" /></a></td>
				</tr>
                <% next %>
				</tbody>
				</table>
			</td>
		</tr>
		<tr>
			<td style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; padding:10px 0; margin:0; line-height:1.6">
				<strong>* ��������</strong><br />
				&nbsp;&nbsp;&nbsp;1.��༭ �ٿ�ε� / �� 2�� ���<br />
				&nbsp;&nbsp;&nbsp;2.���޻翡�� ��༭ Ȯ���� ���� (���� ���ʿ�) / 1�� ����߼�
			</td>
		</tr>
		<tr>
			<td style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; padding:10px 0; margin:0; line-height:1.6">
				<strong>* �����ֽ� ����</strong><br />
				&nbsp;&nbsp;&nbsp;- �⺻��༭, �μ����Ǽ�(�귣�庰), ���޻� �������������� ���� ����<br />
				&nbsp;&nbsp;&nbsp;- �������� �纻 1��<br />
				&nbsp;&nbsp;&nbsp;- ����� ����� �纻 1��<br />
				&nbsp;&nbsp;&nbsp;- �ΰ����� ���� (��༭�� ������ ����, ���� ��࿡ ����)

			</td>
		</tr>
	<tr>
			<td style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; padding:10px 0; margin:0; line-height:1.6">
				<strong>* ��༭ �����ֽǰ�</strong><br />
				&nbsp;&nbsp;&nbsp;- �ּ� : (03082) ����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ����� ���޻� ��༭ ����� ��
			</td>
		</tr>
		<% if oMdInfoList.FResultCount>0 then %>

		<tr>
			<td style="font-size:12px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#333; padding:10px 0; margin:0; line-height:1.6">
				<strong>* ��翥��</strong><br />
				<% for i=0 to oMdInfoList.FResultCount-1 %>
				&nbsp;&nbsp;&nbsp;- <%= oMdInfoList.FItemList(i).Fusername%>&nbsp;<%= oMdInfoList.FItemList(i).Fposit_name%> <%=CHKIIF(oMdInfoList.FItemList(i).isMaybeOffMD,"&nbsp;(�������� ���)","") %>
				<br />&nbsp;&nbsp;&nbsp;- tel : 02-554-2033 <%= CHKIIF(oMdInfoList.FItemList(i).Fextension="","","(���� "&oMdInfoList.FItemList(i).Fextension&")")%> <%= CHKIIF(oMdInfoList.FItemList(i).Fdirect070="",""," / ���� :"&oMdInfoList.FItemList(i).Fdirect070)  %>
				<% if (oMdInfoList.FItemList(i).Fusermail<>"") then %>
				<br />&nbsp;&nbsp;&nbsp;- �̸��� : <a href="mailto:<%= oMdInfoList.FItemList(i).Fusermail %>" style="color:#333;"><%= oMdInfoList.FItemList(i).Fusermail %></a>
				<% end if %>
				<br /><br />
				<% next %>
			</td>
		</tr>
		<% end if %>
		</table>
	</td>
</tr>
<tr>
	<td style="font-size:11px; font-family:dotum, dotumche, '����', '����ü', sans-serif; color:#666; background:#eee; padding:15px 10px; margin:0; line-height:1.8">
		(03082) ����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ����� <a href="" target="_blank" style="color:#666;">10X10.co.kr</a><br />
		��ǥ�̻� : ������ <span style="color:#bbb;">|</span> ����ڵ�Ϲ�ȣ : 211-87-00620 <span style="color:#bbb;">|</span> 
		����Ǹž� �Ű��ȣ : �� 01-1968ȣ <span style="color:#bbb;">|</span> �������� ��ȣ �� û�ҳ� ��ȣå���� : �̹���
	</td>
</tr>
</tbody>
</table>
</body>
</html>
<% end if %>

<%
set ocontract = nothing
set oMdInfoList = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->