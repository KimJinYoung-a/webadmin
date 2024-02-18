<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 업체 계약 관리
' History : 정윤정 생성
'			2017.12.08 한용민 수정
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
    response.write "오픈할 계약서가 없습니다."
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
	<title>계약서 이메일</title>
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
			<td style="font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; line-height:1.6; padding:0; margin:0"><strong>안녕하세요. 텐바이텐 입니다.</strong><br />
				 <%if isEcContract  then%>
				신규 계약서가 생성 되었습니다.
				협력사 어드민(http://scm.10x10.co.kr)에 로그인 후 업체 계약관리 메뉴에서 전자 서명을 진행해주세요
				<%else%>
				신규 계약서가 발송 되었습니다.<br />
				아래 계약서를 다운로드 받으신 후 출력/날인 하시어 담당자에게 우편으로 발송해 주시기바랍니다.<br />
				(아래 내용은 제휴사어드민(scm.10x10.co.kr) 로그인후 업체계약관리 메뉴에서 확인 가능합니다.)
				<%end if%>
			</td>
		</tr>
		<tr>
			<td style="padding:10px 0; margin:0;">
				<table cellspacing="0" cellpadding="0" style="width:100%; border-collapse:collapse; empty-cells:show; padding:0; margin:0;">
				<thead>
				<tr>
					<th style="font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#eee; padding:5px; margin:0; border:1px solid #ccc;">계약서 명</th>
					<th style="font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#eee; padding:5px; margin:0; border:1px solid #ccc;">계약서번호</th>
					<th style="font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#eee; padding:5px; margin:0; border:1px solid #ccc;">브랜드ID</th>
					<th style="font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#eee; padding:5px; margin:0; border:1px solid #ccc;">판매처</th>
					<th style="font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#eee; padding:5px; margin:0; border:1px solid #ccc;">계약일</th>
					<th style="font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#eee; padding:5px; margin:0; border:1px solid #ccc;">다운로드</th>
				</tr>
				</thead>
				<tbody>
				<% for i=0 to ocontract.FResultCount - 1 %>
				<tr>
					<td style="font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"><%= ocontract.FITemList(i).FContractName %></td>
					<td style="font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"><%= ocontract.FITemList(i).FctrNo %></td>
					<td style="font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"><%= ocontract.FITemList(i).FMakerid %></td>
					<td style="font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"><%= ocontract.FITemList(i).getMajorSellplaceName %></td>
					<td style="font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"><%= ocontract.FITemList(i).FcontractDate %></td>
					<td style="font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; background:#fff; padding:5px; margin:0; border:1px solid #ccc; text-align:center;"><a target="_blank" href="<%= ocontract.FITemList(i).getPdfDownLinkUrlAdm %>"><img src="http://scm.10x10.co.kr/images/pdficon.gif" style="border:0;" /></a></td>
				</tr>
                <% next %>
				</tbody>
				</table>
			</td>
		</tr>
		<tr>
			<td style="font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; padding:10px 0; margin:0; line-height:1.6">
				<strong>* 진행절차</strong><br />
				&nbsp;&nbsp;&nbsp;1.계약서 다운로드 / 각 2부 출력<br />
				&nbsp;&nbsp;&nbsp;2.제휴사에서 계약서 확인후 날인 (간인 불필요) / 1부 우편발송
			</td>
		</tr>
		<tr>
			<td style="font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; padding:10px 0; margin:0; line-height:1.6">
				<strong>* 보내주실 서류</strong><br />
				&nbsp;&nbsp;&nbsp;- 기본계약서, 부속합의서(브랜드별), 제휴사 개인정보수집에 관한 사항<br />
				&nbsp;&nbsp;&nbsp;- 결제통장 사본 1부<br />
				&nbsp;&nbsp;&nbsp;- 사업자 등록증 사본 1부<br />
				&nbsp;&nbsp;&nbsp;- 인감증명서 원본 (계약서에 날인한 도장, 최초 계약에 한함)

			</td>
		</tr>
	<tr>
			<td style="font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; padding:10px 0; margin:0; line-height:1.6">
				<strong>* 계약서 보내주실곳</strong><br />
				&nbsp;&nbsp;&nbsp;- 주소 : (03082) 서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐 제휴사 계약서 담당자 앞
			</td>
		</tr>
		<% if oMdInfoList.FResultCount>0 then %>

		<tr>
			<td style="font-size:12px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#333; padding:10px 0; margin:0; line-height:1.6">
				<strong>* 담당엠디</strong><br />
				<% for i=0 to oMdInfoList.FResultCount-1 %>
				&nbsp;&nbsp;&nbsp;- <%= oMdInfoList.FItemList(i).Fusername%>&nbsp;<%= oMdInfoList.FItemList(i).Fposit_name%> <%=CHKIIF(oMdInfoList.FItemList(i).isMaybeOffMD,"&nbsp;(오프라인 담당)","") %>
				<br />&nbsp;&nbsp;&nbsp;- tel : 02-554-2033 <%= CHKIIF(oMdInfoList.FItemList(i).Fextension="","","(내선 "&oMdInfoList.FItemList(i).Fextension&")")%> <%= CHKIIF(oMdInfoList.FItemList(i).Fdirect070="",""," / 직통 :"&oMdInfoList.FItemList(i).Fdirect070)  %>
				<% if (oMdInfoList.FItemList(i).Fusermail<>"") then %>
				<br />&nbsp;&nbsp;&nbsp;- 이메일 : <a href="mailto:<%= oMdInfoList.FItemList(i).Fusermail %>" style="color:#333;"><%= oMdInfoList.FItemList(i).Fusermail %></a>
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
	<td style="font-size:11px; font-family:dotum, dotumche, '돋움', '돋움체', sans-serif; color:#666; background:#eee; padding:15px 10px; margin:0; line-height:1.8">
		(03082) 서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐 <a href="" target="_blank" style="color:#666;">10X10.co.kr</a><br />
		대표이사 : 최은희 <span style="color:#bbb;">|</span> 사업자등록번호 : 211-87-00620 <span style="color:#bbb;">|</span> 
		통신판매업 신고번호 : 제 01-1968호 <span style="color:#bbb;">|</span> 개인정보 보호 및 청소년 보호책임자 : 이문재
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