<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 업체 로그인
' History : 이상구 생성
'			2021.01.12 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

if Not (C_SYSTEM_Part or C_OP or C_ADMIN_AUTH or C_CSPowerUser or C_MngPart) then
	response.write "<script>alert('관리자권한이 없습니다.')</script>"
	response.write "관리자권한이 없습니다."
	dbget.close() : response.end
end if

dim makerid, groupid, pcuserdiv
makerid = requestCheckvar(request("makerid"),32)

dim opartner
set opartner = new CPartnerUser
	opartner.FRectDesignerID = makerid
	opartner.GetOnePartnerNUser

if opartner.FResultCount<=0 then
	response.write "<script>alert('존재하지 않는 브랜드 아이디입니다.')</script>"
	response.write "존재하지 않는 브랜드 아이디입니다."
	dbget.close() : response.end
end if

pcuserdiv = opartner.FOneItem.Fpcuserdiv
groupid = opartner.FOneItem.FGroupid


if (pcuserdiv <> "9999_02") then
	response.write "<script>alert('일반 매입처만 로그인 가능합니다.')</script>"
	response.write "일반 매입처만 로그인 가능합니다."
	dbget.close() : response.end
end if

session("ssBctId") = makerid
session("ssGroupid") = groupid
session("ssBctDiv") = trim("9999")

''강사의 경우
''session("ssUserCDiv")="14"

     'session("ssBctId") = "ban8"
     'session("ssGroupid") = "G01488"
     'session("ssBctDiv") = "9999"

     'session("ssBctId") = "dashanddot"
     'session("ssGroupid") = "G02424"
     'session("ssBctDiv") = "9999"

'' offshop
    'session("ssBctId") = "wholesale1003"
    'session("ssGroupid") = "G05971"
    'session("ssBctDiv") = "503"

    ''/offshop/index.asp

'' 강사(/lectureadmin/)
    'session("ssBctId") = "92mir"
    'session("ssGroupid") = "G02433"
    'session("ssBctDiv") = "14"

%>
<script>document.location.href = '/partner/';</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
