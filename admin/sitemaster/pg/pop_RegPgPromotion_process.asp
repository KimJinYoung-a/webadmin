<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<%
'#######################################################
'	History	: 2013.09.30 서동석 생성
'			  2022.07.04 한용민 수정(isms취약점수정)
'	Description : 신용카드 프로모션 관리(결제단 무이자 display)
'#######################################################
session.codePage = 65001		'세션코드 UTF-8 강제 설정

%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/sitemasterclass/pgPromotionCls.asp"-->
<%
dim idx     : idx = requestCheckvar(getNumeric(request("idx")),10)
dim cimage  : cimage = requestCheckvar(request("cimage"),200)
dim pgprogbn: pgprogbn = requestCheckvar(request("pgprogbn"),10)
dim cardcd  : cardcd = requestCheckvar(request("cardcd"),10)
dim sDt     : sDt = requestCheckvar(request("sDt"),10)
dim eDt     : eDt = requestCheckvar(request("eDt"),10) & " 23:59:59"
dim conts   : conts = html2DB(request("conts"))
dim contlink: contlink = html2DB(requestCheckvar(request("contlink"),1000))
dim isusing : isusing = html2DB(requestCheckvar(request("isusing"),1))
'rw idx
'rw cimage
'rw pgprogbn
'rw cardcd
'rw sDt
'rw eDt
'rw conts
'rw contlink

Dim sqlStr, AssignedRow
if (idx="") then
    if conts <> "" and not(isnull(conts)) then
        conts = ReplaceBracket(conts)
    end If
    if contlink <> "" and not(isnull(contlink)) then
        contlink = ReplaceBracket(contlink)
    end If

    sqlStr = "insert into db_sitemaster.dbo.tbl_pg_promotion"
    sqlStr = sqlStr&"(cimage,pgprogbn,cardcd,sDt,eDt,conts,contlink,isusing)"
    sqlStr = sqlStr&"values('"&cimage&"'"&VbCRLF
    sqlStr = sqlStr&",'"&pgprogbn&"'"&VbCRLF
    sqlStr = sqlStr&",'"&cardcd&"'"&VbCRLF
    sqlStr = sqlStr&",'"&sDt&"'"&VbCRLF
    sqlStr = sqlStr&",'"&eDt&"'"&VbCRLF
    sqlStr = sqlStr&",'"&conts&"'"&VbCRLF
    sqlStr = sqlStr&",'"&contlink&"'"&VbCRLF
    sqlStr = sqlStr&",'"&isusing&"'"&VbCRLF
    sqlStr = sqlStr&")"
    dbget.Execute sqlStr,AssignedRow
else
    if conts <> "" and not(isnull(conts)) then
        'conts = ReplaceBracket(conts)
    end If
    if contlink <> "" and not(isnull(contlink)) then
        contlink = ReplaceBracket(contlink)
    end If

    sqlStr = "update db_sitemaster.dbo.tbl_pg_promotion"
    sqlStr = sqlStr&" set cimage='"&cimage&"'"&VbCRLF
    sqlStr = sqlStr&" , pgprogbn='b'"&VbCRLF
    sqlStr = sqlStr&" , cardcd='"&cardcd&"'"&VbCRLF
    sqlStr = sqlStr&" , sDt='"&sDt&"'"&VbCRLF
    sqlStr = sqlStr&" , eDt='"&eDt&"'"&VbCRLF
    sqlStr = sqlStr&" , conts='"&conts&"'"&VbCRLF
    sqlStr = sqlStr&" , contlink='"&contlink&"'"&VbCRLF
    sqlStr = sqlStr&" , isusing='"&isusing&"'"'"&VbCRLF
    sqlStr = sqlStr&" where idx="&idx&VbCRLF
    dbget.Execute sqlStr,AssignedRow
end if

if (AssignedRow>0) then
%>
<script language='javascript'>
    alert('저장되었습니다.');
    opener.location.reload();
    window.close();
</script>
<%
end if

session.codePage = 949		'세션코드 EUC-KR 원복
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->