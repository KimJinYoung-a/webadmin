<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2010.10.19 eastone 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/lecturer/lecUserCls.asp"-->
<%

Dim lecturer_id : lecturer_id = requestCheckVar(request("lecturer_id"),32)
Dim lecturer_name : lecturer_name = Left(request("lecturer_name"),32)


Dim lec_yn                      : lec_yn = requestCheckVar(request("lec_yn"),10)
Dim lec_margin                  : lec_margin = requestCheckVar(request("lec_margin"),10)
Dim mat_margin                  : mat_margin = requestCheckVar(request("mat_margin"),10)
Dim diy_yn                      : diy_yn = requestCheckVar(request("diy_yn"),10)
Dim diy_margin                  : diy_margin = requestCheckVar(request("diy_margin"),10)
Dim diy_dlv_gubun               : diy_dlv_gubun = requestCheckVar(request("diy_dlv_gubun"),10)
Dim DefaultFreebeasongLimit     : DefaultFreebeasongLimit = requestCheckVar(request("DefaultFreebeasongLimit"),10)
Dim DefaultDeliveryPay          : DefaultDeliveryPay = requestCheckVar(request("DefaultDeliveryPay"),10)
Dim en_name                     : en_name = Left(request("en_name"),32)

if (lec_yn="N") then
    lec_margin=0
    mat_margin=0
end if

if (diy_yn="N") then
    diy_margin="0"
    diy_dlv_gubun=0
    DefaultFreebeasongLimit=0
    DefaultDeliveryPay=0
end if

Dim sqlStr, AssignedRow

sqlStr = " If Exists (select * from db_academy.dbo.tbl_lec_user where lecturer_id='"&lecturer_id&"')"
sqlStr = sqlStr & " BEGIN"
sqlStr = sqlStr & "     update db_academy.dbo.tbl_lec_user" & VbCRLF
sqlStr = sqlStr & "     set lecturer_name=convert(varchar(32),'"&HTML2DB(lecturer_name)&"')" & VbCRLF
sqlStr = sqlStr & "     ,en_name=convert(varchar(32),'"&HTML2DB(en_name)&"')" & VbCRLF
sqlStr = sqlStr & "     ,lec_yn='"&lec_yn&"'" & VbCRLF
sqlStr = sqlStr & "     ,diy_yn='"&diy_yn&"'" & VbCRLF
sqlStr = sqlStr & "     ,lec_margin="&lec_margin & VbCRLF
sqlStr = sqlStr & "     ,mat_margin="&mat_margin & VbCRLF
sqlStr = sqlStr & "     ,diy_margin="&diy_margin & VbCRLF
sqlStr = sqlStr & "     ,diy_dlv_gubun="&diy_dlv_gubun & VbCRLF
sqlStr = sqlStr & "     ,DefaultFreebeasongLimit="&DefaultFreebeasongLimit & VbCRLF
sqlStr = sqlStr & "     ,DefaultDeliveryPay="&DefaultDeliveryPay & VbCRLF
sqlStr = sqlStr & "     where lecturer_id='"&lecturer_id&"'"
sqlStr = sqlStr & " END"
sqlStr = sqlStr & " ELSE"
sqlStr = sqlStr & " BEGIN"
sqlStr = sqlStr & "     insert into db_academy.dbo.tbl_lec_user" & VbCRLF
sqlStr = sqlStr & "     (lecturer_id,lecturer_name,en_name, lec_yn,diy_yn,lec_margin,mat_margin,diy_margin,diy_dlv_gubun,DefaultFreebeasongLimit,DefaultDeliveryPay)"& VbCRLF
sqlStr = sqlStr & "     values('"&lecturer_id&"'"
sqlStr = sqlStr & "     ,convert(varchar(32),'"&HTML2DB(lecturer_name)&"')" & VbCRLF
sqlStr = sqlStr & "     ,convert(varchar(32),'"&HTML2DB(en_name)&"')" & VbCRLF
sqlStr = sqlStr & "     ,'"&lec_yn&"'" & VbCRLF
sqlStr = sqlStr & "     ,'"&diy_yn&"'" & VbCRLF
sqlStr = sqlStr & "     ,"&lec_margin&"" & VbCRLF
sqlStr = sqlStr & "     ,"&mat_margin&"" & VbCRLF
sqlStr = sqlStr & "     ,"&diy_margin&"" & VbCRLF
sqlStr = sqlStr & "     ,"&diy_dlv_gubun&"" & VbCRLF
sqlStr = sqlStr & "     ,"&DefaultFreebeasongLimit&"" & VbCRLF
sqlStr = sqlStr & "     ,"&DefaultDeliveryPay&"" & VbCRLF
sqlStr = sqlStr & "     )" & VbCRLF
sqlStr = sqlStr & " END"

''rw sqlStr
dbAcademyget.Execute sqlStr,AssignedRow


%>
<script language='javascript'>
    alert('<%= AssignedRow %>건 적용되었습니다.')
    location.replace('<%= Request.ServerVariables("HTTP_REFERER") %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
