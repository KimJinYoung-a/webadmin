<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim chkidx, songjangdiv, songjangno
dim makerid : makerid     = session("ssBctID")  ''2017/04/10 추가
chkidx      = request("chkidx")
songjangdiv = request("songjangdiv")
songjangno  = request("songjangno")


chkidx      = split(chkidx,",")
songjangdiv = split(songjangdiv,",")
songjangno  = split(songjangno,",")


  
dim sqlStr, i, cnt
cnt = UBound(chkidx)

for i=0 to cnt
    if (Trim(chkidx(i))<>"") then
        sqlStr = "update [db_sitemaster].[dbo].tbl_etc_songjang" + VbCrlf
        sqlStr = sqlStr + " set songjangdiv=" + Trim(songjangdiv(i)) + VbCrlf
        sqlStr = sqlStr + " ,songjangno='" + Trim(songjangno(i)) + "'" + VbCrlf
        sqlStr = sqlStr + " ,issended='Y'"  + VbCrlf
        sqlStr = sqlStr + " ,senddate=getdate()"  + VbCrlf
        sqlStr = sqlStr + " where id=" + requestCheckvar(CStr(chkidx(i)),10)
        sqlStr = sqlStr + " and delivermakerid='"&makerid&"'"   ''2017/04/10 추가
        
        dbget.Execute sqlStr
    end if
Next


dim referer
referer = request.ServerVariables("HTTP_REFERER")
%>
<script language='javascript'>
alert('저장 되었습니다.');
location.replace('<%= referer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->