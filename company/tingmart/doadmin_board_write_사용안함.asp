<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<%
dim site_name
site_name = request("site_name")
if (site_name = "") then
            response.write("<script>window.alert('Site �����ڰ� �Ѿ���� �ʾҽ��ϴ�.');</script>")
            response.write("<script>history.back();</script>")
            dbget.close()	:	response.End
end if
dim table_name
table_name = request("table_name")
if (table_name = "") then
            response.write("<script>window.alert('table �����ڰ� �Ѿ���� �ʾҽ��ϴ�.');</script>")
            response.write("<script>history.back();</script>")
            dbget.close()	:	response.End
end if

dim name
name = request("name")

dim id,thread,pos,depth
id = request("id")
thread = request("thread")
pos = request("pos")
depth = request("depth")

dim resulthtml
dim title, body, mail,send_mail
title = request.form("title")
body = request.form("body")
mail = request.form("mail")
send_mail = request.form("send_mail")
title = html2db(title)
body = html2db(body)
mail = html2db(mail)

if( (body = "") or (title = "") )then
            response.write("<script>window.alert('����,������ �� �Է��ϼž��մϴ�.');</script>")
            response.write("<script>history.go(-1);</script>")
            dbget.close()	:	response.End
end if

if send_mail = "Y" then
    ' ������ �ҷ��ͼ� ������ ����
    dim fs,objFile,dirPath,fileName,mailtitle,mailcontent
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    dirPath = server.mappath("/admin/board")
    fileName = dirPath&"\\mail_form.html"

    Set objFile = fs.OpenTextFile(fileName,1)
    mailcontent = objFile.readall
    mailcontent = replace(mailcontent,":MAILCONTENT:",request.form("body"))
    mailtitle = "[10x10]���Բ��� �����Ͻ� ���뿡 ���� �亯�Դϴ�."
    call sendmail("customer@10x10.co.kr", request.form("mail"),mailtitle,mailcontent)
end if

'����Ÿ���̽��� �Է�

dim sqlput,sqlput1

sqlput = "update "+table_name+" set pos=pos+1 where pos >= "&pos
rsput.Open sqlput,dbput,1

dim user_file,user_ip,face

sqlput1 = " insert into "+table_name+" (site_name,name,mail,title,body, "
sqlput1 = sqlput1 + " count,reg_date,thread,depth,pos,check_flag) "
sqlput1 = sqlput1 + " values( '"+site_name+"','10x10', '" + mail + "', "
sqlput1 = sqlput1 + " '" + title + "', '" + body + "', "
sqlput1 = sqlput1 + " 1,getdate(),'"+thread+"','"+depth+"', '"+pos+"','Y' ) "

rsput.Open sqlput1,dbput,1

sqlput1 = " update "+table_name+" set check_flag = 'Y' where site_name = '"+site_name+"' and thread = "+thread

rsput.Open sqlput1,dbput,1

response.redirect "boardlist.asp?table_name="+table_name+"&site_name="+site_name

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
