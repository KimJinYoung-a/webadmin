<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/bct_admin_header.asp"-->
<%
const MenuPos1 = "Admin"
const MenuPos2 = "매출 분석"
%>
<!-- #include virtual="/admin/bct_admin_menupos.asp"-->

<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" WIDTH=580 HEIGHT=220>
<PARAM NAME=movie VALUE="/graph/get_graph_day.swf?pMaxHit=4726&pMaxAvg=6051&y=2002&m=10&d=28&pHitCnt0=4726&pAvgCnt0=5484&pHitCnt1=4257&pAvgCnt1=4024&pHitCnt2=2673&pAvgCnt2=2498&pHitCnt3=654&pAvgCnt3=1409&pHitCnt4=491&pAvgCnt4=834&pHitCnt5=437&pAvgCnt5=542&pHitCnt6=0&pAvgCnt6=411&pHitCnt7=0&pAvgCnt7=513&pHitCnt8=0&pAvgCnt8=1229&pHitCnt9=0&pAvgCnt9=3288&pHitCnt10=0&pAvgCnt10=4670&pHitCnt11=0&pAvgCnt11=5329&pHitCnt12=0&pAvgCnt12=4567&pHitCnt13=0&pAvgCnt13=5396&pHitCnt14=0&pAvgCnt14=5602&pHitCnt15=0&pAvgCnt15=5393&pHitCnt16=0&pAvgCnt16=5482&pHitCnt17=0&pAvgCnt17=5257&pHitCnt18=0&pAvgCnt18=4259&pHitCnt19=0&pAvgCnt19=3644&pHitCnt20=0&pAvgCnt20=4307&pHitCnt21=0&pAvgCnt21=5368&pHitCnt22=0&pAvgCnt22=6051&pHitCnt23=0&pAvgCnt23=5940">
<PARAM NAME=loop VALUE=false>
<PARAM NAME=menu VALUE=false>
<PARAM NAME=quality VALUE=high>
<PARAM NAME=bgcolor VALUE=#FFFFFF>
<PARAM NAME=wmode VALUE=Transparent>
<EMBED src="graph/get_graph_day.swf?pMaxHit=4726&pMaxAvg=6051&y=2002&m=10&d=28&pHitCnt0=4726&pAvgCnt0=5484&pHitCnt1=4257&pAvgCnt1=4024&pHitCnt2=2673&pAvgCnt2=2498&pHitCnt3=654&pAvgCnt3=1409&pHitCnt4=491&pAvgCnt4=834&pHitCnt5=437&pAvgCnt5=542&pHitCnt6=0&pAvgCnt6=411&pHitCnt7=0&pAvgCnt7=513&pHitCnt8=0&pAvgCnt8=1229&pHitCnt9=0&pAvgCnt9=3288&pHitCnt10=0&pAvgCnt10=4670&pHitCnt11=0&pAvgCnt11=5329&pHitCnt12=0&pAvgCnt12=4567&pHitCnt13=0&pAvgCnt13=5396&pHitCnt14=0&pAvgCnt14=5602&pHitCnt15=0&pAvgCnt15=5393&pHitCnt16=0&pAvgCnt16=5482&pHitCnt17=0&pAvgCnt17=5257&pHitCnt18=0&pAvgCnt18=4259&pHitCnt19=0&pAvgCnt19=3644&pHitCnt20=0&pAvgCnt20=4307&pHitCnt21=0&pAvgCnt21=5368&pHitCnt22=0&pAvgCnt22=6051&pHitCnt23=0&pAvgCnt23=5940" loop=false menu=false quality=high bgcolor=#FFFFFF wmode=Transparent  WIDTH=580 HEIGHT=220 TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"></EMBED>
</OBJECT>

<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" WIDTH=580 HEIGHT=220>
<PARAM NAME=movie VALUE="/graph/get_graph_month.swf?pMaxHit=128368&pMaxAvg=106331&pMonthLen=31&y=2002&m=10&d=05&pHitCnt0=88525&pAvgCnt0=35704&pHitCnt1=83940&pAvgCnt1=0&pHitCnt2=67743&pAvgCnt2=0&pHitCnt3=72760&pAvgCnt3=41779&pHitCnt4=65807&pAvgCnt4=38223&pHitCnt5=59151&pAvgCnt5=0&pHitCnt6=98819&pAvgCnt6=0&pHitCnt7=103517&pAvgCnt7=28488&pHitCnt8=127269&pAvgCnt8=46985&pHitCnt9=110787&pAvgCnt9=45431&pHitCnt10=128368&pAvgCnt10=42622&pHitCnt11=90651&pAvgCnt11=47906&pHitCnt12=76881&pAvgCnt12=42718&pHitCnt13=98645&pAvgCnt13=33289&pHitCnt14=73123&pAvgCnt14=32131&pHitCnt15=97129&pAvgCnt15=44757&pHitCnt16=91910&pAvgCnt16=31038&pHitCnt17=91800&pAvgCnt17=39408&pHitCnt18=99552&pAvgCnt18=32948&pHitCnt19=83864&pAvgCnt19=25860&pHitCnt20=95589&pAvgCnt20=64341&pHitCnt21=98664&pAvgCnt21=74154&pHitCnt22=108440&pAvgCnt22=82737&pHitCnt23=114599&pAvgCnt23=99641&pHitCnt24=120578&pAvgCnt24=85818&pHitCnt25=88194&pAvgCnt25=98812&pHitCnt26=72681&pAvgCnt26=96876&pHitCnt27=13320&pAvgCnt27=106331&pHitCnt28=0&pAvgCnt28=90686&pHitCnt29=0&pAvgCnt29=102856&pHitCnt30=0&pAvgCnt30=47996">
<PARAM NAME=loop VALUE=false>
<PARAM NAME=menu VALUE=false>
<PARAM NAME=quality VALUE=high>
<PARAM NAME=bgcolor VALUE=#FFFFFF>
<PARAM NAME=wmode VALUE=Transparent>
<EMBED src="/graph/get_graph_month.swf?pMaxHit=128368&pMaxAvg=106331&pMonthLen=31&y=2002&m=10&d=05&pHitCnt0=88525&pAvgCnt0=35704&pHitCnt1=83940&pAvgCnt1=0&pHitCnt2=67743&pAvgCnt2=0&pHitCnt3=72760&pAvgCnt3=41779&pHitCnt4=65807&pAvgCnt4=38223&pHitCnt5=59151&pAvgCnt5=0&pHitCnt6=98819&pAvgCnt6=0&pHitCnt7=103517&pAvgCnt7=28488&pHitCnt8=127269&pAvgCnt8=46985&pHitCnt9=110787&pAvgCnt9=45431&pHitCnt10=128368&pAvgCnt10=42622&pHitCnt11=90651&pAvgCnt11=47906&pHitCnt12=76881&pAvgCnt12=42718&pHitCnt13=98645&pAvgCnt13=33289&pHitCnt14=73123&pAvgCnt14=32131&pHitCnt15=97129&pAvgCnt15=44757&pHitCnt16=91910&pAvgCnt16=31038&pHitCnt17=91800&pAvgCnt17=39408&pHitCnt18=99552&pAvgCnt18=32948&pHitCnt19=83864&pAvgCnt19=25860&pHitCnt20=95589&pAvgCnt20=64341&pHitCnt21=98664&pAvgCnt21=74154&pHitCnt22=108440&pAvgCnt22=82737&pHitCnt23=114599&pAvgCnt23=99641&pHitCnt24=120578&pAvgCnt24=85818&pHitCnt25=88194&pAvgCnt25=98812&pHitCnt26=72681&pAvgCnt26=96876&pHitCnt27=13320&pAvgCnt27=106331&pHitCnt28=0&pAvgCnt28=90686&pHitCnt29=0&pAvgCnt29=102856&pHitCnt30=0&pAvgCnt30=47996" loop=false menu=false quality=high bgcolor=#FFFFFF wmode=Transparent  WIDTH=580 HEIGHT=220 TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"></EMBED>

</OBJECT>

<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" WIDTH=580 HEIGHT=220>
<PARAM NAME=movie VALUE="/graph/get_graph_year.swf?pMaxHit=2522284&pMaxAvg=0&y=2002&m=01&d=01&pHitCnt0=0&pAvgCnt0=0&pHitCnt1=0&pAvgCnt1=0&pHitCnt2=0&pAvgCnt2=0&pHitCnt3=0&pAvgCnt3=0&pHitCnt4=0&pAvgCnt4=0&pHitCnt5=0&pAvgCnt5=0&pHitCnt6=0&pAvgCnt6=0&pHitCnt7=1097362&pAvgCnt7=0&pHitCnt8=2333189&pAvgCnt8=0&pHitCnt9=2522284&pAvgCnt9=0&pHitCnt10=0&pAvgCnt10=0&pHitCnt11=0&pAvgCnt11=0">
<PARAM NAME=loop VALUE=false>
<PARAM NAME=menu VALUE=false>
<PARAM NAME=quality VALUE=high>
<PARAM NAME=bgcolor VALUE=#FFFFFF>
<PARAM NAME=wmode VALUE=Transparent>
<EMBED src="graph/get_graph_year.swf?pMaxHit=2522284&pMaxAvg=0&y=2002&m=01&d=01&pHitCnt0=0&pAvgCnt0=0&pHitCnt1=0&pAvgCnt1=0&pHitCnt2=0&pAvgCnt2=0&pHitCnt3=0&pAvgCnt3=0&pHitCnt4=0&pAvgCnt4=0&pHitCnt5=0&pAvgCnt5=0&pHitCnt6=0&pAvgCnt6=0&pHitCnt7=1097362&pAvgCnt7=0&pHitCnt8=2333189&pAvgCnt8=0&pHitCnt9=2522284&pAvgCnt9=0&pHitCnt10=0&pAvgCnt10=0&pHitCnt11=0&pAvgCnt11=0" loop=false menu=false quality=high bgcolor=#FFFFFF wmode=Transparent  WIDTH=580 HEIGHT=220 TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"></EMBED>
</OBJECT>
<!-- #include virtual="/admin/bct_admin_tail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
