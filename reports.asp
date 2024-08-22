<%@ Language=VBScript %>
<%
Response.Expires = -1
'gscore=0
%>

<!--#include file="xfunctions.asp" -->

<%
   if ulev="" then Response.Redirect("default.asp")
Session.Timeout=30 %>
<%
reprthdr

'get all form values
srch=getreq("srch")

if srch = "1" then 
pt=getreq("pt")
ept=getreq("ept")   
lwf=getreq("lwf")
pf=getreq("pf")
esi=getreq("esi")
GST=getreq("GST")    
IT=getreq("IT")
TCA=getreq("TCA")
registration=getreq("registration")
returns=getreq("returns")
payrolls=getreq("payrolls")
    
else
pt="on"
ept="on"
lwf="on"
pf="on"
esi="on"
GST="on"
IT="on"
TCA="on"
registration="on"
returns="on"        
payrolls="on"

end if

tpa=getreq("tpa")

tpp=getreq("tpp")
status=getreq("status")
if tpp = "" then tpp = 1
scyear = getreq("scyear") 'get scope year
if scyear = "" then scyear = year(now)
scperiod = getreq("scperiod") 'get scope period
if scperiod = "" then scperiod = 0
uolist=getreq("uolist") 'get scope organisation
ualist=getreq("ualist") 'get scope location
stlist=getreq("stlist") 'get scope state
rglist=getreq("rglist") 'get scope region
styear=year(now)-3
nxtyear=year(now)+1

   

tpod=getreq("tpod")       'get type of overdue item
if tpod = "" then tpod = 1

'


'generate location query string
locsql=""
if uolist <> "" then locsql = locsql&" and c.oid = '"&uolist&"' "
if ualist <> "" then locsql = locsql&" and c.lcode = '"&ualist&"' "
if stlist <> "" then locsql = locsql&" and c.lstate = '"&stlist&"' "
if rglist  <> "" then locsql = locsql&" and c.lregion = '"&rglist&"' "

if getreq("tpa") = "adap" then
atit=getreq("title")
oid=getreq("oid")
lcode=getreq("lcode")
aclink=getreq("aclink")

actp="A"
isq="insert into ncaction (oid, aclink, lcode,actp,actitle,acshow,acstatus,acistatus,acidate,acrdate,acruser) "
isq=isq&" values ('"&oid&"','"&aclink&"','"&lcode&"','"&actp&"','"&atit&"','0','O','N',getdate(),getdate(),"&uno&")"
conndb.execute(isq)
response.redirect("reports.asp?"&urladd)

end if
     'response.Write("hiii"&cyear)
    urladd="&scyear="&scyear&"&scperiod="&scperiod&"&uolist="&uolist&"&ualist="&ualist&"&stlist="&stlist&"&rglist="&rglist
%>
  
       
          
<div class="app-main">
                 <div class="app-main__outer">
                    <div class="app-main__inner">
                        <div class="app-page-title">
                            <div class="page-title-wrapper">
                                <div class="page-title-heading">
                                    <div class="page-title-icon">
                                        <i class="pe-7s-car icon-gradient bg-mean-fruit">
                                        </i>
                                    </div>      
                                    <div>Compliance Dashboard
                                          <a href="#" class="btn btn-warning m-btn m-btn--icon m-btn--wide m-btn--md m--margin-right-10" data-toggle="modal" data-target="#m_modal_1_2">
                                                  Filter
                                                </a>
                                        <div class="page-title-subheading">Scope of Service and Status of Labour Compliance.
                                        </div>
                                    </div>
                                </div>
                                <div class="page-title-actions">
                                
                                </div>    </div>
                        </div>      

<!--     CONTRIBUTION-->                                        
                                                
                        <% 
                           coscore=0  'contribution score
cotot=0    'no of items totally
COTOTAL=0    'no of co and upcoming
cototc=0   'no of items compliant
cototnc=0  'no of items non compliant

asqC1="select count(*) from nccontr a, ncmorg b, ncmloc c  where a.lcode = c.lcode "
asqC2="select count(*) from nccontr a, ncmorg b, ncmloc c  where a.lcode = c.lcode and status in (1,3,4) " 
asqC4="select count(*) from nccontr a, ncmorg b, ncmloc c  where a.lcode = c.lcode and (status = 1 OR STATUS = 0) " 
asqC3="select a.contid, a.oid ,a.lcode, tp, period,cyear,lastdate,status, b.oname, c.lname from nccontr a, ncmorg b, ncmloc c where contid <> 0 "
tsqc=tsqc&" and (status <> 0 or (status = 0 and datediff(day,GETDATE(),lastdate)  <= 30))"
tsqc=tsqc&" and a.oid = b.oid and a.lcode = c.lcode and b.oid = c.oid "

tsqc=tsqc&" and a.lcode in  (select lcode from ncumap where oid = a.oid and uno = "&uno&") and c.lactive='1' "
if locsql <> "" then tsqc=tsqc&locsql

if scyear <> 0 then tsqc=tsqc&" and cyear = "&scyear
if scperiod <> "0" then
 if scperiod = "H1" then tsqc=tsqc&" and period <= 6"
 if scperiod = "H2" then tsqc=tsqc&" and period > 6"
 if scperiod = "Q1" then tsqc=tsqc&" and period between 1 and 3"
 if scperiod = "Q2" then tsqc=tsqc&" and period between 4 and 6"
 if scperiod = "Q3" then tsqc=tsqc&" and period between 7 and 9"
 if scperiod = "Q4" then tsqc=tsqc&" and period between 10 and 12"
 if scperiod = "M1" then tsqc=tsqc&" and period = 1"
  if scperiod = "M2" then tsqc=tsqc&" and period = 2"
  if scperiod = "M3" then tsqc=tsqc&" and period = 3"
  if scperiod = "M4" then tsqc=tsqc&" and period = 4"
  if scperiod = "M5" then tsqc=tsqc&" and period = 5"
  if scperiod = "M6" then tsqc=tsqc&" and period = 6"
  if scperiod = "M7" then tsqc=tsqc&" and period = 7"
  if scperiod = "M8" then tsqc=tsqc&" and period = 8"
  if scperiod = "M9" then tsqc=tsqc&" and period = 9"
  if scperiod = "M10" then tsqc=tsqc&" and period = 10"
  if scperiod = "M11" then tsqc=tsqc&" and period = 11"
  if scperiod = "M12" then tsqc=tsqc&" and period = 12"


end if

asqC1=asqC1&tsqc
asqC2=asqC2&tsqc
asqC4=asqC4&tsqc
asqC3=asqC3&tsqc&" order by lastdate desc "
    rs.open asqC1,conndb
                           '  response.write(asqc1)
cotot=rs(0)
rs.close

rs.open asqC2,conndb
                              'response.write(asqc2)
cototc=rs(0)
rs.close


rs.open asqC4,conndb
                             ' response.write(asqc4)
cototal=rs(0) 
rs.close

cototnc=cotot-cotoc

if cotot > 0 then 
                            coscore=round(100*cototc/cotot,2) 
                           
                            else 
                            coscore = 0    
                            bluecont=1
                            end if
                                         
   %>                         
                            
        <!--     REGISTRATION--> 

<%rescore=0  'registrations score
rtot=0    'no of items totally
rtotal=0
rtotup=0
rtotalcom=0
rtotc=0   'no of items compliant
rtotnc=0  'no of items non compliant

asqR1="select count(*) from ncreg a, ncmorg b, ncmloc c where a.lcode = c.lcode  and c.lactive='1' "
asqR2="select count(*) from ncreg a, ncmorg b, ncmloc c where a.lcode = c.lcode and  status  in ('c','a','d','b','ar','aa','su','sc','na','up') and c.lactive='1' "
 asqR9="select count(*) as cnt from ncreg a, ncmorg b, ncmloc c where  (a.doe>a.doi) and a.lcode = c.lcode  and a.oid = b.oid and a.lcode = c.lcode and b.oid = c.oid and a.lcode in (select lcode from ncumap where oid = a.oid and uno = "&uno&") and c.lactive='1' "

asqR12="select count(*) as cnt from ncreg a, ncmorg b, ncmloc c where  (datediff(day,GETDATE(),doe) <=45 ) and a.lcode = c.lcode  and a.oid = b.oid and a.lcode = c.lcode and b.oid = c.oid and a.lcode in (select lcode from ncumap where oid = a.oid and uno = "&uno&") and c.lactive='1' "

asqR3="select a.uid, a.oid ,a.lcode, tp, doe, status, b.oname, c.lname from ncreg a, ncmorg b, ncmloc c where status <> 'XX' "
tsqR=tsqR&" and a.oid = b.oid and a.lcode = c.lcode and a.oid = c.oid"


tsqR=tsqR&" and a.lcode in  (select lcode from ncumap where oid = a.oid and uno = "&uno&")  "
if locsql <> "" then tsqR=tsqR&locsql
'if scyear <> 0 then tsqR=tsqR&" and ( (year(a.doi) = "&scyear&" or year(doe) = "&scyear&") or a.doi is null or year(a.doi) = 1900)"
if scperiod <> "0" then
 if scperiod = "H1" then tsqR=tsqR&" and month(doe) <= 6"
 if scperiod = "H2" then tsqR=tsqR&" and month(doe) > 6"
 if scperiod = "Q1" then tsqR=tsqR&" and month(doe) between 1 and 3"
 if scperiod = "Q2" then tsqR=tsqR&" and month(doe) between 4 and 6"
 if scperiod = "Q3" then tsqR=tsqR&" and month(doe) between 7 and 9"
 if scperiod = "Q4" then tsqR=tsqR&" and month(doe) between 10 and 12"
  if scperiod = "M1" then tsqR=tsqR&" and month(doe) = 1"
  if scperiod = "M2" then tsqR=tsqR&" and month(doe) = 2"
  if scperiod = "M3" then tsqR=tsqR&" and month(doe) = 3"
  if scperiod = "M4" then tsqR=tsqR&" and month(doe) = 4"
  if scperiod = "M5" then tsqR=tsqR&" and month(doe) = 5"
  if scperiod = "M6" then tsqR=tsqR&" and month(doe) = 6"
  if scperiod = "M7" then tsqR=tsqR&" and month(doe) = 7"
  if scperiod = "M8" then tsqR=tsqR&" and month(doe) = 8"
  if scperiod = "M9" then tsqR=tsqR&" and month(doe) = 9"
  if scperiod = "M10" then tsqR=tsqR&" and month(doe) = 10"
  if scperiod = "M11" then tsqR=tsqR&" and month(doe) = 11"
  if scperiod = "M12" then tsqR=tsqR&" and month(doe) = 12"

end if

asqR1=asqR1&tsqR
asqR2=asqR2&tsqR
asqR9=asqR9&tsqR
asqR12=asqR12&tsqR
asqR3=asqR3&tsqR&" order by doe asc, lname "
'response.write(asq3)
    ' response.write(asqR9)
     ' response.write(asqR12)
rs.open asqR1,conndb
  
rtot=rs(0)
rs.close

rs.open asqR2,conndb
rtotc=rs(0)
rs.close



rs.open asqR9,conndb
rtotalcom=rs(0)
rs.close

rs.open asqR12,conndb
rtotup=rs(0)
rs.close

rtotal=rtotalcom+rtotup

rtotnc=rtot-rtoc

    
if rtot > 0 then 
    rescore=round(100*rtotc/rtot,2) 
    else 
    rescore = 0
    bluereg=1
    end if
  '  response.Write("hiii"&rescore)
'gscore=44.4
    %>                                              
             
                                                
      <!--     RETURNS-->                                            
                                                
                                                                
     <%RTscore=0  'returns score
RTtot=0    'no of items totally
RTtotal=0

RTtotc=0   'no of items compliant
RTtotnc=0  'no of items non compliant

asqRT1="select count(*) from ncret a, ncmorg b, ncmloc c, nctempret d  where a.lcode = c.lcode  "
asqRT2="select count(*) from ncret a, ncmorg b, ncmloc c, nctempret d where a.lcode = c.lcode and status in( 1,3,4) "
asqRT4="select count(*) from ncret a, ncmorg b, ncmloc c, nctempret d where a.lcode = c.lcode and (status = 1 or status = 0) "
asqRT3="select a.rtid, a.oid ,a.lcode, a.rcode,  ryear ,lastdate,status, b.oname, c.lname, d.rform, d.rtitle from ncret a, ncmorg b, ncmloc c, nctempret d where status <> 99 "
tsqRT=tsqRT&" and a.oid = b.oid and a.oid = c.oid and a.lcode = c.lcode and a.rcode = d.rcode"
'asq3=asq3&" and (status <> 0 or status = 0 and (lastdate-getdate() <= 45))"

'tsqRT=tsqRT&" and (status <> 0 or (status = 0 and datediff(day,GETDATE(),lastdate)  <= 30))"



tsqRT=tsqRT&" and a.lcode in  (select lcode from ncumap where oid = a.oid and uno = "&uno&") and c.lactive='1' "
if locsql <> "" then tsqRT=tsqRT&locsql

if scyear <> 0 then tsqRT=tsqRT&" and ryear = "&scyear

if scperiod <> "0" then
 if scperiod = "H1" then tsqRT=tsqRT&" and month(lastdate) <= 6"
 if scperiod = "H2" then tsqRT=tsqRT&" and month(lastdate) > 6"
 if scperiod = "Q1" then tsqRT=tsqRT&" and month(lastdate) between 1 and 3"
 if scperiod = "Q2" then tsqRT=tsqRT&" and month(lastdate) between 4 and 6"
 if scperiod = "Q3" then tsqRT=tsqRT&" and month(lastdate) between 7 and 9"
 if scperiod = "Q4" then tsqRT=tsqRT&" and month(lastdate) between 10 and 12"
  if scperiod = "M1" then tsqRT=tsqRT&" and month(lastdate) = 1"
  if scperiod = "M2" then tsqRT=tsqRT&" and month(lastdate) = 2"
  if scperiod = "M3" then tsqRT=tsqRT&" and month(lastdate) = 3"
  if scperiod = "M4" then tsqRT=tsqRT&" and month(lastdate) = 4"
  if scperiod = "M5" then tsqRT=tsqRT&" and month(lastdate) = 5"
  if scperiod = "M6" then tsqRT=tsqRT&" and month(lastdate) = 6" 
  if scperiod = "M7" then tsqRT=tsqRT&" and month(lastdate) = 7"
  if scperiod = "M8" then tsqRT=tsqRT&" and month(lastdate) = 8"
  if scperiod = "M9" then tsqRT=tsqRT&" and month(lastdate) = 9"
  if scperiod = "M10" then tsqRT=tsqRT&" and month(lastdate) = 10"
  if scperiod = "M11" then tsqRT=tsqRT&" and month(lastdate) = 11"
  if scperiod = "M12" then tsqRT=tsqRT&" and month(lastdate) = 12"
end if

asqRT1=asqRT1&tsqRT
asqRT2=asqRT2&tsqRT
asqRT4=asqRT4&tsqRT
asqRT3=asqRT3&tsqRT&" order by lastdate desc "
'response.write(asqRT1)
'response.write(asqRT2)
'response.write(asq3)
rs.open asqRT1,conndb
RTtot=rs(0)
rs.close

rs.open asqRT2,conndb
RTtotc=rs(0)
rs.close

rs.open asqRT4,conndb
RTtotal=rs(0)
rs.close

RTtotnc=RTtot-RTtoc
if RTtot > 0 then 
         RTscore=round(100*RTtotc/RTtot,2) 
         else
         RTscore = 0
         blue=1
         end if
         

         %>      
                                                
                                                                 
 







                        <div class="row">
                            <div class="col-md-6 col-lg-4">
                                <div class="card-shadow-danger mb-3 widget-chart widget-chart2 text-left card">
                                    <div class="widget-content">
                                        <div class="widget-content-outer">
                                            <div class="widget-content-wrapper">
                                                <div class="widget-content-left pr-2 fsize-1">
                                                    
                                                    <%if  bluecont=1 or coscore<=1 then %>
                                                    <div class="widget-numbers mt-0 fsize-3 text-info"><%=coscore%>%</div>
                                                      <%end if %>
                                                     <%if coscore>1 and coscore<=50 then %>
                                                    <div class="widget-numbers mt-0 fsize-3 text-danger"><%=coscore%>%</div>
                                                      <%end if %>
                                                     <%if coscore>50 and coscore<=75 then%>
                                                    <div class="widget-numbers mt-0 fsize-3 text-warning"><%=coscore%>%</div>
                                                      <%end if %>
                                                    <%if coscore>75  then %>
                                                    <div class="widget-numbers mt-0 fsize-3 text-success"><%=coscore%>%</div>
                                                      <%end if %>
                                                
                                                </div>
                                                <div class="widget-content-right w-100">
                                                    <div class="progress-bar-xs progress">
                                                        <%if coscore<=50 then %>
                                                    <div class="progress-bar bg-danger"role="progressbar" aria-valuenow="<%=coscore%>" aria-valuemin="0" aria-valuemax="100" style="width: <%=coscore%>%;"></div>
                                                      <%end if %>
                                                     <%if coscore>50 and coscore<=75 then%>
                                                    <div class="progress-bar bg-warning"role="progressbar" aria-valuenow="<%=coscore%>" aria-valuemin="0" aria-valuemax="100" style="width: <%=coscore%>%;"></div>
                                                      <%end if %>
                                                   
                                                    <%if coscore>75 then%>
                                                    <div class="progress-bar bg-success"role="progressbar" aria-valuenow="<%=coscore%>" aria-valuemin="0" aria-valuemax="100" style="width: <%=coscore%>%;"></div>
                                                    <%end if %>
                                                         
                                                        
                                                    </div>
                                                </div>
                                            </div>
                                            <div class="widget-content-left fsize-1">
                                                <div class="text-muted opacity-6"><b>Payroll Health Index</b></div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-6 col-lg-4">
                                <div class="card-shadow-success mb-3 widget-chart widget-chart2 text-left card">
                                    <div class="widget-content">
                                        <div class="widget-content-outer">
                                            <div class="widget-content-wrapper">
                                                <div class="widget-content-left pr-2 fsize-1">
                                                    <%if  blue=1 or RTscore<=1 then %>
                                                    <div class="widget-numbers mt-0 fsize-3 text-info"><%=RTscore%>%</div>
                                                      <%end if %>
                                                    <%if RTscore>1 and RTscore<=50 then %>
                                                    <div class="widget-numbers mt-0 fsize-3 text-danger"><%=RTscore%>%</div>
                                                      <%end if %>
                                                     <%if RTscore>50 and RTscore<=75 then%>
                                                    <div class="widget-numbers mt-0 fsize-3 text-warning"><%=RTscore%>%</div>
                                                      <%end if %>
                                                   
                                                    <%if RTscore>75 then%>
                                                    <div class="widget-numbers mt-0 fsize-3 text-success"><%=RTscore%>%</div>
                                                    <%end if %>
                                                </div>
                                                <div class="widget-content-right w-100">
                                                    <div class="progress-bar-xs progress">
  <%if RTscore<=50 then %>
                                                    <div class="progress-bar bg-danger"role="progressbar" aria-valuenow="<%=RTscore%>" aria-valuemin="0" aria-valuemax="100" style="width: <%=RTscore%>%;"></div>
                                                      <%end if %>
                                                     <%if RTscore>50 and RTscore<=75 then%>
                                                    <div class="progress-bar bg-warning"role="progressbar" aria-valuenow="<%=RTscore%>" aria-valuemin="0" aria-valuemax="100" style="width: <%=RTscore%>%;"></div>
                                                      <%end if %>
                                                   
                                                    <%if RTscore>75 then%>
                                                    <div class="progress-bar bg-success"role="progressbar" aria-valuenow="<%=RTscore%>" aria-valuemin="0" aria-valuemax="100" style="width: <%=RTscore%>%;"></div>
                                                    <%end if %>                                                    </div>
                                                </div>
                                            </div>
                                            <div class="widget-content-left fsize-1">
                                                <div class="text-muted opacity-6"><b>Returns Health Index</b></div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-6 col-lg-4">
                                <div class="card-shadow-warning mb-3 widget-chart widget-chart2 text-left card">
                                    <div class="widget-content">
                                        <div class="widget-content-outer">
                                            <div class="widget-content-wrapper">
                                                <div class="widget-content-left pr-2 fsize-1">
                                                    <%if  bluereg=1 or rescore <=1 then %>
                                                    <div class="widget-numbers mt-0 fsize-3 text-info"><%=rescore%>%</div>
                                                      <%end if %>  
                                                    <%if rescore>1 and rescore<=50 then %>
                                                    <div class="widget-numbers mt-0 fsize-3 text-danger"><%=rescore%>%</div>
                                                      <%end if %>
                                                    <%if rescore>50 and rescore<=75 then %>
                                                    <div class="widget-numbers mt-0 fsize-3 text-warning"><%=rescore%>%</div>
                                                      <%end if %>
                                                  
                                                    <%if rescore>75 then%>
                                                    <div class="widget-numbers mt-0 fsize-3 text-success"><%=rescore%>%</div>
                                                    <%end if %>
                                                </div>
                                                <div class="widget-content-right w-100">
                                                    <div class="progress-bar-xs progress">
  <%if rescore<=50 then %>
                                                    <div class="progress-bar bg-danger"role="progressbar" aria-valuenow="<%=rescore%>" aria-valuemin="0" aria-valuemax="100" style="width: <%=rescore%>%;"></div>
                                                      <%end if %>
                                                     <%if rescore>50 and rescore<=75 then %>
                                                    <div class="progress-bar bg-warning"role="progressbar" aria-valuenow="<%=rescore%>" aria-valuemin="0" aria-valuemax="100" style="width: <%=rescore%>%;"></div>
                                                      <%end if %>
                                                   
                                                    <%if rescore>75 then%>
                                                    <div class="progress-bar bg-success"role="progressbar" aria-valuenow="<%=rescore%>" aria-valuemin="0" aria-valuemax="100" style="width: <%=rescore%>%;"></div>
                                                    <%end if %>                                                    </div>
                                                </div>
                                            </div>
                                            <div class="widget-content-left fsize-1">
                                                <div class="text-muted opacity-6"><b>Registrations Health Index</b></div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            
                              
                        </div>  





                        <div class="row">
                            <div class="col-md-12 col-lg-6">
                                <div class="mb-3 card">
                                    <div class="card-header-tab card-header-tab-animation card-header">
                                        <div class="card-header-title">
                                            <i class="header-icon lnr-apartment icon-gradient bg-love-kiss"> </i>
                                            Compliance Report
                                          
                                        </div>
                                        <ul class="nav">
                                            <!--<li class="nav-item"><a href="javascript:void(0);" class="active nav-link">Last</a></li>
                                            <li class="nav-item"><a href="javascript:void(0);" class="nav-link second-tab-toggle">Current</a></li>-->
                                        </ul>
                                    </div>
                                    <div class="card-body">
                                        <div class="tab-content">
                                            <div class="tab-pane fade show active" id="tabs-eg-77">
                                                <div class="card mb-3 widget-chart widget-chart2 text-left w-100">
                                                    <div class="widget-chat-wrapper-outer">
                                                        <div class="widget-chart-wrapper widget-chart-wrapper-lg opacity-10 m-0">
                                                            <canvas id="stacked-bars-chart1"></canvas>
                                                            
                                                        </div>
                                                    </div>
                                                </div>
                                                <!--<h6 class="text-muted text-uppercase font-size-md opacity-5 font-weight-normal">Top Authors</h6>-->
                                                <div class="card-header-tab card-header-tab-animation card-header">
                                                <div class="card-header-title">
                                            <i class="header-icon lnr-apartment icon-gradient bg-love-kiss"> </i>
                                            Completion Details
                                        </div>
                                                    </div>
                                               
           

                                                <div class="scroll-area-sm">
                                                    <div class="scrollbar-container">
                                                        <ul class="rm-list-borders rm-list-borders-scroll list-group list-group-flush">
                                                            <%if payrolls = "on" then %>
                                                            <li class="list-group-item">
                                                                <div class="widget-content p-0">
                                                                    <div class="widget-content-wrapper">
                                                                        <div class="widget-content-left mr-3">
                                                                            <i class="btn rounded-circle btn-info text-white"><b>P</b></i>
                                                                        </div>
                                                                        <div class="widget-content-left">
                                                                            <div class="widget-heading">Payroll</div>
                                                                            <div class="widget-subheading">PF, ESI, PT, LWF</div>
                                                                        </div>
                                                                        <div class="widget-content-right">
                                                                            <div class="font-size-xlg text-muted">
                                                                                <span><%=cototal%></span>
                                                                                <button type="button" aria-expanded="true" aria-controls="payAccordion" data-toggle="collapse" href="#payroll" class="m-0 p-0 btn btn-link"> 
                                                                                    <strong class="text-success">
                                                                                                                  <i class="btn rounded-circle fa fa-angle-down bg-success"></i>
                                                        </strong>

                                                                                </button>
                                                                                <button type="button" aria-expanded="true" aria-controls="payAccordion" data-toggle="collapse" href="#payrollup" class="m-0 p-0 btn btn-link"> 
                                                                                    <strong class="text-success">
                                                                                                                  <i class="btn rounded-circle fa fa-angle-down bg-info"></i>
                                                        </strong>

                                                                                </button>
                                                                                  <!--  <a  data-toggle="collapse" data-target="#collapseOne1" aria-expanded="true" aria-controls="collapseOne" class="text-left m-0 p-0 btn btn-link btn-block">
                                                   
                                                </a>-->
                                                       
                                                                                
                                                                            </div>
                                                                        </div>
                                                                    </div>
                                                                    
                                                                </div>
                                                                <div id="payAccordion" data-children=".item">
                                                <div class="item">
                                                    
                                                    <div data-parent="#payAccordion" id="payroll" class="collapse">
                                                        <table class="mb-0 table table-bordered table-hover table-responsive">
                                            <thead>
                                                <tr><th colspan="5"  class="text-center">Deposited/Filed 
                                        </th></tr>
                                            <tr>
                                                <th>Site</th>
                                                <th>Type</th>
                                                <th>Period</th>
                                                <th>Due Date</th>
                                               
                                                <th>Deposit Date</th>
                                            </tr>
                                            </thead>
                                            <tbody>
                                            <%showcontr(1)%>
                                               
                                            </tbody>
                                        </table>
                                                    </div>
                                                </div>
                                                <div class="item">
                                                    
                                                    <div data-parent="#payAccordion" id="payrollup" class="collapse">
                                                        <table class="mb-0 table table-bordered table-hover table-responsive">
                                            <thead>
                                                <tr><th colspan="4"  class="text-center">Upcoming 
                                        </th></tr>
                                            <tr>
                                                <th>Site</th>
                                                <th>Type</th>
                                                <th>Period</th>
                                                <th>Due Date</th>
                                               
                                                
                                            </tr>
                                            </thead>
                                                           
                                            <tbody>
                                          
                                               <tr>
                                                    <%showcontr(2)%>
                                                   
                                                </tr>
                                            
                                            
                                            </tbody>
                                                              
                                        </table>
                                                    </div>
                                                </div>
                                            </div>  
                                                            </li>
                                                            <%end if %>
                                                             <%if returns = "on" then %>
                                                            <li class="list-group-item">
                                                                <div class="widget-content p-0">
                                                                    <div class="widget-content-wrapper">
                                                                        <div class="widget-content-left mr-3">
                                                                            <i class="btn rounded-circle btn-info text-white"><b>L</b></i>
                                                                            <!--<img width="42" class="rounded-circle" src="assets/images/avatars/5.jpg" alt="">-->
                                                                        </div>
                                                                        <div class="widget-content-left">
                                                                            <div class="widget-heading">Liasion</div>
                                                                            <div class="widget-subheading">Returns, Inspections, Notices</div>
                                                                        </div>
                                                                        <div class="widget-content-right">
                                                                            <div class="font-size-xlg text-muted">
                                                                                <!--<small class="opacity-5 pr-1">$</small>-->
                                                                                <span><%=RTtotal%></span>
                                                                                <button type="button" aria-expanded="true" aria-controls="liasAccordion" data-toggle="collapse" href="#lias" class="m-0 p-0 btn btn-link"> 
                                                                                    <strong class="text-success">
                                                                                                                   <i class="btn rounded-circle fa fa-angle-down bg-success"></i>
                                                        </strong>

                                                                                </button>
                                                                                 <button type="button" aria-expanded="true" aria-controls="liasAccordion" data-toggle="collapse" href="#liasup" class="m-0 p-0 btn btn-link"> 
                                                                                    <strong class="text-success">
                                                                                                                   <i class="btn rounded-circle fa fa-angle-down bg-info"></i>
                                                        </strong>

                                                                                </button>
                                                                            </div>
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                <div id="liasAccordion" data-children=".item">
                                                <div class="item">
                                                    
                                                    <div data-parent="#liasAccordion" id="lias" class="collapse">
                                                         <table class="mb-0 table table-bordered table-hover table-responsive">
                                            <thead>
                                                <tr><th colspan="5"  class="text-center">Deposited/Filed 
                                       </th></tr>
                                            <tr>
                                                <th>Site</th>
                                                <th>Type</th>
                                                <th>Period</th>
                                                <th>Due Date</th>
                                               
                                                <th>Deposit Date</th>
                                            </tr>
                                            </thead>
                                            <tbody>
                                            
                                               <%showret(1) %>
                                            
                                            
                                            </tbody>
                                        </table>
                                                    </div>
                                                </div>
                                                                    <div class="item">
                                                    
                                                    <div data-parent="#liasAccordion" id="liasup" class="collapse"><table class="mb-0 table table-bordered table-hover table-responsive">
                                            <thead>
                                                <tr><th colspan="4"  class="text-center">Upcoming</th></tr>
                                            <tr>
                                                <th>Site</th>
                                                <th>Type</th>
                                                <th>Period</th>
                                                <th>Due Date</th>
                                               
                                                
                                            </tr>
                                            </thead>
                                            <tbody>
                                            
                                              <%showret(2) %>
                                            </tbody>
                                        </table></div>
                                                </div>
                                                
                                            </div>  
                                                            </li>
                                                            <%end if %>
                                                             <%if registration = "on" then %>
                                                            <li class="list-group-item">
                                                                <div class="widget-content p-0">
                                                                    <div class="widget-content-wrapper">
                                                                        <div class="widget-content-left mr-3">
                                                                            <i class="btn rounded-circle btn-info text-white"><b>R</b></i>
                                                                        </div>
                                                                        <div class="widget-content-left">
                                                                            <div class="widget-heading">Registrations/Licenses</div>
                                                                            <div class="widget-subheading">S&E, CLRA, PT, LWF, PF, ESI, BOCW, W.W.E.</div>
                                                                        </div>
                                                                        <div class="widget-content-right">
                                                                            <div class="font-size-xlg text-muted">
                                                                                <!--<small class="opacity-5 pr-1">$</small>-->
                                                                                <span><%=rtotal%></span>
                                                                                 <button type="button" aria-expanded="true" aria-controls="regAccordion" data-toggle="collapse" href="#regis" class="m-0 p-0 btn btn-link"> 
                                                                                    <strong class="text-success">
                                                                                                                   <i class="btn rounded-circle fa fa-angle-down bg-success"></i>
                                                        </strong>
                                                                                     </button>
                                                                                <button type="button" aria-expanded="true" aria-controls="regAccordion" data-toggle="collapse" href="#regisup" class="m-0 p-0 btn btn-link"> 
                                                                                    <strong class="text-success">
                                                                                                                   <i class="btn rounded-circle fa fa-angle-down bg-info"></i>
                                                        </strong>
                                                                                     </button>
                                                                            </div>
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                <div id="regAccordion" data-children=".item">
                                                <div class="item">
                                                    
                                                    <div data-parent="#regAccordion" id="regis" class="collapse">
                                                        <table class="mb-0 table table-bordered table-hover table-responsive">
                                            <thead>
                                                <tr><th colspan="3"  class="text-center">Compliant</th></tr>
                                            <tr>
                                                <th>Site</th>
                                                <th>Type</th>
                                               
                                               
                                                <th>Expiry Date</th>
                                            </tr>
                                            </thead>
                                            <tbody>
                                            
                                              <%showreg(1) %>
                                            
                                            
                                            </tbody>
                                        </table>
                                                    </div>
                                                </div>
                                                <div class="item">
                                                    
                                                    <div data-parent="#regAccordion" id="regisup" class="collapse">
                                                        <table class="mb-0 table table-bordered table-hover table-responsive">
                                            <thead>
                                                <tr><th colspan="3"  class="text-center">Expiring/Expired</th></tr>
                                            <tr>
                                                <th>Site</th>
                                                <th>Type</th>
                                               
                                               
                                                <th>Expiry Date</th>
                                            </tr>
                                            </thead>
                                            <tbody>
                                            
                                              <%showreg(2) %>
                                            
                                            
                                            </tbody>
                                        </table>
                                                    </div>
                                                </div>

                                            </div>  
                                                            </li>
                                                            <%end if %>
                                                          
                                                           
                                                        </ul>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-12 col-lg-6">
                                <div class="mb-3 card container-height-div">
                                   
                                    <div class="card-header-tab card-header">
                                        <div class="card-header-title">
                                            <i class="header-icon lnr-rocket icon-gradient bg-tempting-azure"> </i>
                                            Flags
                                        </div>
                                        <div class="btn-actions-pane-right">
                                            <div class="nav">
                                                <a class="mb-2 mr-2 btn-transition btn btn-outline-info" href="details_Dev.asp?<%=urladd %>" target="_blank">Detailed View
                                        </a>
                                                <a data-toggle="tab" href="#overdue" class="border-0 btn-pill btn-wide btn-transition  active btn btn-outline-alternate">OverDue</a>
                                                <a data-toggle="tab" href="#delays" class="ml-1 btn-pill btn-wide border-0 btn-transition   btn btn-outline-alternate second-tab-toggle-alt">Delayed</a>
                                            </div>
                                        </div>
                                    </div>
                                   
                                    <div class="tab-content">
                                        <div class="tab-pane fade active show" id="overdue">
                                            <table class="mb-0 table table-bordered table-hover table-responsive">
                                            <thead>
                                            <tr>
                                                <th>Site</th>
                                                <th>Type</th>
                                                <th>Due Date</th>
                                                <th>Delay</th>
                                            </tr>
                                            </thead>
                                            <tbody>
                                                <%if payrolls="on" then %>
                                            <%ocont(1) %>
                                                <%end if %>
                                                 <%if returns="on" then%>
                                            <%oret(1) %>
                                                <%end if %>
                                           
                                            
                                           
                                            </tbody>
                                        </table>
                                                                                    </div>
                                        <div class="tab-pane fade show" id="delays">
                                             <table class="mb-0 table table-bordered table-hover table-responsive">
                                            <thead>
                                            <tr>
                                                <th>Site</th>
                                                <th>Type</th>
                                                <th>Due Date</th>
                                                <th>Delay</th>
                                                <th>Deposit Date</th>
                                            </tr>
                                            </thead>
                                            <tbody>
                                            
                                              <%if payrolls="on" then%>
                                            <%ocont(2) %>
                                                <%end if %>
                                              <%if returns="on" then%>
                                            <%oret(2) %>
                                                <%end if %>
                                          
                                            </tbody>
                                        </table>
                                                                                    </div>
                                    </div>
                                </div>
                            </div>
                        </div>

                 





                      
                        <div class="row">
                            <div class="col-md-12">
                                <div class="main-card mb-3 card">
                                   
                                    <div class="card-header">On-Going Activities
                                     
                                        <div class="btn-actions-pane-right">
                                           
                                                    <div role="group" class="btn-group-sm nav btn-group">
                                                        
                                                          
                                                        <a data-toggle="tab" href="#activity" class="btn-pill pl-1 active btn btn-focus" >INSPECTIONS</a>
                                                        
                                                        <a data-toggle="tab" href="#license"  class="btn btn-focus" >REG./AMMEND.</a>
                                                        <a data-toggle="tab" href="#pftracker" class="btn btn-focus"   >PF TRACKER</a>
                                                        <a data-toggle="tab" href="#other" class="btn btn-focus" >OTHER</a>
                                                        <a data-toggle="tab" href="#audit" class=" btn btn-focus" >Site Audit</a>
                                       
                                                    </div>
                                                </div>
                                                                           </div>
                                    <div class="tab-content">
                                                    <div class="tab-pane active" id="activity" role="tabpanel"><div class="table-responsive" >
                                        <table class="align-middle mb-0 table table-borderless table-striped table-hover">
                                            <thead>
                                            <tr>
                                                <th class="text-center">#</th>
                                                <th>ORG.Name</th>
                                                <th class="text-center">Title</th>
                                                <th class="text-center">Site Name</th>
                                                <th class="text-center">City</th>
                                                <th class="text-center">Status</th>
                                                <th class="text-center">Date of Opening</th>
                                                <th class="text-center">Actions</th>

                                            </tr>
                                            </thead>
                                            <tbody>
                                          <%showactivity("I") %>
                                                
                                                
                                            </tbody>
                                        </table>
                                    </div></div>
                                         <div class="tab-pane active1" id="license" role="tabpanel"><div class="table-responsive" >
                                        <table class="align-middle mb-0 table table-borderless table-striped table-hover">
                                            <thead>
                                            <tr>
                                                <th class="text-center">#</th>
                                                <th>ORG.Name</th>
                                                <th class="text-center">Title</th>
                                                <th class="text-center">Site Name</th>
                                                <th class="text-center">City</th>
                                                <th class="text-center">Status</th>
                                                <th class="text-center">Date of Opening</th>
                                                <th class="text-center">Actions</th>

                                            </tr>
                                            </thead>
                                            <tbody>
                                          <%showactivity("A") %>
                                                
                                                
                                            </tbody>
                                        </table>
                                    </div></div>
                                         <div class="tab-pane active2" id="pftracker" role="tabpanel"><div class="table-responsive" >
                                        <table class="align-middle mb-0 table table-borderless table-striped table-hover">
                                            <thead>
                                            <tr>
                                                <th class="text-center">#</th>
                                                <th>ORG.Name</th>
                                                <th class="text-center">Title</th>
                                                <th class="text-center">Site Name</th>
                                                <th class="text-center">City</th>
                                                <th class="text-center">Status</th>
                                                <th class="text-center">Date of Opening</th>
                                                <th class="text-center">Actions</th>

                                            </tr>
                                            </thead>
                                            <tbody>
                                          <%showactivity("P") %>
                                                
                                                
                                            </tbody>
                                        </table>
                                    </div></div>
                                         <div class="tab-pane" id="other" role="tabpanel"><div class="table-responsive" >
                                        <table class="align-middle mb-0 table table-borderless table-striped table-hover">
                                            <thead>
                                            <tr>
                                                <th class="text-center">#</th>
                                                <th>ORG.Name</th>
                                                <th class="text-center">Title</th>
                                                <th class="text-center">Site Name</th>
                                                <th class="text-center">City</th>
                                                <th class="text-center">Status</th>
                                                <th class="text-center">Date of Opening</th>
                                                <th class="text-center">Actions</th>

                                            </tr>
                                            </thead>
                                            <tbody>
                                          <%showactivity("O") %>
                                                
                                                
                                            </tbody>
                                        </table>
                                    </div></div>
                                      
                                                    <div class="tab-pane" id="audit" role="tabpanel">
                                                        <div class="table-responsive" >
                                        <table class="align-middle mb-0 table table-borderless table-striped table-hover">
                                            <thead>
                                            <tr>
                                                <th class="text-center">#</th>
                                                <th>ORG.Name</th>
                                                
                                                <th class="text-center">Site Name</th>
                                                
                                                <th class="text-center">City</th>
                                                <th class="text-center">Date of Audit/Visit</th>
                                                <th class="text-center">Score</th>
                                                <th class="text-center">Actions</th>

                                            </tr>
                                            </thead>
                                            <tbody>
                                            
                                             <%showaud%>
                                            </tbody>
                                        </table>
                                    </div>
                                                    </div>
                                    
                                </div>
                                
                            </div>
                        </div>
                        
                    </div>
                       </div>
                <script src="http://maps.google.com/maps/api/js?sensor=true"></script>
                      
        </div>
    </div>

<div class="modal  fade" id="m_modal_1_2" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true"  >
                            
                                <div class="modal-dialog modal-lg" role="document" style="max-width: 1000px;">
 
                                    <form action="reports.asp" name="fq"  id="fq"  method="post" class="m-form m-form--fit  m-form--label-align-right m-form--group-seperator-dashed">
                      <div class="modal-content">
                                        <div class="modal-header">
                                            <h5 class="modal-title" id="exampleModalLabel">Filter</h5>
                                            <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                                               <span aria-hidden="true">&times;</span>
                                            </button>
                                        </div>
                                        <div class="modal-body p1">
                                            <div class="m-scrollable" data-scrollbar-shown="true" data-scrollable="true" data-height="400">
                                                  <%if ulev < 4 then%>
            <!--                                                <div class="m-portlet__body">-->
                                                <div class="form-group m-form__group row p1">
                                                    <div class="col-lg-6">
                                                      <b>  <label for="">Select Client </label></b>
                                                 
                                                   <%=getuoauth(uolist,"uolist",uno)%>
                                                    </div>
                                                   
                                                    
													
												 </div>   
                                                  <div class="divider"></div>
                                                 <%end if%>


                                                <div class="form-group m-form__group row p1">
                                                    <div class="col-lg-6">
                                                    <b>    <label for="">State </label></b>
                                                    		
                                               <%=getostates(stlist,"stlist", uolist,ualist)%>
                                                    </div>
                                                    <div class="col-lg-6">
                                                 <b>         <label for="">City </label>  </b>
                                                <%=getocity(ct,"ct",uolist,ualist)%>
                                                    </div>
													
                                                    		
												 </div>
                                           
													
												  <div class="divider"></div>
                                               
                                                  
                                                   
                                                   <div class="form-group m-form__group row p1">
                                                        <div class="col-lg-6">
                                                       <b>      <label for="">Location </label>  </b>
                                                      <%=getuauth(ualist,"ualist",uno,uolist)%>
                                                        </div>
                                                     
                                                   
                                                 </div>

                                                 <div class="divider"></div>

                                                 <div class="form-group m-form__group row p1">
                                                        <div class="col-lg-6">
                                                     <b>        <label for="">Choose Category </label><br/>  </b>
                                                     <div class="position-relative form-check form-check-inline" ><label class="form-check-label" STYLE="color:black"><input type="checkbox" class="form-check-input" id="payrolls" name="payrolls" onclick="PAY()"  <%if payrolls="on" then Response.Write(" checked ")%> /><B>Payroll</B></label>&nbsp;&nbsp;(&nbsp;&nbsp;
                                                 
                                                 <div class="position-relative form-check form-check-inline"><label class="form-check-label"><input type="checkbox" class="form-check-input" id=pf name=pf <%if pf="on" then Response.Write(" checked ")%>>P.F.</label></div>
                                                <div class="position-relative form-check form-check-inline"><label class="form-check-label"><input type="checkbox" class="form-check-input" id=esi name=esi <%if esi="on" then Response.Write(" checked ")%>>E.S.I.</label></div>
                                                <div class="position-relative form-check form-check-inline"><label class="form-check-label"><input type="checkbox" class="form-check-input" id=pt name=pt  <%if pt="on" then Response.Write(" checked ")%>>P.T.</label></div>
                                                <div class="position-relative form-check form-check-inline"><label class="form-check-label"><input type="checkbox" class="form-check-input" id=lwf name=lwf <%if lwf="on" then Response.Write(" checked ")%>>L.W.F.</label></div>)
                                                  </div><br/>
                          <div class="position-relative form-check form-check-inline">
                              <label class="form-check-label" STYLE="color:black">
                                  <input type="checkbox" class="form-check-input" id=returns name=returns <%if returns="on" then Response.Write(" checked ")%>><B>Labour Returns</B></label>

                          </div><br/>
            <div class="position-relative form-check form-check-inline">
                <label class="form-check-label" STYLE="color:black"><input type="checkbox" class="form-check-input"  id=registration name=registration <%if registration="on" then Response.Write(" checked ")%>><B>Registrations</B></label>

            </div>

                                                        </div>
                                                
           
                                                   
                                                 </div>
                                           



                                                 <div class="divider"></div>

                                                  <div class="form-group m-form__group row p1">
                                                    <div class="col-lg-6">
                                               <b>         <label for="">Year </label>  </b>
                                                    <select name="SCYEAR" class="form-control-sm form-control">

<option value = "" <%if scyear = 0 then response.write(" selected ")%>>All</option>
<%for x = styear to nxtyear%>
<option value = <%=x%>  <%if cint(scyear) = x then response.write(" selected ")%>><%=x%></option>

<%next%>
</select>
                                                    </div>
                                                    <div class="col-lg-6">
                                                   <b>       <label for="">Period </label>  </b>
                                               <select name="SCPERIOD" class="form-control-sm form-control">
 <option value = 0 <%if scperiod = "0" then response.write(" selected ")%>>Full</option>
 <option value = "M1"  <%if scperiod = "M1" then response.write(" selected ")%>>January</option>
  <option value = "M2"  <%if scperiod = "M2" then response.write(" selected ")%>>Febrary</option>
 <option value = "M3"  <%if scperiod = "M3" then response.write(" selected ")%>>March</option>
 <option value = "M4"  <%if scperiod = "M4" then response.write(" selected ")%>>April</option>
 <option value = "M5"  <%if scperiod = "M5" then response.write(" selected ")%>>May</option>
 <option value = "M6"  <%if scperiod = "M6" then response.write(" selected ")%>>June</option>
 <option value = "M7"  <%if scperiod = "M7" then response.write(" selected ")%>>July</option>
 <option value = "M8"  <%if scperiod = "M8" then response.write(" selected ")%>>August</option>
 <option value = "M9"  <%if scperiod = "M9" then response.write(" selected ")%>>September</option>
 <option value = "M10"  <%if scperiod = "M10" then response.write(" selected ")%>>October</option>
 <option value = "M11"  <%if scperiod = "M11" then response.write(" selected ")%>>November</option>
  <option value = "M12"  <%if scperiod = "M12" then response.write(" selected ")%>>December</option>


 <option value = "H1"  <%if scperiod = "H1" then response.write(" selected ")%>>H1 - First Half</option>
 <option value = "H2"  <%if scperiod = "H2" then response.write(" selected ")%>>H2 - Second Half</option>
 <option value = "Q1"  <%if scperiod = "Q1" then response.write(" selected ")%>>Q1 - First Quarter</option>
 <option value = "Q2"  <%if scperiod = "Q2" then response.write(" selected ")%>>Q2 - Second Quarter</option>
 <option value = "Q3"  <%if scperiod = "Q3" then response.write(" selected ")%>>Q3 - Third Quarter</option>
 <option value = "Q4"  <%if scperiod = "Q4" then response.write(" selected ")%>>Q4 - Fourth Quarter</option>
</select>
                                                    </div>
													
                                                    		
												 </div>


   <div class="divider"></div>
                                                     
                                  <div class="modal-footer p1">
                                             <input type=hidden name="srch" id = srch value=1>
                                                    <input type = hidden name = tp value = 1>
<input type = hidden name = tpp value = "<%=tpp%>">
<input type = hidden name = tpod value = "<%=tpod%>">
     
      
                                                    
                                                    
                    <button type="button" class="btn btn-secondary" data-dismiss="modal">Cancel</button>
                  <input type="submit"   value="Submit" class="btn btn-info" id="Submit" name="Submit"> 
                </div>
                                       		      
                                      </div>
                                                                <!--    </div>-->
                                            </div>
                                    </div>        
                        
                            </form>     
									</div>
                                         </div>
<script type="text/javascript" src="./assets/scripts/maine.js"></script>
    <!--<script type="text/javascript" src="./assets/scripts/dashboard.js"></script>-->
        <script type="text/javascript">
<!--#include file="dashbrdjs.asp" -->
        </script>
</body>
</html>


<%function showcontr(aa)
  
      '  response.write("here"&aa)
   
    
cscore=0  'contribution score
ctot=0    'no of items totally
ctotc=0   'no of items compliant
ctotnc=0  'no of items non compliant

asq1="select count(*) from nccontr a, ncmorg b, ncmloc c  where a.lcode = c.lcode and a.status <> 0 and c.lactive='1' "
asq2="select count(*) from nccontr a, ncmorg b, ncmloc c  where a.lcode = c.lcode and status = 1 and c.lactive='1' "
asq3="select a.contid, a.oid ,a.lcode, tp, period,cyear,lastdate,status,depdate, b.oname, c.lname from nccontr a, ncmorg b, ncmloc c where contid <> 0 and c.lactive='1' "
    
   

tsq=tsq&" and (status <> 0 or (status = 0 and datediff(day,GETDATE(),lastdate)  <= 30))"
tsq=tsq&" and a.oid = b.oid and a.lcode = c.lcode and b.oid = c.oid "

tsq=tsq&" and a.lcode in  (select lcode from ncumap where oid = a.oid and uno = "&uno&") and ("
    
                            
if pt="on" then TSQ=TSQ&"tp = 'PT' or "
if ept="on" then TSQ=TSQ&" tp = 'EPT' or "
if lwf="on" then TSQ=TSQ&" tp = 'LWF' or "
if pf="on" then TSQ=TSQ&" tp = 'PF' or "
if esi="on" then TSQ=TSQ&" tp = 'ESI' or "
TSQ=TSQ&" tp = '')"
     if aa="1" then
    tsq=tsq&"  and status='1'  "
    else
    tsq=tsq&" and status='0' "
    end if
if locsql <> "" then tsq=tsq&locsql

if scyear <> 0 then tsq=tsq&" and cyear = "&scyear
if scperiod <> "0" then
 if scperiod = "H1" then tsq=tsq&" and period <= 6"
 if scperiod = "H2" then tsq=tsq&" and period > 6"
 if scperiod = "Q1" then tsq=tsq&" and period between 1 and 3"
 if scperiod = "Q2" then tsq=tsq&" and period between 4 and 6"
 if scperiod = "Q3" then tsq=tsq&" and period between 7 and 9"
 if scperiod = "Q4" then tsq=tsq&" and period between 10 and 12"
 if scperiod = "M1" then tsq=tsq&" and period = 1"
  if scperiod = "M2" then tsq=tsq&" and period = 2"
  if scperiod = "M3" then tsq=tsq&" and period = 3"
  if scperiod = "M4" then tsq=tsq&" and period = 4"
  if scperiod = "M5" then tsq=tsq&" and period = 5"
  if scperiod = "M6" then tsq=tsq&" and period = 6"
  if scperiod = "M7" then tsq=tsq&" and period = 7"
  if scperiod = "M8" then tsq=tsq&" and period = 8"
  if scperiod = "M9" then tsq=tsq&" and period = 9"
  if scperiod = "M10" then tsq=tsq&" and period = 10"
  if scperiod = "M11" then tsq=tsq&" and period = 11"
  if scperiod = "M12" then tsq=tsq&" and period = 12"


end if

asq1=asq1&tsq
asq2=asq2&tsq
asq3=asq3&tsq&" order by lastdate desc "

'response.write(asq1)
'response.write(asq2)
'response.write(asq3)
rs.open asq1,conndb
ctot=rs(0)
rs.close

rs.open asq2,conndb
ctotc=rs(0)
rs.close

ctotnc=ctot-ctoc
if ctot > 0 then cscore=round(100*ctotc/ctot,2) else cscore = 0
%>



<% 'end score calculations / start showing nclist
   ' response.Write(asq3)
rs.open asq3,conndb
if rs.eof then
if aa="1" then

    %>
<tr>
      <th scope="row" colspan="5" class="text-center">NONE</th>  
    
        


    </tr>
<%
    else

     %>
<tr>
      <th scope="row" colspan="4" class="text-center">NONE</th>  
    
        


    </tr>
<%
  
    end if
rs.close
else
%>
<%

cntr=1
do while not rs.eof
contid=rs("contid")
oid=rs("oid")
lcode=rs("lcode")
tp=rs("tp")
period=rs("period")
cyear=rs("cyear")
lastdate=rs("lastdate")
oname=rs("oname")
lname=rs("lname")
status=rs("status")
depdate=rs("depdate")
speriod = monthname(period)&"-"&cyear
if status = 0 then sstatus = "<font color  = Navy><b>U</b></font>"
if status = 1 then sstatus = "<font color  = green><b>C</b></font>"
if status = 2 then sstatus = "<font color  = red><b>NC</b></font>"
durl="edcont.asp?contid="&contid
atit=lname&" - "&tp&" "&speriod
if atit = "" then atit = "Blank"
%>
<%if aa="1"   then

    %>
<tr>
   
<td><A href = "<%=durl%>" target = _blank><b><%=trim(lname)%></b></a></td>
<td><b><%=tp%></b></td>
<td><b><%=speriod%></b></td>
<td><%=showdate(lastdate,"dmy")%></td>
<td><%=showdate(depdate,"dmy")%></td>

</tr>
<%else %>
<tr>
<td><A href = "<%=durl%>" target = _blank><b><%=trim(lname)%></b></a></td>
<td><b><%=tp%></b></td>
<td><b><%=speriod%></b></td>
<td><%=showdate(lastdate,"dmy")%></td>


</tr>
<%end if %>
<%
cntr=cntr+1
rs.movenext
loop
rs.close
end if
%>

<%end function %>

<%function showreg(aa)
rscore=0  'registrations score
rtot=0    'no of items totally
rtotc=0   'no of items compliant
rtotnc=0  'no of items non compliant

asq1="select count(*) from ncreg a, ncmorg b, ncmloc c where a.lcode = c.lcode and (datediff(day,GETDATE(),doe) <=45 ) and c.lactive='1' "
asq2="select count(*) from ncreg a, ncmorg b, ncmloc c where a.lcode = c.lcode and   (a.doe>a.doi) and c.lactive='1' "
asq3="select a.uid, a.oid ,a.lcode, tp, doe, status, b.oname, c.lname from ncreg a, ncmorg b, ncmloc c where status <> 'XX' and c.lactive='1' "

  

tsq=tsq&" and a.oid = b.oid and a.lcode = c.lcode and a.oid = c.oid"


tsq=tsq&" and a.lcode in  (select lcode from ncumap where oid = a.oid and uno = "&uno&") and ("
 
if registration="on" then TSQ=TSQ&"tp <> '' or "
      
TSQ=TSQ&" tp = '')"
     if aa="1" then
    tsq=tsq&"  and status='C'  "
    else
    tsq=tsq&" and status='E' "
    end if
if locsql <> "" then tsq=tsq&locsql
'if scyear <> 0 then tsq=tsq&" and ( (year(a.doi) = "&scyear&" or year(doe) = "&scyear&") or a.doi is null or year(a.doi) = 1900)"
if scperiod <> "0" then
 if scperiod = "H1" then tsq=tsq&" and month(doe) <= 6"
 if scperiod = "H2" then tsq=tsq&" and month(doe) > 6"
 if scperiod = "Q1" then tsq=tsq&" and month(doe) between 1 and 3"
 if scperiod = "Q2" then tsq=tsq&" and month(doe) between 4 and 6"
 if scperiod = "Q3" then tsq=tsq&" and month(doe) between 7 and 9"
 if scperiod = "Q4" then tsq=tsq&" and month(doe) between 10 and 12"
  if scperiod = "M1" then tsq=tsq&" and month(doe) = 1"
  if scperiod = "M2" then tsq=tsq&" and month(doe) = 2"
  if scperiod = "M3" then tsq=tsq&" and month(doe) = 3"
  if scperiod = "M4" then tsq=tsq&" and month(doe) = 4"
  if scperiod = "M5" then tsq=tsq&" and month(doe) = 5"
  if scperiod = "M6" then tsq=tsq&" and month(doe) = 6"
  if scperiod = "M7" then tsq=tsq&" and month(doe) = 7"
  if scperiod = "M8" then tsq=tsq&" and month(doe) = 8"
  if scperiod = "M9" then tsq=tsq&" and month(doe) = 9"
  if scperiod = "M10" then tsq=tsq&" and month(doe) = 10"
  if scperiod = "M11" then tsq=tsq&" and month(doe) = 11"
  if scperiod = "M12" then tsq=tsq&" and month(doe) = 12"

end if

asq1=asq1&tsq
asq2=asq2&tsq
asq3=asq3&tsq&" order by doe asc, lname "
'response.write(asq1)
'response.write(asq2)
'response.write(asq3)
rs.open asq1,conndb
rtot=rs(0)
rs.close

rs.open asq2,conndb
rtotc=rs(0)
rs.close

rtotnc=rtot-rtoc
if rtot > 0 then rscore=round(100*rtotc/rtot,2) else rscore = 0
gscore=44.4
%>


<% 'end score calculations / start showing nclist
rs.open asq3,conndb
if rs.eof then
if aa="1" then

    %>
<tr>
      <th scope="row" colspan="5" class="text-center">NONE</th>  
    
        


    </tr>
<%
    else

     %>
<tr>
      <th scope="row" colspan="4" class="text-center">NONE</th>  
    
        


    </tr>
<%
  
    end if
rs.close
else
%>

<%
cntr=1
do while not rs.eof
uid=rs("uid")
oid=rs("oid")
lcode=rs("lcode")
tp=trim(rs("tp"))
doe=rs("doe")
oname=rs("oname")
lname=rs("lname")
status=rs("status")
sstatus = status

if status = "C" then sstatus = "<font color  = green>Compliant</font>"
if status = "A" then sstatus = "<font color  = darkorange>Applied</font>"
if status = "B" then sstatus = "<font color  = darkblue>In Process</font>"
    if status = "SC" then sstatus = "<font color  = darkblue>Site Closed</font>"
    if status = "SU" then sstatus = "<font color  = darkblue>Surrenderd</font>"
    if status = "NA" then sstatus = "<font color  = darkblue>Not Applicable</font>"

if status = "N" then sstatus = "<font color  = red>Not Applied</font>"
if status = "E" then sstatus = "<font color  = red>Expired</font>"
if status = "D" then sstatus = "<font color  = red>Docs Awaited</font>"
'if status = 1 then sstatus = "<font color  = greeen>C</font>" else sstatus = "<font color  = red><b>NC</b></font>"
exdate=showdate(doe,"dmy")
if isnull(doe) or doe = "" or year(doe) = 1900 then
exdate=""
else
if dateadd("d",60,date()) > cdate(doe) and status <> "E" and status <> "D" and status <> "B"   and status <> "A"    then sstatus ="<font color  = darkorange>Expiring</font>"
end if


if tp="SE" then tp1="Shops and Establishments Registration"
if tp="CLR" then tp1="CLRA Registration"
if tp="PT" then tp1="PT Registration"
if tp="ESI" then tp1="ESI sub-code"
if tp="LWF" then tp1="LWF code"
if tp="LL" then tp1="CLRA Licence"
if tp="TL" then tp1="Trade Licence"
if tp="247" then tp1="24-7 Exemption"
if tp="WWN" then tp1="Woman Working at Night"
if tp="FL" then tp1="Factories License"
if tp="OTH" then tp1="Other"

if year(doe)>=2050  then  exdate = "31-12-2099"
if year(doe)=1900 then  exdate = ""



durl="edreg.asp?uid="&uid&"&oid="&oid
atit=lname&" - "&tp1
if atit = "" then atit = "Blank"

%>
<%if aa="1"   then %>
<tr>

<td><A href = "<%=durl%>" target = _blank><b><%=trim(lname)%></b></a></td>
<td><A href = "<%=durl%>" target = _blank><b><%=tp%></b></a></td>
<td><%=exdate%></td>

</tr>
<%else %>
<tr>

<td><A href = "<%=durl%>" target = _blank><b><%=trim(lname)%></b></a></td>
<td><A href = "<%=durl%>" target = _blank><b><%=tp%></b></a></td>
<td><%=exdate%></td>

</tr>
<%end if %>
   




<%
cntr=cntr+1
rs.movenext
loop
rs.close
end if
%>



<%end function %>


<%function showret(aa)
cscore=0  'returns score
ctot=0    'no of items totally
ctotc=0   'no of items compliant
ctotnc=0  'no of items non compliant

asq1="select count(*) from ncret a, ncmorg b, ncmloc c, nctempret d  where a.lcode = c.lcode and a.status <> 0  and c.lactive='1' "
asq2="select count(*) from ncret a, ncmorg b, ncmloc c, nctempret d where a.lcode = c.lcode and status = 1 and c.lactive='1' "
asq3="select a.rtid, a.oid ,a.lcode, a.rcode,  ryear ,lastdate,status,depdate, b.oname, c.lname, d.rform, d.rtitle from ncret a, ncmorg b, ncmloc c, nctempret d where status <> 99 and c.lactive='1' "

 if aa="1" then
    tsq=tsq&"  and status='1'  "
    else
    tsq=tsq&" and status='0' "
    end if

tsq=tsq&" and a.oid = b.oid and a.oid = c.oid and a.lcode = c.lcode and a.rcode = d.rcode"
'asq3=asq3&" and (status <> 0 or status = 0 and (lastdate-getdate() <= 45))"

tsq=tsq&" and (status <> 0 or (status = 0 and datediff(day,GETDATE(),lastdate)  <= 30))"



tsq=tsq&" and a.lcode in  (select lcode from ncumap where oid = a.oid and uno = "&uno&") and ("
if returns="on" then TSQ=TSQ&"rtype <> '' or "
TSQ=TSQ&" rtype = '')"

if locsql <> "" then tsq=tsq&locsql

if scyear <> 0 then tsq=tsq&" and ryear = "&scyear

if scperiod <> "0" then
 if scperiod = "H1" then tsq=tsq&" and month(lastdate) <= 6"
 if scperiod = "H2" then tsq=tsq&" and month(lastdate) > 6"
 if scperiod = "Q1" then tsq=tsq&" and month(lastdate) between 1 and 3"
 if scperiod = "Q2" then tsq=tsq&" and month(lastdate) between 4 and 6"
 if scperiod = "Q3" then tsq=tsq&" and month(lastdate) between 7 and 9"
 if scperiod = "Q4" then tsq=tsq&" and month(lastdate) between 10 and 12"
  if scperiod = "M1" then tsq=tsq&" and month(lastdate) = 1"
  if scperiod = "M2" then tsq=tsq&" and month(lastdate) = 2"
  if scperiod = "M3" then tsq=tsq&" and month(lastdate) = 3"
  if scperiod = "M4" then tsq=tsq&" and month(lastdate) = 4"
  if scperiod = "M5" then tsq=tsq&" and month(lastdate) = 5"
  if scperiod = "M6" then tsq=tsq&" and month(lastdate) = 6"
  if scperiod = "M7" then tsq=tsq&" and month(lastdate) = 7"
  if scperiod = "M8" then tsq=tsq&" and month(lastdate) = 8"
  if scperiod = "M9" then tsq=tsq&" and month(lastdate) = 9"
  if scperiod = "M10" then tsq=tsq&" and month(lastdate) = 10"
  if scperiod = "M11" then tsq=tsq&" and month(lastdate) = 11"
  if scperiod = "M12" then tsq=tsq&" and month(lastdate) = 12"
end if

asq1=asq1&tsq
asq2=asq2&tsq
asq3=asq3&tsq&" order by lastdate desc "
'response.write(asq1)
'response.write(asq2)
'response.write(asq3)
rs.open asq1,conndb
ctot=rs(0)
rs.close

rs.open asq2,conndb
ctotc=rs(0)
rs.close

ctotnc=ctot-ctoc
if ctot > 0 then cscore=round(100*ctotc/ctot,2) else cscore = 0
%>


<% 'end score calculations / start showing nclist
rs.open asq3,conndb
if rs.eof then 
    if aa="1" then

    %>
<tr>
      <th scope="row" colspan="5" class="text-center">NONE</th>  
    
        


    </tr>
<%
    else

     %>
<tr>
      <th scope="row" colspan="4" class="text-center">NONE</th>  
    
        


    </tr>
<%
  
    end if
rs.close
else
%>



<%
cntr=1
do while not rs.eof
rtid=rs("rtid")
oid=rs("oid")
lcode=rs("lcode")
rcode=rs("rcode")
ryear=rs("ryear")
lastdate=rs("lastdate")
depdate=rs("depdate")
oname=rs("oname")
lname=rs("lname")
rform=trim(rs("rform"))
rtitle=trim(rs("rtitle"))

status=rs("status")

if status = 0 then sstatus = "<font color  = navy><b>U</b></font>"
if status = 1 then sstatus = "<font color  = greeen>C</font>"
if status = 2 then sstatus = "<font color  = red><b>NC</b></font>"

durl="edret.asp?rtid="&rtid
atit=lname&" - "&rtitle&" "&rform
if atit = "" then atit = "Blank"

%>

<%if aa="1"   then %>
<tr>

<td><b><%=trim(lname)%></b></td>
<td width = 180 class = fovs><b><A href = "<%=durl%>" target = _blank><%=rtitle%> - <%=rform%></a></b></td>
<td><%=ryear%></td>
<td><%=showdate(lastdate,"dmy")%></td>
<td><%=showdate(depdate,"dmy")%>

</td>
</tr>
<%else %>
<tr>

<td><b><%=trim(lname)%></b></td>
<td width = 180 class = fovs><b><A href = "<%=durl%>" target = _blank><%=rtitle%> - <%=rform%></a></b></td>
<td><%=ryear%></td>
<td><%=showdate(lastdate,"dmy")%></td>


</td>
</tr>
<%end if %>

<%
cntr=cntr+1
rs.movenext
loop
rs.close
end if
%>



<%end function %>


<%function showfin(aa)
cscore=0  'finance score
ctot=0    'no of items totally
ctotc=0   'no of items compliant
ctotnc=0  'no of items non compliant

asq1="select count(*) from ncfin a, ncmorg b, ncmloc c, nctempfin d  where a.lcode = c.lcode and a.status <> 0 "
asq2="select count(*) from ncfin a, ncmorg b, ncmloc c, nctempfin d where a.lcode = c.lcode and status = 1 "
asq3="select a.rtid, a.oid ,a.lcode, a.rcode,  ryear ,lastdate,status,depdate, b.oname, c.lname, d.rform, d.rtitle from ncfin a, ncmorg b, ncmloc c, nctempfin d where status <> 99 "

 if aa="1" then
    tsq=tsq&"  and status='1'  "
    else
    tsq=tsq&" and status='0' "
    end if


tsq=tsq&" and a.oid = b.oid and a.oid = c.oid and a.lcode = c.lcode and a.rcode = d.rcode"
'asq3=asq3&" and (status <> 0 or status = 0 and (lastdate-getdate() <= 45))"

tsq=tsq&" and (status <> 0 or (status = 0 and datediff(day,GETDATE(),lastdate)  <= 30))"



tsq=tsq&" and a.lcode in  (select lcode from ncumap where oid = a.oid and uno = "&uno&") and ("

if GST="on" then TSQ=TSQ&"rtype = 'GST' or "
if IT="on" then TSQ=TSQ&" rtype = 'IT' or "
if TCA="on" then TSQ=TSQ&" rtype = 'TCA' or "

TSQ=TSQ&" rtype = '')"
if locsql <> "" then tsq=tsq&locsql

if scyear <> 0 then tsq=tsq&" and ryear = "&scyear

if scperiod <> "0" then
 if scperiod = "H1" then tsq=tsq&" and month(lastdate) <= 6"
 if scperiod = "H2" then tsq=tsq&" and month(lastdate) > 6"
 if scperiod = "Q1" then tsq=tsq&" and month(lastdate) between 1 and 3"
 if scperiod = "Q2" then tsq=tsq&" and month(lastdate) between 4 and 6"
 if scperiod = "Q3" then tsq=tsq&" and month(lastdate) between 7 and 9"
 if scperiod = "Q4" then tsq=tsq&" and month(lastdate) between 10 and 12"
  if scperiod = "M1" then tsq=tsq&" and month(lastdate) = 1"
  if scperiod = "M2" then tsq=tsq&" and month(lastdate) = 2"
  if scperiod = "M3" then tsq=tsq&" and month(lastdate) = 3"
  if scperiod = "M4" then tsq=tsq&" and month(lastdate) = 4"
  if scperiod = "M5" then tsq=tsq&" and month(lastdate) = 5"
  if scperiod = "M6" then tsq=tsq&" and month(lastdate) = 6"
  if scperiod = "M7" then tsq=tsq&" and month(lastdate) = 7"
  if scperiod = "M8" then tsq=tsq&" and month(lastdate) = 8"
  if scperiod = "M9" then tsq=tsq&" and month(lastdate) = 9"
  if scperiod = "M10" then tsq=tsq&" and month(lastdate) = 10"
  if scperiod = "M11" then tsq=tsq&" and month(lastdate) = 11"
  if scperiod = "M12" then tsq=tsq&" and month(lastdate) = 12"
end if

asq1=asq1&tsq
asq2=asq2&tsq
asq3=asq3&tsq&" order by lastdate desc "
'response.write(asq1)
'response.write(asq2)
'response.write(asq3)
rs.open asq1,conndb
ctot=rs(0)
rs.close

rs.open asq2,conndb
ctotc=rs(0)
rs.close

ctotnc=ctot-ctoc
if ctot > 0 then cscore=round(100*ctotc/ctot,2) else cscore = 0
%>


<% 'end score calculations / start showing nclist
rs.open asq3,conndb
if rs.eof then
if aa="1" then

    %>
<tr>
      <th scope="row" colspan="5" class="text-center">NONE</th>  
    
        


    </tr>
<%
    else

     %>
<tr>
      <th scope="row" colspan="4" class="text-center">NONE</th>  
    
        


    </tr>
<%
  
    end if
rs.close
else
%>

<%
cntr=1
do while not rs.eof
rtid=rs("rtid")
oid=rs("oid")
lcode=rs("lcode")
rcode=rs("rcode")
ryear=rs("ryear")
lastdate=rs("lastdate")
depdate=rs("depdate")
oname=rs("oname")
lname=rs("lname")
rform=trim(rs("rform"))
rtitle=trim(rs("rtitle"))

status=rs("status")

if status = 0 then sstatus = "<font color  = navy><b>U</b></font>"
if status = 1 then sstatus = "<font color  = greeen>C</font>"
if status = 2 then sstatus = "<font color  = red><b>NC</b></font>"

durl="edfin.asp?rtid="&rtid
atit=lname&" - "&rtitle&" "&rform
if atit = "" then atit = "Blank"

%>

<%if aa="1"   then %>
<tr>

<td><b><%=trim(lname)%></b></td>
<td width = 180 class = fovs><b><A href = "<%=durl%>" target = _blank><%=rtitle%> - <%=rform%></a></b></td>
<td><%=ryear%></td>
<td><%=showdate(lastdate,"dmy")%></td>
<td><%=showdate(depdate,"dmy")%>

</td>
</tr>
<%else %>
<tr>

<td><b><%=trim(lname)%></b></td>
<td width = 180 class = fovs><b><A href = "<%=durl%>" target = _blank><%=rtitle%> - <%=rform%></a></b></td>
<td><%=ryear%></td>
<td><%=showdate(lastdate,"dmy")%></td>


</td>
</tr>
<%end if %>

<%
cntr=cntr+1
rs.movenext
loop
rs.close
end if
%>



<%end function %>



<% function showactivity(tpa)
sq1="select a.*, c.lname, c.lstate, c.lcity, c.lregion , cc.oname from ncaction a, ncmloc c, ncmorg cc "
  '  sq1=sq1& "where a.lcode = c.lcode and a.oid = c.oid" 
   

sq1=sq1&" where a.lcode = c.lcode and a.oid = c.oid and c.oid=cc.oid and a.oid in (select distinct oid from ncumap where uno = "&uno&") "
if locsql <> "" then sq1=sq1&locsql
     if uolist <> "" then sq1=sq1&" and a.oid = '"&uolist&"'"
sq1=sq1&" and a.lcode in (select distinct lcode from ncumap where uno = "&uno&") "
if tpa="I" then sq1=sq1&" and a.actp = 'I' "
if tpa="A" then sq1=sq1&" and a.actp = 'A'"
if tpa="P" then sq1=sq1&" and a.actp = 'P' "
if tpa="O" then sq1=sq1&" and a.actp = 'O' "
'if esi="on" then sq1=sq1&" tp = 'ESI' or "
'sq1=sq1&" tp = '')"
'if uolist <> "" then sq1=sq1&" and a.oid = '"&uolist&"'"
'if status <=2 then sq1=sq1&" and status = "&status
'if status <> "0" then sq1=sq1& " and (status <> 0 or (status = 0 and datediff(d,getdate(),lastdate) <= 15))"
'if timely="1" then sq1=sq1&" and (depdate <= lastdate) and status = 1 "
'if timely="2" then sq1=sq1&" and (depdate > lastdate) and status = 1"
'if status = 0 then sby = 2
if scyear <> 0 then sq1=sq1&" and year(acidate)= "&scyear
if ust<> "" and ust <> "1" then sq1=sq1&" and lstate = '"&ust&"'"
if reg<> "" then sq1=sq1&" and lregion = '"&reg&"'"
if loc<> "" then sq1=sq1&" and a.lcode = '"&loc&"'"
if ct<> "" then sq1=sq1&" and lcity = '"&ct&"'"
sq1=sq1&" order by acidate desc "
rs.CursorLocation = adUseClient
    
rs.open sq1,conndb
    if tpa="P" then
    rs.pagesize = 200
    else
rs.pagesize = 30
    end if
if rs.eof then %>
<%if ulev="2" then %> <th scope="row" colspan="8" class="text-center">
                                        <div class="btn">
                                        <div class="nav">
                                              <a class="mb-2 mr-2 btn-transition btn btn-outline-info" href="activity.asp?acid=0&tp=<%=tpa %>" target="_blank">New
                                        </a>
                                       </div>
                                           </div>
         </th>  
                                        <%end if %>
    <tr>
      <th scope="row" colspan="7" class="text-center">No Data Found</th>  
     
        


    </tr>
<%

else
cnt=0
cbud=0
page=getreq("page")
if trim(page)="" then page=1
np=pg+1
pp=pg-1
TotalPages = rs.PageCount
TotalFound = rs.recordcount
rs.MoveFirst
rs.AbsolutePage = page
'showsum page,TotalPages,TotalFound

%>
 <%if ulev="2" then %> <th scope="row" colspan="7" class="text-center">
                                        <div class="btn">
                                        <div class="nav">
                                              <a class="mb-2 mr-2 btn-transition btn btn-outline-info" href="activity.asp?acid=0&tp=<%=tpa %>" target="_blank">New
                                        </a>
                                       </div>
                                           </div>
         </th>  
                                        <%end if %>
<%
do while not rs.eof and cnt < rs.PageSize
cnt=cnt+1
acid=trim(rs("acid"))
oid=trim(rs("oid"))
oname=trim(rs("oname"))
lcode=trim(rs("lcode"))
actp=trim(rs("actp"))
actitle=trim(rs("actitle"))
    acdetail=trim(rs("acdetail"))
opdate=rs("ACIDATE")
status=rs("acstatus")

cldate=rs("ACCLDATE")

remarks=rs("acremarks")
lname=rs("lname")
lstate=rs("lstate")
lcity=rs("lcity")
lregion=rs("lregion")
sno=(rs.pagesize*page)-rs.pagesize+cnt
if status = "C" then sstatus = "Closed"
if status = "I" then sstatus = "Documents Submitted"
if status = "O" then sstatus = "Open"

if status = "D" then sstatus = "Documents Awaited"

'per=monthname(period)&"-"&cyear
if actp="I" then fre="Inspections/Notices"
if actp="A" then fre="Renewals/Applications"
if actp="P" then fre="PF Tracker"
if actp="O" then fre="Misc. Activities"

    if ulev < "3" then
    durl="activity.asp?acid="&acid&"&oid="&oid&""
    else
     durl="activity.asp?acid="&acid&"&oid="&oid&"&stype=st"
    end if
%>
<tr>
     
                                                <td class="text-center text-muted">#<%=sno%></td>
                                                <td ><%=oname%></td>
                                                <td>
                                                    <div class="widget-content p-0">
                                                        <div class="widget-content-wrapper">
                                                            <div class="widget-content-left mr-3">
                                                                <div class="widget-content-left">
                                                                      <i class="btn rounded-circle btn-info text-white"><b><%=ACTP %></b></i>
                                                                </div>
                                                            </div>
                                                            <div class="widget-content-left flex2">
                                                                <div class="widget-heading"><%=actitle%></div>
                                                                <div class="widget-subheading opacity-7"><%=fre%></div>
                                                            </div>
                                                        </div>
                                                    </div>
                                                </td>
 
      <td class="text-center"><%=lname%></td>
                                               

                                                <td class="text-center"><%=lcity%></td>
                                                <td class="text-center">
                                                    <%if status="O" then %>
                                                    <div class="badge badge-danger"><%=sstatus%></div>
                                                      <%end if %>
                                                     <%if status="D" then%>
                                                    <div class="badge badge-warning"><%=sstatus%></div>
                                                      <%end if %>
                                                    <%if status="I" then %>
                                                    <div class="badge badge-info"><%=sstatus%></div>
                                                      <%end if %>
                                                    <%if status="C" then%>
                                                    <div class="badge badge-success"><%=sstatus%></div>
                                                    <%end if %>
                                                     
                                                    
                                                </td>
                                                 <td class="text-center"><%=showdate(opdate,"dmmy")%></td>

                                                <td class="text-center">
                                                  <a href="<%=durl%>" class="btn btn-primary btn-sm" target="_blank"> Details</a>
                                                </td>
                                            </tr>





<%Response.flush

rs.movenext
loop
end if
rs.close

%>


<%'showsum page,TotalPages,TotalFound

end function%>


<%function showaud()
caudscore=0  'all audits score
caudtot=0    'no of audits
caudtotc=0   'average score of audits


sqaud1="select count(*) from ncaudmast a, ncmloc c  where a.lcode = c.lcode "
sqaud2="select avg(ascore) from ncaudmast a, ncmloc c  where a.lcode = c.lcode "
sqaud3="select a.aid, a.oid ,a.lcode, a.aperiod,  a.ayear ,a.aschdate, a.acomplete, a.ascore, b.oname, c.lname,c.lcity from ncaudmast a, ncmorg b, ncmloc c where ayear <> 0 "
sqaud3=sqaud3&" and a.oid = b.oid and a.lcode = c.lcode "


tsqaud=tsqaud&" and c.lcode in  (select lcode from ncumap where uno = "&uno&") "
if locsql <> "" then tsqaud=tsqaud&locsql

if scyear <> 0 then tsqaud=tsqaud&" and ayear = "&scyear

if scperiod <> "0" then
 if scperiod = "H1" then tsqaud=tsqaud&" and month(aschdate) <= 6"
 if scperiod = "H2" then tsqaud=tsqaud&" and month(aschdate) > 6"
 if scperiod = "Q1" then tsqaud=tsqaud&" and month(aschdate) between 1 and 3"
 if scperiod = "Q2" then tsqaud=tsqaud&" and month(aschdate) between 4 and 6"
 if scperiod = "Q3" then tsqaud=tsqaud&" and month(aschdate) between 7 and 9"
 if scperiod = "Q4" then tsqaud=tsqaud&" and month(aschdate) between 10 and 12"
end if

sqaud1=sqaud1&tsqaud
sqaud2=sqaud2&tsqaud
sqaud3=sqaud3&tsqaud&" order by aschdate desc "
'response.write(sqaud1)
'response.write(sqaud2)
'response.write(sqaud3)
rs.open sqaud1,conndb
caudtot=rs(0)
rs.close

rs.open sqaud2,conndb
caudtotc=rs(0)
rs.close

'ctotnc=ctot-ctoc
'if ctot > 0 then cscore=round(100*ctotc/ctot,2) else cscore = 0
%>

<% 'end score calculations / start showing nclist
rs.open sqaud3,conndb
if rs.eof then%>
    <tr>
      <th scope="row" colspan="6" class="text-center">NONE</th>  
    
        


    </tr>
<%

rs.close
else
%>

<%
cntr=1
'asq3="select a.aid, a.oid ,a.lcode, a.aperiod,  a.ayear ,a.aschdate, a.acomplete, a.ascore, b.oname, c.lname from ncaudmast a, ncmorg b, ncmloc c where ayear <> 0 "

do while not rs.eof
aid=rs("aid")
oid=rs("oid")
lcode=rs("lcode")
aperiod=rs("aperiod")
ayear=rs("ayear")
aschdate=rs("aschdate")
acomplete=rs("acomplete")
ascore=rs("ascore")

oname=rs("oname")
lname=rs("lname")
lcity=rs("lcity")



durl="edaud.asp?aid="&aid


%>

    <tr>
                                                <td class="text-center text-muted">#<%=cntr %></td>
                                                <td>
                                                    <div class="widget-content p-0">
                                                        <div class="widget-content-wrapper">
                                                            <div class="widget-content-left mr-3">
                                                                <div class="widget-content-left">
                                                                      <i class="btn rounded-circle btn-info text-white"><b>I</b></i>
                                                                </div>
                                                            </div>
                                                            <div class="widget-content-left flex2">
                                                                <div class="widget-heading"><%=oname%></div>
                                                                
                                                            </div>
                                                        </div>
                                                    </div>
                                                </td>
                                                 
         <td class="text-center"><%=trim(lname)%></td>
                                                <td class="text-center"><%=lcity%></td>
                                                <td class="text-center">
                                                 <%=showdate(aschdate,"dmmy")%>
                                                </td>
                                                 <td class="text-center"><%=ascore%>%</td>

                                                <td class="text-center">
                                                  <!--<button type="button" id="Button5" class="btn btn-primary btn-sm">Details</button>-->
                                                    <a href="edaud.asp?aid=<%=aid %>" id="Button5" class="btn btn-primary btn-sm">Details</a>
                                                </td>
                                            </tr>


<%
cntr=cntr+1
rs.movenext
loop
rs.close
end if
%>

<%end function %>









<!--FLAGS-->






<%function ocont(aa)
page=getreq("page")
if page = "" then page = 1
' contributions
asq="select a.contid, a.oid ,a.lcode, tp, period,cyear,lastdate,status, depdate, amount, chqno, chqdate, b.oname, c.lname from nccontr a, ncmorg b, ncmloc c where contid <> 0 "
'asq3=asq3&" and (status <> 0 or status = 0 and (lastdate-getdate() <= 15))"

    if aa="1" then
    tsq=tsq&" and status = 2  "
    else
    tsq=tsq&" and status='1' and depdate='' "
    end if

asq=asq&" and a.oid = b.oid and a.lcode = c.lcode and b.oid = c.oid "
asq=asq&" and c.lcode in  (select lcode from ncumap where oid = a.oid and uno = "&uno&")  AND ("
    
                            
if pt="on" then ASQ=ASQ&"tp = 'PT' or "
if ept="on" then ASQ=ASQ&" tp = 'EPT' or "
if lwf="on" then ASQ=ASQ&" tp = 'LWF' or "
if pf="on" then ASQ=ASQ&" tp = 'PF' or "
if esi="on" then ASQ=ASQ&" tp = 'ESI' or "
ASQ=ASQ&" tp = '')"

if locsql <> "" then tsq=tsq&locsql

if scyear <> 0 then tsq=tsq&" and cyear = "&scyear
if scperiod <> "0" then
 if scperiod = "H1" then tsq=tsq&" and period <= 6"
 if scperiod = "H2" then tsq=tsq&" and period > 6"
 if scperiod = "Q1" then tsq=tsq&" and period between 1 and 3"
 if scperiod = "Q2" then tsq=tsq&" and period between 4 and 6"
 if scperiod = "Q3" then tsq=tsq&" and period between 7 and 9"
 if scperiod = "Q4" then tsq=tsq&" and period between 10 and 12"
  if scperiod = "M1" then tsq=tsq&" and period = 1"
  if scperiod = "M2" then tsq=tsq&" and period = 2"
  if scperiod = "M3" then tsq=tsq&" and period = 3"
  if scperiod = "M4" then tsq=tsq&" and period = 4"
  if scperiod = "M5" then tsq=tsq&" and period = 5"
  if scperiod = "M6" then tsq=tsq&" and period = 6"
  if scperiod = "M7" then tsq=tsq&" and period = 7"
  if scperiod = "M8" then tsq=tsq&" and period = 8"
  if scperiod = "M9" then tsq=tsq&" and period = 9"
  if scperiod = "M10" then tsq=tsq&" and period = 10"
  if scperiod = "M11" then tsq=tsq&" and period = 11"
  if scperiod = "M12" then tsq=tsq&" and period = 12"
end if

'if srch <> "" then tsq=tsq&" and (lname like '%"&srch&"%' or oname like '%"&srch&"%' or tp like '%"&srch&"%')"

asq=asq&tsq
asq=asq&" order by lastdate, oid "
%>


<%
    rs.CursorLocation = adUseClient
    rs.open asq, conndb
  '  response.Write(asq)
if rs.eof then
    if aa="1" then
'response.Write("No Overdue Payroll items") 
    %>
<tr>
      <th scope="row" colspan="4" class="text-center">No Overdue Payroll items</th>  
    
        


    </tr>
<%
    else
'response.Write("No Delayed Payroll items") 
     %>
<tr>
      <th scope="row" colspan="5" class="text-center">No Delayed Payroll items</th>  
    
        


    </tr>
<%
  
    end if
rs.close
else
%>
<%rs.pagesize = 5
TotalPages = rs.PageCount
TotalFound = rs.recordcount
rs.MoveFirst
rs.AbsolutePage = page
ps=rs.pagesize %>
<%
    cnt=1
    do while not rs.eof and cnt <= ps
    contid=rs("contid")
    oid=rs("oid")
    lcode=rs("lcode")
    tp=rs("tp")
    period=rs("period")
    cyear=rs("cyear")
    lastdate=rs("lastdate")
    oname=rs("oname")
    lname=rs("lname")
    status=rs("status")
    depdate=rs("depdate")
    depdate1=showdate(depdate,"dmy")
    amount=rs("amount")
    chqno=rs("chqno")
    chqdate=rs("chqdate")
    chqdate1=showdate(chqdate,"dmy")
    speriod = monthname(period)&"-"&cyear
    if status = 0 then sstatus = "<font color  = Navy><b>U</b></font>"
    if status = 1 then sstatus = "<font color  = green><b>C</b></font>"
    if status = 2 then sstatus = "<font color  = red><b>NC</b></font>"
    durl="edcont.asp?contid="&contid
    'atit=lname&" - "&tp&" "&speriod
    %>
<%if aa="1"   then %>
    <tr>
      <th scope="row"><A target = _blank href = "<%=durl%>"><%=lname%></a></th>  
    
        <td><%=tp%></td>
        <td><%=showdate(lastdate,"dmy")%></td>
            <td><%=datediff("d",lastdate,now())%></td>


    </tr>
  <%else %>
 <tr>
        <th scope="row"><A target = _blank href = "<%=durl%>"><%=lname%></a></th>  
    
        <td><%=tp%></td>
        <td><%=showdate(lastdate,"dmy")%></td>
            <td><%=datediff("d",lastdate,depdate)%></td>
     <td><%=showdate(depdate,"dmy")%></td>
  
    </tr>
<%end if %>  


    <%cnt=cnt+1
    rs.movenext
    loop
        rs.close
end if 'if rs.eof

%>

<%end function%>



<%function oret(aa)
page=getreq("page")
if page = "" then page = 1
' returns
asq="select a.rtid, a.oid ,a.lcode, ryear, lastdate,status, depdate, a.rcode, d.rtitle, d.rdesc, b.oname, c.lname from ncret a, ncmorg b, ncmloc c, nctempret d where rtid <> 0"
'asq3=asq3&" and (status <> 0 or status = 0 and (lastdate-getdate() <= 15))"
     
asq=asq&" and a.oid = b.oid and a.oid = c.oid and a.lcode = c.lcode and a.rcode=d.rcode"
asq=asq&" and c.lcode in  (select lcode from ncumap where oid = a.oid and uno = "&uno&") and ("
if returns="on" then TSQ=TSQ&"rtype <> '' or "
TSQ=TSQ&" rtype = '')"
 if aa="1" then
    tsq=tsq&" and status = 2 and lastdate < getdate() "
    else
    tsq=tsq&" and status='1' and ((depdate='') or (depdate > lastdate)) "
   end if
if locsql <> "" then tsq=tsq&locsql

if scyear <> 0 then tsq=tsq&" and ryear = "&scyear
if scperiod <> "0" then
 if scperiod = "H1" then tsq=tsq&" and month(lastdate) <= 6"
 if scperiod = "H2" then tsq=tsq&" and month(lastdate) > 6"
 if scperiod = "Q1" then tsq=tsq&" and month(lastdate) between 1 and 3"
 if scperiod = "Q2" then tsq=tsq&" and month(lastdate) between 4 and 6"
 if scperiod = "Q3" then tsq=tsq&" and month(lastdate) between 7 and 9"
 if scperiod = "Q4" then tsq=tsq&" and month(lastdate) between 10 and 12"
  if scperiod = "M1" then tsq=tsq&" and month(lastdate) = 1"
  if scperiod = "M2" then tsq=tsq&" and month(lastdate) = 2"
  if scperiod = "M3" then tsq=tsq&" and month(lastdate) = 3"
  if scperiod = "M4" then tsq=tsq&" and month(lastdate) = 4"
  if scperiod = "M5" then tsq=tsq&" and month(lastdate) = 5"
  if scperiod = "M6" then tsq=tsq&" and month(lastdate) = 6"
  if scperiod = "M7" then tsq=tsq&" and month(lastdate) = 7"
  if scperiod = "M8" then tsq=tsq&" and month(lastdate) = 8"
  if scperiod = "M9" then tsq=tsq&" and month(lastdate) = 9"
  if scperiod = "M10" then tsq=tsq&" and month(lastdate) = 10"
  if scperiod = "M11" then tsq=tsq&" and month(lastdate) = 11"
  if scperiod = "M12" then tsq=tsq&" and month(lastdate) = 12"
end if

'if srch <> "" then tsq=tsq&" and (lname like '%"&srch&"%' or oname like '%"&srch&"%' or rtitle like '%"&srch&"%' or rdesc like '%"&srch&"%')"

asq=asq&tsq
asq=asq&" order by lastdate, oid, lname "
%>


<%
    
'response.end

    rs.CursorLocation = adUseClient
rs.open asq, conndb
   ' response.write(asq)
if rs.eof then
   
   
   if aa="1" then
'response.Write("No Overdue Payroll items") 
    %>
<tr>
      <th scope="row" colspan="4" class="text-center">No Overdue Liasion items</th>  
    
        


    </tr>
<%
    else
'response.Write("No Delayed Payroll items") 
     %>
<tr>
      <th scope="row" colspan="5" class="text-center">No Delayed Liasion items</th>  
    
        


    </tr>
<%
  
    end if
rs.close
else
%>
<%rs.pagesize = 5

TotalPages = rs.PageCount
TotalFound = rs.recordcount
rs.MoveFirst
rs.AbsolutePage = page
ps=rs.pagesize
     %>
<%
    cnt=1
    do while not rs.eof and cnt <= ps
    rtid=rs("rtid")
    oid=rs("oid")
    lcode=rs("lcode")
    rcode=rs("rcode")
    rtitle=rs("rtitle")
    rdesc=rs("rdesc")
    ryear=rs("ryear")
    lastdate=rs("lastdate")
    oname=rs("oname")
    lname=rs("lname")
    status=rs("status")
    depdate=rs("depdate")
    depdate1=showdate(depdate,"dmy")
    if status = 0 then sstatus = "<font color  = Navy><b>U</b></font>"
    if status = 1 then sstatus = "<font color  = green><b>C</b></font>"
    if status = 2 then sstatus = "<font color  = red><b>NC</b></font>"
    durl="edret.asp?rtid="&rtid

    %>
<%if aa="1"   then %>
    <tr>
        <th scope="row"><A target = _blank href = "<%=durl%>"><%=lname%></a></th>  
       
        <td><%=rtitle%></td>
        <td><%=showdate(lastdate,"dmy")%></td>
        <td><%=datediff("d",lastdate,now())%></td>

  
    </tr>
    <%else %>
 <tr>
         <th scope="row"><A target = _blank href = "<%=durl%>"><%=lname%></a></th> 
   
        <td><%=rtitle%></td>
        <td><%=showdate(lastdate,"dmy")%></td>
        <td><%=datediff("d",lastdate,depdate)%></td>
     <td><%=showdate(depdate,"dmy")%></td>
  
    </tr>
<%end if %>  


    <%cnt=cnt+1
    rs.movenext
    loop
         rs.close
end if 'if rs.eof

%>

<%end function%>