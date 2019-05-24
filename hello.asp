<%
Set objekConn=Server.CreateObject("ADODB.Connection")
objekConn.Open "dsn=mpsp11g;uid=kehadiran;pwd=kehadiran;"
''objConn.Open  "Driver={Microsoft ODBC for Oracle};Server=//10.10.1.37/mpspv1.mpsp.gov.my;uid=kehadiran;pwd=kehadiran;"
'objekConn.Open "dsn=mpspv1;uid=kehadiran;pwd=kehadiran;"

  
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "dsn=12c;uid=payroll;pwd=gaji;"



s1="select hadir1,hadir2 from payroll.proses_ot where no_pekerja=40956 and tkh_ot=to_date('03012016','ddmmyyyy')"
set qs1 = objConn.Execute(s1)

response.write("<br>hadir1--->"& qs1("hadir1") & "  hadir2--->"& qs1("hadir2"))


q9="select nvl(time1,'0') time1,nvl(time2,'0') time2,nvl(time3,'0') time3,nvl(time4,'0') time4 from kehadiran.sispot_kehadiran where "
q9 = q9 & " no_pekerja=40956 and tarikh like to_date('03012016','ddmmyyyy') "
set oq9 = objekConn.Execute(q9)

response.write("hadir1--->"& oq9("time1") &" - "& oq9("time2") & "   hadir2--->"& oq9("time3") &" - "& oq9("time4") &"<br>")

%>