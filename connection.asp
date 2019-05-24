
<% 
Set objekConn=Server.CreateObject("ADODB.Connection")
objekConn.Open "dsn=MPSPV01;uid=kehadiran;pwd=kehadiran;"
''objConn.Open  "Driver={Microsoft ODBC for Oracle};Server=//10.10.1.37/mpspv1.mpsp.gov.my;uid=kehadiran;pwd=kehadiran;"
'objekConn.Open "dsn=mpspv1;uid=kehadiran;pwd=kehadiran;"

  
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "dsn=mpspDSN1;uid=payroll;pwd=gaji;"
%>




