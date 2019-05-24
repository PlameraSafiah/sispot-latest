<!--#include file="connection.asp" -->
<%
a="SELECT table_name, column_name, data_type, data_length FROM USER_TAB_COLUMNS WHERE table_name = upper('lokasi_caw')"
set qa = objekConn.Execute(a)

do while not qa.eof 
	response.Write(qa("column_name") & ", "& qa("data_type") &", "& qa("data_length") &"<br>")
qa.movenext
loop

b="SELECT * FROM lokasi_caw order by NAMA_LOKASI"
set qb = objekConn.Execute(b)

do while not qb.eof 
	'response.Write(qb("KODLOKASI_CAW") & ", "& qb("NAMA_LOKASI") &"<br>")

	
	c="select count(*) as totalstaff,lokasi from kehadiran.paymas where kodlokasi_caw='"& qb("kodlokasi_caw") &"' group by lokasi order by lokasi"
	set qc = objekConn.Execute(c)
	do while not qc.eof
	response.Write(qb("NAMA_LOKASI") & ", "& qc("totalstaff") &", "& qc("lokasi") &"<br>")

	qc.movenext
	loop
qb.movenext
loop

	c2="select count(*) as totalstaff from kehadiran.paymas where kodlokasi_caw is not null"
	set qc2 = objekConn.Execute(c2)
	response.Write(qc2("totalstaff")&"<br>")
%>