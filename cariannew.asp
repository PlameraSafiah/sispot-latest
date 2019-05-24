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
	
		if Cint(qc("totalstaff")) > 0 then
		nama_lokasi=mid(qb("NAMA_LOKASI"),2,len(qb("NAMA_LOKASI")))
		
		'insert unit_sispot
		q0="select * from payroll.unit_sispot where upper(nama_unit) like upper('"& nama_lokasi &"') and ptj="& qc("lokasi") &""
		set r0=objConn.Execute(q0)
		if r0.eof then
			a1="select (nvl(max(substr(id_unit,2,4)),0)) + 1 as siri from payroll.unit_sispot  where "
			a1 = a1 & "upper(substr(nama_unit,1,1)) like upper(substr('"& nama_lokasi &"',1,1))"
			set qa1 = objConn.Execute(a1)
			'siri = Cint(qa("siri") + 1			
			
			d1 = "select substr( upper('"& nama_lokasi &"'),1,1)||lpad('"&qa1("siri")&"',4,'0') as unit_id from dual "
			response.write(d1 &"<br>")
			Set od1 = objConn.Execute(d1)
			
			q2="Insert into payroll.unit_sispot (id_unit,nama_unit,ptj) values ('"& od1("unit_id") &"',upper('"& nama_lokasi &"'),"& qc("lokasi") &")"
			'objConn.Execute(q2)
		end if
		
		 
		c1="select no_pekerja from kehadiran.paymas where kodlokasi_caw='"& qb("kodlokasi_caw") &"' and lokasi="& qc("lokasi") &""
		set qc1 = objekConn.Execute(c1)
	do while not qc1.eof
		q01="delete from payroll.unit_kakitangan where no_pekerja = "& qc1("no_pekerja") &" and ptj="& qc("lokasi") &""
		'objConn.Execute(q01)
		
		'check kakitangan ada di jabatan berkenaan
		q0="select no_pekerja,nama from payroll.paymas where no_pekerja="& qc1("no_pekerja") &" and lokasi="& qc("lokasi") &" and status_gaji<8"
		set r0 = objConn.Execute(q0)
			if not r0.bof and not r0.eof then
				'get nama unit
				q02="select upper(nama_unit) as nama_unit from payroll.unit_sispot where ptj="& qc("lokasi") &" and id_unit like '"& od1("unit_id") &"'"
				response.write(q02 &"<br>")
				set oq2 = objConn.Execute(q02)				
				if not oq2.bof and not oq2.eof then
					q2="Insert into payroll.unit_kakitangan (no_pekerja,nama,ptj,id_unit,nama_unit) values "
					q2 = q2 & "("& qc1("no_pekerja") &",'"& replace(r0("nama"),"'","''") &"',"& qc("lokasi") &",'"& od1("unit_id") &"','"& oq2("nama_unit") &"')"
					'objConn.Execute(q2)
				end if
			end if
					qc1.movenext
	loop
	end if
	'response.Write(qb("NAMA_LOKASI") & ", "& qb("NAMA_LOKASI") &"<br>")

	qc.movenext
	loop
qb.movenext
loop
%>