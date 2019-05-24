<script language="Javascript">
var new_win

function open_staf(p_item)
	{
	 if(!win_opened())
	  {
	   new_win = window.open("staff_ot.asp?pemilik="+p_item,"", "height=500, width=700, left=45,top=10")  	    	   	   
	  } 
	 else
	  {
	   new_win.focus()
	  }
    }
	
function open_staf_development(p_item,m_item)
	{
	alert(p_item);
	alert(m_item);
	 if(!win_opened())
	  {
	   new_win = window.open("staff_ot.asp?blnthnot="+m_item+"&pemilik="+p_item,"", "height=500, width=700, left=45,top=10")  	    	   	   
	  } 
	 else
	  {
	   new_win.focus()
	  }
    }	
	
function win_opened()
	{
	 if (!new_win)
	  {
	   return false
	  } 
	 else if (new_win.closed)
	  {
	   return false
	  }
	 else
	  {
	   return true
	  }
	}
	var new_cagar


	
function win_opened()
	{
	 if (!new_cagar)
	  {
	   return false
	  } 
	 else if (new_cagar.closed)
	  {
	   return false
	  }
	 else
	  {
	   return true
	  }
	}	
	
</script>