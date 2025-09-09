<?php

class oracle{
        var $queries = 0;
        var $user;
        var $pwd;
        var $base;
        var $conn;

 function getUser()
 {
   return $this->user;
 }

    function connect($db_host, $db_user, $db_pwd, $db_name, $pconnect)
    {
      //$this->conn = OCILogon($db_user,$db_pwd,$db_name);
      $this->conn = oci_connect($db_user, $db_pwd, $db_name);
      $this->user = $db_user;
      $this->pwd = $db_pwd;
      $this->base = $db_name;

      //ocilogoff($this->conn);
      //$this->conn = OCILogon($db_user,$db_pwd,$db_name);

      $cmd = "alter session set nls_date_format = 'DD/MM/YYYY'";
      $query = OCI_Parse($this->conn,$cmd);
      OCI_Execute($query);

      $cmd = "alter session set nls_numeric_characters = ',.'";
      $query = OCI_Parse($this->conn,$cmd);
      OCI_Execute($query);

      OCI_commit($this->conn);
      return $this->conn ;
  }

  function query($sql)
  {
    //echo "<!-- <$sql> -->";
    $query = OCI_Parse($this->conn,$sql);
    OCI_Execute($query);
    OCI_commit($this->conn);

    return $query;
  }
  
  function queryWithErrors($sql)
  {
	$errors = array();
    $stmt = OCI_Parse($this->conn, $sql);
    if (!$stmt){
    	$errors = oci_error();
    	throw new Exception("Erro ao executar comando: " . htmlentities($errors['message']));
    }

	$executa = OCI_Execute($stmt);
	if (!$executa){
    	$errors = oci_error($stmt);
    	throw new Exception("Erro ao executar comando: " . htmlentities($errors['message']));
	}
    
    OCI_Commit($this->conn);

    return $stmt;
  }

  function fetch_array($query)
  {
     $query = OCIFetchInto($query,$row,OCI_BOTH);
     return $row;
  }

  function query1($sql)
  {
  /*
      $cmd = "SELECT * FROM NLS_SESSION_PARAMETERS n
                where n.parameter = 'NLS_DATE_FORMAT' ";
      $result = $this->query($cmd);
      $row = $this->fetch_array($result);
      echo "<!-- Vc está com problemas -> $row[VALUE]-->";

    */
   // echo "<!-- [$sql] -->";
    $sql = str_replace(chr(13),' ',$sql);
    $query = OCI_Parse($this->conn,$sql);
    return $query;
  }

  function num_rows($query,$sql)
  {

     $sql = "select Count(*) NUM from ($sql)";

     $query = OCI_Parse($this->conn,$sql);
     OCI_Execute($query);
     ocifetchinto($query, $row, OCI_BOTH);

     return $row[NUM] ;

  }

  function fetch_row($query)
  {
     $query = ocifetchinto($query, $row, OCI_BOTH);
     return $row;
   }

  function insert_id($table)
  {
    $sql = "SELECT max(id) ID FROM $table";
    $query = OCI_Parse($this->conn, $sql);
    OCI_Execute($query);
    ocifetchinto($query, $row, OCI_BOTH);

    return $row[ID];
  }

}

?>
