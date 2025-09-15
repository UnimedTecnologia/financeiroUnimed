<?php
putenv('NLS_LANG=AMERICAN_AMERICA.AL32UTF8'); //* Formatação UTF-8
header('Content-Type: application/json; charset=utf-8'); // Certifique-se de que o PHP avisa ao navegador sobre o charset
ini_set('default_charset', 'UTF-8');                    // Configure o charset padrão para UTF-8 no PHP

$db_host = "10.10.10.23";
$db_name = "INFO";
$db_user = "AUTOWEB";
$db_pwd  = "EBAD";

// Conexão direta com Easy Connect
$db_name = "(DESCRIPTION =
  (ADDRESS = (PROTOCOL = TCP)(HOST = 10.10.10.23)(PORT = 1521))
  (CONNECT_DATA =
    (SERVICE_NAME = info)
  )
)";

$pconnect = 0;


?>
