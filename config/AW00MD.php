<?php

// Criação da conexão com codificação
$conn = oci_connect($db_user, $db_pwd, $db_name, 'UTF8');

if (!$conn) {
    $e = oci_error();
    // echo "Erro de conexão: " . htmlentities($e['message'], ENT_QUOTES, 'UTF-8');
    error_log("Erro de conexão: " . $e['message']);
    // $_SESSION['errodb'] = $e['message'];
    header("Location: ../error?");
    exit;
}

// Inclua classes adicionais
require_once("formata_data.class.php");

?>
