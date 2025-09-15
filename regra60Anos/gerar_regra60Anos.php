<?php
// gerar_fluxo.php (versão ajustada para evitar setCellValueByColumnAndRow)
set_time_limit(600);
ini_set('memory_limit', '768M');

if (!file_exists(__DIR__ . '/../vendor/autoload.php')) {
    die("Erro: vendor/autoload.php não encontrado. Rode 'composer install' no diretório do projeto ou verifique a pasta vendor.");
}

require __DIR__ . '/../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Recebe parâmetros do POST (inputs type=date)
// $data_inicial = $_POST['data_inicial'] ?? null;
// $data_final   = $_POST['data_final'] ?? null;
// if (!$data_inicial || !$data_final) {
//     die("Período inválido. Informe Data Inicial e Data Final.");
// }

// $data_ini = date("d/m/Y", strtotime($data_inicial));
// $data_fim = date("d/m/Y", strtotime($data_final));

// === Conexão Oracle (ajuste usuário/senha/conn string) ===
// Conexão Oracle
require_once "../config/AW00DB.php";
require_once "../config/oracle.class.php";
require_once "../config/AW00MD.php";

// === SQL com binds ===
$sql = "
SELECT US.NM_USUARIO, TO_CHAR(US.DT_NASCIMENTO, 'DD/MM/YYYY') as DT_NASCIMENTO, IDADE(US.DT_NASCIMENTO, SYSDATE) AS IDADE, US.CD_MODALIDADE, MD.DS_MODALIDADE, 
       US.NR_PROPOSTA, US.NR_TER_ADESAO, US.CD_USUARIO, PT.CD_PLANO, PS.NM_PLANO, PT.CD_TIPO_PLANO, TS.NM_TIPO_PLANO, 
       TS.NM_TIPO_PLANO_REDUZ, CT.NM_CONTRATANTE, RM.DES_REGRA
  FROM GP.USUARIO US INNER JOIN GP.PROPOST               PT  ON US.CD_MODALIDADE   = PT.CD_MODALIDADE
                                                            AND US.NR_PROPOSTA     = PT.NR_PROPOSTA
                                                            AND US.LOG_17         <> 1
                      LEFT JOIN GP.CONTRAT               CT  ON PT.CD_CONTRATANTE  = CT.CD_CONTRATANTE
                     INNER JOIN GP.MODALID               MD  ON US.CD_MODALIDADE   = MD.CD_MODALIDADE
                     INNER JOIN GP.PLA_SAU               PS  ON US.CD_MODALIDADE   = PS.CD_MODALIDADE
                                                            AND PT.CD_PLANO        = PS.CD_PLANO
                     INNER JOIN GP.TI_PL_SA              TS  ON US.CD_MODALIDADE   = TS.CD_MODALIDADE
                                                            AND PT.CD_PLANO        = TS.CD_PLANO
                                                            AND PT.CD_TIPO_PLANO   = TS.CD_TIPO_PLANO
                      LEFT JOIN GP.REGRA_MENSLID_PROPOST RP  ON US.CD_MODALIDADE   = RP.CD_MODALIDADE
                                                            AND US.NR_PROPOSTA     = RP.NR_PROPOSTA
                                                            AND US.CD_USUARIO      = RP.CD_USUARIO
                      LEFT JOIN GP.REGRA_MENSLID         RM  ON RP.CDD_REGRA       = RM.CDD_REGRA
 WHERE US.CD_MODALIDADE = 40 -- COLETIVO ADESAO JUNDIAI
   AND EXTRACT(MONTH FROM US.DT_NASCIMENTO) = CASE WHEN EXTRACT(MONTH FROM SYSDATE) + 1 = 13 THEN 1 ELSE EXTRACT(MONTH FROM SYSDATE) + 1 END -- MES SUBSEQUENTE
   AND EXTRACT(YEAR  FROM US.DT_NASCIMENTO) = EXTRACT(YEAR  FROM SYSDATE) - 60 -- 60 ANOS
 GROUP BY US.NM_USUARIO, US.DT_NASCIMENTO, US.CD_MODALIDADE, MD.DS_MODALIDADE, US.NR_PROPOSTA, US.NR_TER_ADESAO, US.CD_USUARIO,
          PT.CD_PLANO, PS.NM_PLANO, PT.CD_TIPO_PLANO, TS.NM_TIPO_PLANO, TS.NM_TIPO_PLANO_REDUZ, CT.NM_CONTRATANTE, RM.DES_REGRA
 ORDER BY US.CD_MODALIDADE, US.NR_PROPOSTA, US.CD_USUARIO, PT.CD_PLANO, PT.CD_TIPO_PLANO, US.NM_USUARIO
";

$stid = oci_parse($conn, $sql);
// oci_bind_by_name($stid, ":data_ini", $data_ini);
// oci_bind_by_name($stid, ":data_fim", $data_fim);
$r = @oci_execute($stid);
if (!$r) {
    $err = oci_error($stid);
    die("Erro ao executar query: " . ($err['message'] ?? json_encode($err)));
}

// === Helpers ===
function colLetter($col) {
    $col--; // 1 => A
    $letter = '';
    while ($col >= 0) {
        $letter = chr($col % 26 + 65) . $letter;
        $col = intdiv($col, 26) - 1;
    }
    return $letter;
}

$headers = [
    "NM_USUARIO","DT_NASCIMENTO","IDADE", "CD_MODALIDADE", "DS_MODALIDADE", 
    "NR_PROPOSTA", "NR_TER_ADESAO", "CD_USUARIO", "CD_PLANO", "NM_PLANO", "CD_TIPO_PLANO", "NM_TIPO_PLANO", 
    "NM_TIPO_PLANO_REDUZ", "NM_CONTRATANTE", "DES_REGRA"
];

// === Cria planilha e escreve cabeçalhos usando coordenadas A1, B1, ... ===
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Fluxo de Caixa');

$col = 1;
foreach ($headers as $h) {
    $cell = colLetter($col) . '1';
    $sheet->setCellValue($cell, $h);
    $col++;
}

// === Preenche linhas ===
$row = 2;
while ($rowData = oci_fetch_assoc($stid)) {
    $col = 1;
    foreach ($headers as $field) {
        $value = $rowData[$field] ?? '';
        $cell = colLetter($col) . $row;
        $sheet->setCellValue($cell, $value);
        $col++;
    }
    $row++;
}


// libera e fecha
oci_free_statement($stid);
oci_close($conn);

// Ajusta colunas (auto width por letra)
$totalCols = count($headers);
for ($i = 1; $i <= $totalCols; $i++) {
    $sheet->getColumnDimension(colLetter($i))->setAutoSize(true);
}

// Gera e envia arquivo
$filename = "regra60Anos.xlsx";
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;
