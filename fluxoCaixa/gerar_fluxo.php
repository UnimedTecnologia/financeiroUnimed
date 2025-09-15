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
$data_inicial = $_POST['data_inicial'] ?? null;
$data_final   = $_POST['data_final'] ?? null;
if (!$data_inicial || !$data_final) {
    die("Período inválido. Informe Data Inicial e Data Final.");
}

$data_ini = date("d/m/Y", strtotime($data_inicial));
$data_fim = date("d/m/Y", strtotime($data_final));

// === Conexão Oracle (ajuste usuário/senha/conn string) ===
// Conexão Oracle
require_once "../config/AW00DB.php";
require_once "../config/oracle.class.php";
require_once "../config/AW00MD.php";

// === SQL com binds ===
$sql = "
SELECT 
    COD_ESTAB,
    IND_FLUXO_MOVTO_CX,
    TO_CHAR(DAT_MOVTO_FLUXO_CX, 'DD/MM/YYYY') AS DAT_MOVTO_FLUXO_CX,
    TO_CHAR(DAT_EMIS_DOCTO, 'DD/MM/YYYY')     AS DAT_EMIS_DOCTO,
    COD_MODUL_DTSUL,
    NUM_FLUXO_CX,
    DES_FLUXO_CX,
    TP_FLUXO,
    NIVEL_1,
    NIVEL_2,
    NIVEL_3,
    NIVEL_4,
    DES_TIP_FLUXO_FINANC,
    TP_MOVIMENTO,
    COD_UNID_NEGOC,
    IND_ORIGIN_TIT_AP,
    COD_ESPEC_DOCTO,
    COD_SER_DOCTO,
    COD_TIT_AP,
    NUMIDORIGMOVTOFLUXOCX,
    VL_FLUXO,
    VL_MOVIMENTO,
    COD_PARCELA,
    COD_REFER,
    CDN_FORNECEDOR,
    NOM_PESSOA,
    IND_TRANS_AP_ABREV
FROM ems5.v_FLUXO_CAIXA
WHERE DAT_MOVTO_FLUXO_CX 
      BETWEEN TO_DATE(:data_ini, 'DD/MM/YYYY') 
          AND TO_DATE(:data_fim, 'DD/MM/YYYY') AND TP_MOVIMENTO <> 'PR'
ORDER BY DAT_MOVTO_FLUXO_CX
";

$stid = oci_parse($conn, $sql);
oci_bind_by_name($stid, ":data_ini", $data_ini);
oci_bind_by_name($stid, ":data_fim", $data_fim);
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
    "COD_ESTAB", "IND_FLUXO_MOVTO_CX", "DAT_MOVTO_FLUXO_CX", "DAT_EMIS_DOCTO", "COD_MODUL_DTSUL",
    "NUM_FLUXO_CX", "DES_FLUXO_CX", "TP_FLUXO", "NIVEL_1", "NIVEL_2", "NIVEL_3", "NIVEL_4",
    "DES_TIP_FLUXO_FINANC", "TP_MOVIMENTO", "COD_UNID_NEGOC", "IND_ORIGIN_TIT_AP",
    "COD_ESPEC_DOCTO", "COD_SER_DOCTO", "COD_TIT_AP", "NUMIDORIGMOVTOFLUXOCX",
    "VL_FLUXO", "VL_MOVIMENTO", "COD_PARCELA", "COD_REFER", "CDN_FORNECEDOR",
    "NOM_PESSOA", "IND_TRANS_AP_ABREV"
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

        // Regra: se IND_FLUXO_MOVTO_CX = 'SAI' → valores negativos
        if (in_array($field, ['VL_FLUXO', 'VL_MOVIMENTO'])) {
            $fluxo = $rowData['IND_FLUXO_MOVTO_CX'] ?? '';
            if ($fluxo === 'SAI' && is_numeric($value)) {
                $value = -abs($value); // garante negativo
            } elseif ($fluxo === 'ENT' && is_numeric($value)) {
                $value = abs($value);  // garante positivo
            }
        }

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
$filename = "Fluxo_Caixa_{$data_ini}_a_{$data_fim}.xlsx";
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;
