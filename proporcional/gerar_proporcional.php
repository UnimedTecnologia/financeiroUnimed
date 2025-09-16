<?php
// gerar_proporcional.php
set_time_limit(3600);
ini_set('memory_limit', '768M');

if (!file_exists(__DIR__ . '/../vendor/autoload.php')) {
    die("Erro: vendor/autoload.php não encontrado. Rode 'composer install' no diretório do projeto ou verifique a pasta vendor.");
}

require __DIR__ . '/../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// === Recebe parâmetros do POST ===
$modalidade     = $_POST['cdModalidadeCopart']   ?? null;
$proposta       = $_POST['nrPropostaCopart']    ?? null;
$ano            = $_POST['anoCopart']           ?? null;
$mes            = $_POST['mesCopart']           ?? null;
$cdContratante  = $_POST['cdContratanteCopart'] ?? null;
$eventoStr      = $_POST['eventoProp']          ?? '';

if (!$modalidade || !$proposta || !$ano || !$mes || !$cdContratante || !$eventoStr) {
    die("Informe todos os parametros");
}

// === Converte eventos (string -> array) ===
$eventos = array_map('trim', explode(',', $eventoStr));
$eventos = array_filter($eventos, 'is_numeric'); // só números

if (empty($eventos)) {
    die("Nenhum evento válido informado.");
}

// placeholders dinâmicos :evento0, :evento1...
$placeholders = [];
foreach ($eventos as $i => $ev) {
    $placeholders[] = ":evento{$i}";
}
$inClause = implode(',', $placeholders);

// === Conexão Oracle ===
require_once "../config/AW00DB.php";
require_once "../config/oracle.class.php";
require_once "../config/AW00MD.php";

// === SQL ===
$sql = "
SELECT ANO,MES,MODALIDADE,TERMO,PROPOSTA,TITULAR,CD_USUARIO,NOME,CARTEIRA,to_char(DT_INCLUSAO_PLANO, 'dd/mm/yyyy') as DT_INCLUSAO_PLANO, 
       to_char(DT_NASCIMENTO, 'dd/mm/yyyy') as DT_NASCIMENTO, IDADE,CPF,NR_IDENTIDADE,PIS,DT_EXCLUSAO_PLANO,NOME_MAE,TP_PESSOA,
       DS_GRAU_PARENTESCO,TP_PLANO,ENDERECO,BAIRRO,CIDADE,ESTADO,CEP,DECL_NASCIDO_VIVO,
        sum(VL_MENSALIDADE) MENSALIDADE 
  FROM gp.V_VLBENEF_MENSALIDADE
 WHERE modalidade   = :modalidade
   AND termo        = :termo
   AND mes          = :mes
   AND cd_contratante = :contratante
   AND CD_EVENTO IN ($inClause)
   AND ano = :ano
   GROUP BY 
ANO,MES,MODALIDADE,TERMO,PROPOSTA,TITULAR,CD_USUARIO,NOME,CARTEIRA,DT_INCLUSAO_PLANO, 
DT_NASCIMENTO,IDADE,CPF,NR_IDENTIDADE,PIS,DT_EXCLUSAO_PLANO,NOME_MAE,TP_PESSOA,DS_GRAU_PARENTESCO,
TP_PLANO,ENDERECO,BAIRRO,CIDADE,ESTADO,CEP,DECL_NASCIDO_VIVO
";

$stid = oci_parse($conn, $sql);

// binds fixos
oci_bind_by_name($stid, ":modalidade",    $modalidade);
oci_bind_by_name($stid, ":termo",         $proposta);
oci_bind_by_name($stid, ":contratante",   $cdContratante);
oci_bind_by_name($stid, ":ano",           $ano);
oci_bind_by_name($stid, ":mes",           $mes);

// binds dinâmicos para eventos
foreach ($eventos as $i => $ev) {
    oci_bind_by_name($stid, ":evento{$i}", $eventos[$i]);
}

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

// === Cria planilha ===
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Proporcional');

$rowData = oci_fetch_assoc($stid);
if (!$rowData) {
    if (!headers_sent()) {
        header('Content-Type: application/json; charset=utf-8');
    }
    echo json_encode([
        "success" => false,
        "message" => "Nenhum dado encontrado para os parâmetros informados."
    ]);
    exit;
}

// Cabeçalhos
$headers = array_keys($rowData);
$col = 1;
foreach ($headers as $h) {
    $cell = colLetter($col) . '1';
    $sheet->setCellValue($cell, $h);
    $col++;
}

// Preenche dados
$row = 2;
do {
    $col = 1;
    foreach ($headers as $field) {
        $value = $rowData[$field] ?? '';
        $cell = colLetter($col) . $row;
        $sheet->setCellValue($cell, $value);
        $col++;
    }
    $row++;
} while ($rowData = oci_fetch_assoc($stid));

// libera
oci_free_statement($stid);
oci_close($conn);

// Auto width colunas
$totalCols = count($headers);
for ($i = 1; $i <= $totalCols; $i++) {
    $sheet->getColumnDimension(colLetter($i))->setAutoSize(true);
}

// Download
$filename = "proporcional{$ano}_{$mes}.xlsx";
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;
