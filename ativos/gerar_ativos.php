<?php
// Inicia buffer imediatamente
ob_start();

// gerar_fluxo.php (corrigido)
set_time_limit(600);
ini_set('memory_limit', '768M');

// Desativa erros para não corromper o output
error_reporting(0);
ini_set('display_errors', 0);

if (!file_exists(__DIR__ . '/../vendor/autoload.php')) {
    ob_end_clean();
    if (!headers_sent()) {
        header('Content-Type: application/json; charset=utf-8');
    }
    echo json_encode(["success" => false, "message" => "Erro: vendor/autoload.php não encontrado. Rode 'composer install' no diretório do projeto ou verifique a pasta vendor."]);
    exit;
}

require __DIR__ . '/../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

// Recebe parâmetros do POST
$modalidade = $_POST['cdModalidadeAtivos'] ?? null;
$proposta   = $_POST['nrPropostaAtivos'] ?? null;
$ano        = $_POST['anoAtivos'] ?? null;
$mes        = $_POST['mesAtivos'] ?? null;

if (!$modalidade || !$proposta || !$ano || !$mes) {
    ob_end_clean();
    if (!headers_sent()) {
        header('Content-Type: application/json; charset=utf-8');
    }
    echo json_encode(["success" => false, "message" => "Informe todos os parametros"]);
    exit;
}

// Transforma em array, limpa espaços e só deixa números
$propostas = array_filter(array_map('trim', explode(',', $proposta)), 'is_numeric');
if (empty($propostas)) {
    ob_end_clean();
    if (!headers_sent()) {
        header('Content-Type: application/json; charset=utf-8');
    }
    echo json_encode(["success" => false, "message" => "Nenhuma proposta válida informada"]);
    exit;
}

// Monta placeholders :p0, :p1, ...
$placeholders = [];
foreach ($propostas as $i => $val) {
    $placeholders[] = ":p{$i}";
}
$placeholdersStr = implode(',', $placeholders);

// Conexão Oracle
require_once "../config/AW00DB.php";
require_once "../config/oracle.class.php";
require_once "../config/AW00MD.php";

// Limpa possível output dos includes
ob_clean();

// SQL com binds
$sql = "
SELECT DISTINCT
    u.cd_modalidade       AS modalidade,
    u.nr_proposta         AS proposta,
    u2.nm_usuario         AS titular,
    u2.cd_cpf             AS cpf_titular,
    u.cd_usuario,
    TO_CHAR(ci.cd_unimed) || ci.cd_carteira_inteira AS carteirinha,
    u.nm_usuario          AS nome,
    TO_CHAR(u.dt_inclusao_plano,'DD/MM/YYYY') AS data_inclusao_plano,
    TO_CHAR(u.dt_nascimento, 'DD/MM/YYYY')    AS data_nascimento,
    TRUNC(MONTHS_BETWEEN(SYSDATE, u.dt_nascimento) / 12) AS idade,
    CASE TO_CHAR(u.lg_sexo) WHEN '0' THEN 'F' WHEN '1' THEN 'M' END AS sexo,
    u.cd_cpf,
    u.nr_identidade       AS identidade,
    u.cd_pis_pasep        AS pis,
    u.dt_exclusao_plano   AS data_de_exclusao_do_plano,
    u.nm_mae              AS nome_mae,
    'PJE'                 AS tipo_pessoa,
    gp.ds_grau_parentesco AS grau_parentesco,
    CASE TO_NUMBER(p.cd_tipo_plano) WHEN 1 THEN 'ENFERMARIA' WHEN 2 THEN 'APARTAMENTO' END AS tipo_de_plano,
    u.en_rua              AS endereco,
    u.en_bairro           AS bairro,
    dc.nm_cidade          AS cidade,
    u.en_uf               AS uf,
    u.en_cep              AS cep,
    pf.char_2             AS declaracao_nascido_vivo,
    vb.vl_total           AS valor
FROM gp.usuario u
LEFT JOIN gp.usuario u2
       ON u2.cd_usuario = u.cd_titular
      AND u2.cd_modalidade = u.cd_modalidade
      AND u2.nr_proposta = u.nr_proposta
LEFT JOIN gp.pessoa_fisica pf
       ON pf.id_pessoa = u.id_pessoa
LEFT JOIN gp.vlbenef vb
       ON vb.cd_modalidade = u.cd_modalidade
      AND vb.cd_usuario = u.cd_usuario
      AND vb.nr_ter_adesao = u.nr_ter_adesao
LEFT JOIN gp.dzcidade dc
       ON dc.cd_cidade = u.cd_cidade
LEFT JOIN gp.car_ide ci
       ON ci.cd_usuario = u.cd_usuario
      AND ci.cd_modalidade = u.cd_modalidade
      AND ci.nr_ter_adesao = u.nr_ter_adesao
INNER JOIN gp.gra_par gp
       ON gp.cd_grau_parentesco = u.cd_grau_parentesco
LEFT JOIN gp.propost p
       ON p.cd_modalidade = u.cd_modalidade
      AND p.nr_proposta = u.nr_proposta
      AND p.nr_ter_adesao = u.nr_ter_adesao
WHERE u.cd_modalidade = :modalidade
  AND u.nr_proposta IN ($placeholdersStr)
  AND vb.cd_evento IN ('10','11')
  AND vb.aa_referencia = :ano
  AND vb.mm_referencia = :mes
  AND u.log_17 <> 1
  AND vb.vl_total <> 0
  AND (ci.dt_cancelamento IS NULL OR ci.dt_cancelamento >= TO_DATE(:data_ref, 'YYYY-MM-DD'))
";

$stid = oci_parse($conn, $sql);

// binds
oci_bind_by_name($stid, ":modalidade", $modalidade);
oci_bind_by_name($stid, ":ano", $ano);
oci_bind_by_name($stid, ":mes", $mes);

// Data referência (primeiro dia do mês)
$primeiroDia = sprintf("%04d-%02d-01", $ano, $mes);
oci_bind_by_name($stid, ":data_ref", $primeiroDia);

// binds dinâmicos das propostas
foreach ($propostas as $i => $val) {
    oci_bind_by_name($stid, ":p{$i}", $propostas[$i]);
}

// executa
$r = @oci_execute($stid);
if (!$r) {
    $err = oci_error($stid);
    ob_end_clean();
    if (!headers_sent()) {
        header('Content-Type: application/json; charset=utf-8');
    }
    echo json_encode([
        "success" => false,
        "message" => $err['message'] ?? 'Erro desconhecido'
    ]);
    exit;
}

// helper: converte número em letra de coluna
function colLetter($col) {
    $col--; 
    $letter = '';
    while ($col >= 0) {
        $letter = chr($col % 26 + 65) . $letter;
        $col = intdiv($col, 26) - 1;
    }
    return $letter;
}

// pega todos os resultados
$rows = [];
$fetchCount = 0;
while ($rowData = oci_fetch_assoc($stid)) {
    $rows[] = $rowData;
    $fetchCount++;
}

// contador de registros retornados pelo SELECT
$recordCount = count($rows);

oci_free_statement($stid);
oci_close($conn);

if (empty($rows)) {
    ob_end_clean();
    if (!headers_sent()) {
        header('Content-Type: application/json; charset=utf-8');
    }
    echo json_encode([
        "success" => false,
        "message" => "Nenhum dado encontrado para os parâmetros informados."
    ]);
    exit;
}

// === GERAÇÃO DO EXCEL ===
// Limpa completamente o buffer antes de gerar Excel
ob_end_clean();

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Ativos');

// cabeçalhos
$headers = array_keys($rows[0]);
$col = 1;
foreach ($headers as $h) {
    $cell = colLetter($col) . '1';
    $sheet->setCellValue($cell, $h);
    $col++;
}

// escreve linhas
$row = 2;
foreach ($rows as $rowData) {
    $col = 1;
    foreach ($headers as $field) {
        $value = $rowData[$field] ?? '';
        $cell = colLetter($col) . $row;
        $sheet->setCellValueExplicit($cell, $value, DataType::TYPE_STRING);
        $col++;
    }
    $row++;
}

// ajusta colunas
$totalCols = count($headers);
for ($i = 1; $i <= $totalCols; $i++) {
    $sheet->getColumnDimension(colLetter($i))->setAutoSize(true);
}

// envia Excel
$filename = "ativos{$ano}_{$mes}.xlsx";

// Headers para download do Excel
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');
header('Pragma: public');
header('Expires: 0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;
