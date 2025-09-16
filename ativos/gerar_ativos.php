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
$modalidade = $_POST['cdModalidadeAtivos'] ?? null;
$proposta   = $_POST['nrPropostaAtivos'] ?? null;
$ano        = $_POST['anoAtivos'] ?? null;
$mes        = $_POST['mesAtivos'] ?? null;
if (!$modalidade || !$proposta || !$ano  || !$mes ) {
    die("Informe todos os parametros");
}

// Transforma em array, limpa espaços e só deixa números
$propostas = array_filter(array_map('trim', explode(',', $proposta)), 'is_numeric');

// Garante que tem pelo menos uma
if (empty($propostas)) {
    die("Nenhuma proposta válida informada");
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

// === SQL com binds ===
$sql = "
SELECT u.CD_MODALIDADE AS MODALIDADE, u.NR_PROPOSTA AS PROPOSTA, u2.NM_USUARIO AS TITULAR, u2.CD_CPF AS CPF_TITULAR, 
u.cd_usuario, CI.CD_UNIMED||ci.CD_CARTEIRA_INTEIRA AS CARTEIRINHA, u.NM_USUARIO AS NOME, to_char(u.DT_INCLUSAO_PLANO,'dd/mm/yyyy') AS DATA_INCLUSAO_PLANO,
TO_CHAR(u.DT_NASCIMENTO, 'DD/MM/YYYY') AS DATA_NASCIMENTO, TRUNC(MONTHS_BETWEEN(SYSDATE, u.DT_NASCIMENTO) / 12) AS IDADE, 
CASE 
	WHEN u.lg_sexo = 0 THEN 'F'
	WHEN u.lg_sexo = 1 THEN 'M'
END AS SEXO, U.CD_CPF CPF, u.NR_IDENTIDADE AS IDENTIDADE, u.CD_PIS_PASEP PIS, u.DT_EXCLUSAO_PLANO DATA_DE_EXCLUSAO_DO_PLANO, u.NM_MAE NOME_MAE, 'PJE' AS TIPO_PESSOA,
gp.ds_grau_parentesco GRAU_PARENTESCO, 'ENFERMARIA' AS TIPO_PLANO,
u.EN_RUA AS ENDERECO, u.EN_BAIRRO AS BAIRRO, DC.NM_CIDADE CIDADE, u.EN_UF AS UF, u.EN_CEP AS CEP, PF.CHAR_2 as DECLARACAO_NASCIDO_VIVO, VB.VL_TOTAL VALOR
	FROM gp.usuario u
	LEFT JOIN gp.USUARIO u2 
       ON u2.CD_USUARIO = u.CD_TITULAR
      AND u2.CD_MODALIDADE = u.CD_MODALIDADE   
      AND u2.NR_PROPOSTA = u.NR_PROPOSTA 
  left join GP.PESSOA_FISICA PF    
      on pf.id_pessoa = u.id_pessoa  
  left join gp.VLBENEF vb  
      ON vb.cd_modalidade = U.CD_MODALIDADE      
      AND VB.CD_USUARIO = U.CD_USUARIO 
      AND VB.NR_TER_ADESAO = U.NR_TER_ADESAO      
      AND U.LOG_17  <> 1     
	left JOIN gp.DZCIDADE dc ON dc.CD_CIDADE = u.CD_CIDADE 
	left JOIN gp.car_ide ci ON ci.CD_USUARIO = u.CD_USUARIO AND ci.CD_MODALIDADE = u.CD_MODALIDADE AND ci.NR_TER_ADESAO = u.NR_TER_ADESAO 
	INNER JOIN gp.GRA_PAR gp ON gp.CD_GRAU_PARENTESCO = u.CD_GRAU_PARENTESCO 
	WHERE u.CD_MODALIDADE = :modalidade AND u.NR_PROPOSTA IN ($placeholdersStr) AND VB.CD_EVENTO in('10','11') AND VB.AA_REFERENCIA = :ano AND  VB.MM_REFERENCIA = :mes
    AND VB.VL_TOTAL <> 0
  GROUP BY u.CD_MODALIDADE,u.NR_PROPOSTA,u2.NM_USUARIO,u2.CD_CPF,u.cd_usuario,CI.CD_UNIMED,ci.CD_CARTEIRA_INTEIRA,
           u.NM_USUARIO,u.DT_INCLUSAO_PLANO,u.DT_NASCIMENTO,u.DT_NASCIMENTO,u.lg_sexo,U.CD_CPF,u.NR_IDENTIDADE,
           u.CD_PIS_PASEP,u.DT_EXCLUSAO_PLANO,u.NM_MAE,gp.ds_grau_parentesco,u.EN_RUA,u.EN_BAIRRO,
           DC.NM_CIDADE,u.EN_UF,u.EN_CEP,PF.CHAR_2,VB.VL_TOTAL
";

$stid = oci_parse($conn, $sql);
oci_bind_by_name($stid, ":modalidade", $modalidade);
// oci_bind_by_name($stid, ":proposta", $proposta);
oci_bind_by_name($stid, ":ano", $ano);
oci_bind_by_name($stid, ":mes", $mes);
// Binds dinâmicos
foreach ($propostas as $i => $val) {
    oci_bind_by_name($stid, ":p{$i}", $propostas[$i]);
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

// === Cria planilha e escreve cabeçalhos usando coordenadas A1, B1, ... ===
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Ativos');


$rowData = oci_fetch_assoc($stid);
if (!$rowData) {
    // limpa qualquer header de Excel que já tenha sido enviado
    if (!headers_sent()) {
        header('Content-Type: application/json; charset=utf-8');
    }
    echo json_encode([
        "success" => false,
        "message" => "Nenhum dado encontrado para os parâmetros informados."
    ]);
    exit;
}


$headers = array_keys($rowData); // nomes das colunas (UPPERCASE)

$col = 1;
foreach ($headers as $h) {
    $cell = colLetter($col) . '1';
    $sheet->setCellValue($cell, $h);
    $col++;
}

// Preenche primeira linha (já lida) e as demais
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


// libera e fecha
oci_free_statement($stid);
oci_close($conn);

// Ajusta colunas (auto width por letra)
$totalCols = count($headers);
for ($i = 1; $i <= $totalCols; $i++) {
    $sheet->getColumnDimension(colLetter($i))->setAutoSize(true);
}

// Gera e envia arquivo
$filename = "ativos{$ano}_{$mes}.xlsx";
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;
