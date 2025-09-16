<?php
// gerar_fluxo.php (versão ajustada para evitar setCellValueByColumnAndRow)
set_time_limit(3600);
ini_set('memory_limit', '768M');

if (!file_exists(__DIR__ . '/../vendor/autoload.php')) {
    die("Erro: vendor/autoload.php não encontrado. Rode 'composer install' no diretório do projeto ou verifique a pasta vendor.");
}

require __DIR__ . '/../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Recebe parâmetros do POST (inputs type=date)
$modalidade         = $_POST['cdModalidadeCopart'] ?? null;
$proposta           = $_POST['nrPropostaCopart'] ?? null;
$ano                = $_POST['anoCopart'] ?? null;
$mes                = $_POST['mesCopart'] ?? null;
$cdContratante      = $_POST['cdContratanteCopart'] ?? null;
if (!$modalidade || !$proposta || !$ano  || !$mes || !$cdContratante) {
    die("Informe todos os parametros");
}

// Conexão Oracle
require_once "../config/AW00DB.php";
require_once "../config/oracle.class.php";
require_once "../config/AW00MD.php";

// === SQL com binds ===
$sql = "
SELECT TITULAR.NM_USUARIO TITULAR, TITULAR.CD_CPF CPF_TITULAR, PP.CD_CONTRATANTE CONTRATANTE, ct.nm_contratante EMPRESA,PF.NM_PESSOA USUARIO, 'ENFERMARIA' TIPO_DE_PLANO,
CASE WHEN US.lg_sexo = 0 THEN 'F' WHEN US.lg_sexo = 1 THEN 'M' END AS SEXO, to_char(PF.DT_NASCIMENTO, 'dd/mm/yyyy') as NASCIMENTO, 
GU.DS_GRAU_PARENTESCO, PF.CD_CPF CPF, MP.NR_TER_ADESAO, MP.NR_DOC_ORIGINAL, MP.CD_UNIDADE||MP.CD_CARTEIRA_USUARIO CARTEIRINHA, to_char(MP.DT_REALIZACAO, 'dd/mm/yyyy') as UTILIZACAO,
        MP.QT_PROCEDIMENTOS QUANTIDADE, AP.CDPROCEDIMENTOCOMPLETO PROCEDIMENTO, MP.NM_PRESTADOR_VALIDA PRESTADOR, PP.VL_EVENTO VALOR, ROUND(PP.PC_EVENTO)||'%' PERCENTUAL 
 FROM GP.FATEVECO PP INNER JOIN GP.MOVIPROC MP  ON MP.CD_UNIDADE             = PP.CD_UNIDADE
                                                AND MP.CD_UNIDADE_PRESTADORA = PP.CD_UNIDADE_PRESTADOR
                                                AND MP.CD_TRANSACAO          = PP.CD_TRANSACAO
                                                AND MP.NR_SERIE_DOC_ORIGINAL = PP.NR_SERIE_DOC_ORIGINAL
                                                AND MP.NR_DOC_ORIGINAL       = PP.NR_DOC_ORIGINAL
                                                AND MP.NR_DOC_SISTEMA        = PP.NR_DOC_SISTEMA
                                                AND MP.NR_PROCESSO           = PP.NR_PROCESSO
                                                AND MP.NR_SEQ_DIGITACAO      = PP.NR_SEQ_DIGITACAO
                      INNER JOIN GP.AMBPROCE AP  ON MP.CD_ESP_AMB            = AP.CD_ESP_AMB
                                                AND MP.CD_GRUPO_PROC_AMB     = AP.CD_GRUPO_PROC_AMB
                                                AND MP.CD_PROCEDIMENTO       = AP.CD_PROCEDIMENTO
                                                AND MP.DV_PROCEDIMENTO       = AP.DV_PROCEDIMENTO
                      INNER JOIN GP.USUARIO US    ON  US.CD_MODALIDADE           = MP.CD_MODALIDADE 
                                                AND US.NR_TER_ADESAO      = MP.NR_TER_ADESAO  
                                                AND  US.CD_USUARIO              = MP.CD_USUARIO
                                                AND US.LOG_17            <> 1 -- USUARIO EVENTUAL
            left JOIN GP.USUARIO TITULAR ON TITULAR.CD_MODALIDADE   = US.CD_MODALIDADE  
                                                AND TITULAR.NR_TER_ADESAO     = US.NR_TER_ADESAO  
                                                AND TITULAR.cd_usuario        = US.cd_titular             
            INNER JOIN GP.PESSOA_FISICA PF ON PF.ID_PESSOA           = US.ID_PESSOA 

                      INNER JOIN GP.GRA_PAR       GU ON GU.CD_GRAU_PARENTESCO = US.CD_GRAU_PARENTESCO           
                      INNER JOIN GP.PROPOST PT   ON PT.CD_MODALIDADE         = US.CD_MODALIDADE
                                                AND PT.NR_PROPOSTA           = US.NR_PROPOSTA
                      INNER JOIN GP.TI_PL_SA TP  ON TP.CD_MODALIDADE         = US.CD_MODALIDADE
                                                AND TP.CD_PLANO              = PT.CD_PLANO
                                                AND TP.CD_TIPO_PLANO         = PT.CD_TIPO_PLANO
                      INNER JOIN GP.CONTRAT CT   ON PT.CD_CONTRATANTE        = CT.CD_CONTRATANTE                          
  WHERE PP.CD_EVENTO IN ( 70, 71 ) 
     AND PP.CD_CONTRATANTE = :contratante AND PP.AA_REFERENCIA  = :ano AND PP.MM_REFERENCIA  = :mes  
     AND US.CD_MODALIDADE = :modalidade AND us.NR_PROPOSTA = :proposta ";

$stid = oci_parse($conn, $sql);
oci_bind_by_name($stid, ":contratante", $cdContratante);
oci_bind_by_name($stid, ":modalidade", $modalidade);
oci_bind_by_name($stid, ":proposta", $proposta);
oci_bind_by_name($stid, ":ano", $ano);
oci_bind_by_name($stid, ":mes", $mes);
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
$sheet->setTitle('Coparticipação');


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
$filename = "coparticipacao{$ano}_{$mes}.xlsx";
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;
