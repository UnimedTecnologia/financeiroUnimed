<?php

set_time_limit(3600);
ini_set('memory_limit', '768M');

if (!file_exists(__DIR__ . '/../vendor/autoload.php')) {
    die("Erro: vendor/autoload.php não encontrado. Rode 'composer install' no diretório do projeto ou verifique a pasta vendor.");
}

require __DIR__ . '/../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Recebe parâmetros do POST
$cdContratante  = $_POST['cdContratante'] ?? null;
$ano            = $_POST['anoAnalitico'] ?? null;
$mes            = $_POST['mesAnalitico'] ?? null;

if (!$cdContratante || !$ano || !$mes) {
    echo json_encode([
        "success" => false,
        "message" => "Informe todos os parâmetros: ano, mês e contratante."
    ]);
    exit;
}

// Converte em array e filtra apenas números
$contratantes = array_filter(array_map('trim', explode(',', $cdContratante)), 'is_numeric');
$meses        = array_filter(array_map('trim', explode(',', $mes)), 'is_numeric');

if (empty($contratantes)) {
    echo json_encode([
        "success" => false,
        "message" => "Nenhum contratante válido informado."
    ]);
    exit;
}
if (empty($meses)) {
    echo json_encode([
        "success" => false,
        "message" => "Nenhum mês válido informado."
    ]);
    exit;
}

// Cria placeholders para binds
$placeholdersContratantes = [];
foreach ($contratantes as $i => $val) {
    $placeholdersContratantes[] = ":pContratante{$i}";
}
$arrayContratanteStr = implode(',', $placeholdersContratantes);

$placeholdersMes = [];
foreach ($meses as $i => $val) {
    $placeholdersMes[] = ":pMes{$i}";
}
$arrayMesStr = implode(',', $placeholdersMes);

// Conexão Oracle
require_once "../config/AW00DB.php";
require_once "../config/oracle.class.php";
require_once "../config/AW00MD.php";

// === SQL ORIGINAL (com UNION ALL) ===
$sql = "
SELECT CT.CD_CONTRATANTE CONTRATANTE, CT.NM_CONTRATANTE EMPRESA, PT.CD_CONTRAT_ORIGEM, CO.NM_CONTRATANTE,
       US.NM_USUARIO, LPAD(US.CD_UNIMED,4,'0')||MP.CD_CARTEIRA_USUARIO CARTEIRINHA, MP.NR_DOC_ORIGINAL, MP.DT_REALIZACAO UTILIZACAO, MP.NR_PERREF, MP.DT_ANOREF,
       CASE WHEN SUBSTR(D.CD_TRANSACAO,3,2) = 10 THEN 'CONSULTA'   WHEN SUBSTR(D.CD_TRANSACAO,3,2) = 20 THEN 'EXAME'
            WHEN SUBSTR(D.CD_TRANSACAO,3,2) = 30 THEN 'INTERNACAO' WHEN SUBSTR(D.CD_TRANSACAO,3,2) = 40 THEN 'HONORARIO'
       ELSE TO_CHAR(D.CD_TRANSACAO) END TIPO, E.CDPROCEDIMENTOCOMPLETO CODIGO, E.DES_PROCEDIMENTO DESCRICAO,
       SUM(MP.QT_PROCEDIMENTOS) QUANTIDADE, MP.CD_PRESTADOR, PV.NM_PRESTADOR, UPPER(EP.DS_CBO) ESPECIALIDADE,
       CASE WHEN SUBSTR(D.CD_TRANSACAO,3,2) = 40 THEN SUM(MP.VL_HONORARIOS_MEDICOS + MP.VLTAXAOUTUNIHONORARIOS )
       ELSE SUM(MP.VL_REAL_PAGO + MP.VLTAXAOUTUNICOBRADO ) END VALOR, MP.IN_LIBERADO_PAGTO
  FROM GP.MOVIPROC MP INNER JOIN GP.USUARIO  US  ON MP.CD_MODALIDADE      = US.CD_MODALIDADE
                                                AND MP.CD_USUARIO         = US.CD_USUARIO
                                                AND MP.NR_TER_ADESAO      = US.NR_TER_ADESAO
                      INNER JOIN GP.PROPOST  PT  ON PT.NR_PROPOSTA        = US.NR_PROPOSTA
                                                AND PT.CD_MODALIDADE      = US.CD_MODALIDADE
                      INNER JOIN GP.TI_PL_SA TP  ON TP.CD_MODALIDADE      = US.CD_MODALIDADE
                                                AND TP.CD_PLANO           = PT.CD_PLANO
                                                AND TP.CD_TIPO_PLANO      = PT.CD_TIPO_PLANO
                      INNER JOIN GP.CONTRAT  CT  ON PT.CD_CONTRATANTE     = CT.CD_CONTRATANTE
                      INNER JOIN GP.CONTRAT  CO  ON PT.CD_CONTRAT_ORIGEM  = CO.CD_CONTRATANTE

                      INNER JOIN GP.GRA_PAR  GP  ON US.CD_GRAU_PARENTESCO = GP.CD_GRAU_PARENTESCO
                       LEFT JOIN GP.DZ_CBO02 EP  ON TO_CHAR(EP.CD_CBO)    = TRIM(MP.CHAR_11)||'00'
                      INNER JOIN GP.PRESERV  PV  ON MP.CD_PRESTADOR       = PV.CD_PRESTADOR
                      INNER JOIN GP.MODALID   A  ON PT.CD_MODALIDADE      = A.CD_MODALIDADE
                      INNER JOIN GP.PLA_SAU   B  ON A.CD_MODALIDADE       = B.CD_MODALIDADE
                                                AND PT.CD_PLANO           = B.CD_PLANO
                      INNER JOIN GP.TI_PL_SA  C  ON A.CD_MODALIDADE       = C.CD_MODALIDADE
                                                AND B.CD_PLANO            = C.CD_PLANO
                                                AND PT.CD_TIPO_PLANO      = C.CD_TIPO_PLANO
                      INNER JOIN GP.TRANREVI  D  ON MP.CD_TRANSACAO       = D.CD_TRANSACAO
                      INNER JOIN GP.AMBPROCE  E  ON MP.CD_ESP_AMB         = E.CD_ESP_AMB
                                                AND MP.CD_GRUPO_PROC_AMB  = E.CD_GRUPO_PROC_AMB
                                                AND MP.CD_PROCEDIMENTO    = E.CD_PROCEDIMENTO
                                                AND MP.DV_PROCEDIMENTO    = E.DV_PROCEDIMENTO
                      INNER JOIN GP.GRU_PRO   F  ON E.CD_GRUPO_PROC       = F.CD_GRUPO_PROC
 WHERE US.CD_MODALIDADE    IN ( 20,21,22,23,24,25,26,27,28,29,30,31,32,40,41,42,43 )
   AND CT.CD_CONTRATANTE IN ( $arrayContratanteStr)                              
   AND ( MP.IN_LIBERADO_PAGTO  = 1      AND MP.CD_TIPO_PAGAMENTO = 0 )
   AND MP.DT_ANOREF            = :ano   AND MP.NR_PERREF         IN ($arrayMesStr)

 GROUP BY CT.CD_CONTRATANTE, CT.NM_CONTRATANTE, PT.CD_CONTRAT_ORIGEM, CO.NM_CONTRATANTE, US.NM_USUARIO, US.CD_UNIMED, MP.CD_CARTEIRA_USUARIO, MP.NR_DOC_ORIGINAL, MP.DT_REALIZACAO,
  MP.NR_PERREF, MP.DT_ANOREF, D.CD_TRANSACAO, E.CDPROCEDIMENTOCOMPLETO, E.DES_PROCEDIMENTO, MP.CD_PRESTADOR, PV.NM_PRESTADOR, EP.DS_CBO, MP.IN_LIBERADO_PAGTO UNION

SELECT CT.CD_CONTRATANTE CONTRATANTE, CT.NM_CONTRATANTE EMPRESA, PT.CD_CONTRAT_ORIGEM, CO.NM_CONTRATANTE, US.NM_USUARIO,
       LPAD(US.CD_UNIMED,4,'0')||MI.CD_CARTEIRA_USUARIO CARTEIRINHA, MI.NR_DOC_ORIGINAL, MI.DT_REALIZACAO UTILIZACAO, MI.NR_PERREF, MI.DT_ANOREF,
       CASE WHEN SUBSTR(D.CD_TRANSACAO,3,2) = 10 THEN 'CONSULTA'   WHEN SUBSTR(D.CD_TRANSACAO,3,2) = 20 THEN 'EXAME'
            WHEN SUBSTR(D.CD_TRANSACAO,3,2) = 30 THEN 'INTERNACAO' WHEN SUBSTR(D.CD_TRANSACAO,3,2) = 40 THEN 'HONORARIO'
       ELSE NULL END TIPO, E.CD_INSUMO CODIGO, E.DS_INSUMO DESCRICAO, SUM(MI.QT_INSUMO) QUANTIDADE, MI.CD_PRESTADOR, PV.NM_PRESTADOR,
       UPPER(EP.DS_CBO) ESPECIALIDADE, SUM (MI.VL_REAL_PAGO + MI.VLTAXAOUTUNICOBRADO ) VALOR, MI.IN_LIBERADO_PAGTO
  FROM GP.MOV_INSU MI INNER JOIN GP.USUARIO  US  ON MI.CD_MODALIDADE      = US.CD_MODALIDADE
                                                AND MI.CD_USUARIO         = US.CD_USUARIO
                                                AND MI.NR_TER_ADESAO      = US.NR_TER_ADESAO

                      INNER JOIN GP.PROPOST  PT  ON PT.NR_PROPOSTA        = US.NR_PROPOSTA
                                                AND PT.CD_MODALIDADE      = US.CD_MODALIDADE

                      INNER JOIN GP.TI_PL_SA TP  ON TP.CD_MODALIDADE      = US.CD_MODALIDADE
                                                AND TP.CD_PLANO           = PT.CD_PLANO
                                                AND TP.CD_TIPO_PLANO      = PT.CD_TIPO_PLANO
                      INNER JOIN GP.CONTRAT  CT  ON PT.CD_CONTRATANTE     = CT.CD_CONTRATANTE
                      INNER JOIN GP.CONTRAT  CO  ON PT.CD_CONTRAT_ORIGEM  = CO.CD_CONTRATANTE
                      INNER JOIN GP.GRA_PAR  GP  ON US.CD_GRAU_PARENTESCO = GP.CD_GRAU_PARENTESCO
                       LEFT JOIN GP.DZ_CBO02 EP  ON TO_CHAR(EP.CD_CBO)    = TRIM(MI.CHAR_11)||'00'
                      INNER JOIN GP.PRESERV  PV  ON MI.CD_PRESTADOR       = PV.CD_PRESTADOR
                      INNER JOIN GP.MODALID   A  ON PT.CD_MODALIDADE      = A.CD_MODALIDADE
                      INNER JOIN GP.PLA_SAU   B  ON A.CD_MODALIDADE       = B.CD_MODALIDADE
                                                AND PT.CD_PLANO           = B.CD_PLANO
                      INNER JOIN GP.TI_PL_SA  C  ON A.CD_MODALIDADE       = C.CD_MODALIDADE
                                                AND B.CD_PLANO            = C.CD_PLANO
                                                AND PT.CD_TIPO_PLANO      = C.CD_TIPO_PLANO
                      INNER JOIN GP.TRANREVI  D  ON MI.CD_TRANSACAO       = D.CD_TRANSACAO
                      INNER JOIN GP.INSUMOS   E  ON MI.CD_INSUMO          = E.CD_INSUMO
                                                AND MI.CD_TIPO_INSUMO     = E.CD_TIPO_INSUMO
                      INNER JOIN GP.TIPOINSU  F  ON E.CD_TIPO_INSUMO      = F.CD_TIPO_INSUMO
 WHERE US.CD_MODALIDADE    IN ( 20,21,22,23,24,25,26,27,28,29,30,40,41,42,43 )

   AND CT.CD_CONTRATANTE IN ( $arrayContratanteStr)                            
   AND MI.DT_ANOREF            = :ano   AND MI.NR_PERREF         IN ($arrayMesStr)

  GROUP BY CT.CD_CONTRATANTE, CT.NM_CONTRATANTE, PT.CD_CONTRAT_ORIGEM, CO.NM_CONTRATANTE,  US.NM_USUARIO,  US.CD_UNIMED, MI.CD_CARTEIRA_USUARIO, MI.NR_DOC_ORIGINAL,
        MI.DT_REALIZACAO, MI.NR_PERREF, MI.DT_ANOREF, D.CD_TRANSACAO, E.CD_INSUMO, E.DS_INSUMO, MI.CD_PRESTADOR, PV.NM_PRESTADOR, EP.DS_CBO, MI.IN_LIBERADO_PAGTO
";

// Prepara statement
$stid = @oci_parse($conn, $sql);
if (!$stid) {
    $err = oci_error($conn);
    echo json_encode([
        "success" => false,
        "message" => "Erro ao preparar query: " . ($err['message'] ?? 'Desconhecido')
    ]);
    exit;
}

oci_bind_by_name($stid, ":ano", $ano, -1, OCI_B_INT);

// Bind contratantes
foreach ($contratantes as $i => $val) {
    oci_bind_by_name($stid, ":pContratante{$i}", $contratantes[$i], -1, OCI_B_INT);
}

// Bind meses
foreach ($meses as $i => $val) {
    oci_bind_by_name($stid, ":pMes{$i}", $meses[$i], -1, OCI_B_INT);
}

// Executa
$r = @oci_execute($stid);
if (!$r) {
    $err = oci_error($stid);
    echo json_encode([
        "success" => false,
        "message" => "Erro ao executar query: " . ($err['message'] ?? 'Desconhecido')
    ]);
    exit;
}

// === Se não houver resultados
$rowData = oci_fetch_assoc($stid);
if (!$rowData) {
    echo json_encode([
        "success" => false,
        "message" => "Nenhum dado encontrado para os parâmetros informados."
    ]);
    exit;
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
$filename = "analitico{$ano}_{$mes}.xlsx";
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"{$filename}\"");
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

