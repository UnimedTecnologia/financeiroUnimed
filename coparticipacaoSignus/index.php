<?php

session_start();

// tempo de vida da sessão (em segundos) — 1 hora
$session_lifetime = 600; // 10 minutos

$senha_correta = 'Signus540@'; // senha fixa

// logout via ?logout=1
if (isset($_GET['logout'])) {
    // remove apenas os dados de autenticação e timestamp
    unset($_SESSION['autenticado'], $_SESSION['login_time']);
    // opcional: destrói toda a sessão
    // session_unset();
    // session_destroy();
    header('Location: ' . $_SERVER['PHP_SELF']);
    exit;
}
// se já estiver autenticado, checar expiração
if (isset($_SESSION['autenticado']) && $_SESSION['autenticado']) {
    if (empty($_SESSION['login_time']) || (time() - $_SESSION['login_time'] > $session_lifetime)) {
        // expirada: desloga e solicita novo login
        unset($_SESSION['autenticado'], $_SESSION['login_time']);
        $erro = 'Sessão expirada. Faça login novamente.';
    } else {
        // renova timestamp de atividade
        $_SESSION['login_time'] = time();
    }
}

if (isset($_POST['senha'])) {
    if ($_POST['senha'] === $senha_correta) {
        $_SESSION['autenticado'] = true;
        $_SESSION['login_time'] = time(); // marca hora do login
        header('Location: ' . $_SERVER['PHP_SELF']);
        exit;
    } else {
        $erro = 'Senha incorreta.';
    }
}

if (!isset($_SESSION['autenticado']) || !$_SESSION['autenticado']) {
    // formulário de senha simples
?>
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <link rel="icon" href="icon.png" type="image/png">
        <title>Entrar - Coparticipação</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>
            body { background:#f8f9fa; }
            .card { max-width: 420px; margin: 80px auto; }
        </style>
    </head>
    <body>
    <div class="card shadow-sm">
        <div class="card-body">
            <h5 class="card-title text-center mb-3">Acesso — Coparticipação</h5>
            <?php if (!empty($erro)): ?>
                <div class="alert alert-danger"><?php echo htmlspecialchars($erro); ?></div>
            <?php endif; ?>
            <form method="post" novalidate>
                <div class="mb-3">
                    <label for="senha" class="form-label">Senha</label>
                    <input type="password" class="form-control" id="senha" name="senha" required autofocus>
                </div>
                <div class="d-grid">
                    <button class="btn btn-primary" type="submit">Entrar</button>
                </div>
            </form>
        </div>
    </div>
    </body>
    </html>
    <?php
    exit;
}
?>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <link rel="icon" href="icon.png" type="image/png">
    <title>Coparticipação - Relatório</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="../style.css" rel="stylesheet">
</head>
<body class="bg-light">
    <div class="d-flex justify-content-end" style="width:100%; background:#f8f9fa; height:40px;">
        <a href="?logout=1" class="btn btn-link m-3">Logout</a>
    </div>
    <div class="container mt-5">
        <div class="row justify-content-center">
            <div class="col-md-8 col-lg-6">
                <div class="card shadow-lg border-0 rounded-4">
                    <div class="card-header text-center bg-primary text-white rounded-top-4">
                        <h4 class="mb-0">Relatório - Coparticipação</h4>
                    </div>
                    <div class="card-body">
                        <form id="formCopart">
                            <div class="row justify-content-center">
                                <div class="col-md-5 mb-3">
                                    <div class="form-floating-label">
                                        <input type="text" pattern="[0-9]*" class="form-control inputback" placeholder=" " id="anoCopart" name="anoCopart"  required maxlength="4">
                                        <label for="anoCopart" class="form-label">Ano</label>
                                    </div>
                                </div>
                                <div class="col-md-5 mb-3">
                                    <div class="form-floating-label">
                                        <input type="text" pattern="[0-9]*" class="form-control inputback" placeholder=" " id="mesCopart" name="mesCopart"  required maxlength="2">
                                        <label for="mesCopart" class="form-label">Mês</label>
                                    </div>
                                </div>
                            </div>
                            <div class="row justify-content-center">
                                <div class="col-md-12 mb-3">
                                    <div class="form-floating-label">
                                        <input type="text" pattern="[0-9]*" class="form-control inputback" placeholder=" " id="cdModalidadeCopart" name="cdModalidadeCopart"  required >
                                        <label for="cdModalidadeCopart" class="form-label">Código Modalidade</label>
                                    </div>
                                </div>
                            </div>
                            <div class="row justify-content-center">
                                <div class="col-md-12 mb-3">
                                    <div class="form-floating-label">
                                        <input type="text" pattern="[0-9]*" class="form-control inputback" placeholder=" " id="nrPropostaCopart" name="nrPropostaCopart"  required >
                                        <label for="nrPropostaCopart" class="form-label">Número Proposta</label>
                                    </div>
                                </div>
                            </div>
                            <div class="row justify-content-center">
                                <div class="col-md-12 mb-3">
                                    <div class="form-floating-label">
                                        <input type="text" pattern="[0-9]*" class="form-control inputback" placeholder=" " id="cdContratanteCopart" name="cdContratanteCopart"  required >
                                        <label for="cdContratanteCopart" class="form-label">Código Contratante</label>
                                    </div>
                                </div>
                            </div>
                
                            <div class="d-grid gap-2 mt-3">
                                <button id="btnRelCopartExcel" type="submit" class="btn btn-success btn-lg" name="tipo" value="excel" style="display:none">
                                    Gerar Relatório em Excel
                                </button>
                                <button id="btnRelCopartTxt" type="submit" class="btn btn-info btn-lg" name="tipo" value="txt">
                                    Gerar Relatório em TXT
                                </button>
                            </div>
                        </form>
                        <!-- Loader -->
                        <div id="loader" class="text-center mt-3 d-none">
                            <div class="spinner-border text-primary" role="status">
                                <span class="visually-hidden">Gerando relatório...</span>
                            </div>
                            <p class="mt-2">Gerando relatório, aguarde...</p>
                        </div>
                    </div>
                    <div class="card-footer text-muted text-center small">
                    Informe os parâmetros para exportar os dados.
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
document.getElementById("formCopart").addEventListener("submit", async function(e) {
    e.preventDefault();

    const form = e.target;
    const loader = document.getElementById("loader");
    const btnExcel = document.getElementById("btnRelCopartExcel");
    const btnTxt = document.getElementById("btnRelCopartTxt");
    
    // Identifica qual botão foi clicado
    const clickedButton = e.submitter;
    const tipoRelatorio = clickedButton.value;
    
    console.log("Tipo do relatório:", tipoRelatorio); // Para debug
    
    // Cria FormData manualmente e adiciona o tipo
    const formData = new FormData(form);
    formData.append('tipo', tipoRelatorio); // ← ADICIONA O TIPO MANUALMENTE
    
    // Desativa ambos os botões e mostra loader
    btnExcel.disabled = true;
    btnTxt.disabled = true;
    clickedButton.innerText = "Gerando...";
    loader.classList.remove("d-none");

    try {
        const response = await fetch("gerar_coparticipacao.php", {
            method: "POST",
            body: formData
        });

        if (!response.ok) {
            throw new Error("Erro ao gerar relatório.");
        }

        // Se a resposta for JSON (erro ou sem dados)
        const contentType = response.headers.get("Content-Type") || "";
        if (contentType.includes("application/json")) {
            const data = await response.json();
            alert(data.message || "Nenhum dado encontrado.");
            return;
        }

        // Caso contrário, é o arquivo (Excel ou TXT)
        const disposition = response.headers.get("Content-Disposition");
        let filename = tipoRelatorio === 'txt' ? "relatorio_coparticipacao.txt" : "relatorio_coparticipacao.xlsx";
        if (disposition && disposition.includes("filename=")) {
            filename = disposition.split("filename=")[1].replace(/["']/g, "");
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);

        const a = document.createElement("a");
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);

    } catch (err) {
        alert("Erro: " + err.message);
    } finally {
        // Reativa botões e esconde loader
        btnExcel.disabled = false;
        btnTxt.disabled = false;
        btnExcel.innerText = "Gerar Relatório em Excel";
        btnTxt.innerText = "Gerar Relatório em TXT";
        loader.classList.add("d-none");
    }
});
</script>
</body>
</html>