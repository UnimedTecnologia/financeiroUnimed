<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Proporcional - Relatório</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="../style.css" rel="stylesheet">
</head>
<body class="bg-light">

<div class="container mt-5">
    <div class="row justify-content-center">
        <div class="col-md-8 col-lg-6">
            <div class="card shadow-lg border-0 rounded-4">
                <div class="card-header text-center bg-primary text-white rounded-top-4">
                    <h4 class="mb-0">Relatório - Proporcional</h4>
                </div>
                <div class="card-body">
                    <form id="formProporcional">
                        <div class="row justify-content-center">
                            <div class="col-md-5 mb-3">
                                <div class="form-floating-label">
                                    <input type="text" pattern="[0-9]*" class="form-control inputback" placeholder=" " id="anoCopart" name="anoCopart"  required maxlength="4">
                                    <label for="anoCProp" class="form-label">Ano</label>
                                </div>
                            </div>
                            <div class="col-md-5 mb-3">
                                <div class="form-floating-label">
                                    <input type="text" pattern="[0-9]*" class="form-control inputback" placeholder=" " id="mesCopart" name="mesCopart"  required maxlength="2">
                                    <label for="mesProp" class="form-label">Mês</label>
                                </div>
                            </div>
                        </div>
                        <div class="row justify-content-center">
                            <div class="col-md-12 mb-3">
                                <div class="form-floating-label">
                                    <input type="text" pattern="[0-9]*" class="form-control inputback" placeholder=" " id="cdModalidadeCopart" name="cdModalidadeCopart"  required >
                                    <label for="cdModalidadeProp" class="form-label">Código Modalidade</label>
                                </div>
                            </div>
                        </div>
                        <div class="row justify-content-center">
                            <div class="col-md-12 mb-3">
                                <div class="form-floating-label">
                                    <input type="text" pattern="[0-9]*" class="form-control inputback" placeholder=" " id="nrPropostaCopart" name="nrPropostaCopart"  required >
                                    <label for="nrPropostaProp" class="form-label">Número Proposta</label>
                                </div>
                            </div>
                        </div>
                        <div class="row justify-content-center">
                            <div class="col-md-6 mb-3">
                                <div class="form-floating-label">
                                    <input type="text" pattern="[0-9]*" class="form-control inputback" placeholder=" " id="cdContratanteCopart" name="cdContratanteCopart"  required >
                                    <label for="cdContratanteProp" class="form-label">Código Contratante</label>
                                </div>
                            </div>
                            <div class="col-md-6 mb-3">
                                <div class="form-floating-label">
                                    <input type="text" pattern="[0-9,]*" class="form-control inputback" placeholder=" " id="eventoProp" name="eventoProp" required>
                                    <label for="eventoProp" class="form-label">Evento (separado por vírgula)</label>
                                </div>
                            </div>
                        </div>
            
                        <div class="d-grid mt-3">
                            <button id="btnRelProp" type="submit" class="btn btn-success btn-lg">
                                Gerar Relatório em Excel
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
document.getElementById("formProporcional").addEventListener("submit", async function(e) {
    e.preventDefault();

    const form = e.target;
    const formData = new FormData(form);
    const loader = document.getElementById("loader");
    const btn = document.getElementById("btnRelProp");

    // Desativa botão e mostra loader
    btn.disabled = true;
    btn.innerText = "Gerando...";
    loader.classList.remove("d-none");

    try {
    const response = await fetch("gerar_proporcional.php", {
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
        btn.disabled = false;
        btn.innerText = "Gerar Relatório em Excel";
        loader.classList.add("d-none");
        alert(data.message || "Nenhum dado encontrado.");
        return; // não tenta baixar
        

    }

    // Caso contrário, é o Excel
    const disposition = response.headers.get("Content-Disposition");
    let filename = "relatorio_proporcional.xlsx";
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
        btn.disabled = false;
        btn.innerText = "Gerar Relatório em Excel";
        loader.classList.add("d-none");
    }

    btn.disabled = false;
    btn.innerText = "Gerar Relatório em Excel";
    loader.classList.add("d-none");

});

</script>

</body>
</html>
