<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Regra 60 anos - Relatório</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">

<div class="container mt-5">
    <div class="row justify-content-center">
        <div class="col-md-8 col-lg-6">
            <div class="card shadow-lg border-0 rounded-4">
                <div class="card-header text-center bg-primary text-white rounded-top-4">
                    <h4 class="mb-0">Relatório - Regra 60 anos</h4>
                </div>
                <div class="card-body">
                    <form id="formRelatorio">
                        <!-- <div class="mb-3">
                            <label for="data_inicial" class="form-label">Data Inicial</label>
                            <input type="date" class="form-control" id="data_inicial" name="data_inicial" required>
                        </div>
                        <div class="mb-3">
                            <label for="data_final" class="form-label">Data Final</label>
                            <input type="date" class="form-control" id="data_final" name="data_final" required>
                        </div> -->
                        <div class="d-grid">
                            <button id="btnRel" type="submit" class="btn btn-success btn-lg">
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
                <!-- <div class="card-footer text-muted text-center small">
                    Selecione o período desejado para exportar os dados.
                </div> -->
            </div>
        </div>
    </div>
</div>

<script>
document.getElementById("formRelatorio").addEventListener("submit", async function(e) {
    e.preventDefault();

    const form = e.target;
    const formData = new FormData(form);
    const loader = document.getElementById("loader");
    const btn = document.getElementById("btnRel");

    // Desativa botão e mostra loader
    btn.disabled = true;
    btn.innerText = "Gerando...";
    loader.classList.remove("d-none");

    try {
        const response = await fetch("gerar_regra60Anos.php", {
            method: "POST",
            body: formData
        });

        if (!response.ok) {
            throw new Error("Erro ao gerar relatório.");
        }

        const disposition = response.headers.get("Content-Disposition");
        let filename = "relatorio.xlsx";
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
        // Reativa botão e esconde loader
        btn.disabled = false;
        btn.innerText = "Gerar Relatório em Excel";
        loader.classList.add("d-none");
    }
});

</script>

</body>
</html>
