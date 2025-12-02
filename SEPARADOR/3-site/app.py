<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Filtrador de Excel</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        /* Estilo básico para a fonte Inter */
        body { font-family: 'Inter', sans-serif; }
    </style>
</head>
<body class="bg-gray-100 min-h-screen flex items-center justify-center p-4">
    <div id="app" class="bg-white p-6 sm:p-8 rounded-xl shadow-2xl w-full max-w-lg transition-all duration-300">
        <h1 class="text-3xl font-extrabold text-gray-800 text-center mb-2">Filtrador de Excel</h1>
        <p class="text-gray-500 text-center mb-6">Filtre e formate seus documentos automaticamente.</p>

        <form id="uploadForm" action="/processar" method="post" enctype="multipart/form-data" class="space-y-6">
            
            <div class="space-y-2">
                <label for="file" class="block text-lg font-medium text-gray-700">Anexar Documento Excel</label>
                <div class="flex items-center space-x-4">
                    <label class="block">
                        <span class="sr-only">Escolher arquivo</span>
                        <input type="file" id="file" name="file" required
                            class="block w-full text-sm text-gray-500
                                file:mr-4 file:py-2 file:px-4
                                file:rounded-full file:border-0
                                file:text-sm file:font-semibold
                                file:bg-blue-50 file:text-blue-700
                                hover:file:bg-blue-100 cursor-pointer
                            "/>
                    </label>
                </div>
            </div>

            <div class="grid grid-cols-2 gap-4">
                <div class="space-y-2">
                    <label for="mes" class="block text-lg font-medium text-gray-700">Mês</label>
                    <select id="mes" name="mes" required
                        class="mt-1 block w-full pl-3 pr-10 py-2 text-base border-gray-300 focus:outline-none focus:ring-blue-500 focus:border-blue-500 sm:text-sm rounded-lg shadow-sm">
                        <!-- Opções de Mês (1 a 12) -->
                        <option value="1">Janeiro</option>
                        <option value="2">Fevereiro</option>
                        <option value="3">Março</option>
                        <option value="4">Abril</option>
                        <option value="5">Maio</option>
                        <option value="6">Junho</option>
                        <option value="7">Julho</option>
                        <option value="8">Agosto</option>
                        <option value="9" selected>Setembro</option>
                        <option value="10">Outubro</option>
                        <option value="11">Novembro</option>
                        <option value="12">Dezembro</option>
                    </select>
                </div>

                <div class="space-y-2">
                    <label for="ano" class="block text-lg font-medium text-gray-700">Ano</label>
                    <select id="ano" name="ano" required
                        class="mt-1 block w-full pl-3 pr-10 py-2 text-base border-gray-300 focus:outline-none focus:ring-blue-500 focus:border-blue-500 sm:text-sm rounded-lg shadow-sm">
                        <!-- Opções de Ano (Exemplo: 2024 a 2026) -->
                        <option value="2024">2024</option>
                        <option value="2025" selected>2025</option>
                        <option value="2026">2026</option>
                    </select>
                </div>
            </div>

            <button type="submit"
                class="w-full flex justify-center py-3 px-4 border border-transparent rounded-lg shadow-md text-lg font-medium text-white bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 transition duration-150 ease-in-out transform hover:scale-[1.01] active:scale-[0.99] disabled:bg-gray-400"
                id="submitButton">
                Consertar e Baixar
            </button>
        </form>

        <div id="messageBox" class="mt-6 hidden p-4 rounded-lg text-center" role="alert">
            <p id="messageText" class="font-medium"></p>
        </div>

    </div>

    <script>
        const form = document.getElementById('uploadForm');
        const submitButton = document.getElementById('submitButton');
        const messageBox = document.getElementById('messageBox');
        const messageText = document.getElementById('messageText');

        // Funções para manipulação da caixa de mensagem
        function showMessage(text, isError) {
            messageText.textContent = text;
            messageBox.classList.remove('hidden', 'bg-red-100', 'text-red-700', 'bg-green-100', 'text-green-700');
            if (isError) {
                messageBox.classList.add('bg-red-100', 'text-red-700');
            } else {
                messageBox.classList.add('bg-green-100', 'text-green-700');
            }
        }

        function hideMessage() {
            messageBox.classList.add('hidden');
        }

        // Listener para o formulário
        form.addEventListener('submit', async function(e) {
            e.preventDefault();
            hideMessage();
            submitButton.disabled = true;
            submitButton.textContent = 'Processando...';

            const formData = new FormData(form);

            try {
                const response = await fetch(form.action, {
                    method: 'POST',
                    body: formData,
                });

                if (response.ok) {
                    // Se a resposta for bem-sucedida (status 200), inicie o download
                    const blob = await response.blob();
                    const contentDisposition = response.headers.get('Content-Disposition');
                    
                    // Extrai o nome do arquivo do cabeçalho Content-Disposition
                    let filename = 'processado.xlsx';
                    if (contentDisposition) {
                        const match = contentDisposition.match(/filename="(.+?)"/);
                        if (match && match[1]) {
                            filename = match[1];
                        }
                    }

                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.style.display = 'none';
                    a.href = url;
                    a.download = filename;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    
                    showMessage('Sucesso! O arquivo processado foi baixado.', false);

                } else {
                    // Se houver erro (status != 200), leia a mensagem de erro
                    const errorText = await response.text();
                    showMessage(`Erro: ${errorText}`, true);
                }
            } catch (error) {
                console.error('Erro na requisição:', error);
                showMessage(`Erro de conexão: ${error.message}`, true);
            } finally {
                submitButton.disabled = false;
                submitButton.textContent = 'Consertar e Baixar';
            }
        });

        // Configura o mês e ano padrão para o atual
        document.addEventListener('DOMContentLoaded', () => {
            const now = new Date();
            document.getElementById('mes').value = now.getMonth() + 1;
            document.getElementById('ano').value = now.getFullYear();
        });
    </script>
</body>
</html>
