# Template HTML do Dashboard Web para o Workflow Fiscal
# Desenvolvido com Estética Premium, Tailwind CSS, Google Fonts e Vanilla JS.

HTML_CONTENT = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Painel Fiscal - Staging & Apuração</title>
    <!-- Google Fonts (Inter) -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome Icons -->
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: {
                        sans: ['Inter', 'sans-serif'],
                    },
                    colors: {
                        brand: {
                            50: '#f5f6ff',
                            100: '#ebedff',
                            500: '#4f46e5',
                            600: '#4338ca',
                            700: '#3730a3',
                            900: '#1e1b4b',
                        }
                    }
                }
            }
        }
    </script>
    <style>
        body {
            background-color: #f8fafc;
        }
        .glass {
            background: rgba(255, 255, 255, 0.85);
            backdrop-filter: blur(12px);
            border: 1px solid rgba(226, 232, 240, 0.8);
        }
        .toast {
            animation: slideIn 0.3s ease-out forwards;
        }
        @keyframes slideIn {
            from { transform: translateY(1rem); opacity: 0; }
            to { transform: translateY(0); opacity: 1; }
        }
    </style>
</head>
<body class="font-sans antialiased text-slate-800">

    <!-- HEADER / NAVIGATION -->
    <header class="sticky top-0 z-40 w-full bg-brand-900 text-white shadow-md">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
            <div class="flex items-center space-x-3">
                <div class="bg-brand-500 p-2 rounded-lg text-white">
                    <i class="fa-solid fa-scale-balanced text-xl"></i>
                </div>
                <div>
                    <h1 class="text-lg font-bold leading-tight tracking-tight">Antigravity Fiscal</h1>
                    <p class="text-xs text-brand-100 opacity-80">Workflow Modular de Notas Fiscais</p>
                </div>
            </div>
            <div class="flex items-center space-x-4">
                <button onclick="abrirModalLogs()" class="inline-flex items-center px-3 py-1.5 rounded-lg text-xs font-medium bg-slate-800 border border-slate-700 hover:bg-slate-700 text-slate-200 transition-colors">
                    <i class="fa-solid fa-terminal mr-2 text-indigo-400"></i>
                    Logs do Servidor
                </button>
                <span id="db-status-badge" class="inline-flex items-center px-3 py-1 rounded-full text-xs font-semibold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                    <span class="w-2 h-2 mr-2 rounded-full bg-emerald-400 animate-pulse"></span>
                    Conectado (Neon / SQLite)
                </span>
            </div>
        </div>
    </header>

    <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        
        <!-- TOP GRID: CADASTRAR EMPRESA & UPLOAD XML -->
        <div class="grid grid-cols-1 lg:grid-cols-3 gap-8 mb-8">
            
            <!-- CARD 1: SELECIONAR & CADASTRAR EMPRESA -->
            <div class="glass p-6 rounded-2xl shadow-sm lg:col-span-1 flex flex-col justify-between">
                <div>
                    <div class="flex items-center justify-between mb-4">
                        <h2 class="text-md font-semibold text-slate-900 flex items-center">
                            <i class="fa-solid fa-building mr-2 text-brand-500"></i> Empresa Ativa
                        </h2>
                        <button onclick="toggleModal('modal-empresa', true)" class="text-xs text-brand-500 hover:text-brand-600 font-medium">
                            <i class="fa-solid fa-plus mr-1"></i> Nova Empresa
                        </button>
                    </div>
                    <label class="block text-xs font-medium text-slate-500 mb-2">Filtrar por Empresa ativa:</label>
                    <select id="select-empresa" onchange="carregarDocumentos()" class="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-brand-500">
                        <option value="">-- Nenhuma empresa cadastrada --</option>
                    </select>
                    <div class="mt-4">
                        <label class="block text-[11px] font-semibold text-slate-500 mb-1">Período Fiscal (Mês/Ano):</label>
                        <div class="grid grid-cols-2 gap-2">
                            <select id="select-mes" onchange="carregarDocumentos()" class="w-full bg-white border border-slate-200 rounded-lg px-2 py-1.5 text-xs focus:outline-none focus:ring-2 focus:ring-brand-500">
                                <option value="">Todos os meses</option>
                                <option value="1">Janeiro</option>
                                <option value="2">Fevereiro</option>
                                <option value="3">Março</option>
                                <option value="4">Abril</option>
                                <option value="5">Maio</option>
                                <option value="6">Junho</option>
                                <option value="7">Julho</option>
                                <option value="8">Agosto</option>
                                <option value="9">Setembro</option>
                                <option value="10">Outubro</option>
                                <option value="11">Novembro</option>
                                <option value="12">Dezembro</option>
                            </select>
                            <select id="select-ano" onchange="carregarDocumentos()" class="w-full bg-white border border-slate-200 rounded-lg px-2 py-1.5 text-xs focus:outline-none focus:ring-2 focus:ring-brand-500">
                                <option value="">Todos os anos</option>
                                <option value="2024">2024</option>
                                <option value="2025">2025</option>
                                <option value="2026">2026</option>
                                <option value="2027">2027</option>
                            </select>
                        </div>
                    </div>
                </div>
                
                <div id="info-empresa-box" class="mt-4 p-3 bg-slate-50 rounded-xl border border-slate-100 hidden">
                    <div class="flex justify-between items-start">
                        <div>
                            <p class="text-xs text-slate-500">CNPJ: <span id="info-empresa-cnpj" class="font-mono text-slate-700 font-medium"></span></p>
                            <p class="text-xs text-slate-500 mt-1">Regime Fiscal: <span id="info-empresa-regime" class="inline-flex px-2 py-0.5 ml-1 text-[10px] font-semibold bg-brand-100 text-brand-600 rounded"></span></p>
                            <p class="text-xs text-slate-500 mt-1" id="info-empresa-cnae-wrapper">CNAE: <span id="info-empresa-cnae" class="font-mono text-slate-700 font-medium"></span></p>
                        </div>
                        <button onclick="abrirEditarEmpresa()" class="text-xs text-amber-600 hover:text-amber-700 font-semibold flex items-center mt-0.5 transition-colors">
                            <i class="fa-solid fa-pen-to-square mr-1"></i>Editar
                        </button>
                    </div>
                </div>
            </div>

            <!-- CARD 2: ÁREA DE UPLOAD DE XML -->
            <div class="glass p-6 rounded-2xl shadow-sm lg:col-span-2">
                <h2 class="text-md font-semibold text-slate-900 mb-4 flex items-center">
                    <i class="fa-solid fa-file-import mr-2 text-brand-500"></i> Ingestão & Normalização (XML)
                </h2>
                
                <div id="dropzone" 
                     class="border-2 border-dashed border-slate-200 hover:border-brand-500 rounded-2xl p-6 text-center cursor-pointer transition-colors flex flex-col items-center justify-center bg-slate-50/50"
                     onclick="document.getElementById('file-input').click()"
                     ondragover="event.preventDefault(); this.classList.add('border-brand-500', 'bg-brand-50/20')"
                     ondragleave="this.classList.remove('border-brand-500', 'bg-brand-50/20')"
                     ondrop="lidarComDrop(event)">
                    
                    <i class="fa-solid fa-cloud-arrow-up text-4xl text-slate-300 mb-3" id="upload-icon"></i>
                    <p class="text-sm font-medium text-slate-700" id="upload-text">Arraste e solte arquivos XML (NF-e, NFC-e ou NFS-e) aqui ou clique para buscar</p>
                    <p class="text-xs text-slate-400 mt-1">Os emitentes serão autodetectados e cadastrados automaticamente no banco de dados</p>
                    
                    <input type="file" id="file-input" class="hidden" accept=".xml" onchange="lidarComSelecaoArquivo(event)" multiple>
                </div>
                
                <!-- Opções de Forçar Notas de Entrada de Terceiros -->
                <div class="mt-4 p-4 bg-slate-50 rounded-xl border border-slate-200/60 flex flex-col md:flex-row md:items-center justify-between gap-4">
                    <div class="flex items-center space-x-3">
                        <input type="checkbox" id="chk-forcar-entrada" class="w-4 h-4 rounded text-brand-500 focus:ring-brand-500 border-slate-300 cursor-pointer text-xs" onchange="toggleForcarEntradaOptions()">
                        <div class="cursor-pointer select-none" onclick="document.getElementById('chk-forcar-entrada').click()">
                            <span class="text-xs font-bold text-slate-700 block">Importar como Notas de Entrada (Compras) da empresa ativa</span>
                            <span class="text-[10px] text-slate-400 block">Associa todas as notas deste lote à empresa selecionada forçando a competência e o tipo de operação</span>
                        </div>
                    </div>
                    <div id="forcar-entrada-opcoes" class="flex items-center space-x-2 hidden opacity-0 transition-opacity duration-200">
                        <label class="text-[10px] font-bold text-slate-400 uppercase select-none">Competência:</label>
                        <select id="import-mes" class="bg-white border border-slate-200 rounded-lg px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-brand-500 font-medium">
                            <option value="1">Janeiro</option>
                            <option value="2">Fevereiro</option>
                            <option value="3">Março</option>
                            <option value="4">Abril</option>
                            <option value="5">Maio</option>
                            <option value="6">Junho</option>
                            <option value="7">Julho</option>
                            <option value="8">Agosto</option>
                            <option value="9">Setembro</option>
                            <option value="10">Outubro</option>
                            <option value="11">Novembro</option>
                            <option value="12">Dezembro</option>
                        </select>
                        <select id="import-ano" class="bg-white border border-slate-200 rounded-lg px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-brand-500 font-medium">
                            <option value="2024">2024</option>
                            <option value="2025">2025</option>
                            <option value="2026">2026</option>
                            <option value="2027">2027</option>
                        </select>
                    </div>
                </div>
            </div>
        </div>

        <!-- BOARD TRIBUTÁRIO CONSOLIDADO -->
        <div id="consolidado-section" class="mb-8 hidden">
            <div class="glass p-6 rounded-2xl shadow-sm border border-slate-100">
                <div class="flex items-center justify-between mb-6">
                    <h2 class="text-md font-semibold text-slate-900 flex items-center">
                        <i class="fa-solid fa-chart-line mr-2 text-brand-500"></i> Apuração Consolidada de Impostos (Período)
                    </h2>
                    <div class="flex items-center space-x-2">
                        <button onclick="abrirDrawerAuditoria()" class="inline-flex items-center px-3 py-1.5 text-xs font-semibold bg-brand-500 hover:bg-brand-600 text-white rounded-lg transition-colors shadow-sm">
                            <i class="fa-solid fa-square-poll-vertical mr-1.5"></i> Auditar Memória de Cálculo
                        </button>
                        <span id="regime-badge" class="inline-flex px-3 py-1 text-[10px] font-bold bg-brand-100 text-brand-700 rounded-full uppercase tracking-wider"></span>
                    </div>
                </div>
                
                <div class="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-4 gap-6">
                    <!-- Faturamento -->
                    <div class="bg-gradient-to-br from-indigo-50/50 to-indigo-100/20 p-5 rounded-2xl border border-indigo-100 relative overflow-hidden transition-all hover:shadow-md">
                        <p class="text-xs text-indigo-500 font-semibold uppercase tracking-wider">Faturamento Staging Area</p>
                        <h3 id="c-faturamento" class="text-xl font-black text-indigo-900 mt-2 font-mono">R$ 0,00</h3>
                        <p class="text-[10px] text-indigo-600 mt-1 opacity-80">Receita Bruta recalculada (com ajustes)</p>
                    </div>

                    <!-- Total Imposto -->
                    <div class="bg-gradient-to-br from-emerald-50/50 to-emerald-100/20 p-5 rounded-2xl border border-emerald-100 relative overflow-hidden transition-all hover:shadow-md">
                        <p class="text-xs text-emerald-600 font-semibold uppercase tracking-wider">Total de Impostos Calculados</p>
                        <h3 id="c-imposto" class="text-xl font-black text-emerald-950 mt-2 font-mono">R$ 0,00</h3>
                        <p class="text-[10px] text-emerald-600 mt-1 opacity-80">Gerado via Strategy Pattern</p>
                    </div>

                    <!-- Alíquota Efetiva -->
                    <div class="bg-gradient-to-br from-amber-50/50 to-amber-100/20 p-5 rounded-2xl border border-amber-100 relative overflow-hidden transition-all hover:shadow-md">
                        <p class="text-xs text-amber-600 font-semibold uppercase tracking-wider">Alíquota Efetiva Consolidada</p>
                        <h3 id="c-aliquota" class="text-xl font-black text-amber-950 mt-2 font-mono">0,00%</h3>
                        <p class="text-[10px] text-amber-700 mt-1 opacity-80">Relação direta Imposto / Receita</p>
                    </div>

                    <!-- Detalhamento Individual -->
                    <div class="bg-white/60 p-5 rounded-2xl border border-slate-100 transition-all hover:shadow-md flex flex-col justify-center">
                        <p class="text-[10px] text-slate-400 font-bold uppercase tracking-wider mb-2">Abertura por Tributo</p>
                        <div id="c-detalhes-list" class="space-y-1 text-xs text-slate-600">
                            <!-- Injetado dinamicamente -->
                        </div>
                    </div>
                </div>
                
                <!-- Apuração de Entradas (Compras) -->
                <div class="border-t border-slate-200/60 pt-6 mt-6">
                    <h3 class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-4 flex items-center">
                        <i class="fa-solid fa-cart-shopping mr-2 text-brand-500"></i> Apuração de Compras (Entradas)
                    </h3>
                    <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                        <!-- Total Compras -->
                        <div class="bg-slate-50 p-4 rounded-xl border border-slate-100 relative overflow-hidden transition-all hover:shadow-md">
                            <p class="text-xs text-slate-500 font-semibold uppercase tracking-wider">Total de Compras</p>
                            <h3 id="c-compras-total" class="text-lg font-bold text-slate-800 mt-1 font-mono">R$ 0,00</h3>
                            <p class="text-[10px] text-slate-400 mt-1" id="c-compras-count">0 notas de compra no período</p>
                        </div>
                        
                        <!-- DIFAL Acumulado -->
                        <div class="bg-gradient-to-br from-blue-50/50 to-blue-100/20 p-4 rounded-xl border border-blue-100 relative overflow-hidden transition-all hover:shadow-md">
                            <p class="text-xs text-blue-600 font-semibold uppercase tracking-wider">DIFAL Interestadual Acumulado</p>
                            <h3 id="c-compras-difal" class="text-lg font-bold text-blue-900 mt-1 font-mono">R$ 0,00</h3>
                            <p class="text-[10px] text-blue-600 mt-1 opacity-80">Diferencial de alíquota interestadual</p>
                        </div>
                        
                        <!-- ICMS-ST Compra -->
                        <div class="bg-gradient-to-br from-cyan-50/50 to-cyan-100/20 p-4 rounded-xl border border-cyan-100 relative overflow-hidden transition-all hover:shadow-md">
                            <p class="text-xs text-cyan-600 font-semibold uppercase tracking-wider">ICMS-ST destacado (Compra)</p>
                            <h3 id="c-compras-icms-st" class="text-lg font-bold text-cyan-900 mt-1 font-mono">R$ 0,00</h3>
                            <p class="text-[10px] text-cyan-600 mt-1 opacity-80">Substituição Tributária de entrada</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- MAIN CARD: STAGING AREA (TABELA DE DOCUMENTOS) -->
        <div class="glass rounded-2xl shadow-sm overflow-hidden">
            <!-- Tabela Header -->
            <div class="px-6 py-5 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
                <div>
                    <h2 class="text-lg font-bold text-slate-900 flex items-center">
                        <i class="fa-solid fa-list-check mr-2 text-brand-500"></i> Staging Area (Revisão & Edição)
                    </h2>
                    <p class="text-xs text-slate-500">Qualquer ajuste manual é auditado e gravado sem alterar o XML bruto original.</p>
                </div>
                <div class="flex items-center space-x-2">
                    <button onclick="carregarDocumentos()" class="p-2 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-lg text-sm transition-colors" title="Atualizar Tabela">
                        <i class="fa-solid fa-rotate"></i>
                    </button>
                    <button onclick="excluirPeriodo()" id="btn-excluir-periodo" class="inline-flex items-center px-3 py-2 bg-amber-50 hover:bg-amber-100 text-amber-700 border border-amber-200/60 rounded-lg text-xs font-semibold transition-colors" title="Excluir notas da competência selecionada">
                        <i class="fa-solid fa-calendar-minus mr-1.5"></i> Limpar Período
                    </button>
                    <button onclick="excluirNotasEmpresa()" id="btn-excluir-empresa-notes" class="inline-flex items-center px-3 py-2 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200/60 rounded-lg text-xs font-semibold transition-colors" title="Excluir todas as notas desta empresa">
                        <i class="fa-solid fa-eraser mr-1.5"></i> Limpar Notas da Empresa
                    </button>
                    <button onclick="resetarBancoDados()" class="p-2 bg-rose-50 hover:bg-rose-100 text-rose-600 border border-rose-200/50 rounded-lg text-sm transition-colors" title="Resetar/Limpar Todo o Banco de Dados">
                        <i class="fa-solid fa-trash-can"></i>
                    </button>
                    <span id="doc-counter" class="bg-brand-100 text-brand-600 text-xs font-semibold px-3 py-1 rounded-full">
                        0 documentos
                    </span>
                </div>
            </div>
            <!-- Abas de Tipo de Operação -->
            <div class="px-6 border-b border-slate-100 bg-slate-50/20 flex space-x-6 text-xs font-semibold">
                <button onclick="setTipoOperacaoFilter(this, '')" class="tab-btn py-3 border-b-2 border-brand-500 text-brand-600 transition-all focus:outline-none">
                    Todas as Notas
                </button>
                <button onclick="setTipoOperacaoFilter(this, 'Saída')" class="tab-btn py-3 border-b-2 border-transparent text-slate-500 hover:text-slate-700 transition-all focus:outline-none">
                    Saídas (Vendas / Serviços)
                </button>
                <button onclick="setTipoOperacaoFilter(this, 'Entrada')" class="tab-btn py-3 border-b-2 border-transparent text-slate-500 hover:text-slate-700 transition-all focus:outline-none">
                    Entradas (Compras)
                </button>
            </div>
            <!-- Painel de Ações em Lote (oculto por padrão) -->
            <div id="batch-actions-panel" class="hidden bg-slate-50 border-b border-slate-100 px-6 py-3 flex items-center justify-between text-xs transition-all duration-300">
                <div class="flex items-center space-x-2">
                    <span class="font-semibold text-slate-600"><span id="batch-selected-count">0</span> notas selecionadas:</span>
                    <button onclick="excluirSelecionadas()" class="px-2.5 py-1.5 bg-rose-600 hover:bg-rose-700 text-white rounded font-medium transition-colors flex items-center ml-2">
                        <i class="fa-solid fa-trash mr-1.5"></i> Excluir Selecionadas
                    </button>
                    <button onclick="encerrarSelecionadas()" class="px-2.5 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded font-medium transition-colors flex items-center ml-2">
                        <i class="fa-solid fa-lock mr-1.5"></i> Encerrar Selecionadas
                    </button>
                </div>
                <button onclick="desmarcarTodos()" class="text-slate-400 hover:text-slate-600">
                    Desmarcar todas
                </button>
            </div>

            <!-- Tabela de Documentos -->
            <div class="overflow-x-auto">
                <table class="w-full text-left border-collapse">
                    <thead>
                        <tr class="bg-slate-50 text-slate-500 text-xs font-semibold border-b border-slate-100 uppercase tracking-wider">
                            <th class="px-6 py-4 w-12"><input type="checkbox" id="select-all-docs" onclick="toggleSelectAllDocs(this)" class="rounded text-brand-500 focus:ring-brand-500"></th>
                            <th class="px-6 py-4">Número NF</th>
                            <th class="px-6 py-4">Emissão</th>
                            <th class="px-6 py-4">Parceiro</th>
                            <th class="px-6 py-4">Chave de Acesso</th>
                            <th class="px-6 py-4">Tipo</th>
                            <th class="px-6 py-4 text-right">Valor XML (Bruto)</th>
                            <th class="px-6 py-4 text-right">Valor Final (Staging)</th>
                            <th class="px-6 py-4">Status</th>
                            <th class="px-6 py-4 text-center">Ações</th>
                        </tr>
                    </thead>
                    <tbody id="documentos-tbody" class="divide-y divide-slate-100 text-sm text-slate-600">
                        <tr>
                            <td colspan="10" class="px-6 py-12 text-center text-slate-400">
                                <i class="fa-solid fa-building-circle-exclamation text-3xl mb-2 text-slate-300 block"></i>
                                Selecione uma Empresa acima para carregar a Staging Area.
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
    </main>

    <!-- ==========================================
         MODAIS DA APLICAÇÃO
         ========================================== -->

    <!-- MODAL 1: CADASTRAR EMPRESA -->
    <div id="modal-empresa" class="fixed inset-0 z-50 flex items-center justify-center bg-brand-900/40 backdrop-blur-sm hidden">
        <div class="bg-white rounded-2xl shadow-xl w-full max-w-md mx-4 overflow-hidden border border-slate-100 animate-slide-in">
            <div class="px-6 py-4 bg-brand-900 text-white flex items-center justify-between">
                <h3 class="font-bold text-md"><i class="fa-solid fa-building mr-2"></i> Cadastrar Nova Empresa</h3>
                <button onclick="toggleModal('modal-empresa', false)" class="text-white/80 hover:text-white"><i class="fa-solid fa-xmark"></i></button>
            </div>
            <form id="form-empresa" onsubmit="salvarEmpresa(event)" class="p-6 space-y-4">
                <div>
                    <label class="block text-xs font-semibold text-slate-500 mb-1">CNPJ (apenas números, 14 dígitos):</label>
                    <input type="text" id="empresa-cnpj" required pattern="\d{14}" class="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-brand-500" placeholder="Ex: 12345678000199">
                </div>
                <div>
                    <label class="block text-xs font-semibold text-slate-500 mb-1">Razão Social:</label>
                    <input type="text" id="empresa-razao" required class="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-brand-500" placeholder="Ex: Stenio Software Ltda">
                </div>
                <div class="relative">
                    <div class="flex items-center justify-between mb-1">
                        <label class="block text-xs font-semibold text-slate-500">CNAE Principal (Opcional):</label>
                        <button type="button" onclick="sincronizarCnaes(event)" class="text-[10px] text-brand-500 hover:text-brand-600 font-semibold focus:outline-none"><i class="fa-solid fa-rotate mr-1"></i>Sincronizar IBGE</button>
                    </div>
                    <input type="text" id="empresa-cnae" class="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-brand-500" placeholder="Digite ou busque o CNAE..." autocomplete="off">
                    <div id="empresa-cnae-results" class="absolute left-0 right-0 mt-1 max-h-60 overflow-y-auto bg-white border border-slate-200 rounded-lg shadow-lg hidden z-50"></div>
                    <div id="empresa-cnae-info" class="mt-1.5 p-2 bg-indigo-50 border border-indigo-100 rounded text-[11px] text-indigo-700 hidden"></div>
                </div>
                <div>
                    <label class="block text-xs font-semibold text-slate-500 mb-1">Regime Tributário:</label>
                    <select id="empresa-regime" required class="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-brand-500">
                        <option value="Simples Nacional">Simples Nacional</option>
                        <option value="Lucro Presumido">Lucro Presumido</option>
                        <option value="Lucro Real">Lucro Real</option>
                    </select>
                </div>
                <div>
                    <label class="block text-xs font-semibold text-slate-500 mb-1">Categoria do Simples Nacional:</label>
                    <select id="empresa-categoria" class="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-brand-500">
                        <option value="Serviços (Anexo III)">Serviços (Anexo III)</option>
                        <option value="Serviços (Anexo IV)">Serviços (Anexo IV)</option>
                        <option value="Serviços (Anexo V - Fator R)">Serviços (Anexo V - Fator R)</option>
                        <option value="Comércio (Anexo I)">Comércio (Anexo I)</option>
                        <option value="Indústria (Anexo II)">Indústria (Anexo II)</option>
                    </select>
                </div>
                <div class="border-t border-slate-100 pt-3 mt-3">
                    <p class="text-xs font-bold text-slate-400 mb-2">Simples Nacional & Fator R (Opcional)</p>
                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block text-[10px] font-semibold text-slate-500 mb-1">RBT12 Acumulado (R$):</label>
                            <input type="number" step="0.01" id="empresa-rbt12" class="w-full border border-slate-200 rounded-lg px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-brand-500" placeholder="Ex: 150000.00">
                        </div>
                        <div>
                            <label class="block text-[10px] font-semibold text-slate-500 mb-1">Folha Salários 12m (R$):</label>
                            <input type="number" step="0.01" id="empresa-folha12" class="w-full border border-slate-200 rounded-lg px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-brand-500" placeholder="Ex: 45000.00">
                        </div>
                    </div>
                    <div class="flex items-center space-x-2 mt-2">
                        <input type="checkbox" id="empresa-fator-r" class="rounded text-brand-500 focus:ring-brand-500 text-xs">
                        <label for="empresa-fator-r" class="text-[11px] font-semibold text-slate-500 select-none">Atividade Sujeita ao Fator R</label>
                    </div>
                </div>
                <div class="flex justify-end space-x-2 pt-2">
                    <button type="button" onclick="toggleModal('modal-empresa', false)" class="px-4 py-2 border border-slate-200 hover:bg-slate-50 text-slate-600 rounded-lg text-sm font-medium">Cancelar</button>
                    <button type="submit" class="px-4 py-2 bg-brand-500 hover:bg-brand-600 text-white rounded-lg text-sm font-medium">Salvar Empresa</button>
                </div>
            </form>
        </div>
    </div>

    <!-- MODAL 2: AJUSTE MANUAL (OVERRIDE AUDITÁVEL) -->
    <div id="modal-ajuste" class="fixed inset-0 z-50 flex items-center justify-center bg-brand-900/40 backdrop-blur-sm hidden">
        <div class="bg-white rounded-2xl shadow-xl w-full max-w-md mx-4 overflow-hidden border border-slate-100 animate-slide-in">
            <div class="px-6 py-4 bg-amber-600 text-white flex items-center justify-between">
                <h3 class="font-bold text-md"><i class="fa-solid fa-pen-to-square mr-2"></i> Ajuste Manual (Override Auditável)</h3>
                <button onclick="toggleModal('modal-ajuste', false)" class="text-white/80 hover:text-white"><i class="fa-solid fa-xmark"></i></button>
            </div>
            <form id="form-ajuste" onsubmit="salvarAjuste(event)" class="p-6 space-y-4">
                <input type="hidden" id="ajuste-doc-id">
                <div>
                    <label class="block text-xs font-semibold text-slate-500 mb-1">Valor do Ajuste (R$):</label>
                    <input type="number" step="0.01" id="ajuste-valor" required class="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500" placeholder="Ex: 150.00 ou -50.00">
                    <p class="text-[10px] text-slate-400 mt-1">Valores positivos somam; valores negativos subtraem do total original da nota.</p>
                </div>
                <div>
                    <label class="block text-xs font-semibold text-slate-500 mb-1">Justificativa do Ajuste (Auditoria):</label>
                    <input type="text" id="ajuste-justificativa" required minlength="5" class="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500" placeholder="Ex: Glosa de frete faturado erroneamente">
                </div>
                <div>
                    <label class="block text-xs font-semibold text-slate-500 mb-1">Nome do Auditor / Usuário:</label>
                    <input type="text" id="ajuste-usuario" required class="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500" placeholder="Ex: Stenio Cardoso">
                </div>
                <div class="flex justify-end space-x-2 pt-2">
                    <button type="button" onclick="toggleModal('modal-ajuste', false)" class="px-4 py-2 border border-slate-200 hover:bg-slate-50 text-slate-600 rounded-lg text-sm font-medium">Cancelar</button>
                    <button type="submit" class="px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-sm font-medium">Salvar Ajuste</button>
                </div>
            </form>
        </div>
    </div>

    <!-- MODAL 3: RELATÓRIO DE APURAÇÃO (STRATEGY PATTERN REPORT) -->
    <div id="modal-apuracao" class="fixed inset-0 z-50 flex items-center justify-center bg-brand-900/40 backdrop-blur-sm hidden">
        <div class="bg-white rounded-2xl shadow-xl w-full max-w-lg mx-4 overflow-hidden border border-slate-100 animate-slide-in">
            <div class="px-6 py-4 bg-indigo-900 text-white flex items-center justify-between">
                <h3 class="font-bold text-md"><i class="fa-solid fa-calculator mr-2"></i> Apuração Tributária (Strategy Engine)</h3>
                <button onclick="toggleModal('modal-apuracao', false)" class="text-white/80 hover:text-white"><i class="fa-solid fa-xmark"></i></button>
            </div>
            
            <!-- CONTROLES DE ALÍQUOTA CUSTOMIZADA (DINÂMICO) -->
            <div class="px-6 py-3 bg-slate-50 border-b border-slate-100 flex items-center justify-between gap-4">
                <div class="flex items-center space-x-2">
                    <label class="text-xs font-semibold text-slate-500">Alíquota Aplicada (%):</label>
                    <input type="number" step="0.01" id="apuracao-aliquota-input" class="w-20 border border-slate-200 rounded px-2 py-1 text-xs focus:outline-none focus:ring-1 focus:ring-indigo-500 font-mono" placeholder="Ex: 6.00">
                </div>
                <button id="btn-recalcular-apuracao" class="px-3 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold rounded-lg transition-colors flex items-center">
                    <i class="fa-solid fa-rotate mr-1"></i>Recalcular
                </button>
            </div>

            <div id="apuracao-content" class="p-6 space-y-4">
                <!-- Carregado Dinamicamente -->
            </div>
            <div class="px-6 py-4 bg-slate-50 border-t border-slate-100 flex justify-end">
                <button onclick="toggleModal('modal-apuracao', false)" class="px-4 py-2 bg-indigo-900 hover:bg-indigo-950 text-white rounded-lg text-sm font-medium">Concluído</button>
            </div>
        </div>
    </div>

    <!-- MODAL 4: EDITAR CADASTRO DA EMPRESA -->
    <div id="modal-editar-empresa" class="fixed inset-0 z-50 flex items-center justify-center bg-brand-900/40 backdrop-blur-sm hidden">
        <div class="bg-white rounded-2xl shadow-xl w-full max-w-md mx-4 overflow-hidden border border-slate-100 animate-slide-in">
            <div class="px-6 py-4 bg-brand-900 text-white flex items-center justify-between">
                <h3 class="font-bold text-md"><i class="fa-solid fa-pen-to-square mr-2"></i> Editar Cadastro da Empresa</h3>
                <button onclick="toggleModal('modal-editar-empresa', false)" class="text-white/80 hover:text-white"><i class="fa-solid fa-xmark"></i></button>
            </div>
            <form id="form-editar-empresa" onsubmit="salvarEdicaoEmpresa(event)" class="p-6 space-y-4">
                <div>
                    <label class="block text-xs font-semibold text-slate-500 mb-1">CNPJ (Não alterável):</label>
                    <input type="text" id="editar-empresa-cnpj" disabled class="w-full bg-slate-100 border border-slate-200 rounded-lg px-3 py-2 text-sm text-slate-500 font-mono">
                </div>
                <div>
                    <label class="block text-xs font-semibold text-slate-500 mb-1">Razão Social:</label>
                    <input type="text" id="editar-empresa-razao" required class="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-brand-500" placeholder="Ex: Stenio Software Ltda">
                </div>
                <div class="relative">
                    <div class="flex items-center justify-between mb-1">
                        <label class="block text-xs font-semibold text-slate-500">CNAE Principal (Opcional):</label>
                        <button type="button" onclick="sincronizarCnaes(event)" class="text-[10px] text-brand-500 hover:text-brand-600 font-semibold focus:outline-none"><i class="fa-solid fa-rotate mr-1"></i>Sincronizar IBGE</button>
                    </div>
                    <input type="text" id="editar-empresa-cnae" class="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-brand-500" placeholder="Digite ou busque o CNAE..." autocomplete="off">
                    <div id="editar-empresa-cnae-results" class="absolute left-0 right-0 mt-1 max-h-60 overflow-y-auto bg-white border border-slate-200 rounded-lg shadow-lg hidden z-50"></div>
                    <div id="editar-empresa-cnae-info" class="mt-1.5 p-2 bg-indigo-50 border border-indigo-100 rounded text-[11px] text-indigo-700 hidden"></div>
                </div>
                <div>
                    <label class="block text-xs font-semibold text-slate-500 mb-1">Regime Tributário:</label>
                    <select id="editar-empresa-regime" required class="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-brand-500">
                        <option value="Simples Nacional">Simples Nacional</option>
                        <option value="Lucro Presumido">Lucro Presumido</option>
                        <option value="Lucro Real">Lucro Real</option>
                    </select>
                </div>
                <div>
                    <label class="block text-xs font-semibold text-slate-500 mb-1">Categoria do Simples Nacional:</label>
                    <select id="editar-empresa-categoria" class="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-brand-500">
                        <option value="Serviços (Anexo III)">Serviços (Anexo III)</option>
                        <option value="Serviços (Anexo IV)">Serviços (Anexo IV)</option>
                        <option value="Serviços (Anexo V - Fator R)">Serviços (Anexo V - Fator R)</option>
                        <option value="Comércio (Anexo I)">Comércio (Anexo I)</option>
                        <option value="Indústria (Anexo II)">Indústria (Anexo II)</option>
                    </select>
                </div>
                <div class="border-t border-slate-100 pt-3 mt-3">
                    <p class="text-xs font-bold text-slate-400 mb-2">Simples Nacional & Fator R (Opcional)</p>
                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block text-[10px] font-semibold text-slate-500 mb-1">RBT12 Acumulado (R$):</label>
                            <input type="number" step="0.01" id="editar-empresa-rbt12" class="w-full border border-slate-200 rounded-lg px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-brand-500" placeholder="Ex: 150000.00">
                        </div>
                        <div>
                            <label class="block text-[10px] font-semibold text-slate-500 mb-1">Folha Salários 12m (R$):</label>
                            <input type="number" step="0.01" id="editar-empresa-folha12" class="w-full border border-slate-200 rounded-lg px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-brand-500" placeholder="Ex: 45000.00">
                        </div>
                    </div>
                    <div class="flex items-center space-x-2 mt-2">
                        <input type="checkbox" id="editar-empresa-fator-r" class="rounded text-brand-500 focus:ring-brand-500 text-xs">
                        <label for="editar-empresa-fator-r" class="text-[11px] font-semibold text-slate-500 select-none">Atividade Sujeita ao Fator R</label>
                    </div>
                </div>
                <div class="flex justify-end space-x-2 pt-2">
                    <button type="button" onclick="toggleModal('modal-editar-empresa', false)" class="px-4 py-2 border border-slate-200 hover:bg-slate-50 text-slate-600 rounded-lg text-sm font-medium">Cancelar</button>
                    <button type="submit" class="px-4 py-2 bg-brand-500 hover:bg-brand-600 text-white rounded-lg text-sm font-medium">Salvar Alterações</button>
                </div>
            </form>
        </div>
    </div>

    <!-- DRAWER LATERAL DE AUDITORIA (MEMÓRIA DE CÁLCULO) -->
    <div id="drawer-auditoria" class="fixed right-0 top-0 h-full w-[450px] bg-white shadow-2xl z-50 transform translate-x-full transition-all duration-300 ease-in-out border-l border-slate-200 flex flex-col hidden">
        <!-- Drawer Header -->
        <div class="p-6 border-b border-slate-100 flex items-center justify-between bg-brand-900 text-white">
            <div>
                <h3 class="text-md font-bold flex items-center"><i class="fa-solid fa-calculator mr-2 text-brand-400"></i> Memória de Cálculo</h3>
                <p class="text-[11px] text-brand-100 opacity-80">Rastreabilidade e validação matemática de impostos</p>
            </div>
            <button onclick="fecharDrawerAuditoria()" class="text-white hover:text-slate-300 p-1">
                <i class="fa-solid fa-xmark text-xl"></i>
            </button>
        </div>
        <!-- Drawer Content -->
        <div class="flex-1 overflow-y-auto p-6 space-y-6 bg-slate-50/50" id="drawer-auditoria-content">
            <!-- Injetado via Javascript -->
        </div>
    </div>

    <!-- TOAST CONTAINER -->
    <div id="toast-container" class="fixed bottom-4 right-4 z-50 flex flex-col space-y-2"></div>

    <!-- ==========================================
         SCRIPT DE LÓGICA E INTEGRAÇÃO (API FETCH)
         ========================================== -->
    <script>
        const API_URL = '';
        const cnaeCache = {};
        let consolidadoDataGlobal = null;
        let activeTipoOperacao = '';

        function toggleForcarEntradaOptions() {
            const chk = document.getElementById('chk-forcar-entrada');
            const opts = document.getElementById('forcar-entrada-opcoes');
            if (chk.checked) {
                opts.classList.remove('hidden');
                setTimeout(() => opts.classList.remove('opacity-0'), 50);
                
                // Pre-enche os seletores de importação com a competência ativa
                const selectMes = document.getElementById('select-mes').value;
                const selectAno = document.getElementById('select-ano').value;
                if (selectMes) document.getElementById('import-mes').value = selectMes;
                if (selectAno) document.getElementById('import-ano').value = selectAno;
            } else {
                opts.classList.add('opacity-0');
                setTimeout(() => opts.classList.add('hidden'), 200);
            }
        }

        function fecharDrawerAuditoria() {
            const drawer = document.getElementById('drawer-auditoria');
            drawer.classList.add('translate-x-full');
            setTimeout(() => {
                drawer.classList.add('hidden');
            }, 300);
        }

        function setTipoOperacaoFilter(btn, value) {
            activeTipoOperacao = value;
            const container = btn.parentElement;
            container.querySelectorAll('.tab-btn').forEach(b => {
                b.classList.remove('border-brand-500', 'text-brand-600');
                b.classList.add('border-transparent', 'text-slate-500');
            });
            btn.classList.add('border-brand-500', 'text-brand-600');
            btn.classList.remove('border-transparent', 'text-slate-500');
            carregarDocumentos();
        }

        function abrirDrawerAuditoria() {
            if (!consolidadoDataGlobal || !consolidadoDataGlobal.memoria_calculo) {
                showToast("Memória de cálculo indisponível para esta empresa.", "warning");
                return;
            }
            
            const drawer = document.getElementById('drawer-auditoria');
            const content = document.getElementById('drawer-auditoria-content');
            drawer.classList.remove('hidden');
            // Timeout para transição CSS funcionar
            setTimeout(() => {
                drawer.classList.remove('translate-x-full');
            }, 50);
            
            const data = consolidadoDataGlobal;
            const mc = data.memoria_calculo;
            
            let html = '';
            
            if (data.regime === "Simples Nacional") {
                const rbt12 = Number(mc.rbt12);
                const folha12 = Number(mc.folha12);
                const fatorR = Number(mc.fator_r);
                const aliqNom = Number(mc.aliq_nom);
                const deducao = Number(mc.deducao);
                const aliqEfet = Number(mc.aliq_efetiva);
                const issShare = Number(mc.iss_share);
                const icmsShare = Number(mc.icms_share);
                
                const valComIss = Number(mc.valor_com_iss_retido);
                const valSemIss = Number(mc.valor_sem_iss_retido);
                const valComSt = Number(mc.valor_com_st);
                const valSemSt = Number(mc.valor_sem_st);
                const faturamentoTotal = Number(data.total_faturamento);
                
                // Anexo e Faixa
                html += `
                    <div class="space-y-4">
                        <div class="p-4 bg-indigo-50 rounded-2xl border border-indigo-100/60">
                            <h4 class="text-xs text-indigo-500 uppercase font-semibold">Enquadramento</h4>
                            <p class="text-md font-bold text-indigo-900 mt-1">${mc.enquadramento}</p>
                            <p class="text-xs text-indigo-600/80 mt-1">Categoria: ${mc.categoria_simples}</p>
                        </div>
                `;
                
                // Fator R
                if (mc.sujeito_fator_r) {
                    html += `
                        <div class="p-4 bg-white rounded-2xl border border-slate-100 space-y-2">
                            <h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider">Cálculo do Fator R</h4>
                            <div class="flex justify-between text-xs text-slate-600">
                                <span>Folha 12 meses:</span>
                                <span class="font-mono">${formatarMoeda(folha12)}</span>
                            </div>
                            <div class="flex justify-between text-xs text-slate-600">
                                <span>RBT12 meses:</span>
                                <span class="font-mono">${formatarMoeda(rbt12)}</span>
                            </div>
                            <div class="border-t border-slate-100 pt-2 flex justify-between items-center text-xs">
                                <span class="font-semibold text-slate-700">Equação Fator R:</span>
                                <span class="font-mono text-brand-600 font-bold">${formatarMoeda(folha12)} / ${formatarMoeda(rbt12)} = ${fatorR.toFixed(2)}%</span>
                            </div>
                            <div class="p-2 text-[10px] ${fatorR >= 28 ? 'bg-emerald-50 text-emerald-700 border-emerald-100' : 'bg-amber-50 text-amber-700 border-amber-100'} rounded border mt-2">
                                <i class="fa-solid fa-circle-info mr-1"></i> Fator R é de ${fatorR.toFixed(2)}% (${fatorR >= 28 ? '>= 28%: Tributado no Anexo III' : '< 28%: Tributado no Anexo V'}).
                            </div>
                        </div>
                    `;
                }
                
                // Alíquota Efetiva
                html += `
                    <div class="p-4 bg-white rounded-2xl border border-slate-100 space-y-2">
                        <h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider">Fórmula da Alíquota Efetiva</h4>
                        <div class="text-[11px] text-slate-400 font-mono py-1.5 text-center bg-slate-50 rounded border border-slate-100">
                            AE = (RBT12 * AliqNom - Deducao) / RBT12
                        </div>
                        <div class="flex justify-between text-xs text-slate-600">
                            <span>RBT12:</span>
                            <span class="font-mono">${formatarMoeda(rbt12)}</span>
                        </div>
                        <div class="flex justify-between text-xs text-slate-600">
                            <span>Alíquota Nominal (Anexo):</span>
                            <span class="font-mono">${aliqNom.toFixed(2)}%</span>
                        </div>
                        <div class="flex justify-between text-xs text-slate-600">
                            <span>Parcela a Deduzir:</span>
                            <span class="font-mono">${formatarMoeda(deducao)}</span>
                        </div>
                        <div class="border-t border-slate-100 pt-2 text-xs space-y-1">
                            <div class="flex justify-between items-center text-slate-700">
                                <span>Substituição numérica:</span>
                                <span class="font-mono text-[11px] text-slate-500">(${rbt12.toFixed(2)} * ${(aliqNom/100).toFixed(4)}) - ${deducao.toFixed(2)} / ${rbt12.toFixed(2)}</span>
                            </div>
                            <div class="flex justify-between items-center">
                                <span class="font-semibold text-slate-800">Alíquota Efetiva Resultante:</span>
                                <span class="font-mono font-bold text-brand-600">${aliqEfet.toFixed(4)}%</span>
                            </div>
                        </div>
                    </div>
                `;
                
                // Segregações
                html += `
                    <div class="p-4 bg-white rounded-2xl border border-slate-100 space-y-3">
                        <h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider">Segregações & Deduções</h4>
                `;
                
                if (mc.categoria_simples === "Comércio (Anexo I)" || mc.categoria_simples === "Indústria (Anexo II)") {
                    const ipiShare = Number(mc.ipi_share) || 0;
                    html += `
                        <div class="space-y-2 text-xs">
                            ${mc.categoria_simples === "Indústria (Anexo II)" ? `
                            <div class="flex justify-between text-slate-600">
                                <span>Fração IPI no Simples (Fixo):</span>
                                <span class="font-mono">${ipiShare.toFixed(2)}%</span>
                            </div>
                            ` : ''}
                            <div class="flex justify-between text-slate-600">
                                <span>Fração ICMS no Simples (Faixa):</span>
                                <span class="font-mono">${icmsShare.toFixed(2)}%</span>
                            </div>
                            <div class="flex justify-between text-slate-600">
                                <span>Faturamento sem ST:</span>
                                <span class="font-mono">${formatarMoeda(valSemSt)}</span>
                            </div>
                            <div class="flex justify-between text-slate-600">
                                <span>Faturamento com ST (Dedutível):</span>
                                <span class="font-mono">${formatarMoeda(valComSt)}</span>
                            </div>
                            <div class="bg-indigo-50/50 p-2.5 rounded-xl border border-indigo-100 text-[11px] text-indigo-950 space-y-1">
                                <p class="font-bold flex items-center"><i class="fa-solid fa-percent mr-1"></i> Alíquotas Aplicadas:</p>
                                <p>• Sem ST: ${aliqEfet.toFixed(4)}%</p>
                                <p>• Com ST (ICMS Deduzido): ${(aliqEfet * (1 - icmsShare/100)).toFixed(4)}%</p>
                                ${mc.categoria_simples === "Indústria (Anexo II)" ? `<p class="text-[10px] text-indigo-600 mt-1">Obs: IPI de ${ipiShare.toFixed(2)}% incluído na alíquota unificada.</p>` : ''}
                            </div>
                        </div>
                    `;
                } else {
                    html += `
                        <div class="space-y-2 text-xs">
                            <div class="flex justify-between text-slate-600">
                                <span>Fração ISS no Simples (Faixa):</span>
                                <span class="font-mono">${issShare.toFixed(2)}%</span>
                            </div>
                            <div class="flex justify-between text-slate-600">
                                <span>Faturamento sem Retenção ISS:</span>
                                <span class="font-mono">${formatarMoeda(valSemIss)}</span>
                            </div>
                            <div class="flex justify-between text-slate-600">
                                <span>Faturamento com Retenção ISS:</span>
                                <span class="font-mono">${formatarMoeda(valComIss)}</span>
                            </div>
                            <div class="bg-indigo-50/50 p-2.5 rounded-xl border border-indigo-100 text-[11px] text-indigo-950 space-y-1">
                                <p class="font-bold flex items-center"><i class="fa-solid fa-percent mr-1 text-brand-600"></i> Alíquotas Aplicadas:</p>
                                <p>• Operações Normais: <span class="font-mono font-bold">${aliqEfet.toFixed(4)}%</span></p>
                                <p>• Com ISS Retido na fonte: <span class="font-mono font-bold text-emerald-600">${(aliqEfet * (1 - issShare/100)).toFixed(4)}%</span> (ISS ${issShare.toFixed(1)}% deduzido)</p>
                            </div>
                        </div>
                    `;
                }
                
                html += `</div>`;
                
                // Lembrete Previdenciário do Anexo IV (CPP por fora)
                if (mc.categoria_simples && mc.categoria_simples.includes("Anexo IV")) {
                    html += `
                        <div class="p-4 bg-indigo-50 rounded-2xl border border-indigo-100/60 space-y-2 text-xs text-indigo-900">
                            <h4 class="font-bold flex items-center text-indigo-800"><i class="fa-solid fa-circle-exclamation mr-1.5 text-indigo-600"></i> Lembrete Previdenciário (CPP)</h4>
                            <p class="text-[11px] leading-relaxed opacity-90">
                                Para empresas do <strong>Anexo IV do Simples Nacional</strong>, a Contribuição Previdenciária Patronal (CPP) de 20% não está incluída na guia do DAS. Ela deve ser calculada e recolhida separadamente sobre a folha de pagamento da empresa.
                            </p>
                        </div>
                    `;
                }
                
                // Comparativo de Arredondamento
                html += `
                    <div class="p-4 bg-amber-50/50 rounded-2xl border border-amber-100/60 space-y-2 text-xs text-amber-900">
                        <h4 class="font-bold flex items-center text-amber-800"><i class="fa-solid fa-triangle-exclamation mr-1.5"></i> Diferença de Arredondamento</h4>
                        <p class="text-[11px] leading-relaxed opacity-90">
                            Sistemas de ERP (como o Domínio) costumam realizar a apuração aplicando a alíquota efetiva consolidada sobre a soma total das notas fiscais. O Antigravity faz o cálculo nota a nota e soma os resultados. Podem surgir discrepâncias de centavos de arredondamento.
                        </p>
                    </div>
                `;
                
            } else if (data.regime === "Lucro Presumido") {
                const valComSt = Number(mc.valor_com_st);
                const valSemSt = Number(mc.valor_sem_st);
                const aliqPis = Number(mc.aliquota_pis);
                const aliqCofins = Number(mc.aliquota_cofins);
                const aliqIrpj = Number(mc.aliquota_irpj);
                const aliqCsll = Number(mc.aliquota_csll);
                const aliqIss = Number(mc.aliquota_iss);
                
                const pis = Number(mc.pis);
                const cofins = Number(mc.cofins);
                const irpj = Number(mc.irpj);
                const csll = Number(mc.csll);
                const iss = Number(mc.iss);
                const faturamentoTotal = Number(data.total_faturamento);
                
                html += `
                    <div class="space-y-4">
                        <div class="p-4 bg-indigo-50 rounded-2xl border border-indigo-100/60">
                            <h4 class="text-xs text-indigo-500 uppercase font-semibold">Regime Tributário</h4>
                            <p class="text-md font-bold text-indigo-900 mt-1">Lucro Presumido (Serviços)</p>
                            <p class="text-xs text-indigo-600/80 mt-1">Base de Presunção Padrão: 32.00%</p>
                        </div>
                        
                        <div class="p-4 bg-white rounded-2xl border border-slate-100 space-y-3">
                            <h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider">Abertura por Tributo e Alíquotas</h4>
                            <div class="space-y-2 text-xs">
                                <div class="flex justify-between items-center py-1 border-b border-slate-100">
                                    <div>
                                        <p class="font-semibold text-slate-800">PIS (${aliqPis.toFixed(2)}%)</p>
                                        <p class="text-[10px] text-slate-400 font-mono">${formatarMoeda(faturamentoTotal)} * ${(aliqPis/100).toFixed(4)}</p>
                                    </div>
                                    <span class="font-mono font-bold text-slate-900">${formatarMoeda(pis)}</span>
                                </div>
                                <div class="flex justify-between items-center py-1 border-b border-slate-100">
                                    <div>
                                        <p class="font-semibold text-slate-800">COFINS (${aliqCofins.toFixed(2)}%)</p>
                                        <p class="text-[10px] text-slate-400 font-mono">${formatarMoeda(faturamentoTotal)} * ${(aliqCofins/100).toFixed(4)}</p>
                                    </div>
                                    <span class="font-mono font-bold text-slate-900">${formatarMoeda(cofins)}</span>
                                </div>
                                <div class="flex justify-between items-center py-1 border-b border-slate-100">
                                    <div>
                                        <p class="font-semibold text-slate-800">IRPJ (${aliqIrpj.toFixed(2)}%)</p>
                                        <p class="text-[10px] text-slate-400 font-mono">${formatarMoeda(faturamentoTotal)} * ${(aliqIrpj/100).toFixed(4)} (32% * 15%)</p>
                                    </div>
                                    <span class="font-mono font-bold text-slate-900">${formatarMoeda(irpj)}</span>
                                </div>
                                <div class="flex justify-between items-center py-1 border-b border-slate-100">
                                    <div>
                                        <p class="font-semibold text-slate-800">CSLL (${aliqCsll.toFixed(2)}%)</p>
                                        <p class="text-[10px] text-slate-400 font-mono">${formatarMoeda(faturamentoTotal)} * ${(aliqCsll/100).toFixed(4)} (32% * 9%)</p>
                                    </div>
                                    <span class="font-mono font-bold text-slate-900">${formatarMoeda(csll)}</span>
                                </div>
                                ${aliqIss > 0 ? `
                                <div class="flex justify-between items-center py-1 border-b border-slate-100">
                                    <div>
                                        <p class="font-semibold text-slate-800 text-brand-600">ISS Municipal (${aliqIss.toFixed(2)}%)</p>
                                        <p class="text-[10px] text-slate-400 font-mono">${formatarMoeda(faturamentoTotal)} * ${(aliqIss/100).toFixed(4)}</p>
                                    </div>
                                    <span class="font-mono font-bold text-brand-700">${formatarMoeda(iss)}</span>
                                </div>
                                ` : ''}
                            </div>
                        </div>
                    </div>
                `;
            }
            
            content.innerHTML = html;
        }

        // Sincroniza os CNAEs com o IBGE a partir da API
        async function sincronizarCnaes(event) {
            const btn = event.currentTarget;
            const originalHTML = btn.innerHTML;
            btn.disabled = true;
            btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin mr-1"></i>Sincronizando...`;
            
            try {
                const res = await fetch('/api/cnaes/sync', { method: 'POST' });
                const data = await res.json();
                if (res.ok) {
                    showToast(`Sincronização concluída! ${data.total} CNAEs carregados.`, 'success');
                } else {
                    showToast(data.detail || 'Erro ao sincronizar CNAEs', 'error');
                }
            } catch (err) {
                console.error(err);
                showToast('Falha na requisição de sincronização', 'error');
            } finally {
                btn.disabled = false;
                btn.innerHTML = originalHTML;
            }
        }

        // Inicializa o componente de autocomplete dinâmico
        function initCnaeAutocomplete(prefix) {
            const input = document.getElementById(`${prefix}-cnae`);
            const resultsDiv = document.getElementById(`${prefix}-cnae-results`);
            const infoBox = document.getElementById(`${prefix}-cnae-info`);
            const regimeSelect = document.getElementById(`${prefix}-regime`);
            const categoriaSelect = document.getElementById(`${prefix}-categoria`);
            const fatorRCheckbox = document.getElementById(`${prefix}-fator-r`);
            
            let debounceTimeout = null;

            input.addEventListener('input', () => {
                const query = input.value.trim();
                
                clearTimeout(debounceTimeout);

                if (!query) {
                    resultsDiv.innerHTML = '';
                    resultsDiv.classList.add('hidden');
                    infoBox.classList.add('hidden');
                    return;
                }

                debounceTimeout = setTimeout(() => {
                    fetch(`/api/cnaes?q=${encodeURIComponent(query)}`)
                        .then(res => res.json())
                        .then(data => {
                            resultsDiv.innerHTML = '';
                            if (data && data.length > 0) {
                                data.forEach(item => {
                                    const option = document.createElement('div');
                                    option.className = "px-3 py-2 hover:bg-slate-50 cursor-pointer text-xs border-b border-slate-100 last:border-0 transition-colors";
                                    option.innerHTML = `<span class="font-mono font-bold text-brand-600">${item.codigo}</span> - <span class="text-slate-600 font-medium">${item.descricao}</span> <span class="ml-1 text-[9px] bg-indigo-50 border border-indigo-100 px-1 py-0.5 rounded text-indigo-500">${item.anexo ? 'Anexo ' + item.anexo : ''}</span>`;
                                    
                                    option.addEventListener('click', () => {
                                        input.value = item.codigo;
                                        cnaeCache[item.codigo] = item;
                                        
                                        if (regimeSelect) regimeSelect.value = "Simples Nacional";
                                        if (categoriaSelect) {
                                            if (item.anexo === "I") {
                                                categoriaSelect.value = "Comércio (Anexo I)";
                                            } else if (item.anexo === "II") {
                                                categoriaSelect.value = "Indústria (Anexo II)";
                                            } else if (item.anexo === "V") {
                                                categoriaSelect.value = "Serviços (Anexo V - Fator R)";
                                            } else if (item.anexo === "IV") {
                                                categoriaSelect.value = "Serviços (Anexo IV)";
                                            } else {
                                                categoriaSelect.value = "Serviços (Anexo III)";
                                            }
                                        }
                                        if (fatorRCheckbox) {
                                            fatorRCheckbox.checked = item.fator_r;
                                        }
                                        
                                        let extraInfo = item.fator_r ? " (Sujeito a Fator R)" : "";
                                        infoBox.innerHTML = `
                                            <span class="font-semibold"><i class="fa-solid fa-circle-info mr-1"></i>${item.descricao}</span><br>
                                            <span>Anexo ${item.anexo} | Alíquota nominal inicial: <strong>${item.aliquota.toFixed(2)}%</strong>${extraInfo}</span>
                                        `;
                                        infoBox.classList.remove('hidden');
                                        resultsDiv.classList.add('hidden');
                                    });
                                    resultsDiv.appendChild(option);
                                });
                                resultsDiv.classList.remove('hidden');
                            } else {
                                resultsDiv.innerHTML = '<div class="px-3 py-2 text-xs text-slate-400">Nenhum CNAE encontrado</div>';
                                resultsDiv.classList.remove('hidden');
                            }
                        })
                        .catch(err => console.error("Erro ao buscar CNAEs:", err));
                }, 300);
            });

            document.addEventListener('click', (e) => {
                if (e.target !== input && e.target !== resultsDiv && !resultsDiv.contains(e.target)) {
                    resultsDiv.classList.add('hidden');
                }
            });

            input.addEventListener('focus', () => {
                if (input.value.trim() && resultsDiv.children.length > 0) {
                    resultsDiv.classList.remove('hidden');
                }
            });
        }

        async function aplicarRegrasCNAE(prefix) {
            const cnaeInput = document.getElementById(`${prefix}-cnae`);
            const infoBox = document.getElementById(`${prefix}-cnae-info`);
            
            const rawCnae = cnaeInput.value.trim().replace(/\D/g, "");
            if (!rawCnae) {
                infoBox.classList.add('hidden');
                return;
            }

            if (cnaeCache[rawCnae]) {
                const item = cnaeCache[rawCnae];
                let extraInfo = item.fator_r ? " (Sujeito a Fator R)" : "";
                infoBox.innerHTML = `
                    <span class="font-semibold"><i class="fa-solid fa-circle-info mr-1"></i>${item.descricao}</span><br>
                    <span>Anexo ${item.anexo} | Alíquota nominal inicial: <strong>${item.aliquota.toFixed(2)}%</strong>${extraInfo}</span>
                `;
                infoBox.classList.remove('hidden');
            } else {
                try {
                    const res = await fetch(`/api/cnaes?q=${rawCnae}`);
                    const data = await res.json();
                    if (data && data.length > 0) {
                        const item = data.find(c => c.codigo === rawCnae);
                        if (item) {
                            cnaeCache[rawCnae] = item;
                            let extraInfo = item.fator_r ? " (Sujeito a Fator R)" : "";
                            infoBox.innerHTML = `
                                <span class="font-semibold"><i class="fa-solid fa-circle-info mr-1"></i>${item.descricao}</span><br>
                                <span>Anexo ${item.anexo} | Alíquota nominal inicial: <strong>${item.aliquota.toFixed(2)}%</strong>${extraInfo}</span>
                            `;
                            infoBox.classList.remove('hidden');
                        } else {
                            infoBox.classList.add('hidden');
                        }
                    } else {
                        infoBox.classList.add('hidden');
                    }
                } catch (err) {
                    console.error("Erro ao carregar regras do CNAE:", err);
                    infoBox.classList.add('hidden');
                }
            }
        }

        // Executado no carregamento
        window.addEventListener('DOMContentLoaded', () => {
            carregarEmpresas();
            initCnaeAutocomplete('empresa');
            initCnaeAutocomplete('editar-empresa');
        });

        // Exibe alertas elegantes na tela (Toasts)
        function showToast(message, type = 'success') {
            const container = document.getElementById('toast-container');
            const toast = document.createElement('div');
            toast.className = `toast glass flex items-center space-x-3 px-4 py-3 rounded-xl shadow-lg border-l-4 text-sm font-medium ${
                type === 'success' ? 'border-emerald-500 text-emerald-950 bg-emerald-50/90' : 
                type === 'error' ? 'border-rose-500 text-rose-950 bg-rose-50/90' : 
                'border-amber-500 text-amber-950 bg-amber-50/90'
            }`;
            
            const icon = type === 'success' ? 'fa-circle-check text-emerald-500' : 
                         type === 'error' ? 'fa-triangle-exclamation text-rose-500' : 
                         'fa-circle-exclamation text-amber-500';

            toast.innerHTML = `<i class="fa-solid ${icon} text-lg"></i><span>${message}</span>`;
            container.appendChild(toast);
            
            setTimeout(() => {
                toast.classList.add('opacity-0', 'transition-opacity', 'duration-300');
                setTimeout(() => toast.remove(), 300);
            }, 4000);
        }

        // Abre ou fecha modais
        function toggleModal(modalId, show) {
            const modal = document.getElementById(modalId);
            if (show) {
                modal.classList.remove('hidden');
            } else {
                modal.classList.add('hidden');
                // Limpa formulários ao fechar
                const form = modal.querySelector('form');
                if (form) form.reset();
            }
        }

        // ==========================================
        // FLUXO DE EMPRESAS
        // ==========================================

        async function carregarEmpresas() {
            try {
                const resEmp = await fetch(`${API_URL}/empresas`);
                if (resEmp.ok) {
                    const empresas = await resEmp.json();
                    const select = document.getElementById('select-empresa');
                    select.innerHTML = '';
                    
                    // Opção padrão de seleção (sem filtro ativo)
                    const placeholderOpt = document.createElement('option');
                    placeholderOpt.value = "";
                    placeholderOpt.textContent = "-- Selecione uma Empresa --";
                    select.appendChild(placeholderOpt);
                    
                    empresas.forEach(emp => {
                        const opt = document.createElement('option');
                        opt.value = emp.id;
                        opt.textContent = emp.razao_social;
                        opt.dataset.cnpj = emp.cnpj;
                        opt.dataset.regime = emp.regime_tributario;
                        opt.dataset.rbt12 = emp.rbt12;
                        opt.dataset.folha12 = emp.folha12;
                        opt.dataset.sujeito_fator_r = emp.sujeito_fator_r;
                        opt.dataset.cnae = emp.cnae || "";
                        opt.dataset.categoria = emp.categoria_simples || "";
                        select.appendChild(opt);
                    });
                    
                    atualizarInfoEmpresa();
                }
            } catch (err) {
                console.error("Erro ao carregar empresas:", err);
            }
        }

        function atualizarInfoEmpresa() {
            const select = document.getElementById('select-empresa');
            const opt = select.selectedOptions[0];
            const infoBox = document.getElementById('info-empresa-box');
            
            if (opt && opt.value && opt.value !== "auto") {
                document.getElementById('info-empresa-cnpj').textContent = formatarCNPJ(opt.dataset.cnpj);
                
                let regimeText = opt.dataset.regime;
                if (opt.dataset.regime === "Simples Nacional") {
                    const rbt12Val = parseFloat(opt.dataset.rbt12 || 0);
                    const folha12Val = parseFloat(opt.dataset.folha12 || 0);
                    const fatorR = rbt12Val > 0 ? (folha12Val / rbt12Val) * 100 : 0;
                    const sujeitoFatorR = opt.dataset.sujeito_fator_r === "true";
                    
                    regimeText += ` (RBT12: ${formatarMoeda(rbt12Val)}`;
                    if (sujeitoFatorR) {
                        regimeText += ` | Fator R: ${fatorR.toFixed(2)}%`;
                    }
                    regimeText += `)`;
                }
                
                document.getElementById('info-empresa-regime').textContent = regimeText;
                
                const cnaeVal = opt.dataset.cnae;
                const cnaeWrapper = document.getElementById('info-empresa-cnae-wrapper');
                if (cnaeVal) {
                    const cleanCnae = cnaeVal.replace(/\D/g, "");
                    document.getElementById('info-empresa-cnae').textContent = cnaeVal;
                    cnaeWrapper.classList.remove('hidden');
                    
                    if (cnaeCache[cleanCnae]) {
                        document.getElementById('info-empresa-cnae').textContent = `${cnaeVal} - ${cnaeCache[cleanCnae].descricao}`;
                    } else {
                        fetch(`/api/cnaes?q=${cleanCnae}`)
                            .then(res => res.json())
                            .then(data => {
                                if (data && data.length > 0) {
                                    const match = data.find(c => c.codigo === cleanCnae);
                                    if (match) {
                                        cnaeCache[cleanCnae] = match;
                                        document.getElementById('info-empresa-cnae').textContent = `${cnaeVal} - ${match.descricao}`;
                                    }
                                }
                            })
                            .catch(err => console.error("Erro ao buscar descrição do CNAE:", err));
                    }
                } else {
                    cnaeWrapper.classList.add('hidden');
                }
                
                infoBox.classList.remove('hidden');
            } else {
                infoBox.classList.add('hidden');
            }
        }

        async function salvarEmpresa(event) {
            event.preventDefault();
            const cnpj = document.getElementById('empresa-cnpj').value;
            const razao = document.getElementById('empresa-razao').value;
            const cnae = document.getElementById('empresa-cnae').value.trim().replace(/\D/g, "");
            const regime = document.getElementById('empresa-regime').value;
            const categoria_simples = document.getElementById('empresa-categoria').value;
            
            const rbt12 = parseFloat(document.getElementById('empresa-rbt12').value || 0);
            const folha12 = parseFloat(document.getElementById('empresa-folha12').value || 0);
            const sujeito_fator_r = document.getElementById('empresa-fator-r').checked;

            try {
                const res = await fetch(`${API_URL}/empresas`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ 
                        cnpj, 
                        razao_social: razao, 
                        regime_tributario: regime,
                        rbt12,
                        folha12,
                        sujeito_fator_r,
                        categoria_simples,
                        cnae: cnae || null
                    })
                });

                if (res.ok) {
                    showToast("Empresa cadastrada com sucesso!");
                    toggleModal('modal-empresa', false);
                    await carregarEmpresas();
                    
                    // Foca na empresa criada
                    const empresas = await (await fetch(`${API_URL}/empresas`)).json();
                    const novaEmp = empresas.find(e => e.cnpj === cnpj);
                    if (novaEmp) {
                        document.getElementById('select-empresa').value = novaEmp.id;
                        await carregarDocumentos();
                    }
                } else {
                    const err = await res.json();
                    showToast(err.detail || "Erro ao cadastrar empresa", "error");
                }
            } catch (err) {
                showToast("Erro de rede ao cadastrar empresa", "error");
            }
        }

        // ==========================================
        // FLUXO DE DOCUMENTOS (XMLs)
        // ==========================================

        async function carregarDocumentos() {
            atualizarInfoEmpresa();
            const select = document.getElementById('select-empresa');
            const empresaId = select.value;
            const tbody = document.getElementById('documentos-tbody');

            if (!empresaId) {
                document.getElementById('consolidado-section').classList.add('hidden');
                tbody.innerHTML = `
                    <tr>
                        <td colspan="10" class="px-6 py-12 text-center text-slate-400">
                            <i class="fa-solid fa-filter text-3xl mb-2 text-slate-300 block"></i>
                            Selecione uma empresa no filtro para visualizar a Staging Area e a apuração fiscal.
                        </td>
                    </tr>
                `;
                document.getElementById('doc-counter').textContent = "0 documentos";
                return;
            }

            const mesVal = document.getElementById('select-mes').value;
            const anoVal = document.getElementById('select-ano').value;
            let url = `${API_URL}/documentos?empresa_id=${empresaId}`;
            if (mesVal) url += `&mes=${mesVal}`;
            if (anoVal) url += `&ano=${anoVal}`;
            if (activeTipoOperacao) url += `&tipo_operacao=${activeTipoOperacao}`;

            try {
                const res = await fetch(url);
                if (res.ok) {
                    const docs = await res.json();
                    tbody.innerHTML = '';
                    document.getElementById('select-all-docs').checked = false;
                    atualizarAcoesLote();
                    document.getElementById('doc-counter').textContent = `${docs.length} documento(s)`;

                    if (docs.length === 0) {
                        document.getElementById('consolidado-section').classList.add('hidden');
                        tbody.innerHTML = `
                            <tr>
                                <td colspan="10" class="px-6 py-12 text-center text-slate-400">
                                    <i class="fa-solid fa-folder-open text-3xl mb-2 text-slate-300 block"></i>
                                    Nenhum XML importado na Staging Area desta empresa.
                                </td>
                            </tr>
                        `;
                        return;
                    }

                    docs.forEach(doc => {
                        const tr = document.createElement('tr');
                        tr.className = "hover:bg-slate-50/80 transition-colors";
                        
                        // Formatação de valores
                        const valXml = formatarMoeda(doc.valor_total);
                        const valFinal = formatarMoeda(doc.valor_final);
                        const temAjuste = Number(doc.valor_total) !== Number(doc.valor_final);
                        
                        // Formatação de emissão e parceiro
                        const dtEmi = doc.data_emissao ? new Date(doc.data_emissao).toLocaleDateString('pt-BR', { timeZone: 'UTC' }) : '---';
                        const parceiroNome = doc.tipo_operacao === "Entrada" ? (doc.emitente_nome || 'Não informado') : (doc.destinatario_nome || 'Não informado');
                        const parceiroLabel = doc.tipo_operacao === "Entrada" ? 'Fornecedor' : 'Cliente';
                        const labelBadgeColor = doc.tipo_operacao === "Entrada" ? 'bg-blue-50 text-blue-700 border-blue-100' : 'bg-violet-50 text-violet-700 border-violet-100';
                        
                        const parceiroHtml = `
                            <div class="flex flex-col">
                                <span class="inline-flex items-center px-1.5 py-0.5 w-max rounded text-[9px] font-bold border ${labelBadgeColor} uppercase tracking-wider mb-1">${parceiroLabel}</span>
                                <span class="text-xs font-semibold text-slate-700 max-w-[200px] truncate" title="${parceiroNome}">${parceiroNome}</span>
                            </div>
                        `;

                        // Badge de status
                        let statusColor = "bg-amber-100 text-amber-700 border border-amber-200/50";
                        if (doc.status_apuracao === "Em Revisão") {
                            statusColor = "bg-orange-100 text-orange-700 border border-orange-200/50";
                        } else if (doc.status_apuracao === "Encerrado") {
                            statusColor = "bg-emerald-100 text-emerald-700 border border-emerald-200/50";
                        }

                        // Ações desativadas se encerrado
                        const isEncerrado = doc.status_apuracao === "Encerrado";

                        tr.innerHTML = `
                            <td class="px-6 py-4"><input type="checkbox" class="doc-checkbox rounded text-brand-500 focus:ring-brand-500" value="${doc.id}" onchange="atualizarAcoesLote()"></td>
                            <td class="px-6 py-4 font-bold text-slate-900">${doc.numero_nf || '---'}</td>
                            <td class="px-6 py-4 text-slate-500 font-medium">${dtEmi}</td>
                            <td class="px-6 py-4">${parceiroHtml}</td>
                            <td class="px-6 py-4 font-mono font-medium text-slate-400" title="${doc.chave_acesso}">
                                ${doc.chave_acesso.substring(0, 8)}...${doc.chave_acesso.substring(36)}
                                <button onclick="navigator.clipboard.writeText('${doc.chave_acesso}'); showToast('Chave copiada!', 'info')" class="ml-1 text-slate-300 hover:text-slate-500">
                                    <i class="fa-regular fa-copy"></i>
                                </button>
                            </td>
                            <td class="px-6 py-4">
                                <span class="inline-flex items-center px-2 py-0.5 rounded text-xs font-semibold bg-slate-100 text-slate-600">
                                    ${doc.tipo_documento}
                                </span>
                            </td>
                            <td class="px-6 py-4 text-right font-medium text-slate-500">${valXml}</td>
                            <td class="px-6 py-4 text-right font-bold ${temAjuste ? 'text-amber-600' : 'text-slate-900'}">
                                ${valFinal}
                                ${temAjuste ? `<span class="block text-[10px] font-medium text-amber-500">Com override</span>` : ''}
                            </td>
                            <td class="px-6 py-4">
                                ${doc.cstat === "101" ? 
                                    `<span class="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-bold bg-rose-100 text-rose-700 border border-rose-200/50"><i class="fa-solid fa-ban mr-1"></i>Cancelada</span>` :
                                  ["110", "301", "302"].includes(doc.cstat) ?
                                    `<span class="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-bold bg-slate-100 text-slate-700 border border-slate-200/50"><i class="fa-solid fa-triangle-exclamation mr-1"></i>Denegada</span>` :
                                    `<span class="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-semibold ${statusColor}">${doc.status_apuracao}</span>`
                                }
                            </td>
                            <td class="px-6 py-4 text-center">
                                <div class="flex items-center justify-center space-x-1">
                                    <button onclick="abrirAjuste(${doc.id})" ${isEncerrado ? 'disabled' : ''} class="px-2.5 py-1.5 bg-amber-500 hover:bg-amber-600 disabled:opacity-30 text-white text-xs font-semibold rounded-lg transition-colors" title="Ajuste Manual">
                                        <i class="fa-solid fa-pen-to-square"></i>
                                    </button>
                                    <button onclick="abrirApuracao(${doc.id})" class="px-2.5 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold rounded-lg transition-colors" title="Ver Apuração">
                                        <i class="fa-solid fa-calculator"></i>
                                    </button>
                                    <button onclick="encerrarDocumento(${doc.id})" ${isEncerrado ? 'disabled' : ''} class="px-2.5 py-1.5 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-30 text-white text-xs font-semibold rounded-lg transition-colors" title="Encerrar / Snapshot">
                                        <i class="fa-solid fa-lock"></i>
                                    </button>
                                    <button onclick="deletarDocumento(${doc.id}, '${doc.chave_acesso}')" class="px-2.5 py-1.5 bg-rose-600 hover:bg-rose-700 text-white text-xs font-semibold rounded-lg transition-colors" title="Excluir Nota Fiscal">
                                        <i class="fa-solid fa-trash-can"></i>
                                    </button>
                                </div>
                            </td>
                        `;
                        tbody.appendChild(tr);
                    });

                    // Chama a apuração consolidada ao renderizar a tabela
                    await carregarApuracaoConsolidada(empresaId);
                }
            } catch (err) {
                showToast("Erro ao carregar documentos.", "error");
            }
        }

        async function carregarApuracaoConsolidada(empresaId) {
            const section = document.getElementById('consolidado-section');
            const mesVal = document.getElementById('select-mes').value;
            const anoVal = document.getElementById('select-ano').value;
            let url = `${API_URL}/empresas/${empresaId}/consolidado`;
            let params = [];
            if (mesVal) params.push(`mes=${mesVal}`);
            if (anoVal) params.push(`ano=${anoVal}`);
            if (params.length > 0) {
                url += `?${params.join('&')}`;
            }

            try {
                const res = await fetch(url);
                if (res.ok) {
                    const data = await res.json();
                    consolidadoDataGlobal = data; // Armazena globalmente
                    
                    document.getElementById('regime-badge').textContent = data.regime;
                    document.getElementById('c-faturamento').textContent = formatarMoeda(data.total_faturamento);
                    document.getElementById('c-imposto').textContent = formatarMoeda(data.total_imposto);
                    document.getElementById('c-aliquota').textContent = `${(data.aliquota_efetiva_consolidada * 100).toFixed(2)}%`;
                    
                    // Atualiza o contador de notas com divisão ativo/cancelado
                    let counterText = `${data.quantidade_ativos} ativo(s)`;
                    if (data.quantidade_cancelados > 0) {
                        counterText += ` | ${data.quantidade_cancelados} cancelado(s)`;
                    }
                    document.getElementById('doc-counter').textContent = `${data.quantidade_documentos} documento(s) (${counterText})`;
                    
                    const compras = data.compras || { total_compras: 0, total_difal: 0, total_icms_st: 0, quantidade_entradas: 0 };
                    document.getElementById('c-compras-total').textContent = formatarMoeda(compras.total_compras);
                    document.getElementById('c-compras-difal').textContent = formatarMoeda(compras.total_difal);
                    document.getElementById('c-compras-icms-st').textContent = formatarMoeda(compras.total_icms_st);
                    document.getElementById('c-compras-count').textContent = `${compras.quantidade_entradas} nota(s) de compra no período`;
                    
                    const list = document.getElementById('c-detalhes-list');
                    list.innerHTML = '';
                    
                    for (const [key, value] of Object.entries(data.detalhes)) {
                        const div = document.createElement('div');
                        div.className = "flex justify-between items-center";
                        div.innerHTML = `
                            <span class="font-bold text-slate-400 uppercase mr-2">${key}:</span>
                            <span class="font-mono font-bold text-slate-800">${formatarMoeda(value)}</span>
                        `;
                        list.appendChild(div);
                    }
                    
                    section.classList.remove('hidden');
                } else {
                    section.classList.add('hidden');
                }
            } catch (err) {
                console.error("Erro ao carregar consolidado:", err);
                section.classList.add('hidden');
            }
        }

        // Ingestão via File Selection
        function lidarComSelecaoArquivo(event) {
            const files = event.target.files;
            if (files && files.length > 0) fazerUploadXML(files);
        }

        // Ingestão via Drag & Drop
        function lidarComDrop(event) {
            event.preventDefault();
            document.getElementById('dropzone').classList.remove('border-brand-500', 'bg-brand-50/20');
            const files = event.dataTransfer.files;
            if (files && files.length > 0) fazerUploadXML(files);
        }

        async function fazerUploadXML(files) {
            // Filtra os arquivos válidos (apenas .xml)
            const xmlFiles = Array.from(files).filter(file => file.name.toLowerCase().endsWith('.xml'));

            if (xmlFiles.length === 0) {
                showToast("Nenhum arquivo XML válido selecionado.", "error");
                return;
            }

            const chkForcar = document.getElementById('chk-forcar-entrada');
            const selectEmpresa = document.getElementById('select-empresa');
            
            const formData = new FormData();
            
            if (selectEmpresa && selectEmpresa.value) {
                formData.append("empresa_id", selectEmpresa.value);
            }
            
            if (chkForcar && chkForcar.checked) {
                const empresaId = selectEmpresa.value;
                if (!empresaId) {
                    showToast("Por favor, selecione uma empresa ativa antes de importar notas de entrada vinculadas.", "error");
                    return;
                }
                
                const mes = document.getElementById('import-mes').value;
                const ano = document.getElementById('import-ano').value;
                const paddedMes = mes.padStart(2, '0');
                const competencia = `${ano}-${paddedMes}-01`;
                
                formData.set("empresa_id", empresaId);
                formData.append("tipo_operacao_forcada", "Entrada");
                formData.append("data_competencia", competencia);
            }
            
            // Adiciona múltiplos arquivos no form-data sob a chave "files"
            xmlFiles.forEach(file => {
                formData.append("files", file);
            });

            const uploadIcon = document.getElementById('upload-icon');
            const uploadText = document.getElementById('upload-text');

            // Feedback visual de progresso
            uploadIcon.className = "fa-solid fa-circle-notch text-4xl text-brand-500 animate-spin mb-3";
            uploadText.textContent = `Processando e Normalizando ${xmlFiles.length} XML(s)...`;

            try {
                const res = await fetch(`${API_URL}/documentos/upload`, {
                    method: 'POST',
                    body: formData
                });

                if (res.ok) {
                    const docsCriados = await res.json();
                    
                    if (docsCriados.length > 0) {
                        showToast(`${docsCriados.length} Nota(s) Fiscal(is) importada(s) com sucesso!`);
                        
                        // Recarrega as empresas (no caso de autocadastro)
                        await carregarEmpresas();
                        
                        // Seleciona automaticamente a empresa do primeiro documento importado se nenhuma estiver selecionada
                        const select = document.getElementById('select-empresa');
                        if (!select.value) {
                            select.value = docsCriados[0].empresa_id;
                        }
                        await carregarDocumentos();
                    } else {
                        showToast("Nenhuma nova nota foi importada (todas já existiam ou eram inválidas).", "warning");
                    }
                } else {
                    const err = await res.json();
                    showToast(err.detail || "Erro ao processar XML", "error");
                }
            } catch (err) {
                showToast("Erro de rede ao carregar arquivo(s).", "error");
            } finally {
                // Restaura o dropzone
                uploadIcon.className = "fa-solid fa-cloud-arrow-up text-4xl text-slate-300 mb-3";
                uploadText.textContent = "Arraste e solte arquivos XML (NF-e, NFC-e ou NFS-e) aqui ou clique para buscar";
                document.getElementById('file-input').value = '';
            }
        }

        // ==========================================
        // FLUXO DE AJUSTES E APURAÇÃO
        // ==========================================

        function abrirAjuste(docId) {
            document.getElementById('ajuste-doc-id').value = docId;
            toggleModal('modal-ajuste', true);
        }

        async function salvarAjuste(event) {
            event.preventDefault();
            const id = document.getElementById('ajuste-doc-id').value;
            const valor = document.getElementById('ajuste-valor').value;
            const justificativa = document.getElementById('ajuste-justificativa').value;
            const usuario = document.getElementById('ajuste-usuario').value;

            try {
                const res = await fetch(`${API_URL}/documentos/${id}/ajustes`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ valor_total_ajuste: valor, justificativa, usuario })
                });

                if (res.ok) {
                    showToast("Ajuste manual gravado e auditado!");
                    toggleModal('modal-ajuste', false);
                    await carregarDocumentos();
                } else {
                    const err = await res.json();
                    showToast(err.detail || "Erro ao registrar ajuste", "error");
                }
            } catch (err) {
                showToast("Erro de rede.", "error");
            }
        }

        async function encerrarDocumento(id) {
            if (!confirm("Tem certeza que deseja encerrar o período desta nota? Isso impedirá qualquer novo ajuste manual (gerando um Snapshot permanente).")) {
                return;
            }

            try {
                const res = await fetch(`${API_URL}/documentos/${id}/encerrar`, { method: 'POST' });
                if (res.ok) {
                    showToast("Período fiscal encerrado e congelado!");
                    await carregarDocumentos();
                } else {
                    showToast("Erro ao encerrar período.", "error");
                }
            } catch (err) {
                showToast("Erro de rede.", "error");
            }
        }

        async function deletarDocumento(id, chave) {
            if (!confirm(`Tem certeza que deseja excluir permanentemente a nota fiscal chave: ${chave}? Esta ação é irreversível!`)) {
                return;
            }

            try {
                const res = await fetch(`${API_URL}/documentos/${id}`, { method: 'DELETE' });
                if (res.ok) {
                    showToast("Nota fiscal excluída do banco com sucesso!");
                    await carregarDocumentos();
                } else {
                    const err = await res.json();
                    showToast(err.detail || "Erro ao excluir nota fiscal", "error");
                }
            } catch (err) {
                showToast("Erro de rede ao excluir.", "error");
            }
        }

        async function resetarBancoDados() {
            if (!confirm(`⚠️ ATENÇÃO CRÍTICA!

Você está prestes a apagar absolutamente TODOS os dados do sistema (todas as notas, ajustes manuais, histórico e empresas cadastradas).

Esta ação é definitiva e IRREVERSÍVEL. Deseja prosseguir para a etapa de confirmação final?`)) {
                return;
            }

            const confirmacao = prompt("Para confirmar o RESET TOTAL do banco de dados, digite exatamente a palavra RESETAR no campo abaixo:");
            if (confirmacao !== "RESETAR") {
                showToast("Operação cancelada. A palavra digitada foi inválida ou nula.", "warning");
                return;
            }

            try {
                const res = await fetch(`${API_URL}/system/reset`, { method: 'POST' });
                if (res.ok) {
                    showToast("Banco de dados limpo com sucesso! Reiniciando...", "success");
                    setTimeout(() => window.location.reload(), 1500);
                } else {
                    showToast("Erro ao resetar o banco de dados.", "error");
                }
            } catch (err) {
                showToast("Erro de rede ao resetar banco.", "error");
            }
        }

        async function abrirApuracao(docId, customAliquota = null) {
            const content = document.getElementById('apuracao-content');
            
            // Configura o botão de recalcular para passar o docId e o valor do input
            const btnRecalcular = document.getElementById('btn-recalcular-apuracao');
            btnRecalcular.onclick = () => {
                const val = document.getElementById('apuracao-aliquota-input').value;
                const aliq = val ? parseFloat(val) / 100.0 : null;
                abrirApuracao(docId, aliq);
            };

            content.innerHTML = `
                <div class="text-center py-6">
                    <i class="fa-solid fa-circle-notch text-3xl text-indigo-500 animate-spin mb-2"></i>
                    <p class="text-sm text-slate-500">Calculando impostos via Strategy Pattern...</p>
                </div>
            `;
            
            // Abre o modal na primeira chamada
            const modal = document.getElementById('modal-apuracao');
            if (modal.classList.contains('hidden')) {
                toggleModal('modal-apuracao', true);
                document.getElementById('apuracao-aliquota-input').value = '';
            }

            try {
                let url = `${API_URL}/documentos/${docId}/apurar`;
                if (customAliquota !== null) {
                    url += `?aliquota=${customAliquota}`;
                }
                const res = await fetch(url);
                if (res.ok) {
                    const data = await res.json();
                    
                    // Preenche o input com a alíquota atual se ele estiver vazio
                    const inputAliq = document.getElementById('apuracao-aliquota-input');
                    if (!inputAliq.value) {
                        inputAliq.value = (data.aliquota_aplicada * 100).toFixed(2);
                    }
                    
                    // Renderiza o relatório de Strategy dinamicamente
                    let apuracaoHtml = `
                        <div class="space-y-4">
                            <div class="flex items-center justify-between p-4 bg-indigo-50 rounded-xl border border-indigo-100">
                                <div>
                                    <p class="text-xs text-indigo-500 uppercase font-semibold">Regime Tributário Injetado</p>
                                    <p class="text-lg font-bold text-indigo-900">${data.regime}</p>
                                </div>
                                <div class="text-right">
                                    <p class="text-xs text-indigo-500 uppercase font-semibold">Total Imposto Calculado</p>
                                    <p class="text-xl font-black text-indigo-900">R$ ${data.imposto_calculado.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</p>
                                </div>
                            </div>
                            
                            <div class="p-4 bg-slate-50 rounded-xl border border-slate-100 space-y-2">
                                <h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider">Detalhamento das Bases</h4>
                                <div class="flex justify-between text-xs text-slate-600">
                                    <span>Valor Original da Nota:</span>
                                    <span class="font-mono">R$ ${data.valor_original.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                </div>
                                <div class="flex justify-between text-xs font-bold text-slate-800">
                                    <span>Valor Final da Staging Area (Base de Cálculo):</span>
                                    <span class="font-mono">R$ ${data.valor_final_base.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                </div>
                                ${data.valor_com_st > 0 ? `
                                <div class="border-t border-slate-200 pt-2 mt-2 space-y-1">
                                    <div class="flex justify-between text-[11px] text-slate-600">
                                        <span>Faturamento Normal:</span>
                                        <span class="font-mono">R$ ${data.valor_sem_st.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                    </div>
                                    <div class="flex justify-between text-[11px] font-bold text-amber-700 bg-amber-50/50 p-1.5 rounded border border-amber-100 flex items-center justify-between">
                                        <span><i class="fa-solid fa-circle-info mr-1"></i>Faturamento com ICMS-ST (Segregável):</span>
                                        <span class="font-mono">R$ ${data.valor_com_st.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                    </div>
                                </div>
                                ` : ''}
                            </div>
                            
                            <div class="space-y-2">
                                <h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider">Memória de Cálculo (Módulos Concretos)</h4>
                    `;

                    if (data.mensagem.includes("Entrada") || (data.detalhes && (data.detalhes.difal !== undefined || data.detalhes.icms_st_compra !== undefined))) {
                        const difalVal = data.detalhes.difal || 0;
                        const stCompraVal = data.detalhes.icms_st_compra || 0;
                        apuracaoHtml += `
                                <div class="space-y-2">
                                    <div class="flex justify-between items-center p-3 bg-white border border-slate-100 rounded-lg shadow-sm">
                                        <div class="flex items-center space-x-2">
                                            <span class="p-1.5 bg-blue-100 text-blue-600 rounded-md text-xs font-bold"><i class="fa-solid fa-scale-unbalanced"></i></span>
                                            <span class="text-xs font-medium text-slate-700">DIFAL Interestadual</span>
                                        </div>
                                        <span class="font-mono text-xs font-bold text-slate-900">R$ ${difalVal.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                    </div>
                                    
                                    <div class="flex justify-between items-center p-3 bg-white border border-slate-100 rounded-lg shadow-sm">
                                        <div class="flex items-center space-x-2">
                                            <span class="p-1.5 bg-cyan-100 text-cyan-600 rounded-md text-xs font-bold"><i class="fa-solid fa-truck-ramp-box"></i></span>
                                            <span class="text-xs font-medium text-slate-700">ICMS-ST Destacado (Compra)</span>
                                        </div>
                                        <span class="font-mono text-xs font-bold text-slate-900">R$ ${stCompraVal.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                    </div>
                                </div>
                        `;
                        
                        if (data.memoria_calculo && data.memoria_calculo.detalhes_itens && data.memoria_calculo.detalhes_itens.length > 0) {
                            apuracaoHtml += `
                                <div class="mt-4 pt-4 border-t border-slate-200">
                                    <h5 class="text-xs font-bold text-slate-500 uppercase tracking-wider mb-2 flex items-center">
                                        <i class="fa-solid fa-list mr-1"></i> Memória de Cálculo por Item
                                    </h5>
                                    <div class="space-y-3">
                            `;
                            data.memoria_calculo.detalhes_itens.forEach((it, idx) => {
                                const vTotal = Number(it.valor_total);
                                const vDesc = Number(it.desconto);
                                const vFrete = Number(it.frete);
                                const vIpi = Number(it.valor_ipi) || 0;
                                const difalCalc = Number(it.difal_calculado);
                                const baseDifal = Number(it.base_difal_calculada);
                                const icmsOrig = Number(it.icms_origem_deduzido) || 0;
                                const vSt = Number(it.icms_st_destacado) || 0;
                                const vLiq = vTotal - vDesc + vFrete + vIpi;
                                
                                apuracaoHtml += `
                                    <div class="p-3 bg-white border border-slate-100 rounded-lg space-y-2 text-xs">
                                        <div class="flex justify-between items-center font-semibold text-slate-800">
                                            <span>#${idx + 1} - ${it.descricao}</span>
                                            <span class="font-mono text-blue-600 font-bold">DIFAL: R$ ${difalCalc.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                        </div>
                                        <div class="grid grid-cols-2 gap-x-4 gap-y-1 text-slate-500 text-[11px]">
                                            <div class="flex justify-between">
                                                <span>Valor Item:</span>
                                                <span class="font-mono text-slate-700">R$ ${vTotal.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                            </div>
                                            <div class="flex justify-between">
                                                <span>Desconto / Frete:</span>
                                                <span class="font-mono text-slate-700">-${vDesc.toFixed(2)} / +${vFrete.toFixed(2)}</span>
                                            </div>
                                            <div class="flex justify-between">
                                                <span>Valor IPI:</span>
                                                <span class="font-mono text-slate-700">R$ ${vIpi.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                            </div>
                                            <div class="flex justify-between font-medium">
                                                <span>Base Líquida (c/ IPI):</span>
                                                <span class="font-mono text-slate-800">R$ ${vLiq.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                            </div>
                                            <div class="flex justify-between">
                                                <span>Alíquota Inter / Interna:</span>
                                                <span class="font-mono text-slate-700">${it.aliquota_interestadual.toFixed(1)}% / ${it.aliquota_interna_destino.toFixed(1)}%</span>
                                            </div>
                                            <div class="flex justify-between">
                                                <span>Fórmula DIFAL:</span>
                                                <span class="font-mono font-semibold text-indigo-600 bg-indigo-50 px-1 rounded">Base ${it.tipo_base_difal}</span>
                                            </div>
                                        </div>
                                        
                                        ${it.tipo_base_difal === "Dupla" ? `
                                            <div class="mt-1.5 p-2 bg-slate-50 rounded border border-slate-100 text-[10px] space-y-1 text-slate-600">
                                                <div class="flex justify-between">
                                                    <span>ICMS Origem Destacado:</span>
                                                    <span class="font-mono">R$ ${icmsOrig.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                                </div>
                                                <div class="flex justify-between font-medium">
                                                    <span>Base DIFAL ("por dentro"):</span>
                                                    <span class="font-mono">(${vLiq.toFixed(2)} - ${icmsOrig.toFixed(2)}) / (1 - ${(it.aliquota_interna_destino/100).toFixed(3)}) = R$ ${baseDifal.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                                </div>
                                                <div class="flex justify-between text-slate-700">
                                                    <span>Cálculo final:</span>
                                                    <span class="font-mono">(${baseDifal.toFixed(2)} * ${(it.aliquota_interna_destino/100).toFixed(3)}) - ${icmsOrig.toFixed(2)} = R$ ${difalCalc.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                                </div>
                                            </div>
                                        ` : `
                                            <div class="mt-1.5 p-2 bg-slate-50 rounded border border-slate-100 text-[10px] space-y-1 text-slate-600">
                                                <div class="flex justify-between font-medium">
                                                    <span>Cálculo Base Simples:</span>
                                                    <span class="font-mono">R$ ${vLiq.toLocaleString('pt-BR', { minimumFractionDigits: 2 })} * (${it.aliquota_interna_destino.toFixed(1)}% - ${it.aliquota_interestadual.toFixed(1)}%) = R$ ${difalCalc.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                                </div>
                                            </div>
                                        `}
                                        
                                        ${vSt > 0 ? `
                                            <div class="flex justify-between text-[11px] text-cyan-600 font-semibold mt-1">
                                                <span>ICMS-ST destacado neste item:</span>
                                                <span class="font-mono">R$ ${vSt.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                            </div>
                                        ` : ''}
                                    </div>
                                `;
                            });
                            apuracaoHtml += `
                                    </div>
                                </div>
                            `;
                        }
                    } else if (data.regime === "Simples Nacional") {
                        const obterBadgeAnexo = (anexo) => {
                            if (!anexo) return '';
                            let classes = "px-2 py-0.5 rounded text-[10px] font-bold border ";
                            if (anexo.includes("Anexo III")) {
                                classes += "bg-blue-50 text-blue-700 border-blue-200";
                            } else if (anexo.includes("Anexo II")) {
                                classes += "bg-amber-50 text-amber-700 border-amber-200";
                            } else if (anexo.includes("Anexo I")) {
                                classes += "bg-emerald-50 text-emerald-700 border-emerald-200";
                            } else if (anexo.includes("Anexo IV")) {
                                classes += "bg-purple-50 text-purple-700 border-purple-200";
                            } else if (anexo.includes("Anexo V")) {
                                classes += "bg-indigo-50 text-indigo-700 border-indigo-200";
                            } else if (anexo.includes("Excluído")) {
                                classes += "bg-slate-100 text-slate-600 border-slate-200";
                            } else if (anexo.includes("Ajuste")) {
                                classes += "bg-slate-50 text-slate-600 border-slate-200";
                            } else {
                                classes += "bg-slate-50 text-slate-700 border-slate-200";
                            }
                            return `<span class="${classes}">${anexo}</span>`;
                        };

                        apuracaoHtml += `
                                <div class="flex justify-between items-center p-3 bg-white border border-slate-100 rounded-lg shadow-sm">
                                    <div class="flex items-center space-x-2">
                                        <span class="p-1.5 bg-brand-100 text-brand-600 rounded-md text-xs font-bold"><i class="fa-solid fa-receipt"></i></span>
                                        <span class="text-xs font-medium text-slate-700">DAS (Imposto Unificado)</span>
                                    </div>
                                    <div class="text-right">
                                        <span class="text-[10px] text-slate-400 mr-2">Alíquota Base: ${(data.aliquota_aplicada * 100).toFixed(2)}%</span>
                                        <span class="font-mono text-xs font-bold text-slate-900">R$ ${data.detalhes.das.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                    </div>
                                </div>
                        `;

                        if (data.memoria_calculo && data.memoria_calculo.itens_calculados && data.memoria_calculo.itens_calculados.length > 0) {
                            apuracaoHtml += `
                                <div class="mt-4 pt-4 border-t border-slate-200">
                                    <h5 class="text-xs font-bold text-slate-500 uppercase tracking-wider mb-2 flex items-center">
                                        <i class="fa-solid fa-list mr-1"></i> Memória de Cálculo por Item
                                    </h5>
                                    <div class="space-y-3">
                            `;
                            data.memoria_calculo.itens_calculados.forEach((it, idx) => {
                                const vTotal = Number(it.valor_total);
                                const vLiq = Number(it.valor_liquido);
                                const aliq = Number(it.aliquota_efetiva);
                                const imposto = Number(it.imposto_calculado);
                                
                                let badgeHtml = obterBadgeAnexo(it.anexo_aplicado);
                                
                                apuracaoHtml += `
                                    <div class="p-3 bg-white border border-slate-100 rounded-lg space-y-2 text-xs">
                                        <div class="flex justify-between items-center font-semibold text-slate-800">
                                            <span class="truncate max-w-[250px]" title="${it.descricao}">#${it.sequencia || (idx+1)} - ${it.descricao}</span>
                                            <div class="flex items-center space-x-2">
                                                ${badgeHtml}
                                                <span class="font-mono text-indigo-600 font-bold">Imposto: R$ ${imposto.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                            </div>
                                        </div>
                                        <div class="grid grid-cols-2 gap-x-4 gap-y-1 text-slate-500 text-[11px]">
                                            <div class="flex justify-between">
                                                <span>CFOP / Valor Total:</span>
                                                <span class="font-mono text-slate-700">${it.cfop} / R$ ${vTotal.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                            </div>
                                            <div class="flex justify-between font-medium">
                                                <span>Base Líquida:</span>
                                                <span class="font-mono text-slate-800">R$ ${vLiq.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                            </div>
                                            <div class="flex justify-between">
                                                <span>Alíquota Efetiva:</span>
                                                <span class="font-mono text-slate-700">${aliq.toFixed(2)}%</span>
                                            </div>
                                            <div class="flex justify-between">
                                                <span>Cálculo:</span>
                                                <span class="font-mono text-slate-600">${it.detalhe_calculo || ''}</span>
                                            </div>
                                        </div>
                                    </div>
                                `;
                            });
                            apuracaoHtml += `
                                    </div>
                                </div>
                            `;
                        }
                    } else if (data.regime === "Lucro Presumido") {
                        const det = data.detalhes;
                        apuracaoHtml += `
                                <div class="space-y-2">
                                    <!-- PIS -->
                                    <div class="flex justify-between items-center p-3 bg-white border border-slate-100 rounded-lg shadow-sm">
                                        <span class="text-xs font-medium text-slate-700">PIS (Alíquota 0.65%)</span>
                                        <span class="font-mono text-xs font-bold text-slate-900">R$ ${det.pis.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                    </div>
                                    <!-- COFINS -->
                                    <div class="flex justify-between items-center p-3 bg-white border border-slate-100 rounded-lg shadow-sm">
                                        <span class="text-xs font-medium text-slate-700">COFINS (Alíquota 3.00%)</span>
                                        <span class="font-mono text-xs font-bold text-slate-900">R$ ${det.cofins.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                    </div>
                                    <!-- IRPJ -->
                                    <div class="flex justify-between items-center p-3 bg-white border border-slate-100 rounded-lg shadow-sm">
                                        <span class="text-xs font-medium text-slate-700">IRPJ (Alíquota Efetiva 4.80%)</span>
                                        <span class="font-mono text-xs font-bold text-slate-900">R$ ${det.irpj.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                    </div>
                                    <!-- CSLL -->
                                    <div class="flex justify-between items-center p-3 bg-white border border-slate-100 rounded-lg shadow-sm">
                                        <span class="text-xs font-medium text-slate-700">CSLL (Alíquota Efetiva 2.88%)</span>
                                        <span class="font-mono text-xs font-bold text-slate-900">R$ ${det.csll.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                                    </div>
                                </div>
                        `;
                    }

                    let alertBg = "bg-emerald-50 border-emerald-100 text-emerald-800";
                    let alertIcon = "fa-circle-info text-emerald-600";
                    
                    if (data.mensagem.includes("CANCELADA") || data.mensagem.includes("DENEGADA")) {
                        alertBg = "bg-rose-50 border-rose-100 text-rose-800";
                        alertIcon = "fa-triangle-exclamation text-rose-600";
                    } else if (data.mensagem.includes("ST DETECTADO")) {
                        alertBg = "bg-amber-50 border-amber-100 text-amber-800";
                        alertIcon = "fa-circle-info text-amber-600";
                    }

                    apuracaoHtml += `
                            </div>
                            <div class="p-3 ${alertBg} rounded-xl border text-xs flex items-start space-x-2">
                                <i class="fa-solid ${alertIcon} mt-0.5"></i>
                                <p>${data.mensagem}</p>
                            </div>
                        </div>
                    `;

                    content.innerHTML = apuracaoHtml;
                } else {
                    content.innerHTML = `
                        <div class="text-center py-6 text-rose-500">
                            <i class="fa-solid fa-circle-xmark text-3xl mb-2"></i>
                            <p class="text-sm font-medium">Erro ao calcular impostos para este documento.</p>
                        </div>
                    `;
                }
            } catch (err) {
                content.innerHTML = `
                    <div class="text-center py-6 text-rose-500">
                        <i class="fa-solid fa-wifi text-3xl mb-2"></i>
                        <p class="text-sm font-medium">Erro de rede ao conectar à esteira Strategy.</p>
                    </div>
                `;
            }
        }

        // ==========================================
        // FLUXO DE EDIÇÃO DE CADASTRO DE EMPRESA
        // ==========================================

        function abrirEditarEmpresa() {
            const select = document.getElementById('select-empresa');
            const opt = select.selectedOptions[0];
            
            if (opt && opt.value && opt.value !== "auto") {
                document.getElementById('editar-empresa-cnpj').value = formatarCNPJ(opt.dataset.cnpj);
                document.getElementById('editar-empresa-razao').value = opt.textContent;
                document.getElementById('editar-empresa-regime').value = opt.dataset.regime;
                
                document.getElementById('editar-empresa-rbt12').value = opt.dataset.rbt12 || "0.00";
                document.getElementById('editar-empresa-folha12').value = opt.dataset.folha12 || "0.00";
                document.getElementById('editar-empresa-fator-r').checked = opt.dataset.sujeito_fator_r === "true";
                document.getElementById('editar-empresa-cnae').value = opt.dataset.cnae || "";
                document.getElementById('editar-empresa-categoria').value = opt.dataset.categoria || "Serviços (Anexo III)";
                
                aplicarRegrasCNAE('editar-empresa');
                toggleModal('modal-editar-empresa', true);
            }
        }

        async function salvarEdicaoEmpresa(event) {
            event.preventDefault();
            const select = document.getElementById('select-empresa');
            const empresaId = select.value;
            const razao = document.getElementById('editar-empresa-razao').value;
            const cnae = document.getElementById('editar-empresa-cnae').value.trim().replace(/\D/g, "");
            const regime = document.getElementById('editar-empresa-regime').value;
            const categoria_simples = document.getElementById('editar-empresa-categoria').value;
            
            const rbt12 = parseFloat(document.getElementById('editar-empresa-rbt12').value || 0);
            const folha12 = parseFloat(document.getElementById('editar-empresa-folha12').value || 0);
            const sujeito_fator_r = document.getElementById('editar-empresa-fator-r').checked;

            try {
                const res = await fetch(`${API_URL}/empresas/${empresaId}`, {
                    method: 'PUT',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ 
                        razao_social: razao, 
                        regime_tributario: regime,
                        rbt12,
                        folha12,
                        sujeito_fator_r,
                        categoria_simples,
                        cnae: cnae || null
                    })
                });

                if (res.ok) {
                    showToast("Cadastro da empresa atualizado com sucesso!");
                    toggleModal('modal-editar-empresa', false);
                    await carregarEmpresas();
                    select.value = empresaId;
                    await carregarDocumentos();
                } else {
                    const err = await res.json();
                    showToast(err.detail || "Erro ao editar empresa", "error");
                }
            } catch (err) {
                showToast("Erro de rede ao editar empresa", "error");
            }
        }

        // ==========================================
        // UTILITÁRIOS
        // ==========================================

        function formatarCNPJ(v) {
            return v.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, "$1.$2.$3/$4-$5");
        }

        function formatarMoeda(val) {
            return Number(val).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
        }

        // WebSocket Heartbeat para autodesligamento do servidor
        (function() {
            const protocol = window.location.protocol === 'https:' ? 'wss:' : 'ws:';
            const wsUrl = `${protocol}//${window.location.host}/ws/heartbeat`;
            let ws;

            function connectHeartbeat() {
                ws = new WebSocket(wsUrl);
                ws.onclose = function() {
                    setTimeout(connectHeartbeat, 3000);
                };
                ws.onerror = function() {
                    ws.close();
                };
            }
            connectHeartbeat();
        })();

        // Controle de logs do sistema
        let logsInterval = null;

        function abrirModalLogs() {
            document.getElementById('modal-logs').classList.remove('hidden');
            carregarLogs();
            logsInterval = setInterval(carregarLogs, 3000);
        }

        function fecharModalLogs() {
            document.getElementById('modal-logs').classList.add('hidden');
            if (logsInterval) {
                clearInterval(logsInterval);
                logsInterval = null;
            }
        }

        async function carregarLogs() {
            try {
                const response = await fetch('/api/logs');
                const data = await response.json();
                const terminal = document.getElementById('terminal-logs');
                if (data.logs && data.logs.length > 0) {
                    const formatted = data.logs.map(line => {
                        let colorClass = 'text-slate-300';
                        if (line.includes(' - ERROR - ') || line.includes(' 404 ') || line.includes(' 500 ') || line.includes('[ERRO]')) {
                            colorClass = 'text-rose-400 font-semibold';
                        } else if (line.includes(' - WARNING - ') || line.includes('[ALERTA]') || line.includes(' 307 ')) {
                            colorClass = 'text-amber-400';
                        } else if (line.includes(' - INFO - ') && (line.includes(' 200 OK') || line.includes(' 201 Created') || line.includes('[SUCESSO]'))) {
                            colorClass = 'text-emerald-400';
                        } else if (line.includes('Uvicorn running on') || line.includes('Application startup complete')) {
                            colorClass = 'text-indigo-400 font-bold';
                        } else if (line.includes(' - INFO - ')) {
                            colorClass = 'text-slate-400';
                        }
                        return `<div class="${colorClass}">${escaparHtml(line)}</div>`;
                    }).join('');
                    
                    const isScrolledToBottom = terminal.scrollHeight - terminal.clientHeight <= terminal.scrollTop + 50;
                    terminal.innerHTML = formatted;
                    if (isScrolledToBottom) {
                        terminal.scrollTop = terminal.scrollHeight;
                    }
                } else {
                    terminal.innerHTML = '<div class="text-slate-500">// Nenhum log disponível.</div>';
                }
            } catch (err) {
                console.error("Erro ao buscar logs:", err);
            }
        }

        function escaparHtml(text) {
            const map = {
                '&': '&amp;',
                '<': '&lt;',
                '>': '&gt;',
                '"': '&quot;',
                "'": '&#039;'
            };
            return text.replace(/[&<>"']/g, function(m) { return map[m]; });
        }

        function copiarLogs() {
            const terminal = document.getElementById('terminal-logs');
            const range = document.createRange();
            range.selectNode(terminal);
            window.getSelection().removeAllRanges();
            window.getSelection().addRange(range);
            try {
                document.execCommand('copy');
                showToast("Logs copiados para a área de transferência!");
            } catch (err) {
                showToast("Erro ao copiar logs.", "error");
            }
            window.getSelection().removeAllRanges();
        }

        async function limparLogsArquivo() {
            if (!confirm("Tem certeza que deseja limpar todo o histórico de logs do servidor?")) return;
            try {
                const response = await fetch('/api/logs/clear', { method: 'POST' });
                if (response.ok) {
                    showToast("Histórico de logs limpo com sucesso!");
                    carregarLogs();
                } else {
                    showToast("Falha ao limpar logs.", "error");
                }
            } catch (err) {
                showToast("Erro ao limpar logs.", "error");
            }
        }

        // Funções para Controle e Ações em Lote (Checkboxes)
        function toggleSelectAllDocs(master) {
            const checkboxes = document.querySelectorAll('.doc-checkbox');
            checkboxes.forEach(cb => {
                cb.checked = master.checked;
            });
            atualizarAcoesLote();
        }

        function desmarcarTodos() {
            const master = document.getElementById('select-all-docs');
            if (master) master.checked = false;
            toggleSelectAllDocs({ checked: false });
        }

        function obterIdsSelecionados() {
            const checkboxes = document.querySelectorAll('.doc-checkbox:checked');
            return Array.from(checkboxes).map(cb => parseInt(cb.value));
        }

        function atualizarAcoesLote() {
            const selecionados = obterIdsSelecionados();
            const panel = document.getElementById('batch-actions-panel');
            const counterSpan = document.getElementById('batch-selected-count');
            
            if (selecionados.length > 0) {
                if (counterSpan) counterSpan.textContent = selecionados.length;
                if (panel) {
                    panel.classList.remove('hidden');
                    panel.classList.add('flex');
                }
            } else {
                if (panel) {
                    panel.classList.add('hidden');
                    panel.classList.remove('flex');
                }
            }
        }

        async function excluirSelecionadas() {
            const ids = obterIdsSelecionados();
            if (ids.length === 0) return;
            if (!confirm(`Tem certeza que deseja excluir permanentemente as ${ids.length} nota(s) fiscal(is) selecionada(s)? Esta ação é irreversível!`)) {
                return;
            }

            try {
                const response = await fetch(`${API_URL}/documentos/excluir-em-lote`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ ids })
                });

                if (response.ok) {
                    showToast(`${ids.length} nota(s) excluída(s) com sucesso!`);
                    desmarcarTodos();
                    await carregarDocumentos();
                } else {
                    showToast("Erro ao excluir notas em lote.", "error");
                }
            } catch (err) {
                showToast("Erro de rede ao excluir notas em lote.", "error");
            }
        }

        async function encerrarSelecionadas() {
            const ids = obterIdsSelecionados();
            if (ids.length === 0) return;
            if (!confirm(`Tem certeza que deseja encerrar e congelar o período das ${ids.length} nota(s) fiscal(is) selecionada(s)? Isso impedirá novos ajustes manuais nelas.`)) {
                return;
            }

            try {
                const response = await fetch(`${API_URL}/documentos/encerrar-em-lote`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ ids })
                });

                if (response.ok) {
                    showToast(`${ids.length} nota(s) encerrada(s) com sucesso!`);
                    desmarcarTodos();
                    await carregarDocumentos();
                } else {
                    showToast("Erro ao encerrar notas em lote.", "error");
                }
            } catch (err) {
                showToast("Erro de rede ao encerrar notas em lote.", "error");
            }
        }

        async function excluirPeriodo() {
            const selectEmpresa = document.getElementById('select-empresa');
            const empresaId = selectEmpresa.value;
            if (!empresaId) {
                showToast("Selecione uma empresa antes de limpar o período.", "warning");
                return;
            }

            const mesVal = document.getElementById('select-mes').value;
            const anoVal = document.getElementById('select-ano').value;
            
            if (!mesVal || !anoVal) {
                showToast("Selecione um mês e um ano específicos para limpar a competência.", "warning");
                return;
            }

            const nomeMes = document.getElementById('select-mes').selectedOptions[0].textContent;
            const descFiltro = `${nomeMes}/${anoVal}`;

            if (!confirm(`⚠️ ATENÇÃO CRÍTICA!
            
Você está prestes a excluir permanentemente TODAS as notas fiscais da empresa ativa na competência ${descFiltro}.

Esta ação é definitiva e irreversível! Deseja continuar?`)) {
                return;
            }

            try {
                const response = await fetch(`${API_URL}/documentos?empresa_id=${empresaId}&mes=${mesVal}&ano=${anoVal}`, {
                    method: 'DELETE'
                });

                if (response.ok) {
                    showToast(`Notas da competência ${descFiltro} limpas com sucesso!`);
                    desmarcarTodos();
                    await carregarDocumentos();
                } else {
                    showToast("Erro ao limpar período fiscal.", "error");
                }
            } catch (err) {
                showToast("Erro de rede ao limpar período fiscal.", "error");
            }
        }

        async function excluirNotasEmpresa() {
            const selectEmpresa = document.getElementById('select-empresa');
            const empresaId = selectEmpresa.value;
            if (!empresaId) {
                showToast("Selecione uma empresa ativa antes de limpar suas notas.", "warning");
                return;
            }

            const nomeEmpresa = selectEmpresa.selectedOptions[0].textContent;

            if (!confirm(`⚠️ ALERTA DE SEGURANÇA!
            
Você está prestes a excluir permanentemente ABSOLUTAMENTE TODAS as notas fiscais importadas para a empresa:
"${nomeEmpresa}"

O cadastro da empresa será mantido intacto. Deseja prosseguir?`)) {
                return;
            }

            try {
                const response = await fetch(`${API_URL}/documentos?empresa_id=${empresaId}`, {
                    method: 'DELETE'
                });

                if (response.ok) {
                    showToast(`Todas as notas da empresa "${nomeEmpresa}" foram excluídas!`);
                    desmarcarTodos();
                    await carregarDocumentos();
                } else {
                    showToast("Erro ao limpar notas da empresa.", "error");
                }
            } catch (err) {
                showToast("Erro de rede ao limpar notas da empresa.", "error");
            }
        }
    </script>
    <!-- MODAL DE LOGS DO SISTEMA -->
    <div id="modal-logs" class="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/40 backdrop-blur-sm hidden">
        <div class="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-4xl h-[80vh] flex flex-col shadow-2xl overflow-hidden">
            <!-- Header do Modal -->
            <div class="px-6 py-4 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
                <div class="flex items-center space-x-3">
                    <div class="bg-indigo-500/10 text-indigo-400 p-2 rounded-lg border border-indigo-500/20">
                        <i class="fa-solid fa-terminal text-sm"></i>
                    </div>
                    <div>
                        <h3 class="text-sm font-bold text-slate-100">Logs em Tempo Real do Servidor</h3>
                        <p class="text-[10px] text-slate-500">Últimos eventos do motor de apuração e conexões da API</p>
                    </div>
                </div>
                <button onclick="fecharModalLogs()" class="p-1 text-slate-400 hover:text-slate-200 transition-colors">
                    <i class="fa-solid fa-xmark text-lg"></i>
                </button>
            </div>
            <!-- Corpo / Terminal do Modal -->
            <div id="terminal-logs" class="flex-1 p-6 overflow-y-auto font-mono text-xs text-slate-300 space-y-1.5 selection:bg-indigo-500 selection:text-white bg-slate-950">
                <div class="text-slate-500 text-[11px] mb-2">// Conectando ao fluxo de logs do servidor...</div>
            </div>
            <!-- Footer do Modal -->
            <div class="px-6 py-3 bg-slate-950 border-t border-slate-800 flex items-center justify-between text-[11px] text-slate-500">
                <div class="flex items-center space-x-2">
                    <span class="relative flex h-2 w-2">
                        <span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-indigo-400 opacity-75"></span>
                        <span class="relative inline-flex rounded-full h-2 w-2 bg-indigo-500"></span>
                    </span>
                    <span>Atualizando a cada 3 segundos</span>
                </div>
                <div class="flex items-center space-x-4">
                    <button onclick="copiarLogs()" class="text-slate-400 hover:text-slate-200 flex items-center transition-colors">
                        <i class="fa-solid fa-copy mr-1.5"></i> Copiar Tudo
                    </button>
                    <button onclick="limparLogsArquivo()" class="text-rose-400 hover:text-rose-300 flex items-center transition-colors">
                        <i class="fa-solid fa-trash mr-1.5"></i> Limpar Histórico
                    </button>
                </div>
            </div>
        </div>
    </div>

</body>
</html>
"""
