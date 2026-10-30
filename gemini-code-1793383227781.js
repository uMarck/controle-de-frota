let globalProcessedData = [];
let myChart = null;

function extraireNumeroFrota(val) {
    if (!val) return null;
    const str = String(val).toUpperCase().trim();
    const match = str.match(/(\d{1,3})/);
    if (match) {
        const num = parseInt(match[1], 10);
        if (num >= 1 && num <= 185) {
            return `F-${String(num).padStart(3, '0')}`;
        }
    }
    return null;
}

async function processarPlanilhas() {
    const fCtrl = document.getElementById('fileCtrl').files[0];
    const fMapa = document.getElementById('fileMapa').files[0];
    const fPat = document.getElementById('filePat').files[0];

    if (!fCtrl || !fMapa || !fPat) {
        alert("Por favor, selecione os três arquivos solicitados antes de processar!");
        return;
    }

    try {
        // Ler arquivos com SheetJS
        const dataCtrl = await lerArquivoExcel(fCtrl);
        const dataMapa = await lerArquivoExcel(fMapa);
        const dataPat = await lerArquivoExcel(fPat);

        let frotasMap = {};

        // 1. Carregar Pranchas de F-001 a F-185 da Planilha de Controle
        const wsCtrl = dataCtrl.Sheets['PRANCHAS'] || dataCtrl.Sheets[Object.keys(dataCtrl.Sheets)[0]];
        const jsonCtrl = XLSX.utils.sheet_to_json(wsCtrl, { header: 1 });
        
        for (let i = 1; i < jsonCtrl.length; i++) {
            const row = jsonCtrl[i];
            const pId = row[0];
            const placa = row[1] || '';
            const marca = row[2] || 'BERTOLINI';
            const tamanho = row[3] || 40;
            const peso = row[4] || 'BOA';

            const frotaNorm = extraireNumeroFrota(pId);
            if (frotaNorm) {
                frotasMap[frotaNorm] = {
                    PRANCHA: frotaNorm,
                    PLACA: placa,
                    MARCA: marca,
                    TAMANHO: tamanho,
                    PESO: peso,
                    STATUS: 'LIVRE',
                    CLIENTE: 'BASE ALENCAR',
                    SITUACAO: 'DISPONÍVEL PARA USO',
                    EM_MANUTENCAO: false,
                    NO_PATIO: true
                };
            }
        }

        // Garantir todas de 001 a 185 mesmo se faltarem na aba
        for (let j = 1; j <= 185; j++) {
            const fKey = `F-${String(j).padStart(3, '0')}`;
            if (!frotasMap[fKey]) {
                frotasMap[fKey] = {
                    PRANCHA: fKey,
                    PLACA: 'PARTICULAR',
                    MARCA: 'BERTOLINI',
                    TAMANHO: 40,
                    PESO: 'BOA',
                    STATUS: 'LIVRE',
                    CLIENTE: 'BASE ALENCAR',
                    SITUACAO: 'DISPONÍVEL PARA USO',
                    EM_MANUTENCAO: false,
                    NO_PATIO: true
                };
            }
        }

        // 2. Cruzar com PATIAMENTO DE FROTA (Detectar Manutenção e Pátio)
        if (dataPat.Sheets['PRANCHAS']) {
            const wsPatPranchas = dataPat.Sheets['PRANCHAS'];
            const jsonPat = XLSX.utils.sheet_to_json(wsPatPranchas, { header: 1 });
            for (let i = 2; i < jsonPat.length; i++) {
                const row = jsonPat[i];
                const fVal = row[0];
                const obs = String(row[7] || '').toUpperCase();
                const dtSaida = row[8];

                const fNorm = extraireNumeroFrota(fVal);
                if (fNorm && frotasMap[fNorm]) {
                    if (dtSaida) {
                        frotasMap[fNorm].NO_PATIO = false;
                    } else {
                        frotasMap[fNorm].NO_PATIO = true;
                    }

                    if (obs.includes('MANUT') || obs.includes('REPARO') || obs.includes('OFICINA') || obs.includes('SOLDA') || obs.includes('QUEBRADA')) {
                        frotasMap[fNorm].EM_MANUTENCAO = true;
                    }
                }
            }
        }

        // Checar aba LOCKS / Outras abas de manutenção do patiamento
        ['LOCKS', 'CONTAINERS'].forEach(sheetName => {
            if (dataPat.Sheets[sheetName]) {
                const wsL = dataPat.Sheets[sheetName];
                const jsonL = XLSX.utils.sheet_to_json(wsL, { header: 1 });
                jsonL.forEach(row => {
                    row.forEach(cellVal => {
                        const strVal = String(cellVal || '').toUpperCase();
                        if (strVal.includes('MANUTENÇÃO') || strVal.includes('MANUTENCAO') || strVal.includes('OFICINA')) {
                            row.forEach(subVal => {
                                const fn = extraireNumeroFrota(subVal);
                                if (fn && frotasMap[fn]) {
                                    frotasMap[fn].EM_MANUTENCAO = true;
                                }
                            });
                        }
                    });
                });
            }
        });

        // 3. Cruzar com MAPA GERAL DAS MOVIMENTAÇÕES DA SEMANA
        const mapSheetName = dataMapa.SheetNames[0];
        const wsMapa = dataMapa.Sheets[mapSheetName];
        const jsonMapa = XLSX.utils.sheet_to_json(wsMapa, { header: 1 });

        for (let i = 3; i < jsonMapa.length; i++) {
            const row = jsonMapa[i];
            const cliente = row[0];
            const ref = row[1];
            const fVal = row[13]; // Coluna da Frota no Mapa

            const fNorm = extraireNumeroFrota(fVal);
            if (fNorm && frotasMap[fNorm]) {
                if (!frotasMap[fNorm].EM_MANUTENCAO) {
                    if (cliente) {
                        frotasMap[fNorm].CLIENTE = String(cliente).trim();
                        frotasMap[fNorm].SITUACAO = 'EM USO';
                    }

                    if (ref && String(ref).toUpperCase().includes('DEVOLUÇÃO')) {
                        if (frotasMap[fNorm].NO_PATIO) {
                            frotasMap[fNorm].CLIENTE = 'BASE ALENCAR';
                            frotasMap[fNorm].SITUACAO = 'DISPONÍVEL PARA USO';
                        }
                    }
                }
            }
        }

        // 4. Consolidação Final de Status e Cores
        globalProcessedData = Object.values(frotasMap).map(item => {
            if (item.EM_MANUTENCAO) {
                item.STATUS = 'MANUTENÇÃO';
                item.CLIENTE = 'MANUTENÇÃO / OFICINA';
                item.SITUACAO = 'EM MANUTENÇÃO';
            } else if (item.SITUACAO === 'EM USO') {
                item.STATUS = 'CHEIO';
            } else {
                item.STATUS = 'LIVRE';
                item.CLIENTE = 'BASE ALENCAR';
                item.SITUACAO = 'DISPONÍVEL PARA USO';
            }
            return item;
        });

        globalProcessedData.sort((a, b) => a.PRANCHA.localeCompare(b.PRANCHA, undefined, { numeric: true }));

        atualizarInterface();

    } catch (error) {
        console.error(error);
        alert("Erro ao processar as planilhas. Verifique se os formatos estão corretos.");
    }
}

function lerArquivoExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                resolve(workbook);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function atualizarInterface() {
    const emUso = globalProcessedData.filter(i => i.SITUACAO === 'EM USO').length;
    const disp = globalProcessedData.filter(i => i.SITUACAO === 'DISPONÍVEL PARA USO').length;
    const manut = globalProcessedData.filter(i => i.SITUACAO === 'EM MANUTENÇÃO').length;

    document.getElementById('countEmUso').innerText = emUso;
    document.getElementById('countDisp').innerText = disp;
    document.getElementById('countManut').innerText = manut;

    // Preencher Tabela
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';

    globalProcessedData.forEach(item => {
        let bgClass = '';
        if (item.SITUACAO === 'EM MANUTENÇÃO') bgClass = 'bg-red-50 text-red-900';
        else if (item.SITUACAO === 'DISPONÍVEL PARA USO') bgClass = 'bg-green-50 text-green-900';
        else if (item.SITUACAO === 'EM USO') bgClass = 'bg-yellow-50 text-yellow-900';

        const tr = document.createElement('tr');
        tr.className = `border-b border-slate-200 ${bgClass}`;
        tr.innerHTML = `
            <td class="p-3 font-bold">${item.PRANCHA}</td>
            <td class="p-3">${item.PLACA}</td>
            <td class="p-3">${item.MARCA}</td>
            <td class="p-3">${item.TAMANHO}</td>
            <td class="p-3">${item.PESO}</td>
            <td class="p-3 font-semibold">${item.STATUS}</td>
            <td class="p-3">${item.CLIENTE}</td>
            <td class="p-3 font-bold">${item.SITUACAO}</td>
        `;
        tbody.appendChild(tr);
    });

    document.getElementById('resultSection').classList.remove('hidden');

    // Renderizar Gráfico de Pizza
    const ctx = document.getElementById('pieChart').getContext('2d');
    if (myChart) myChart.destroy();

    myChart = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: ['EM USO', 'DISPONÍVEL PARA USO', 'EM MANUTENÇÃO'],
            datasets: [{
                data: [emUso, disp, manut],
                backgroundColor: ['#FACC15', '#4ADE80', '#F87171'],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { position: 'bottom' },
                title: { display: true, text: 'Distribuição Atual das Pranchas' }
            }
        }
    });
}

function baixarExcel() {
    const wb = XLSX.utils.book_new();
    const wsData = [
        ['PRANCHA', 'PLACA', 'MARCA', 'TAMANHO', 'PESO', 'STATUS', 'CLIENTE', 'SITUAÇÃO']
    ];

    globalProcessedData.forEach(item => {
        wsData.push([item.PRANCHA, item.PLACA, item.MARCA, item.TAMANHO, item.PESO, item.STATUS, item.CLIENTE, item.SITUACAO]);
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, "PRANCHAS");

    XLSX.writeFile(wb, "CONTROLE_DE_PRANCHAS_CONSOLIDADO.xlsx");
}